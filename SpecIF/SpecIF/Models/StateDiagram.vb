Imports Visio = Microsoft.Office.Interop.Visio
'    ---------------------------------------------------------------------------------------
'       Modul   : StateDiagram
'       Datum   : 29.04.2017
'       Author  : Philipp Mochine
'       Zweck   : Das Zustandsdiagramm wird eingelesen und in das Integrationsmodell hinzugefügt.
'                 Beziehungen zwischen Aktivitätsdiagramm und Zustandsdiagramm werden hier automatisch gesetzt.
'       Quellort: Von ExecuteModel(modelPath As String)
'    ---------------------------------------------------------------------------------------
Module StateDiagram
    Private vsoDocument As Visio.Document
    Private vsoPage As Visio.Page
    Private Const modelType As String = "State"
    Private Const diagramFrame As String = "Diagram Overview"
    Private ReadOnly Property InitFinalTypes As String() = New String() {"Initial state", "Final state"}
    Private ReadOnly Property StaTypes As String() = New String() {"Submachine state", "State with internal behavior", "State"}
    Private ReadOnly Property ActTypes As String() = New String() {"Diagram Overview", "Dynamic connector"}
    Private HierarchieList As New List(Of ModelElementsDataStructur)
    Private modelDescription As String = ConvertStringToJson("Sämtliche Systeme, egal ob Software-, Hardware-, soziale oder biologische Systeme, weisen Zustandsübergänge (Transitionen) und Zustände auf." &
                                                             " Zustände verändern sich durch ausgelöste Ereignisse, die wiederum bestimmte Aktivitäten in Gang setzen bis dann wieder ein Zustand eintritt." &
                                                             " Zustandsdiagramme modellieren dieses Verhalten mit Zustandsautomaten, die auf der Arbeit von David Harel sowie der allgemeinen Mealy- und Moore-Automaten basieren. ")
    Private Const frameDescription As String = "ist der Diagrammtitel des Zustandsmodells. <p>Die Syntax eines UML-Diagrammtitels lautet: " &
                                                "<b>Diagrammtyp </b>Diagrammname [Parameter]</p><p>UML bzw. SysML verwenden hierbei das Zustandsdiagramm.</p>"
    Private Const elementDescription As String = "ist ein Element des Zustandsmodells. Es führ Aktionen aus."
    '    ---------------------------------------------------------------------------------------
    '       Function     : StateAnalyse(vsoDocumenttemp As Visio.Document)
    '       Zweck        : Liest das Visiodokument ein, erstellt ein Bild und die JSON Struktur. 
    '       @param       : Pointer auf das Visio Dokument.
    '       Quellort     : Von ExecuteModel(modelPath As String)
    '    ---------------------------------------------------------------------------------------
    Public Function StateAnalyse(vsoDocumenttemp As Visio.Document) As Boolean
        StateAnalyse = True
        vsoDocument = vsoDocumenttemp
        vsoPage = vsoDocument.Pages(1)
        Try
            SvgExport(vsoDocument)
            CreateStateSpecIFBase()
            CreateStateSpecIFCode()
            SetHierarchie()
        Catch ex As Exception
            StateAnalyse = False
            SendError(ex.Message, -1)
        Finally
            ReleaseObject(vsoPage)
            HierarchieList = New List(Of ModelElementsDataStructur)
        End Try
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: CreateStateSpecIFCode()
    '       Zweck        : Hauptcode des Moduls.
    '                       1. Suche nach einem "Diagram Overview". 
    '                       2. Danach werden Zustandsmodellelemente gesucht.
    '       @param       : Pointer auf das Visio Dokument.
    '       Quellort     : Von ExecuteModel(modelPath As String)
    '    ---------------------------------------------------------------------------------------
    Private Sub CreateStateSpecIFCode()
        Dim elementID As String
        Dim elementName As String
        Dim shapeOverview As Visio.Shape = Nothing, tempShape As Visio.Shape = Nothing
        Dim vsoPageID As String = GetIDofMEDS(vsoPage.Name, modelType), diagramOverviewID As String = ""
        Dim onlyOneOverview As Boolean = False 'Nur ein Diagrammrahmen erlaubt.

        For Each tempShape In vsoPage.Shapes
            If tempShape.Master.NameU.Contains(diagramFrame) Then 'Es wird nach dem Diagrammrahmen gesucht. 
                If onlyOneOverview = False Then onlyOneOverview = True Else Throw New System.Exception("Nur ein Diagrammrahmen erlaubt.")
                elementID = "MEl-" & GetGenID(27)
                diagramOverviewID = elementID
                elementName = tempShape.Shapes.Item(1).Text
                If IsNullOrBlank(elementName) Then Throw New System.Exception("Diagrammrahmen muss einen Titel haben.")
                Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                AddRelation_vsoPage_Overview(vsoPageID, "", elementID)
                AddObjectState(elementID, elementName, frameDescription)
                HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Zustand"))
                shapeOverview = tempShape
            End If
        Next

        If Not shapeOverview Is Nothing And shapeOverview.Master.NameU.Contains(diagramFrame) Then
            Dim vsoReturnedSelection As Visio.Selection
            Dim actShapes() As Visio.Shape = Nothing
            Dim onlyOneInitialState As Boolean = False
            Dim internBehavior As String = ""

            vsoReturnedSelection = shapeOverview.SpatialNeighbors(Visio.VisSpatialRelationCodes.visSpatialContain, 0, Visio.VisSpatialRelationFlags.visSpatialIncludeContainerShapes)
            If vsoReturnedSelection.Count = 0 Then
                Throw New System.Exception("Keine Elemente im Diagrammrahmen gefunden.")
            Else
                For Each shape As Visio.Shape In vsoReturnedSelection 'Durch alle Elemente im Diagrammrahmen werden durchgesucht.
                    If StaTypes.ToList().IndexOf(GetModelElementBaseName(shape.Master.NameU)) >= 0 Then 'Zustandselemente
                        CheckElement(shape)
                        elementName = shape.Shapes.Item(1).Text
                        elementID = GetIDofMEDS(elementName, "Activity")
                        'Existiert das Element noch nicht?
                        If elementID = "" Then
                            If Not IsNullOrBlank(GetIDofMEDS(elementName, modelType)) Then Throw New System.Exception("Zustandselement " & elementName & " erscheint doppelt. Bitte Namen ändern.")
                            elementID = "MEl-" & GetGenID(27)
                            Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                            If Not IsNullOrBlank(shape.Shapes.Item(2).Text) Then
                                internBehavior = "Sein internes Verhalten ist: <p>" & shape.Shapes.Item(2).Text & "</p>"
                            End If
                            AddObjectState(elementID, elementName, "ist ein Zustandselement des Zustandsdiagrammes. " & internBehavior & ".")
                            internBehavior = ""
                            HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Zustand"))
                            AddRelation_vsoPage_Overview(vsoPageID, diagramOverviewID, elementID)
                        Else
                            Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                            HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Zustand"))
                            AddRelation_vsoPage_Overview(vsoPageID, diagramOverviewID, elementID)
                        End If
                    ElseIf ActTypes.ToList().IndexOf(GetModelElementBaseName(shape.Master.NameU)) >= 0 Then 'Zustandsübergänge sind in diesem Beispiel Akteure.
                        'Dieser Part ist wichtig. Alle Zustände sind für SpecIF eingetragen.
                        'Jetzt müssen alle Transitionen einzeln vorgespeichert und ausgelesen werden.
                        If (actShapes IsNot Nothing AndAlso actShapes.Count > 0) Then
                            ReDim Preserve actShapes(UBound(actShapes) + 1)
                            actShapes(UBound(actShapes)) = shape
                        Else
                            ReDim actShapes(0)
                            actShapes(0) = shape
                        End If
                    ElseIf InitFinalTypes.ToList().IndexOf(GetModelElementBaseName(shape.Master.NameU)) >= 0 Then 'Anfangs oder Endzustände.
                        CheckElement(shape)
                        If InitFinalTypes(0).IndexOf(GetModelElementBaseName(shape.Master.NameU)) >= 0 AndAlso onlyOneInitialState = True Then Throw New System.Exception("Nur einen Anfangsknoten erlaubt.")
                        If InitFinalTypes(0).IndexOf(GetModelElementBaseName(shape.Master.NameU)) >= 0 AndAlso onlyOneInitialState = False Then onlyOneInitialState = True

                        'Wir schauen nicht mehr die Master Names an
                        elementName = shape.Name
                        elementID = GetIDofMEDS(elementName, "Activity")
                        If elementID = "" Then
                            elementID = "MEl-" & GetGenID(27)
                            Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                            AddObjectState(elementID, elementName, "ist ein " & GetModelElementBaseName(shape.Master.NameU) & " des Zustandsdiagrammes.")
                            HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Zustand"))
                            AddRelation_vsoPage_Overview(vsoPageID, diagramOverviewID, elementID)
                        Else
                            Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                            HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Zustand"))
                            AddRelation_vsoPage_Overview(vsoPageID, diagramOverviewID, elementID)
                        End If
                    Else
                        Throw New System.Exception("Kein gültiges Element.")
                    End If
                Next
                If onlyOneInitialState = False Then Throw New System.Exception("Es muss ein Anfangsknoten geben.")

                If actShapes Is Nothing Then Throw New System.Exception("Keine Verbindungen im Zustandsdiagramm vorhanden.")
                For Each shape As Visio.Shape In actShapes 'Nun werden die gespeicherten Transitionen ausgelesen.
                    elementName = shape.Text
                    If IsNullOrBlank(elementName) Then
                        If shape.Connects.Count < 2 Then Throw New System.Exception("Zustandsübergang " & shape.Text & " ist nicht mit einem Element verbunden.")
                        'Erkennung von dem Vorgänger und Nachgänger Zustand der Transition
                        tempShape = shape.Connects.Item(1).ToSheet 'Anfangsknoten-Bestimmung.
                        If InitFinalTypes(0).IndexOf(GetModelElementBaseName(tempShape.Master.NameU)) >= 0 Then
                            elementName = "Transitionanfang" ' Geht nur bei Anfangsknoten.
                        Else
                            Throw New System.Exception("Zustandsübergang " & shape.Text & " hat kein Text.")
                        End If
                    End If
                    elementID = GetIDofMEDS(elementName, "Activity")
                    If elementID = "" Then
                        elementID = GetIDofMEDS(elementName, "State")
                        If elementID = "" Then
                            elementID = "MEl-" & GetGenID(27)
                            AddObjectActor(elementID, elementName, "ist ein Zustandsübergang. Transitionen können auch Akteure sein und finden sich im Aktivitätsmodell wieder.")
                            AddRelation(elementID, shape)
                            Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                            HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Akteur"))
                            AddRelation_vsoPage_Overview(vsoPageID, diagramOverviewID, elementID)
                        Else
                            'Element existiert schon soll nicht doppelt vorkommmen.
                            AddRelation(elementID, shape)
                        End If
                    Else
                        If GetIDofMEDS(elementName, modelType) = "" Then
                            Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                            HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Akteur"))
                            AddRelation_vsoPage_Overview(vsoPageID, diagramOverviewID, elementID)
                        End If
                        AddRelation(elementID, shape)
                    End If
                Next
            End If
        Else
            Throw New System.Exception("Es wurde kein Diagrammrahmen gefunden.") 'Dieser Fehler wird nicht auftauchen. Wird vorher rausgefiltert.
        End If
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: CheckElement(element As Visio.Shape)
    '       Zweck        : Überprüft, ob das Element entweder eine Verbindung hat.
    '       @param       : Ein Shape-Element
    '       Quellort     : Von CreateStateSpecIFCode
    '    ---------------------------------------------------------------------------------------
    Private Sub CheckElement(element As Visio.Shape)
        Dim shapeIDs As Array
        If StaTypes.ToList().IndexOf(GetModelElementBaseName(element.Master.NameU)) >= 0 Then 'Zustandselemente
            shapeIDs = element.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll1D, "") 'Überprüft ob Verbindungen ein oder ausgehen.
            If UBound(shapeIDs) = -1 Then Throw New System.Exception(element.Name & " besitzt keine Verbindung.")
        ElseIf InitFinalTypes.ToList().IndexOf(GetModelElementBaseName(element.Master.NameU)) >= 0 Then
            If InitFinalTypes(0).IndexOf(GetModelElementBaseName(element.Master.NameU)) >= 0 Then
                shapeIDs = element.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming1D, "")
                If UBound(shapeIDs) >= 0 Then Throw New System.Exception("Ein Anfangsknoten hat keine eingehenden Verbindungen. Bitte korrigieren.")
                shapeIDs = element.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "")
                If UBound(shapeIDs) >= 1 Then Throw New System.Exception("Ein Anfangsknoten hat nur eine ausgehende Verbindung. Bitte korrigieren.")
            Else
                shapeIDs = element.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "")
                If UBound(shapeIDs) >= 0 Then Throw New System.Exception("Ein Endknoten hat keine ausgehenden Verbindungen. Bitte korrigieren")
            End If
        End If
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: SetHierarchie()
    '       Zweck        : Setzt Hierarchie und sortiert nach Alphabet.
    '       @param       : -
    '       Quellort     : Von StateAnalyse(vsoDocumenttemp As Visio.Document)
    '    ---------------------------------------------------------------------------------------
    Private Sub SetHierarchie()
        Dim sortedList As New List(Of ModelElementsDataStructur)
        Dim hrarStr As String, idHr As String

        sortedList = HierarchieList.OrderBy(Function(x) x.Type).ThenBy(Function(x) x.Name).ToList
        For Each Element As ModelElementsDataStructur In sortedList
            If Element.Type = "Activity" Then idHr = Element.ID & "-Act" Else idHr = Element.ID
            hrarStr = "{" & vbCrLf &
            "                   ""id"": ""SH-" & idHr & "-" & modelType & """," & vbCrLf &
            "                   ""object"": """ & Element.ID & """," & vbCrLf &
            "                   ""revision"": 0," & vbCrLf &
            "                   ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
            "                   ""nodes"": []" & vbCrLf &
            "               },##" & Element.Type & "##"
            If Not MELHierarchieElementExist(idHr) Then ProjectData.MELHierarchieDS = Replace(ProjectData.MELHierarchieDS, "##" & Element.Type & "##", hrarStr)
        Next
    End Sub

    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: AddRelation(id As String, shape As Visio.Shape)
    '       Zweck        : Fügt Beziehungen hinzu.
    '       @param       : ID und Shape der Transition
    '       Quellort     : Von CreateStateSpecIFCode()
    '    ---------------------------------------------------------------------------------------
    Private Sub AddRelation(id As String, shape As Visio.Shape)
        Dim fromStateTitle As String, toStateTitle As String
        Dim fromStateID As String, toStateID As String

        If shape.Connects.Count < 2 Then Throw New System.Exception("Zustandsübergang " & shape.Text & " ist nicht mit einem Element verbunden.")
        'Erkennung von dem Vorgänger und Nachgänger Zustand der Transition
        fromStateTitle = shape.Connects.Item(1).ToSheet.Name 'Quellzustandtitel
        If InStr(fromStateTitle, "Sheet") > 0 Then
            fromStateTitle = shape.Connects.Item(1).ToSheet.Parent.Shapes.Item(1).Text
        Else
            If shape.Connects.Item(1).ToSheet.Shapes.Count > 0 Then
                fromStateTitle = shape.Connects.Item(1).ToSheet.Shapes.Item(1).Text
            ElseIf InitFinalTypes.ToList().IndexOf(GetModelElementBaseName(shape.Connects.Item(1).ToSheet.Master.NameU)) >= 0 Then
                fromStateTitle = shape.Connects.Item(1).ToSheet.Name
            End If
        End If
        toStateTitle = shape.Connects.Item(2).ToSheet.Name 'Endzustandtitel.
        If InStr(toStateTitle, "Sheet") > 0 Then
            toStateTitle = shape.Connects.Item(2).ToSheet.Parent.Shapes.Item(1).Text
        Else
            If shape.Connects.Item(2).ToSheet.Shapes.Count > 0 Then
                toStateTitle = shape.Connects.Item(2).ToSheet.Shapes.Item(1).Text
            ElseIf InitFinalTypes.ToList().IndexOf(GetModelElementBaseName(shape.Connects.Item(2).ToSheet.Master.NameU)) >= 0 Then
                toStateTitle = shape.Connects.Item(2).ToSheet.Name
            End If
        End If
        'Überprüfung, ob Zustand aus anderem Modell eingetragen wurde.
        fromStateID = GetIDofMEDS(fromStateTitle, "Activity")
        If fromStateID = "" Then fromStateID = GetIDofMEDS(fromStateTitle, modelType)
        toStateID = GetIDofMEDS(toStateTitle, "Activity")
        If toStateID = "" Then toStateID = GetIDofMEDS(toStateTitle, modelType)

        If fromStateID <> "" Then 'Beziehungen hinzufügen.
            If Not RelationExist(fromStateID, id) Then
                ProjectData.RelationenDS = ProjectData.RelationenDS & "{" & vbCrLf &
            "        ""id"": ""RVis-" & fromStateID & "-" & id & """," & vbCrLf &
            "        ""title"": ""führt zu""," & vbCrLf &
            "        ""revision"": 0," & vbCrLf &
            "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
            "        ""relationType"": ""RT-FuehrtZu""," & vbCrLf &
            "        ""source"": {" & vbCrLf &
            "            ""id"": """ & fromStateID & """," & vbCrLf &
            "            ""revision"": 0" & vbCrLf &
            "        }," & vbCrLf &
            "        ""target"": {" & vbCrLf &
            "            ""id"": """ & id & """," & vbCrLf &
            "            ""revision"": 0" & vbCrLf &
            "        }" & vbCrLf &
            "    },"
            End If
            If Not RelationExist(id, toStateID) Then
                ProjectData.RelationenDS = ProjectData.RelationenDS & "{" & vbCrLf &
            "        ""id"": ""RVis-" & id & "-" & RelationExistOverall(id, toStateID) & """," & vbCrLf &
            "        ""title"": ""resultiert in""," & vbCrLf &
            "        ""revision"": 0," & vbCrLf &
            "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
            "        ""relationType"": ""RT-resultiertIn""," & vbCrLf &
            "        ""source"": {" & vbCrLf &
            "            ""id"": """ & id & """," & vbCrLf &
            "            ""revision"": 0" & vbCrLf &
            "        }," & vbCrLf &
            "        ""target"": {" & vbCrLf &
            "            ""id"": """ & toStateID & """," & vbCrLf &
            "            ""revision"": 0" & vbCrLf &
            "        }" & vbCrLf &
            "    },"
            End If
        End If
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Function     : RelationExistOverall(overviewID As String, elementID As String) As String
    '       Zweck        : Überprüft ob schon Relation im Integrationsmodell existiert und gibt ID wieder
    '       @param       : ID von Diagramm Overview und das ausgewählte Element
    '       @return      : Gibt ID wieder
    '       Quellort     : Von AddRelation(id As String, shape As Visio.Shape)
    '    ---------------------------------------------------------------------------------------
    Private Function RelationExistOverall(overviewID As String, elementID As String) As String
        'Brauche ich, da es unterschiedliche Relationen gibt zur gleichen ID
        'Beispiel-Relationen: act warten enthält Wartung im Activitydiagramm.  Aber warten resultiert in Wartung im Zustandsdiagramm
        If InStr(ProjectData.CompleteDS, "RVis-" & overviewID & "-" & elementID) > 0 Then
            Return elementID & GetGenID(3)
        Else
            Return elementID
        End If
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : RelationExist(overviewID As String, elementID As String) As Boolean
    '       Zweck        : Überprüft ob schon Relation in der Relationsliste existiert.
    '       @param       : ID von Diagramm Overview und das ausgewählte Element
    '       @return      : Falls vorhanden, dann true.
    '       Quellort     : Von AddRelation(id As String, shape As Visio.Shape)
    '    ---------------------------------------------------------------------------------------
    Private Function RelationExist(overviewID As String, elementID As String) As Boolean
        If InStr(ProjectData.RelationenDS, "RVis-" & overviewID & "-" & elementID) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: AddObjectActor(id As String, name As String)
    '       Zweck        : Erzeugt Element als Akteurobjekt einer JSON-Struktur.
    '       @param       : ID und Titel
    '       Quellort     : Von CreateStateSpecIFCode()
    '    ---------------------------------------------------------------------------------------
    Private Sub AddObjectActor(id As String, name As String, text As String)
        ProjectData.ObjekteDS = ProjectData.ObjekteDS & "{" & vbCrLf &
        "        ""id"": """ & id & """," & vbCrLf &
        "        ""title"": """ & name & """," & vbCrLf &
        "        ""attributes"": [{" & vbCrLf &
        "            ""title"": ""dcterms:title""," & vbCrLf &
        "            ""attributeType"": ""AT-Act-Name""," & vbCrLf &
        "            ""value"": """ & name & """" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""dcterms:description""," & vbCrLf &
        "            ""attributeType"": ""AT-Act-Text""," & vbCrLf &
        "            ""value"": ""<div><p>" & text & "</p></div>""" & vbCrLf &
        "        }]," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""objectType"": ""OT-Act""" & vbCrLf &
        "    },"
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: CreateStateSpecIFBase()
    '       Zweck        : Erstellt die JSON Basis-Struktur mit Objekten und Hierarchie
    '       @param       : -
    '       Quellort     : Von RequirementsAnalyse(xlWorkBooktemp As Excel.Workbook)
    '    ---------------------------------------------------------------------------------------
    Private Sub CreateStateSpecIFBase()
        Dim svgName As String = ConvUmlaut(RemoveWhitespace(System.IO.Path.GetFileNameWithoutExtension(vsoDocument.FullName)) & ".svg")
        Modelelements.Add(New ModelElementsDataStructur(vsoPage.Name, "Pln-Zustandsmodell", modelType))
        Dim id As String = GetIDofMEDS(vsoPage.Name, modelType)
        ProjectData.ObjekteDS = ProjectData.ObjekteDS & "{" & vbCrLf &
        "       ""id"": """ & GetIDofMEDS(vsoPage.Name, modelType) & """," & vbCrLf &
        "        ""title"": ""Zustandsdiagramm""," & vbCrLf &
        "        ""attributes"": [{" & vbCrLf &
        "            ""title"": ""dcterms:title""," & vbCrLf &
        "            ""attributeType"": ""AT-Pln-Name""," & vbCrLf &
        "            ""value"": ""Zustandsdiagramm""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""dcterms:description""," & vbCrLf &
        "            ""attributeType"": ""AT-Pln-Text""," & vbCrLf &
        "            ""value"": ""<div><p>" & modelDescription & "</p><p class=\""inline-label\"">Model View: \n</p>\n<div class=\""forImage\"" style=\""max-width: 900px;\"" >\n\t<div class=\""forImage\""><a href=\""/reqif/v0.92/projects/" & ProjectData.projectinfoID & "/files/" & svgName & "\"" type=\""image/svg+xml\"" ><object data=\""/reqif/v0.92/projects/" & ProjectData.projectinfoID & "/files/" & svgName & "\"" type=\""image/svg+xml\"" >files_and_images\\" & svgName & "</object></a></div></div></div>""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""SpecIF:Type""," & vbCrLf &
        "            ""attributeType"": ""AT-Pln-Type""," & vbCrLf &
        "            ""value"": ""UML Zustandsmodell""" & vbCrLf &
        "        }]," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""objectType"": ""OT-Pln""" & vbCrLf &
        "    },"

        ProjectData.HierarchieDS = "{" & vbCrLf &
        "                ""id"": ""SH-" & id & """," & vbCrLf &
        "                ""object"": """ & id & """," & vbCrLf &
        "                ""revision"": 0," & vbCrLf &
        "				""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "                ""nodes"": []" & vbCrLf &
        "		},"
    End Sub
End Module
