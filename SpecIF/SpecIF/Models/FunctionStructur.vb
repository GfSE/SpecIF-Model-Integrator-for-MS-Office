Imports Visio = Microsoft.Office.Interop.Visio
'    ---------------------------------------------------------------------------------------
'       Modul   : FunctionStructur
'       Datum   : 29.04.2017
'       Author  : Philipp Mochine
'       Zweck   : Die Funktionsstruktur wird eingelesen und in das Integrationsmodell hinzugefügt.
'                 Beziehungen zwischen Funktionsstruktur und Anforderungen werden hier automatisch gesetzt.
'       Quellort: Von ExecuteModel(modelPath As String)
'    ---------------------------------------------------------------------------------------
Module FunctionStructur
    Private vsoDocument As Visio.Document
    Private vsoPage As Visio.Page
    Private Const modelType As String = "Function"
    Private Const overviewType As String = "Diagram Overview" 'Müsste ich bei den anderen auch ändern.
    Private Const functionstructurType As String = "Action"
    Private HierarchieList As New List(Of ModelElementsDataStructur)
    'Erweiterte Form: PreID und ID wird gespeichert
    'Wird benötigt um eine Hierarchie darzustellen.
    Private FunctionStructurList As New List(Of ModelElementsDataStructur)
    Private modelDescription As String = ConvertStringToJson("In der Entwicklung von technischen bzw. mechatronischen Systemen werden Funktionsmodelle benutzt, " &
                                               "um möglichst abstrakt Funktionen des Systems zu beschreiben. 'Funktionen beschreiben das Verhalten von Produkten, " &
                                               "oder Teilen des Produktes, vorzugsweise in Form eines Zusammenhangs zwischen Eingangs- und Ausgangsgrößen [...]'(VDI 2222 Blatt 1). " &
                                               "Eine typische Modellform ist die Funktionshierarchie. Sie stellt die hierarchische Abhängigkeit der Funktionen in Form einer Baumstruktur untereinander dar. " &
                                               "In der Wurzel des Baumes steht die Gesamtfunktion. In der Regel besteht sie selber aus mehreren Teilfunktionen. " &
                                               "Mit Hilfe von Funktionsmodellen ist es also möglich komplexe Aufgabenstellungen in weniger komplexe Teilaufgaben zu teilen, um dann die Gesamtaufgabe zu beherrschen.")
    Private frameDescription As String = ConvertStringToJson("ist der Diagrammtitel der Funktionsstruktur. <p>Die Syntax eines UML-Diagrammtitels lautet: " &
                                                "<b>Diagrammtyp </b>Diagrammname [Parameter]</p><p>Da UML kein Funktionsstruktur-Modell mitliefert, wird häufig ein Klassendiagramm verwendet.</p>")
    Private Const elementDescription As String = "ist ein Funktionselement der Funktionstruktur."
    '    ---------------------------------------------------------------------------------------
    '       Function     : FunctionStructurAnalyse(vsoDocumenttemp As Visio.Document) As Boolean
    '       Zweck        : Liest das Visiodokument ein, erstellt ein Bild und die JSON Struktur. 
    '       @param       : Pointer auf das Visio Dokument.
    '       Quellort     : Von ExecuteModel(modelPath As String)
    '    ---------------------------------------------------------------------------------------
    Public Function FunctionStructurAnalyse(vsoDocumenttemp As Visio.Document) As Boolean
        FunctionStructurAnalyse = True
        vsoDocument = vsoDocumenttemp
        vsoPage = vsoDocument.Pages(1)
        Try
            SvgExport(vsoDocument)
            CreateFunctionstructurSpecIFBase()
            CreateFunctiontructurSpecIFCode()
            SetHierarchie()
        Catch ex As Exception
            FunctionStructurAnalyse = False
            SendError(ex.Message, -1)
        Finally
            ReleaseObject(vsoPage)
            HierarchieList = New List(Of ModelElementsDataStructur)
            FunctionStructurList = New List(Of ModelElementsDataStructur)
        End Try
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: FunctionStructurAnalyse(vsoDocumenttemp As Visio.Document)
    '       Zweck        : Hauptcode des Moduls.
    '                       1. Suche nach "Diagram Overview". 
    '                       2. Im Diagram Overview wird nach dem Anfangsknoten gesucht.  
    '                       3. jetzt wird durch jedes Element gegangen
    '       @param       : -
    '       Quellort     : Von FunctionStructurAnalyse(vsoDocumenttemp As Visio.Document)
    '    ---------------------------------------------------------------------------------------
    Private Sub CreateFunctiontructurSpecIFCode()
        Dim ShapeOverview As Visio.Shape = Nothing
        Dim elementID As String, elementName As String
        Dim vsoPageID As String = GetIDofMEDS(vsoPage.Name, modelType), diagramOverviewID As String = ""
        Dim vsoReturnedSelection As Visio.Selection
        Dim shapeElement As Visio.Shape, vsoShapeOnPage As Visio.Shape = Nothing
        Dim shapeIDs As Array
        Dim rootExists As Boolean = False, onlyOneOverview As Boolean = False

        'Finde im Overview das oberste Element. Dieses hat keine Incoming Verbindung
        For Each ShapeOverview In vsoPage.Shapes
            If vsoShapeOnPage Is Nothing Then
                If Not ShapeOverview Is Nothing And ShapeOverview.Master.NameU.Contains(overviewType) Then 'Diagrammrahmen finden.
                    If onlyOneOverview = False Then onlyOneOverview = True Else Throw New System.Exception("Nur ein Diagrammrahmen erlaubt.")
                    elementID = "MEl-" & GetGenID(27)
                    diagramOverviewID = elementID
                    elementName = ShapeOverview.Shapes.Item(1).Text
                    If IsNullOrBlank(elementName) Then Throw New System.Exception("Diagrammrahmen muss einen Titel haben.")
                    Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                    AddRelation_vsoPage_Overview(vsoPageID, "", elementID)
                    AddObjectState(elementID, elementName, frameDescription)
                    HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Zustand"))

                    vsoReturnedSelection = ShapeOverview.SpatialNeighbors(Visio.VisSpatialRelationCodes.visSpatialContain, 0, Visio.VisSpatialRelationFlags.visSpatialIncludeContainerShapes)
                    If Not vsoReturnedSelection Is Nothing And vsoReturnedSelection.Count <> 0 Then 'Ergebnis nicht leer und hat mehr als ein Element.
                        For Each shapeElement In vsoReturnedSelection 'Es wird solange gesucht, bis ein Element gefunden wird, dass keine Incoming Verknüpfung hat.
                            If shapeElement.Master.NameU = functionstructurType Then 'Nur Aktionselemente werden betrachtet.
                                shapeIDs = shapeElement.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming1D, "")
                                If UBound(shapeIDs) = -1 Then
                                    'Root gefunden
                                    If rootExists = False Then
                                        vsoShapeOnPage = shapeElement
                                        rootExists = True
                                    Else
                                        Throw New System.Exception("Mehrere Funktionselemente sind nicht verbunden. Der Fehler trat bei diesem Element auf: " & shapeElement.Text)
                                    End If
                                End If
                            End If
                        Next
                    Else
                        Throw New System.Exception("Keine Elemente im Diagrammrahmen gefunden.")
                    End If
                End If
            Else
                'Root in Diagrammrahmen gefunden gefunden.
                Exit For
            End If
        Next

        If vsoShapeOnPage IsNot Nothing Then
            Dim PreID As String = "", AfterID As String = ""
            Dim TempShape As Visio.Shape
            Dim QueueShape As Collection
            Dim QueueShapePreID As Collection

            TempShape = vsoShapeOnPage 'Knoten ins temporäre Shapeelement hinzufügen.
            QueueShape = New Collection 'Liste von Funktionen in der Warteschlange
            QueueShapePreID = New Collection 'Liste der IDs von Vorknoten.
            QueueShape.Add(TempShape)
            QueueShapePreID.Add("")

            'Knoten wird in die Funktionsstrukturliste hinzugefügt. Der Aufbau erfolgt in Baumweise.
            AfterID = GenerateIDAddRelation(TempShape.Text)
            FunctionStructurList.Add(New ModelElementsDataStructur(TempShape.Text, AfterID, PreID))
            Modelelements.Add(New ModelElementsDataStructur(TempShape.Text, AfterID, modelType))
            Do Until QueueShape.Count = 0 'Nun wird die Warteschleife abgearbeitet bis kein Element vorhanden ist.
                TempShape = QueueShape(QueueShape.Count)
                PreID = QueueShapePreID(QueueShapePreID.Count)
                AddObjectActor(PreID, TempShape.Text, elementDescription)
                HierarchieList.Add(New ModelElementsDataStructur(TempShape.Text, GetAfterID(TempShape.Text, PreID), "Akteur")) 'Fügt Element in die Hierarchieliste ein.
                AddRelation_vsoPage_Overview(vsoPageID, diagramOverviewID, GetAfterID(TempShape.Text, PreID))

                PreID = AddStructurFolder(PreID, TempShape.Text) 'Erstellt die Baumstruktur bzw. fügt das Element hinzu
                QueueShape.Remove(QueueShape.Count)
                QueueShapePreID.Remove(QueueShapePreID.Count)

                shapeIDs = TempShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "") 'Alle ausgehenden Verbindungen werden angeschaut.          
                'Geh durch alle Akteure
                If UBound(shapeIDs) >= 0 Then 'Weitere Blätter vorhanden
                    For i = 0 To UBound(shapeIDs)
                        If vsoPage.Shapes.ItemFromID(shapeIDs(i)).Connects.Count > 1 Then
                            QueueShape.Add(vsoPage.Shapes.ItemFromID(shapeIDs(i)).Connects.Item(2).ToSheet)
                            QueueShapePreID.Add(PreID)
                            If QueueShape(QueueShape.Count).Master.NameU.Contains(functionstructurType) Then
                                'Wichtig um die ID zu erhalten für Pre Knoten
                                AfterID = GenerateIDAddRelation(QueueShape(QueueShape.Count).Text) 'Erhalte ID entweder generiert oder aus der Anwendung.
                                AddRelation(PreID, AfterID) 'Relationen in der Hierarchieebene
                                FunctionStructurList.Add(New ModelElementsDataStructur(QueueShape(QueueShape.Count).Text, AfterID, PreID))
                            Else
                                'Fehlermeldung: Verbindung ist an einer Verbindung befestigt
                                QueueShape.Remove(QueueShape.Count)
                                QueueShapePreID.Remove(QueueShapePreID.Count)
                                Throw New System.Exception("Die ausgehende Verbindung des Elements ist mit einer Verbindung verbunden. Elementname: " & TempShape.Text)
                            End If
                        Else
                            Throw New System.Exception("Die ausgehende Verbindung des Elements ist nicht mit einem Element verbunden. Elementname: " & TempShape.Text)
                        End If
                    Next i
                    'Else -> 'Keine weitere Blätter gefunden. Ende erreicht.
                End If
            Loop

        Else
            Throw New System.Exception("Es wurde keine Wurzel in der Baumhierarchie gefunden. Bitte überprüfen.")
        End If

    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: AddRelation(PreId As String, AfterID As String)
    '       Zweck        : Fügt Beziehung "enthält" hinzu. Wichtig für die Hierarchieebene.
    '       @param       : Id von Vorknoten und Knoten
    '       Quellort     : CreateFunctiontructurSpecIFCode()
    '    ---------------------------------------------------------------------------------------
    Private Sub AddRelation(PreId As String, AfterID As String)
        ProjectData.RelationenDS = ProjectData.RelationenDS & "{" & vbCrLf &
    "        ""id"": ""RVis-" & PreId & "-" & AfterID & """," & vbCrLf &
    "        ""title"": ""enthält""," & vbCrLf &
    "        ""revision"": 0," & vbCrLf &
    "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
    "        ""relationType"": ""RT-Containment""," & vbCrLf &
    "        ""source"": {" & vbCrLf &
    "            ""id"": """ & PreId & """," & vbCrLf &
    "            ""revision"": 0" & vbCrLf &
    "        }," & vbCrLf &
    "        ""target"": {" & vbCrLf &
    "            ""id"": """ & AfterID & """," & vbCrLf &
    "            ""revision"": 0" & vbCrLf &
    "        }" & vbCrLf &
    "    },"
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Function     : AddStructurFolder(PreID As String, elementName As String) As String
    '       Zweck        : Erzeugt Hierarchieobjekt und fügt diese in die aktuelle Hierarchiestruktur ein.
    '       @param       : Id des Vorknotens und Titel der Funktion
    '       @return      : Die erstellte ID des Objektes
    '       Quellort     : CreateFunctiontructurSpecIFCode()
    '    ---------------------------------------------------------------------------------------
    Private Function AddStructurFolder(PreID As String, elementName As String) As String
        Dim hierarchieTemp As String = ""
        Dim AfterID As String = GetAfterID(elementName, PreID)
        hierarchieTemp = hierarchieTemp & vbCrLf & "    {" & vbCrLf &
        "       ""id"": ""SH-" & AfterID & """," & vbCrLf &
        "       ""object"": """ & AfterID & """," & vbCrLf &
        "       ""revision"": 0," & vbCrLf &
        "       ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "       ""nodes"": [##" & AfterID & "##]" & vbCrLf &
        "    },"
        Modelelements.Add(New ModelElementsDataStructur(elementName, AfterID, modelType))

        If PreID = "" Then 'Nur der Fall, wenn das Element der Knoten der Baumstruktur ist. 
            ProjectData.HierarchieDS = Replace(ProjectData.HierarchieDS, "##FunctionStructur##", Left(hierarchieTemp, Len(hierarchieTemp) - 1))
            Return AfterID
        Else
            Dim elementID As String = PreID
            ProjectData.HierarchieDS = Replace(ProjectData.HierarchieDS, "##" & elementID & "##", hierarchieTemp & "##" & elementID & "##")
            Return AfterID
        End If
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : GenerateIDAddRelation(name As String) As String
    '       Zweck        : Erstellt ID, wenn nicht in Anforderung schon vorhanden. Falls doch
    '                      Wird eine Beziehung "erfüllt" erstellt und weißt auf Anforderung hin.
    '       @param       : Id des Vorknotens und Titel dr Funktion
    '       @return      : Die erstellte ID des Objektes
    '       Quellort     : CreateFunctiontructurSpecIFCode()
    '    ---------------------------------------------------------------------------------------
    Private Function GenerateIDAddRelation(name As String) As String
        Dim elementID As String, returnID As String
        elementID = GetIDofMEDS(name, "Requirement")
        If elementID = "" Then
            'Anforderung wurde noch nicht gelesen. 
            Return "Mel-" & GetGenID(27)
        Else
            'Anforderung gelesen.
            returnID = "Mel-" & GetGenID(27)
            ProjectData.RelationenDS = ProjectData.RelationenDS & "{" & vbCrLf &
            "        ""id"": ""RVis-" & elementID & "-" & returnID & """," & vbCrLf &
            "        ""title"": ""erfüllt""," & vbCrLf &
            "        ""revision"": 0," & vbCrLf &
            "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
            "        ""relationType"": ""RT-Satisfaction""," & vbCrLf &
            "        ""source"": {" & vbCrLf &
            "            ""id"": """ & returnID & """," & vbCrLf &
            "            ""revision"": 0" & vbCrLf &
            "        }," & vbCrLf &
            "        ""target"": {" & vbCrLf &
            "            ""id"": """ & elementID & """," & vbCrLf &
            "            ""revision"": 0" & vbCrLf &
            "        }" & vbCrLf &
            "    },"
            Return returnID
        End If
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : GetAfterID(name As String, PreID As String) As String
    '       Zweck        : Ermittelt ID eines Elementes der veränderten Functionsstrukturliste.
    '       @param       : Mit den Titel des Elementes und der ID der Vorknotens wird die ID ermittelt
    '       @return      : Die erstellte ID des Objektes
    '    ---------------------------------------------------------------------------------------
    Private Function GetAfterID(name As String, PreID As String) As String
        For Each Element As ModelElementsDataStructur In FunctionStructurList
            If Element.Name = name And Element.Type = PreID Then
                Return Element.ID
            End If
        Next
        Return ""
    End Function

    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: AddObjectActor(id As String, name As String, text As String)
    '       Zweck        : Fügt Funktion als Objekt in die JSON Struktur ein.
    '       @param       : ID, Titel, Text
    '       Quellort     : CreateFunctiontructurSpecIFCode()
    '    ---------------------------------------------------------------------------------------
    Private Sub AddObjectActor(id As String, name As String, text As String)
        If IsNullOrBlank(name) Then Throw New System.Exception("Es fehlt der Text bzw. die Beschreibung in einem Modellelement. Bitte nicht leer lassen.")
        id = GetAfterID(name, id)
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
    '       Sub-Anweisung: CreateFunctionstructurSpecIFBase()
    '       Zweck        : Erstellt die Json Basisstruktur für Objekt und Hierarchie
    '       @param       : -
    '       Quellort     : FunctionStructurAnalyse()
    '    ---------------------------------------------------------------------------------------
    Private Sub CreateFunctionstructurSpecIFBase()
        Dim svgName As String = ConvUmlaut(RemoveWhitespace(System.IO.Path.GetFileNameWithoutExtension(vsoDocument.FullName)) & ".svg")
        Modelelements.Add(New ModelElementsDataStructur(vsoPage.Name, "Pln-Funktionshierarchie", modelType))
        Dim id As String = GetIDofMEDS(vsoPage.Name, modelType)
        ProjectData.ObjekteDS = ProjectData.ObjekteDS & "{" & vbCrLf &
        "       ""id"": """ & id & """," & vbCrLf &
        "        ""title"": ""Funktionshierarchie""," & vbCrLf &
        "        ""attributes"": [{" & vbCrLf &
        "            ""title"": ""dcterms:title""," & vbCrLf &
        "            ""attributeType"": ""AT-Pln-Name""," & vbCrLf &
        "            ""value"": ""Funktionshierarchie""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""dcterms:description""," & vbCrLf &
        "            ""attributeType"": ""AT-Pln-Text""," & vbCrLf &
        "            ""value"": ""<div><p>" & modelDescription & "</p><p class=\""inline-label\"">Model View: \n</p>\n<div class=\""forImage\"" style=\""max-width: 900px;\"" >\n\t<div class=\""forImage\""><a href=\""/reqif/v0.92/projects/" & ProjectData.projectinfoID & "/files/" & svgName & "\"" type=\""image/svg+xml\"" ><object data=\""/reqif/v0.92/projects/" & ProjectData.projectinfoID & "/files/" & svgName & "\"" type=\""image/svg+xml\"" >files_and_images\\" & svgName & "</object></a></div></div></div>""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""SpecIF:Type""," & vbCrLf &
        "            ""attributeType"": ""AT-Pln-Type""," & vbCrLf &
        "            ""value"": ""Funktionsmodell""" & vbCrLf &
        "        }]," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""objectType"": ""OT-Pln""" & vbCrLf &
        "    },"

        ProjectData.HierarchieDS = "{" & vbCrLf &
        "        ""id"": ""SH-" & id & """," & vbCrLf &
        "        ""object"": """ & id & """," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""nodes"": [##FunctionStructur##]" & vbCrLf &
        "},"
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: SetHierarchie()
    '       Zweck        : Setzt eine sortierte Liste in die Hierarchie
    '       @param       : -
    '       Quellort     : FunctionStructurAnalyse()
    '    ---------------------------------------------------------------------------------------
    Private Sub SetHierarchie()
        Dim sortedList As New List(Of ModelElementsDataStructur)
        Dim hrarStr As String, idHr As String

        sortedList = HierarchieList.OrderBy(Function(x) x.Type).ThenBy(Function(x) x.Name).ToList
        For Each Element As ModelElementsDataStructur In sortedList
            If Element.Type = "Activity" Then idHr = Element.ID & "-Act" Else idHr = Element.ID
            hrarStr = "{" & vbCrLf &
            "                   ""id"": ""SH-H-" & idHr & "-" & modelType & """," & vbCrLf &
            "                   ""object"": """ & Element.ID & """," & vbCrLf &
            "                   ""revision"": 0," & vbCrLf &
            "                   ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
            "                   ""nodes"": []" & vbCrLf &
            "               },##" & Element.Type & "##"
            If Not MELHierarchieElementExist(idHr) Then ProjectData.MELHierarchieDS = Replace(ProjectData.MELHierarchieDS, "##" & Element.Type & "##", hrarStr)
        Next
    End Sub
End Module
