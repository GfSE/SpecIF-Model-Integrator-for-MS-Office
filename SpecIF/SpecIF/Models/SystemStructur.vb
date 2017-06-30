'Falls ich es vergesse, bei der Kommentierung: Jedes Element ist unique und kriegt trotz doppel einträge eine eigene ID
Imports Visio = Microsoft.Office.Interop.Visio
'    ---------------------------------------------------------------------------------------
'       Modul   : SystemStructur
'       Datum   : 29.04.2017
'       Author  : Philipp Mochine
'       Zweck   : Die Systemstruktur wird eingelesen und in das Integrationsmodell hinzugefügt.
'       Quellort: Von ExecuteModel(modelPath As String)
'    ---------------------------------------------------------------------------------------
Module SystemStructur
    Private vsoDocument As Visio.Document
    Private vsoPage As Visio.Page
    Private Const modelType As String = "System"
    'Ist versehentlich passiert. Zeigt aber, dass natürlich auch andere Modellelemente verwendet werden können. Eigentlich müsste es Diagramm Overview heißen.
    Private Const overviewType As String = "Package (expanded)"
    Private Const systemstructurType As String = "Interface"
    Private ReadOnly Property WrongConnections As String() = New String() {"Sheet", "Composition", "Dynamic connector"}
    Private HierarchieList As New List(Of ModelElementsDataStructur)
    Private SystemStructurList As New List(Of ModelElementsDataStructur) 'Erweiterte Form: PreID und ID wird gespeichert
    Private modelDescription As String = ConvertStringToJson("Die Systemstrukturen sind die statischen Bausteine der logischen Architektur, die die technischen Konzepte und Prinzipien des Systems abstrakt beschreiben. " &
                                               "Die Suche nach den Systemelementen folgt nach der Ermittlung von Lösungen in der Funktionsmodellierung. Systemelemente sind Elemente, die die vorgegebenen Funktionen realisieren. " &
                                               "Damit das technische System im Hinblick der funktionalen Anforderungen vollständig beschrieben ist, muss jede Funktion durch Systemelemente erfüllt werden. " &
                                               "Aus der Spitze des Systems bilden sich viele weitere Systeme, die wiederrum aus weiteren Systemen bestehen. Aus der Produktsicht sind dies Baugruppen wobei die größtmögliche Aufteilung als Komponente bezeichnet wird.")
    Private frameDescription As String = ConvertStringToJson("ist der Diagrammtitel der Systemstruktur. <p>Die Syntax eines SysML-Diagrammtitels lautet: " &
                    "<b>Diagrammtyp </b>[Modellelementtyp] Modellelement [Diagrammnamen]</p><p>Da SysML kein Systemstruktur-Modell mitliefert, wird häufig ein Blockdefinitionsdiagramm verwendet.</p>")
    Private Const elementDescription As String = "ist ein Systemelement der Systemstruktur."

    '    ---------------------------------------------------------------------------------------
    '       Function     : SystemstructurAnalyse(vsoDocumenttemp As Visio.Document) As Boolean
    '       Zweck        : Liest das Visiodokument ein, erstellt ein Bild und die JSON Struktur. 
    '       @param       : Pointer auf das Visio Dokument.
    '       Quellort     : Von ExecuteModel(modelPath As String)
    '    ---------------------------------------------------------------------------------------
    Public Function SystemstructurAnalyse(vsoDocumenttemp As Visio.Document) As Boolean
        SystemstructurAnalyse = True
        vsoDocument = vsoDocumenttemp
        vsoPage = vsoDocument.Pages(1)
        Try
            SvgExport(vsoDocument)
            CreateSystemstructurSpecIFBase()
            CreateSystemstructurSpecIFCode()
            SetHierarchie()
        Catch ex As Exception
            SystemstructurAnalyse = False
            SendError(ex.Message, -1)
        Finally
            ReleaseObject(vsoPage)
            HierarchieList = New List(Of ModelElementsDataStructur)
            SystemStructurList = New List(Of ModelElementsDataStructur)
        End Try
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: FunctionStructurAnalyse(vsoDocumenttemp As Visio.Document)
    '       Zweck        : Hauptcode des Moduls.
    '                       1. Suche nach "Diagram Overview". 
    '                       2. Im Diagram Overview wird nach dem Anfangsknoten gesucht.  
    '                       3. jetzt wird durch jedes Element gegangen
    '       @param       : -
    '       Quellort     : Von SystemstructurAnalyse(vsoDocumenttemp As Visio.Document)
    '    ---------------------------------------------------------------------------------------
    Private Sub CreateSystemstructurSpecIFCode()
        Dim ShapeOverview As Visio.Shape = Nothing
        Dim elementID As String, elementName As String
        Dim vsoPageID As String = GetIDofMEDS(vsoPage.Name, modelType), diagramOverviewID As String = ""
        Dim containerProps As Visio.ContainerProperties
        Dim lngContainerMembers() As Int32, shapeIDs() As Int32
        Dim shapeElement As Visio.Shape, vsoShapeOnPage As Visio.Shape = Nothing
        Dim rootExists As Boolean = False, onlyOneOverview As Boolean = False 'Nur ein Diagrammrahmen erlaubt.

        'Finde im Overview das oberste Element. Dieses hat keine Incoming Verbindung
        For Each ShapeOverview In vsoPage.Shapes
            If vsoShapeOnPage Is Nothing Then
                If ShapeOverview IsNot Nothing And ShapeOverview.Master.NameU.Contains(overviewType) Then  'Diagrammrahmen finden.
                    If onlyOneOverview = False Then onlyOneOverview = True Else Throw New System.Exception("Nur ein Diagrammrahmen erlaubt.")
                    elementID = "MEl-" & GetGenID(27)
                    diagramOverviewID = elementID
                    elementName = ShapeOverview.Shapes.Item(1).Text
                    If IsNullOrBlank(elementName) Then Throw New System.Exception("Diagrammrahmen muss einen Titel haben.")
                    Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                    AddRelation_vsoPage_Overview(vsoPageID, "", elementID)
                    AddObjectState(elementID, elementName, frameDescription)
                    HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Zustand"))

                    containerProps = ShapeOverview.ContainerProperties
                    If containerProps IsNot Nothing Then
                        lngContainerMembers = containerProps.GetMemberShapes(Visio.VisContainerFlags.visContainerFlagsDefault)
                        For Each varMember As Long In lngContainerMembers 'Es wird solange gesucht, bis Root gefunden wird.
                            shapeElement = vsoPage.Shapes.ItemFromID(varMember)
                            If shapeElement.Master.NameU.Contains(systemstructurType) Then
                                shapeIDs = shapeElement.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming1D, "")
                                If UBound(shapeIDs) = -1 Then
                                    If rootExists = False Then
                                        vsoShapeOnPage = shapeElement
                                        rootExists = True
                                    Else
                                        Throw New System.Exception("Mehrere Systemelemente sind nicht verbunden. Der Fehler trat bei diesem Element auf: " & shapeElement.Text)
                                    End If
                                End If
                            End If
                        Next
                    Else
                        Throw New System.Exception("Keine Elemente im Diagrammrahmen gefunden.")
                    End If
                End If
            Else
                'Root in Package gefunden.
                Exit For
            End If
        Next

        If Not vsoShapeOnPage Is Nothing Then
            Dim PreID As String = "", AfterID As String = ""

            Dim TempShape As Visio.Shape
            Dim QueueShape As Collection
            Dim QueueShapePreID As Collection

            TempShape = vsoShapeOnPage
            QueueShape = New Collection 'Liste von Systemelemente in der Warteschlange
            QueueShapePreID = New Collection 'Liste der IDs von Vorknoten.
            QueueShape.Add(TempShape)
            QueueShapePreID.Add("")

            'Knoten wird in die Systemsstrukturliste hinzugefügt. Der Aufbau erfolgt in Baumweise.
            AfterID = "Mel-" & GetGenID(27)
            SystemStructurList.Add(New ModelElementsDataStructur(TempShape.Text, AfterID, PreID))
            Modelelements.Add(New ModelElementsDataStructur(TempShape.Text, AfterID, modelType))

            Do Until QueueShape.Count = 0 'Nun wird die Warteschleife abgearbeitet bis kein Element vorhanden ist.
                TempShape = QueueShape(QueueShape.Count)
                PreID = QueueShapePreID(QueueShapePreID.Count)
                AddObjectActor(PreID, TempShape.Text, elementDescription)
                HierarchieList.Add(New ModelElementsDataStructur(TempShape.Text, GetAfterID(TempShape.Text, PreID), "Akteur"))
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
                            If WrongConnections.Contains(QueueShape(QueueShape.Count).Master.NameU) Then
                                'Fehlermeldung: Verbindung ist an einer Verbindung befestigt
                                QueueShape.Remove(QueueShape.Count)
                                QueueShapePreID.Remove(QueueShapePreID.Count)
                                Throw New System.Exception("Die ausgehende Verbindung des Elements ist mit einer Verbindung verbunden. Elementname: " & TempShape.Text)
                            Else
                                'Wichtig um die ID zu erhalten für Pre Knoten
                                AfterID = "Mel-" & GetGenID(27)
                                AddRelation(PreID, AfterID) 'Relationen in der Hierarchieebene
                                SystemStructurList.Add(New ModelElementsDataStructur(QueueShape(QueueShape.Count).Text, AfterID, PreID))
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
    '       Zweck        : Fügt Beziehung "enthält" hinzu 
    '       @param       : Id von Vorknoten und Knoten
    '       Quellort     : CreateSystemstructurSpecIFCode()
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
    '       Sub-Anweisung: AddObjectActor(id As String, name As String, text As String)
    '       Zweck        : Fügt Funktion als Objekt in die JSON Struktur ein.
    '       @param       : ID, Titel, Text
    '       Quellort     : CreateSystemstructurSpecIFCode()
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
    '       Function     : AddStructurFolder(PreID As String, elementName As String) As String
    '       Zweck        : Erzeugt Hierarchieobjekt und fügt diese in die aktuelle Hierarchiestruktur ein.
    '       @param       : Id des Vorknotens und Titel der Funktion
    '       @return      : Die erstellte ID des Objektes
    '       Quellort     : CreateSystemstructurSpecIFCode()
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

        If PreID = "" Then 'Im Fall des Knotens
            ProjectData.HierarchieDS = Replace(ProjectData.HierarchieDS, "##Stueckliste##", Left(hierarchieTemp, Len(hierarchieTemp) - 1))
            Return AfterID
        Else
            Dim elementID As String = PreID
            ProjectData.HierarchieDS = Replace(ProjectData.HierarchieDS, "##" & elementID & "##", hierarchieTemp & "##" & elementID & "##")
            Return AfterID
        End If

    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : GetAfterID(name As String, PreID As String) As String
    '       Zweck        : Ermittelt ID eines Elementes der veränderten Systemstrukturliste.
    '       @param       : Mit den Titel des Elementes und der ID der Vorknotens wird die ID ermittelt
    '       @return      : Die erstellte ID des Objektes
    '    ---------------------------------------------------------------------------------------
    Private Function GetAfterID(name As String, PreID As String) As String
        For Each Element As ModelElementsDataStructur In SystemStructurList
            If Element.Name = name And Element.Type = PreID Then
                Return Element.ID
            End If
        Next
        Return ""
    End Function
    Private Function GetPreID(name As String, AfterID As String) As String
        'PrepreID ist noch mal das Element vor
        For Each Element As ModelElementsDataStructur In SystemStructurList
            If Element.Name = name And Element.ID = AfterID Then
                Return Element.Type
            End If
        Next
        Return ""
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: CreateFunctionstructurSpecIFBase()
    '       Zweck        : Erstellt die Json Basisstruktur für Objekt und Hierarchie
    '       @param       : -
    '       Quellort     : SystemstructurAnalyse()
    '    ---------------------------------------------------------------------------------------
    Private Sub CreateSystemstructurSpecIFBase()
        Dim svgName As String = ConvUmlaut(RemoveWhitespace(System.IO.Path.GetFileNameWithoutExtension(vsoDocument.FullName)) & ".svg")
        Modelelements.Add(New ModelElementsDataStructur(vsoPage.Name, "Pln-Stueckliste", modelType))
        Dim id As String = GetIDofMEDS(vsoPage.Name, modelType)
        ProjectData.ObjekteDS = ProjectData.ObjekteDS & "{" & vbCrLf &
        "       ""id"": """ & id & """," & vbCrLf &
        "        ""title"": ""Systemstruktur""," & vbCrLf &
        "        ""attributes"": [{" & vbCrLf &
        "            ""title"": ""dcterms:title""," & vbCrLf &
        "            ""attributeType"": ""AT-Pln-Name""," & vbCrLf &
        "            ""value"": ""Systemstruktur""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""dcterms:description""," & vbCrLf &
        "            ""attributeType"": ""AT-Pln-Text""," & vbCrLf &
        "            ""value"": ""<div><p>" & modelDescription & "</p><p class=\""inline-label\"">Model View: \n</p>\n<div class=\""forImage\"" style=\""max-width: 900px;\"" >\n\t<div class=\""forImage\""><a href=\""/reqif/v0.92/projects/" & ProjectData.projectinfoID & "/files/" & svgName & "\"" type=\""image/svg+xml\"" ><object data=\""/reqif/v0.92/projects/" & ProjectData.projectinfoID & "/files/" & svgName & "\"" type=\""image/svg+xml\"" >files_and_images\\" & svgName & "</object></a></div></div></div>""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""SpecIF:Type""," & vbCrLf &
        "            ""attributeType"": ""AT-Pln-Type""," & vbCrLf &
        "            ""value"": ""Systemmodell""" & vbCrLf &
        "        }]," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""objectType"": ""OT-Pln""" & vbCrLf &
        "    },"

        ProjectData.HierarchieDS = "{" & vbCrLf &
        "    ""id"": ""SH-Fld-Stueckliste""," & vbCrLf &
        "    ""object"": """ & id & """," & vbCrLf &
        "    ""revision"": 0," & vbCrLf &
        "    ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "    ""nodes"": [##Stueckliste##]" & vbCrLf &
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
