Imports Visio = Microsoft.Office.Interop.Visio
'    ---------------------------------------------------------------------------------------
'       Modul   : ActivityDiagram
'       Datum   : 29.04.2017
'       Author  : Philipp Mochine
'       Zweck   : Das Aktivitätsdiagramm wird eingelesen und in das Integrationsmodell hinzugefügt.
'                 Beziehungen zwischen Aktivitätsdiagramm und Zustandsdiagramm werden hier automatisch gesetzt.
'       Quellort: Von ExecuteModel(modelPath As String)
'    ---------------------------------------------------------------------------------------
Module ActivityDiagram
    Private vsoDocument As Visio.Document
    Private vsoPage As Visio.Page
    Private Const modelType As String = "Activity"
    Private Const initType As String = "Initial node"
    Private Const finalType As String = "Final node"
    Private Const eventType As String = "Fork node"
    Private Const actType As String = "Action"
    Private Const diagramFrame As String = "Diagram Overview"

    Private HierarchieList As New List(Of ModelElementsDataStructur)

    Public Const actShortStr As String = "act "
    Private modelDescription As String = ConvertStringToJson("Aktivitätsmodelle sind Teil der Verhaltensmodelle der UML bzw. SysML. Sie werden eingesetzt, um alle Aktivitäten innerhalb eines Systems zu beschreiben, aber auch für den Ablauf. Sie gehören zu den Verhaltensmodellen. ")
    Private frameDescription As String = ConvertStringToJson("ist der Diagrammtitel des Aktivitätsmodells. <p>Die Syntax eines UML-Diagrammtitels lautet: " &
                                                "<b>Diagrammtyp </b>Diagrammname [Parameter]</p><p>UML bzw. SysML verwenden hierbei das Aktivitätsdiagramm.</p>")
    Private Const elementDescription As String = "ist ein Element des Aktivitätsmodells. Es führt Aktionen aus."
    '    ---------------------------------------------------------------------------------------
    '       Function     : ActivityAnalyse(vsoDocumenttemp As Visio.Document)As Boolean
    '       Zweck        : Liest das Visiodokument ein, erstellt ein Bild und die JSON Struktur. 
    '       @param       : Pointer auf das Visio Dokument.
    '       Quellort     : Von ExecuteModel(modelPath As String)
    '    ---------------------------------------------------------------------------------------
    Public Function ActivityAnalyse(vsoDocumenttemp As Visio.Document) As Boolean
        ActivityAnalyse = True
        vsoDocument = vsoDocumenttemp
        vsoPage = vsoDocument.Pages(1)
        Try
            SvgExport(vsoDocument)
            CreateActivitySpecIFBase()
            CreateActivitySpecIFCode()
            SetHierarchie()
        Catch ex As Exception
            ActivityAnalyse = False
            SendError(ex.Message, -1)
        Finally
            ReleaseObject(vsoPage)
            HierarchieList = New List(Of ModelElementsDataStructur)
        End Try
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: CreateActivitySpecIFCode()
    '       Zweck        : Hauptcode des Moduls.
    '                       1. Suche nach allen "Diagram Overview". Da dieser Akteur schon im
    '                           Zustandsdiagramm enthalten sein kann, muss es überprüft werden.
    '                       2. Im Diagram Overview wird nach dem Anfangsknoten gesucht. Der Anfangsknoten
    '                           ist ein Zustand im Zustandsdiagramm. 
    '                       3. jetzt wird durch jedes Element gegangen
    '       @param       : Pointer auf das Visio Dokument.
    '       Quellort     : Von ExecuteModel(modelPath As String)
    '    ---------------------------------------------------------------------------------------
    Private Sub CreateActivitySpecIFCode()
        Dim elementID As String
        Dim elementName As String
        Dim ShapeOverview As Visio.Shape = Nothing
        Dim vsoPageID As String = GetIDofMEDS(vsoPage.Name, modelType), diagramOverviewID As String = ""
        Dim InitList As Collection 'Eine Liste für Anfangsknoten
        Dim FinalList As Collection 'Eine Liste für Endknoten
        Dim ElementList As Collection 'Eine Liste für Ereignisse
        Dim overviewExists As Boolean = False 'Mindestens ein Rahmen muss existieren.

        For Each ShapeOverview In vsoPage.Shapes 'Es werden durch alle Modellelemente durchgegangen
            If ShapeOverview.Master.NameU.Contains(diagramFrame) Then 'Bis das Diagram Overview gefunden wurde
                overviewExists = True
                elementName = GetBaseName(ShapeOverview.Text) 'Löscht Diagrammkürzel
                elementID = GetIDofMEDS(elementName, "State")
                If elementID = "" Then 'Überprüft ob es schon das Modellelement existiert.
                    elementID = "MEl-" & GetGenID(27)
                    diagramOverviewID = elementID
                    AddObjectActor(diagramOverviewID, elementName, frameDescription) 'Fügt Objekt hinzu
                End If
                diagramOverviewID = elementID
                Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                AddRelation_vsoPage_Overview_Activity(vsoPageID, "", elementID)
                HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Activity"))
                HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Akteur"))

                'Verbindung zwischen Zustand und Acitivity
                'Die nimmt hier die ID vom Zustand wenn es noch nicht exisiert!!
                Dim vsoReturnedSelection As Visio.Selection
                Dim predecessorShape As Visio.Shape, currentShape As Visio.Shape
                Dim shpIDs As Array
                Dim i As Integer
                Dim initialNodeExists As Boolean = False, finalNodeExists As Boolean = False
                InitList = New Collection 'Die Liste wird in jedem "Rahmen" wieder neu erstellt.
                FinalList = New Collection
                ElementList = New Collection 'Eine Liste für Ereignisse. Die "echten" Namen werden gefüllt."

                'Alle Modellelemente innerhalb des Diagramm Overviews werden angesprochen.
                vsoReturnedSelection = ShapeOverview.SpatialNeighbors(Visio.VisSpatialRelationCodes.visSpatialContain, 0, Visio.VisSpatialRelationFlags.visSpatialIncludeContainerShapes)
                If vsoReturnedSelection.Count = 0 Then
                    Throw New System.Exception("Keine Elemente im Diagrammrahmen gefunden.")
                Else
                    For Each shape As Visio.Shape In vsoReturnedSelection 'Es wird durch jedes Element durchgegangen bis Anfangsknoten gefunden wurde.
                        'Suche nach Anfangsknoten
                        If shape.Master.NameU.Contains(initType) Then 'Startelement ist der Startknoten des Diagrammes
                            initialNodeExists = True 'Es existiert mindestens ein Anfangsknoten.
                            elementName = shape.Text
                            CheckElement(shape, ShapeOverview.Text, InitList) 'Überprüft den Namen und ob der Anfangsknoten schon existiert.
                            elementID = GetIDofMEDS(elementName, "State")
                            'Existiert Element in Zustandsdiagramm?
                            If elementID = "" Then
                                'Existiert Element schon?
                                elementID = GetIDofMEDS(elementName, modelType)
                                If elementID = "" Then
                                    elementID = "MEl-" & GetGenID(27)
                                    AddObjectState(elementID, PrepareStr(elementName), "ist ein Anfangsknoten in der Aktivität " & ShapeOverview.Text & ".")
                                    HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Zustand"))
                                    Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                                    AddRelation_vsoPage_Overview_Activity(vsoPageID, diagramOverviewID, elementID)
                                Else
                                    'Der Unterschied zwischen CheckInit und hier ist, dass Anfangsknoten auch in anderen Diagrammrahmen mit dem 
                                    'selben Namen auftauchen.
                                    'Für den Fall, dass die Relation zwischen Overview und Element nicht existiert
                                    If Not RelationExist(diagramOverviewID, elementID) Then
                                        AddRelation_vsoPage_Overview_Activity("", diagramOverviewID, elementID)
                                    End If
                                End If
                            Else
                                If GetIDofMEDS(elementName, modelType) = "" Then 'Existiert noch kein Eintrag?
                                    HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Zustand"))
                                    Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                                    AddRelation_vsoPage_Overview_Activity(vsoPageID, diagramOverviewID, elementID)
                                Else
                                    'Für den Fall, dass die Relation zwischen Overview und Element nicht existiert
                                    If Not RelationExist(diagramOverviewID, elementID) Then
                                        AddRelation_vsoPage_Overview_Activity("", diagramOverviewID, elementID)
                                    End If
                                End If
                            End If
                            currentShape = shape
                            predecessorShape = shape
                            'Nächster Schritt: Folge den Transitionen bis Finalknoten
                            Do While currentShape.Master.NameU <> finalType
                                shpIDs = currentShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "") 'Alle ShapesID verbunden mit CurrentShape
                                For i = 0 To UBound(shpIDs)
                                    If vsoPage.Shapes.ItemFromID(shpIDs(i)).Connects.Count > 1 Then
                                        currentShape = vsoPage.Shapes.ItemFromID(shpIDs(i)).Connects.Item(2).ToSheet
                                        'Falls der nachfolgende Shape ein Ereignis  
                                        If currentShape.Master.NameU.Contains(eventType) Then 'Falls Element ein ForkNode ist.
                                            elementName = currentShape.Text
                                            CheckElement(currentShape, ShapeOverview.Text, ElementList)
                                            elementID = GetIDofMEDS(elementName, modelType)
                                            'Auch ein Ereignis kann häufiger auftreten
                                            If elementID = "" Then
                                                elementID = "MEl-" & GetGenID(27)
                                                AddObjectEvent(elementID, PrepareStr(elementName), "ist ein Ereignis in der Aktivität " & ShapeOverview.Text & ".")
                                                HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Ereignis"))
                                                Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                                                AddRelation(GetIDofMEDS(predecessorShape.Text, modelType), elementID, "ruft auf")
                                                AddRelation_vsoPage_Overview_Activity(vsoPageID, diagramOverviewID, elementID)
                                            Else
                                                'Es müssen nur die Relationen hinzugefügt werden.
                                                If Not RelationExist(GetIDofMEDS(predecessorShape.Text, modelType), elementID) Then
                                                    AddRelation(GetIDofMEDS(predecessorShape.Text, modelType), elementID, "ruft auf")
                                                End If
                                                'Für den Fall, dass die Relation zwischen Overview und Element nicht existiert
                                                If Not RelationExist(diagramOverviewID, elementID) Then
                                                    AddRelation_vsoPage_Overview_Activity("", diagramOverviewID, elementID)
                                                End If
                                            End If

                                            'Falls der nachfolgende Shape der Abschlussknoten ist. Auch Fehler: es muss Act -> Sta
                                        ElseIf currentShape.Master.NameU.Contains(finalType) Then
                                            finalNodeExists = True
                                            elementName = currentShape.Text
                                            CheckElement(currentShape, ShapeOverview.Text, FinalList)
                                            elementID = GetIDofMEDS(elementName, "State")
                                            If elementID = "" Then
                                                elementID = GetIDofMEDS(elementName, modelType)
                                                If elementID = "" Then
                                                    elementID = "MEl-" & GetGenID(27)
                                                    AddObjectState(elementID, PrepareStr(elementName), "ist ein Abschlussknoten in der Aktivität " & ShapeOverview.Text & ".")
                                                    HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Zustand"))
                                                    Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                                                    AddRelation(GetIDofMEDS(predecessorShape.Text, modelType), elementID, "resultiert in")
                                                    AddRelation_vsoPage_Overview_Activity(vsoPageID, diagramOverviewID, elementID)
                                                Else
                                                    'Es müssen nur die Relationen hinzugefügt werden.
                                                    If Not RelationExist(GetIDofMEDS(predecessorShape.Text, modelType), elementID) Then
                                                        AddRelation(GetIDofMEDS(predecessorShape.Text, modelType), elementID, "resultiert in")
                                                    End If
                                                    'Für den Fall, dass die Relation zwischen Overview und Element nicht existiert
                                                    If Not RelationExist(diagramOverviewID, elementID) Then
                                                        AddRelation_vsoPage_Overview_Activity("", diagramOverviewID, elementID)
                                                    End If
                                                End If
                                            Else
                                                'Falls es im Zustandsdiagramm existiert.
                                                If GetIDofMEDS(elementName, modelType) = "" Then
                                                    'Noch kein Eintrag. 
                                                    HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Zustand"))
                                                    Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                                                    AddRelation(GetIDofMEDS(predecessorShape.Text, modelType), elementID, "resultiert in")
                                                    AddRelation_vsoPage_Overview_Activity(vsoPageID, diagramOverviewID, elementID)
                                                Else
                                                    'Es müssen nur die Relationen hinzugefügt werden.
                                                    If Not RelationExist(GetIDofMEDS(predecessorShape.Text, modelType), elementID) Then
                                                        AddRelation(GetIDofMEDS(predecessorShape.Text, modelType), elementID, "resultiert in")
                                                    End If
                                                    'Für den Fall, dass die Relation zwischen Overview und Element nicht existiert
                                                    If Not RelationExist(diagramOverviewID, elementID) Then
                                                        AddRelation_vsoPage_Overview_Activity("", diagramOverviewID, elementID)
                                                    End If
                                                End If
                                            End If

                                            'Jetzt kommen die Action Shapes
                                        ElseIf currentShape.Master.NameU.Contains(actType) Then
                                            elementID = GetIDofMEDS(currentShape.Text, modelType)
                                            If elementID = "" Then
                                                elementName = currentShape.Text
                                                CheckElement(currentShape, ShapeOverview.Text, Nothing)
                                                elementID = "MEl-" & GetGenID(27)
                                                AddObjectActor(elementID, PrepareStr(elementName), elementDescription)
                                                HierarchieList.Add(New ModelElementsDataStructur(elementName, elementID, "Akteur"))
                                                Modelelements.Add(New ModelElementsDataStructur(elementName, elementID, modelType))
                                                If predecessorShape.Master.NameU.Contains(eventType) Then
                                                    AddRelation(GetIDofMEDS(predecessorShape.Text, modelType), elementID, "löst aus")
                                                Else
                                                    AddRelation(GetIDofMEDS(predecessorShape.Text, modelType), elementID, "führt zu")
                                                End If
                                                AddRelation_vsoPage_Overview_Activity(vsoPageID, diagramOverviewID, elementID)
                                            Else
                                                'ElementShape existiert schon. Relation muss noch hinzugefügt werden.
                                                'Falls Relation nicht schon existiert. Braucht man es natürlich nicht mehr.
                                                If Not RelationExist(GetIDofMEDS(predecessorShape.Text, modelType), elementID) Then
                                                    If predecessorShape.Master.NameU.Contains(eventType) Then
                                                        AddRelation(GetIDofMEDS(predecessorShape.Text, modelType), elementID, "löst aus")
                                                    Else
                                                        AddRelation(GetIDofMEDS(predecessorShape.Text, modelType), elementID, "führt zu")
                                                    End If
                                                End If
                                                If Not RelationExist(diagramOverviewID, elementID) Then
                                                    'Element existiert schon, jedoch nicht die Verbindung auf dem Diagramm Overview.
                                                    AddRelation_vsoPage_Overview_Activity("", diagramOverviewID, elementID)
                                                Else
                                                    'Element existiert schon und auf dem Overview. Vorzeitiges beenden
                                                    Exit Do
                                                End If
                                            End If
                                        Else
                                            'Error: Shape ist nicht gültig. Wird nie vorkommen.
                                            Throw New System.Exception("Ein Element in der Aktivität " & ShapeOverview.Text & " ist nicht gültig.")
                                        End If
                                    Else
                                        'Error: Verbindung wurde nicht an Shape richtig gesetzt oder es fehlt eine.
                                        Throw New System.Exception("Verbindung wurde nicht an Shape richtig gesetzt oder es fehlt eine.")
                                        Exit Do
                                    End If
                                    predecessorShape = currentShape
                                Next
                            Loop
                        End If
                    Next
                    If initialNodeExists = False Then Throw New System.Exception("Es wurde in der Aktivität " & ShapeOverview.Text & " kein Anfangsknoten gefunden.")
                    If finalNodeExists = False Then Throw New System.Exception("Es wurde in der Aktivität " & ShapeOverview.Text & " kein Endknoten gefunden.")
                End If

            End If
        Next
        If overviewExists = False Then Throw New System.Exception("Es wurde kein ""Diagram Overview"" gefunden. Bitte überprüfen.")
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Function     : GetBaseName(name As String) As String
    '       Zweck        : Entfernt "act" im Titel des Rahmens
    '       @param       : Titelname
    '       @return      : Titel ohne Diagrammkürzel.
    '    ---------------------------------------------------------------------------------------
    Private Function GetBaseName(name As String) As String
        Dim pos As Integer
        If IsNullOrBlank(name) Then Throw New System.Exception("Diagrammrahmen muss einen Titel haben.")
        'Entfernt zuviele Leerzeichen (falls vorhanden)
        name = String.Join(" ", name.Split(New Char() {}, StringSplitOptions.RemoveEmptyEntries))
        'Kurz Zeichen sind maximal 4 Zeichen lang
        pos = InStr(name, actShortStr)
        If pos = 1 Then
            Return Mid(name, actShortStr.Length + 1)
        End If
        Return name
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : RelationExist(overviewID As String, elementID As String) As Boolean
    '       Zweck        : Überprüft ob schon Relation existiert.
    '       @param       : ID von Diagramm Overview und das ausgewählte Element
    '       @return      : True, wenn vorhanden.
    '    ---------------------------------------------------------------------------------------
    Private Function RelationExist(overviewID As String, elementID As String) As Boolean
        If InStr(ProjectData.RelationenDS, "RVis-" & overviewID & "-" & elementID) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : PrepareStr(s As String) As String
    '       Zweck        : Passt String an.
    '       @param       : Stringinhalt
    '       @return      : Geänderter String
    '    ---------------------------------------------------------------------------------------
    Private Function PrepareStr(s As String) As String
        PrepareStr = Replace(s, ChrW(8232), " ")
        PrepareStr = Replace(PrepareStr, ChrW(10), "")
        PrepareStr = Replace(PrepareStr, ChrW(32), " ")
        PrepareStr = Replace(PrepareStr, """", "\""")
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub          : CheckElement(name As String, list As Collection)
    '       Zweck        : Überprüft, ob Shape.Text leer ist oder schon in der Liste vorhanden ist. 
    '                      Und, dass nur ausgehende Verbindungen existieren.
    '       @param       : text und Liste der Anfangsknoten und Name des Rahmens.
    '       @return      : Übermittelt eine Liste für Anfangsknoten mit dem Syntax: Anfangsknoten- und Ereignisnamen.
    '       Quellort     : CreateActivitySpecIFCode
    '    ---------------------------------------------------------------------------------------
    Private Sub CheckElement(shape As Visio.Shape, name As String, ByRef list As Collection)
        Dim i As Integer
        Dim shapeIDs As Array
        Dim tempsShape As Visio.Shape = Nothing
        'Fehlermöglichkeiten nur für Anfangsknoten.
        If shape.Master.NameU.IndexOf(initType) >= 0 Then
            If IsNullOrBlank(shape.Text) Then Throw New System.Exception("Anfangsknoten in der Aktivität " & name & " hat keinen Zustandstext.")

            shapeIDs = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming1D, "")
            If UBound(shapeIDs) >= 0 Then Throw New System.Exception("Ein Anfangsknoten mit dem Zustandsnamen " & shape.Text & " hat keine eingehenden Verbindungen in der Aktivität " & name & ". Bitte korrigieren")
            shapeIDs = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "")
            If UBound(shapeIDs) >= 1 Then Throw New System.Exception("Ein Anfangsknoten mit dem Zustandsnamen " & shape.Text & " hat nur eine ausgehende Verbindung in der Aktivität " & name & ". Bitte korrigieren.")
            If UBound(shapeIDs) = -1 Then Throw New System.Exception("Ein Anfangsknoten mit dem Zustandsnamen " & shape.Text & " hat keine Verbindung in der Aktivität " & name)

            tempsShape = vsoPage.Shapes.ItemFromID(shapeIDs(0))
            If vsoPage.Shapes.ItemFromID(shapeIDs(0)).Connects.Count >= 2 Then 'Wenn verbunden gibt es immer mindestens zwei Verbindungspunkte
                tempsShape = vsoPage.Shapes.ItemFromID(shapeIDs(0)).Connects.Item(2).ToSheet
                If Not tempsShape.Master.NameU.Contains(eventType) Then
                    Throw New System.Exception("Nach Anfangsknoten folgt immer ein Ereignis mit dem Symbol ""Fork node"", bitte korrigieren. Fehler im Anfangsknoten mit dem Zustandsnamen " & shape.Text & " in der Aktivität " & name)
                End If
            Else
                Throw New System.Exception("Ein Anfangsknoten mit dem Zustandsnamen " & shape.Text & " hat keine Verbindung in der Aktivität " & name)
            End If
            For i = 1 To list.Count
                If (shape.Text & tempsShape.Text) = list(i) Then
                    Throw New System.Exception("Anfangsknoten mit dem Zustandsnamen " & shape.Text & " existiert doppelt in der Aktivität " & name)
                End If
            Next
            list.Add(shape.Text & tempsShape.Text)
            'Fehlermöglichkeiten für Endknoten.
        ElseIf shape.Master.NameU.IndexOf(finalType) >= 0 Then
            'Ereignisknoten
            If IsNullOrBlank(shape.Text) Then Throw New System.Exception("Endknoten in der Aktivität " & name & " hat keinen Zustandstext.")
            If list.Count > 1 Then Throw New System.Exception("Es darf nur ein Endknoten in der Aktivität " & name & " existieren")
            shapeIDs = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "")
            If UBound(shapeIDs) >= 1 Then Throw New System.Exception("Anfangsknoten haben nur eine ausgehende Verbindung. Fehler beim Anfangsknoten mit dem Zustandsnamen " & shape.Text & " in der Aktivität " & name & ". Bitte korrigieren.")
            If UBound(shapeIDs) >= 0 Then Throw New System.Exception("Ein Endknoten darf keine ausgehende Verbindung haben. Fehler: Endknoten mit dem Zustandsnamen " & shape.Text & " in der Aktivität " & name & ". Bitte korrigieren.")
            list.Add(shape.Text)
            'Ereignisse
        ElseIf shape.Master.NameU.IndexOf(eventType) >= 0 Then
            If IsNullOrBlank(shape.Text) Then Throw New System.Exception("Ereigniselement in der Aktivität " & name & " hat keinen Text.")
            'Welche Informationen haben wir. Wir wollen wissen, gibt es das Ereignis doppelt im Overview?
            'Fehler kommt wenn unterschiedliche Namen mit selben Text!
            Dim shapeNameExists As Boolean = False
            For i = 1 To list.Count
                tempsShape = list(i)
                If shape.Text = tempsShape.Text Then
                    If shape.Name <> tempsShape.Name Then
                        Throw New System.Exception("Ereigniselement mit dem Namen " & shape.Text & " existiert doppelt in der Aktivität " & name)
                    End If
                End If
                If shape.Name = tempsShape.Name Then
                    shapeNameExists = True
                End If
            Next
            If shapeNameExists = False Then list.Add(shape) 'Shape wird gespeichert, wenn es noch nicht in der Liste ist.

            shapeIDs = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "")
            If UBound(shapeIDs) >= 1 Then Throw New System.Exception("Ereignisknoten haben nur eine ausgehende Verbindung. Fehler beim Element mit dem Zustandsnamen " & shape.Text & " in der Aktivität " & name & ". Bitte korrigieren.")
            If UBound(shapeIDs) = -1 Then
                Throw New System.Exception("Fehler im Ereignisknoten mit dem Namen " & shape.Text & " in der Aktivität " & name & ". Die Verbindung ist nicht an ein Element verknüpft")
            End If
            If UBound(shapeIDs) = 0 Then
                If vsoPage.Shapes.ItemFromID(shapeIDs(0)).Connects.Count < 2 Then
                    Throw New System.Exception("Verbindung steht leer im Raum. Fehler bei Ereignis mit dem Text " & shape.Text & " in der Aktivität " & name & ".")
                End If
                If Not vsoPage.Shapes.ItemFromID(shapeIDs(0)).Connects.Item(2).ToSheet.Master.NameU.Contains(actType) Then
                    Throw New System.Exception("Nach Ereigbnis folgt immer ein Modellelement der Aktion, bitte korrigieren. Fehler im Ereignisknoten mit dem Namen " & shape.Text & " in der Aktivität " & name)
                End If
            End If
        Else
            'Restliche Elemente
            If IsNullOrBlank(shape.Text) Then Throw New System.Exception("Modellelement " & shape.Name & " in der Aktivität " & name & " hat keinen Text.")
            shapeIDs = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "")
            If UBound(shapeIDs) >= 1 Then Throw New System.Exception("Modellemente wie Aktionen haben nur eine Verbindung. Fehler Element mit dem Namen " & shape.Text & " in der Aktivität " & name & ". Bitte korrigieren.")
            If UBound(shapeIDs) = 0 Then
                If vsoPage.Shapes.ItemFromID(shapeIDs(0)).Connects.Count < 2 Then Throw New System.Exception("Verbindung steht leer im Raum. Fehler bei Modellelement mit dem Text " & shape.Text & " in der Aktivität " & name & ".")
            End If
            If UBound(shapeIDs) = -1 Then
                Throw New System.Exception("Modellelement mit dem Text " & shape.Text & " in der Aktivität " & name & " hat keine ausgehende Verbindung zu einem Element.")
            End If
        End If
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: AddRelation(preID As String, currentID As String, relType As String)
    '       Zweck        : Fügt Beziehung abhängig zu Relationtyp hinzu 
    '       @param       : -
    '       Quellort     : -
    '    ---------------------------------------------------------------------------------------
    Private Sub AddRelation(preID As String, currentID As String, relType As String)
        Select Case relType
            Case "ruft auf"
                ProjectData.RelationenDS = ProjectData.RelationenDS & "{" & vbCrLf &
                "        ""id"": ""RVis-" & preID & "-" & currentID & """," & vbCrLf &
                "        ""title"": ""ruft auf""," & vbCrLf &
                "        ""revision"": 0," & vbCrLf &
                "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
                "        ""relationType"": ""RT-ruftauf""," & vbCrLf &
                "        ""source"": {" & vbCrLf &
                "            ""id"": """ & preID & """," & vbCrLf &
                "            ""revision"": 0" & vbCrLf &
                "        }," & vbCrLf &
                "        ""target"": {" & vbCrLf &
                "            ""id"": """ & currentID & """," & vbCrLf &
                "            ""revision"": 0" & vbCrLf &
                "        }" & vbCrLf &
                "    },"
            Case "resultiert in"
                ProjectData.RelationenDS = ProjectData.RelationenDS & "{" & vbCrLf &
                "        ""id"": ""RVis-" & preID & "-" & currentID & """," & vbCrLf &
                "        ""title"": ""resultiert in""," & vbCrLf &
                "        ""revision"": 0," & vbCrLf &
                "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
                "        ""relationType"": ""RT-resultiertIn""," & vbCrLf &
                "        ""source"": {" & vbCrLf &
                "            ""id"": """ & preID & """," & vbCrLf &
                "            ""revision"": 0" & vbCrLf &
                "        }," & vbCrLf &
                "        ""target"": {" & vbCrLf &
                "            ""id"": """ & currentID & """," & vbCrLf &
                "            ""revision"": 0" & vbCrLf &
                "        }" & vbCrLf &
                "    },"
            Case "führt zu"
                ProjectData.RelationenDS = ProjectData.RelationenDS & "{" & vbCrLf &
                "        ""id"": ""RVis-" & preID & "-" & currentID & """," & vbCrLf &
                "        ""title"": ""führt zu""," & vbCrLf &
                "        ""revision"": 0," & vbCrLf &
                "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
                "        ""relationType"": ""RT-FuehrtZu""," & vbCrLf &
                "        ""source"": {" & vbCrLf &
                "            ""id"": """ & preID & """," & vbCrLf &
                "            ""revision"": 0" & vbCrLf &
                "        }," & vbCrLf &
                "        ""target"": {" & vbCrLf &
                "            ""id"": """ & currentID & """," & vbCrLf &
                "            ""revision"": 0" & vbCrLf &
                "        }" & vbCrLf &
                "    },"
            Case "löst aus"
                ProjectData.RelationenDS = ProjectData.RelationenDS & "{" & vbCrLf &
                "        ""id"": ""RVis-" & preID & "-" & currentID & """," & vbCrLf &
                "        ""title"": ""löst aus""," & vbCrLf &
                "        ""revision"": 0," & vbCrLf &
                "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
                "        ""relationType"": ""RT-loestAus""," & vbCrLf &
                "        ""source"": {" & vbCrLf &
                "            ""id"": """ & preID & """," & vbCrLf &
                "            ""revision"": 0" & vbCrLf &
                "        }," & vbCrLf &
                "        ""target"": {" & vbCrLf &
                "            ""id"": """ & currentID & """," & vbCrLf &
                "            ""revision"": 0" & vbCrLf &
                "        }" & vbCrLf &
                "    },"
        End Select

    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: AddObjectActor(id As String, name As String, text As String)
    '       Zweck        : Fügt Beziehung abhängig zu Relationtyp hinzu 
    '       @param       : -
    '       Quellort     : CreateActivitySpecIFCode()
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
    '       Sub-Anweisung: AddObjectEvent(id As String, name As String, text As String)
    '       Zweck        : Fügt Events hinzu 
    '       @param       : -
    '       Quellort     : CreateActivitySpecIFCode()
    '    ---------------------------------------------------------------------------------------
    Private Sub AddObjectEvent(id As String, name As String, text As String)
        If IsNullOrBlank(name) Then Throw New System.Exception("Es fehlt der Text bzw. die Beschreibung in einem Modellelement. Bitte nicht leer lassen.")
        ProjectData.ObjekteDS = ProjectData.ObjekteDS & "{" & vbCrLf &
        "        ""id"": """ & id & """," & vbCrLf &
        "        ""title"": """ & name & """," & vbCrLf &
        "        ""attributes"": [{" & vbCrLf &
        "            ""title"": ""dcterms:title""," & vbCrLf &
        "            ""attributeType"": ""AT-Evt-Name""," & vbCrLf &
        "            ""value"": """ & name & """" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""dcterms:description""," & vbCrLf &
        "            ""attributeType"": ""AT-Evt-Text""," & vbCrLf &
        "            ""value"": ""<div><p>" & text & "</p></div>""" & vbCrLf &
        "        }]," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""objectType"": ""OT-Evt""" & vbCrLf &
        "    },"

    End Sub

    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: SetHierarchie()
    '       Zweck        : Setzt Hierarchie und sortiert nach Alphabet.
    '       @param       : -
    '       Quellort     : Von RequirementsAnalyse(xlWorkBooktemp As Excel.Workbook)
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
            If Element.Type = "Activity" Then
                ProjectData.HierarchieDS = Replace(ProjectData.HierarchieDS, "##" & Element.Type & "##", hrarStr)
            Else
                If Not MELHierarchieElementExist(idHr) Then ProjectData.MELHierarchieDS = Replace(ProjectData.MELHierarchieDS, "##" & Element.Type & "##", hrarStr)
            End If
        Next
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: AddRelation_vsoPage_Overview_Activity(vsoPageID As String, OverviewID As String, targetID As String)
    '       Zweck        : Ziel: Relationen hinzufügen zum gesamten Diagramm und den einzelnen Aktivitätsdiagrammen
    '       @param       : -
    '       Quellort     : Von RequirementsAnalyse(xlWorkBooktemp As Excel.Workbook)
    '    ---------------------------------------------------------------------------------------
    Private Sub AddRelation_vsoPage_Overview_Activity(vsoPageID As String, OverviewID As String, targetID As String)
        If vsoPageID <> "" Then
            ProjectData.RelationenDS = ProjectData.RelationenDS & "{" & vbCrLf &
            "        ""id"": ""RVis-" & vsoPageID & "-" & targetID & """," & vbCrLf &
            "        ""title"": ""zeigt""," & vbCrLf &
            "        ""revision"": 0," & vbCrLf &
            "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
            "        ""relationType"": ""RT-Visibility""," & vbCrLf &
            "        ""source"": {" & vbCrLf &
            "            ""id"": """ & vsoPageID & """," & vbCrLf &
            "            ""revision"": 0" & vbCrLf &
            "        }," & vbCrLf &
            "        ""target"": {" & vbCrLf &
            "            ""id"": """ & targetID & """," & vbCrLf &
            "            ""revision"": 0" & vbCrLf &
            "        }" & vbCrLf &
            "    },"
        End If
        If OverviewID <> "" Then
            ProjectData.RelationenDS = ProjectData.RelationenDS & "{" & vbCrLf &
            "        ""id"": ""RVis-" & OverviewID & "-" & RelationExistOverall(OverviewID, targetID) & """," & vbCrLf &
            "        ""title"": ""enthält""," & vbCrLf &
            "        ""revision"": 0," & vbCrLf &
            "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
            "        ""relationType"": ""RT-Containment""," & vbCrLf &
            "        ""source"": {" & vbCrLf &
            "            ""id"": """ & OverviewID & """," & vbCrLf &
            "            ""revision"": 0" & vbCrLf &
            "        }," & vbCrLf &
            "        ""target"": {" & vbCrLf &
            "            ""id"": """ & targetID & """," & vbCrLf &
            "            ""revision"": 0" & vbCrLf &
            "        }" & vbCrLf &
            "    },"
        End If
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Function     : RelationExistOverall(overviewID As String, elementID As String) As String
    '       Zweck        : Überprüft ob schon Relation existiert und gibt ID wieder
    '       @param       : ID von Diagramm Overview und das ausgewählte Element
    '       @return      : Gibt ID wieder
    '    ---------------------------------------------------------------------------------------
    Private Function RelationExistOverall(overviewID As String, elementID As String) As String
        'Brauche ich, da es unterschiedliche Relationen gibt zur gleichen ID
        'Beispiel-Relationen: act warten enthält Wartung im Activitydiagramm.  Aber warten resultiert in Wartung im Zustandsdiagramm.
        'Somit wird eine neue ID benötigt.
        If InStr(ProjectData.CompleteDS, "RVis-" & overviewID & "-" & elementID) > 0 Then
            Return elementID & GetGenID(3)
        Else
            Return elementID
        End If
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: CreateActivitySpecIFBase()
    '       Zweck        : Erstellt die JSON Basis-Struktur mit Objekten und Hierarchie
    '       @param       : -
    '       Quellort     : Von RequirementsAnalyse(xlWorkBooktemp As Excel.Workbook)
    '    ---------------------------------------------------------------------------------------
    Private Sub CreateActivitySpecIFBase()
        Dim svgName As String = ConvUmlaut(RemoveWhitespace(System.IO.Path.GetFileNameWithoutExtension(vsoDocument.FullName)) & ".svg")
        Modelelements.Add(New ModelElementsDataStructur(vsoPage.Name, "Pln-Aktmodell", modelType))
        Dim id As String = GetIDofMEDS(vsoPage.Name, modelType)
        ProjectData.ObjekteDS = ProjectData.ObjekteDS & "{" & vbCrLf &
        "       ""id"": """ & id & """," & vbCrLf &
        "        ""title"": ""Aktivitätsdiagramm""," & vbCrLf &
        "        ""attributes"": [{" & vbCrLf &
        "            ""title"": ""dcterms:title""," & vbCrLf &
        "            ""attributeType"": ""AT-Pln-Name""," & vbCrLf &
        "            ""value"": ""Aktivitätsdiagramm""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""dcterms:description""," & vbCrLf &
        "            ""attributeType"": ""AT-Pln-Text""," & vbCrLf &
        "            ""value"": ""<div><p>" & modelDescription & "</p><p class=\""inline-label\"">Model View: \n</p>\n<div class=\""forImage\"" style=\""max-width: 900px;\"" >\n\t<div class=\""forImage\""><a href=\""/reqif/v0.92/projects/" & ProjectData.projectinfoID & "/files/" & svgName & "\"" type=\""image/svg+xml\"" ><object data=\""/reqif/v0.92/projects/" & ProjectData.projectinfoID & "/files/" & svgName & "\"" type=\""image/svg+xml\"" >files_and_images\\" & svgName & "</object></a></div></div></div>""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""SpecIF:Type""," & vbCrLf &
        "            ""attributeType"": ""AT-Pln-Type""," & vbCrLf &
        "            ""value"": ""UML Aktivitätsmodell""" & vbCrLf &
        "        }]," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""objectType"": ""OT-Pln""" & vbCrLf &
        "    },{" & vbCrLf &
        "        ""id"": ""Fld-Aktmodell-Akt""," & vbCrLf &
        "        ""title"": ""Aktivitäten""," & vbCrLf &
        "        ""attributes"": [{" & vbCrLf &
        "            ""title"": ""dcterms:title""," & vbCrLf &
        "            ""attributeType"": ""AT-Fld-Name""," & vbCrLf &
        "            ""value"": ""Aktivitäten""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""dcterms:description""," & vbCrLf &
        "            ""attributeType"": ""AT-Fld-Text""," & vbCrLf &
        "            ""value"": ""<div><p>Im Folgenden befinden sich alle Aktivitäten: </p></div>""" & vbCrLf &
        "        }]," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""objectType"": ""OT-Fld""" & vbCrLf &
        "    },"

        ProjectData.HierarchieDS = ProjectData.HierarchieDS & "{" & vbCrLf &
        "        ""id"": ""SH-" & id & """," & vbCrLf &
        "        ""object"": """ & id & """," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""nodes"": [{" & vbCrLf &
        "           ""id"": ""SH-Fld-Aktmodell-Akt""," & vbCrLf &
        "           ""object"": ""Fld-Aktmodell-Akt""," & vbCrLf &
        "            ""revision"": 0," & vbCrLf &
        "           ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "           ""nodes"": [##Activity##]" & vbCrLf &
        "        }]" & vbCrLf &
        "},"
    End Sub
End Module
