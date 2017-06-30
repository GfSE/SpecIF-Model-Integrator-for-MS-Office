Imports Visio = Microsoft.Office.Interop.Visio
Imports Excel = Microsoft.Office.Interop.Excel
'    ---------------------------------------------------------------------------------------
'       Modul   : MainSpecIF
'       Datum   : 29.04.2017
'       Author  : Philipp Mochine
'       Zweck   : Hauptmodul des Projektes. Ziel ist mehrere Teilmodelle zu einem Integrationsmodell zu integrieren.
'       Quellort: Modul Interface.vb mit der Sub-Anweisung MainSpecIFSub.  
'    ---------------------------------------------------------------------------------------
Module MainSpecIF
    'Zählt alle erzeugte Elemente und gibt diese als Konsole aus.
    'Wichtig für Messung der kritischen Elemente. Wieviel kann der Browser aufnehmen.
    'Ein Element wird in den jeweiligen Modellmodulen erzeugt. 
    'Für eine einfache Berechnung. Wird dieser Wert beim Generieren einer ID erhöht.
    'Erhört von Public Function GetGenID(cb As Integer) As String in Functions.vb
    Public countElements As Integer
    'ProjectData enthält sämtliche Informationen des Projektes, wie ID, Name usw.
    'ProjectData ist Public und somit einlesbar für alle Module.
    'Darüber hinaus enthält die Klasse temporäre Speichermöglichkeiten für die Objekte, Relationen, Hierarchie und das Integrationsmodell. 
    Public ProjectData As New Datastructur
    'Bei der Abstraktion der Modellelemente werden diese in der Gesamtliste "Modelelements" gespeichert.
    'Damit ist es möglich zu sehen, ob Elemente schon eingetragen worden sind oder nicht.
    'Die Konsolidierung und Vernetzung der Integration ist somit möglich.
    Public Modelelements As New List(Of ModelElementsDataStructur)
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: InitProjectData(modelPaths() As String)
    '       Zweck        : Initialisiert Metadaten des Projekts       
    '       @param       : modelPaths() ist ein Array String mit den Angaben der Pfadwege zu den Modellen.     
    '       Quellaufruf  : Von MainSpecIFSub(modelPaths() As String)
    '    ---------------------------------------------------------------------------------------
    Private Sub InitProjectData(modelPaths() As String)
        Try
            Dim modelName As String
            'Siehe Settings.vb
            ProjectData.infoEmail = pdEmail
            ProjectData.infoFamilyName = pdFamilyName
            ProjectData.infoGivenName = pdGivenName
            ProjectData.infoOrganisationName = pdOrganisationName
            'Basisname wird für Projectnamen herausgefiltert.
            '*.projectName enthält nur den Projekt Namen ohne Slash.
            If InStr(ProjectData.projectinfoName, "/") > 0 Then
                ProjectData.projectName = Mid(ProjectData.projectinfoName, InStr(ProjectData.projectinfoName, "/") + 1)
            Else
                ProjectData.projectName = ProjectData.projectinfoName
            End If
            'For-Schleife zur Bestimmung der Modellnamen ohne Datentyp Endung aus ListBox.
            For Each modelpath As String In modelPaths
                modelName = System.IO.Path.GetFileNameWithoutExtension(modelpath)
                If ProjectData.modelName(0) = "" Then
                    ProjectData.modelName(0) = modelName
                Else
                    ReDim Preserve ProjectData.modelName(UBound(ProjectData.modelName) + 1)
                    ProjectData.modelName(UBound(ProjectData.modelName)) = modelName
                End If
            Next
            countElements = 0 'Wird auf 0 Elemente gesetzt.
        Catch ex As Exception
            SendError(ex.Message, -1)
        End Try
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: MainSpecIFSub(modelPaths() As String)
    '       Zweck        : Hauptanweisung des gesamten Codes. Erstellt in einer Schleife mehrere Modelle, die zu einem Integrationsmodell
    '                      hinzugefügt wird. Im letzten Modell wird das Datenformat SpecIFz mit Zip erstellt.  
    '       @param       : modelPaths() ist ein Array String mit den Angaben der Pfadwege zu den Modellen.     
    '       Quellaufruf  : Aus Interface.vb
    '    ---------------------------------------------------------------------------------------       
    Public Sub MainSpecIFSub(modelPaths() As String)
        Dim loopCount As Integer = 0 'Wird für WriteProgressbar verwendet
        Dim modelName As String, modelPath As String

        If IsPathWritable(IO.Path.GetDirectoryName(modelPaths(0))) Then
            WriteProgressbar(30, True)
            SendInfo("Teilmodelle werden überführt!", 4)
            Try
                For Each modelPath In modelPaths 'For-Schleife. Läuft durch alle ausgewählten Modelle durch.
                    loopCount = loopCount + 1
                    modelName = System.IO.Path.GetFileNameWithoutExtension(modelPath)
                    SendInfo("Einlesen: " & modelName, 4)
                    If modelName <> "" Then
                        If ProjectData.modelName(0) = "" Then 'Beim ersten Durchlauf werden Metadaten, SpecIF-Ordner und die SpecIF-Grundstruktur gesetzt.
                            InitProjectData(modelPaths)
                            CreateFolder(modelPath)
                            CreateSpecIFBase()
                        End If
                        'Öffnet Modell und findet den Diagrammtyp heraus. 
                        'Anschließend wird das entsprechende Diagrammtyp-Modul geladen.
                        'Das Ergebnis ist eine Teilintegration mit Hilfe der 4 Prinzipien der Modellintegration (siehe specif.de)
                        WriteProgressbar(loopCount * (50 \ modelPaths.Count), True)
                        If ExecuteModel(modelPath) Then 'Modell wird hier analysiert und integriert. Nur weiter machen, kein Fehler auftritt.
                            JoinSpecIFBaseWithProjectData() 'Ermittelte Objekte/Relationen/Hierarchien werden zum gesamt Modell hinzugefügt.
                            If modelName = ProjectData.modelName(UBound(ProjectData.modelName)) Then
                                JoinMElHierarchie() 'Fügt Modellelemente in die Hierarchie ein.
                                ProjectData.CompleteDS = PrepareStringForExport(ProjectData.CompleteDS) 'Das gesamte Modell wird für den HTML-Bedarf konvertiert.
                                WriteProgressbar(90, True)
                                SendInfo("Specif wird erstellt!", 4)
                                'Die JSON-Struktur wird im erstellten Ornder als ".specif" abgespeichert.
                                WriteStringToSpecIF(ProjectData.CompleteDS, ProjectData.projectFullPath & "\" & ConvUmlaut(RemoveWhitespace(ProjectData.projectName)) & ".specif")
                                ZipFileToSpecIF() 'Das Datenformat ".specif" wird erstellt, in dem der erstellte Ordner gezippt und umbenannt wird.
                                WriteProgressbar(100, True)
                                SendInfo("Überführung abgeschlossen!", 6)
                            End If
                        Else
                            RestartProject()
                            Exit For
                        End If
                    End If
                Next
                Form1.Timer2.Start() 'Um Progressbar zu deaktivieren.
            Catch ex As Exception
                SendError(ex.Message, -1)
            Finally 'Löscht Daten
                ProjectData.ReFreshData()
                ResetCount()
                Modelelements = New List(Of ModelElementsDataStructur)
                Console.WriteLine("Anzahl Elemente: " & countElements)
            End Try
        Else
            SendError("", 3)
        End If
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Function:      ExecuteModel(modelPath As String)as Boolean
    '       Zweck        : Das gewählte Modell wird in dieser Anweisung erkannt und entsprechende Module werden geladen.
    '                      Das Ergebnis ist eine Teilintegration. 
    '       @param       : modelPath As String enthält den Pfad zum Modell.  
    '       @retunr      : Ob das Einlesen erfolgreich war.
    '       Quellaufruf  : Von MainSpecIFSub(modelPaths() As String)
    '    ---------------------------------------------------------------------------------------
    Public iXl As Integer = 0 'Zähler für Excel
    Private Function ExecuteModel(modelPath As String) As Boolean
        Dim modelExtension As String
        modelExtension = System.IO.Path.GetExtension(modelPath) 'Ermittlung des Datenformats.
        ExecuteModel = True 'Annahme, dass alles gut läuft.
        If ExcelExtension.Contains(modelExtension) Then 'Aufruf bei einer Exceldatei.
            Dim xlApp As Excel.Application = Nothing
            Dim xlWorkBook As Excel.Workbook = Nothing

            xlApp = New Excel.Application
            If xlApp Is Nothing Then
                Throw New System.Exception("Excel ist nicht installiert!")
            End If
            Try
                xlWorkBook = xlApp.Workbooks.Open(modelPath) 'Excel wird geöffnet.
            Catch ex As Exception
                ExecuteModel = False
                SendError(ex.Message, -1)
            End Try
            Try
                If iXl > 0 Then Throw New System.Exception("Nur eine Anforderung ist erlaubt.")
                ExecuteModel = RequirementsAnalyse(xlWorkBook) 'Excel wird auf Anforderung analysiert.
                iXl = iXl + 1
            Catch ex As Exception
                ExecuteModel = False
                SendError(ex.Message, -1)
            Finally 'Wird gebraucht, um den Speicher freizusetzen.
                xlWorkBook.Close() 'Idee Mach daraus eine Boolean Funktion. Wenn irgendwas nicht stimmt wird mit GO gesprungen und alles Resetet.
                xlApp.Quit()
                ReleaseObject(xlApp)
                ReleaseObject(xlWorkBook)
            End Try
        ElseIf VisioExtension.Contains(modelExtension) Then 'Aufruf bei Visiodateien.
            Dim vsoApp As Visio.Application = Nothing
            Dim vsoDocument As Visio.Document = Nothing
            Dim modelType As String

            vsoApp = New Visio.Application With {
                .Visible = False 'Visio Wird nicht dargestellt.
                }
            If vsoApp Is Nothing Then
                Throw New System.Exception("Visio ist nicht installiert!")
            End If
            Try
                vsoDocument = vsoApp.Documents.Open(modelPath) 'Visio wird geöffnet.
            Catch ex As Exception
                ExecuteModel = False
                SendError(ex.Message, -1)
            End Try
            Try
                modelType = FindModelType(vsoDocument) 'Diagrammtyp wird ermittelt. Daraufhin werden die entsprechenden Module geöffnet.
                Select Case modelType
                    Case "State"
                        ExecuteModel = StateAnalyse(vsoDocument)
                    Case "Activity"
                        ExecuteModel = ActivityAnalyse(vsoDocument)
                    Case "System"
                        ExecuteModel = SystemstructurAnalyse(vsoDocument)
                    Case "Function"
                        ExecuteModel = FunctionStructurAnalyse(vsoDocument)
                    Case "None"
                        ExecuteModel = False
                        SendError("Der Diagrammtyp konnte nicht erkannt werden. Überprüfen Sie, ob die richtigen Modellelemente verwendet wurden.", 0)
                End Select
            Catch ex As Exception
                ExecuteModel = False
                SendError(ex.Message, -1)
            Finally 'Wird gebraucht, um den Speicher freizusetzen.
                vsoDocument.Close()
                vsoApp.Quit()
                ReleaseObject(vsoApp)
                ReleaseObject(vsoDocument)
            End Try
        Else
            ExecuteModel = False
            SendError("Es ist keine gültige Excel oder Visio Datei.", 0)
        End If
    End Function
End Module
