Imports System.IO.Compression
Imports System.IO
Imports Visio = Microsoft.Office.Interop.Visio
Imports System.Runtime.InteropServices
Imports System.Media
'    ---------------------------------------------------------------------------------------
'       Modul   : Functions
'       Datum   : 29.04.2017
'       Author  : Philipp Mochine
'       Zweck   : Eine Sammlung an Funktionen.                           
'       Quellort: Offene Anweisungen/Funktionen. Offen für alle Module.
'    ---------------------------------------------------------------------------------------
Module Functions
    '    ---------------------------------------------------------------------------------------
    '       Function     : RemoveWhitespace(fullString As String) As String
    '       Zweck        : Löscht aus einem String Leerzeichen.
    '       @param       : String mit Leerzeichen
    '       @return      : String 
    '       Quelle       : http://stackoverflow.com/questions/1645546/remove-spaces-from-a-string-in-vb-net
    '    --------------------------------------------------------------------------------------- 
    Public Function RemoveWhitespace(fullString As String) As String
        Return New String(fullString.Where(Function(x) Not Char.IsWhiteSpace(x)).ToArray())
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : ConvHTML(s As String) As String
    '       Zweck        : Konvertiert Umlaute zu Html Tags
    '       @param       : String 
    '       @return      : String mit Html Tags
    '    --------------------------------------------------------------------------------------- 
    Private Function ConvHTML(s As String) As String
        If Not s Is Nothing Then
            s = Replace(s, "Ä", "&#196;")
            s = Replace(s, "ä", "&#228;")
            s = Replace(s, "Ö", "&#214;")
            s = Replace(s, "ö", "&#246;")
            s = Replace(s, "Ü", "&#220;")
            s = Replace(s, "ü", "&#252;")
            s = Replace(s, "ß", "&#223;")
        End If
        ConvHTML = s
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : ConvUmlaut(s As String) As String
    '       Zweck        : Konvertiert Umlaute. Wichtig für Pfade etc.
    '       @param       : String 
    '       @return      : String mit Html Tags
    '    --------------------------------------------------------------------------------------- 
    Public Function ConvUmlaut(s As String) As String
        If Not s Is Nothing Then
            s = Replace(s, "Ä", "Ae")
            s = Replace(s, "ä", "ae")
            s = Replace(s, "Ö", "Oe")
            s = Replace(s, "ö", "oe")
            s = Replace(s, "Ü", "Ue")
            s = Replace(s, "ü", "ue")
            s = Replace(s, "ß", "ss")
        End If
        ConvUmlaut = s
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: CreateFolder(modelPath As String)
    '       Zweck        : Erstellt Ordner. Speicherort beim ersten angegeben Modell
    '       @param       : Ordnerpfad
    '       Quellort     : Von MainSpecIFSub(modelPaths() As String)
    '    ---------------------------------------------------------------------------------------
    Public Sub CreateFolder(modelPath As String)
        If ProjectData.projectFullPath = "" Then
            Dim folderName As String
            modelPath = System.IO.Path.GetDirectoryName(modelPath)
            folderName = ConvUmlaut(RemoveWhitespace(ProjectData.projectName))
            Form1.lastSpecIFzPath = modelPath 'Setzt den Pfad des Ordners, um später zum Ordner zu klicken
            ProjectData.projectFullPath = modelPath & "\" & folderName
            If Dir(ProjectData.projectFullPath, vbDirectory) <> "" Then
                Dim objFSO As Object
                objFSO = CreateObject("Scripting.FileSystemObject")
                objFSO.DeleteFolder(ProjectData.projectFullPath) 'Löscht einen vorhandenen Ordner (falls einer erstellt wurde mit dem selben Namen
                objFSO = Nothing
            End If
            MkDir(ProjectData.projectFullPath)
            MkDir(ProjectData.projectFullPath & filesandimagesSettings)
        End If
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Function     : GetCreatedAt() As String
    '       Zweck        : Erstellt ein Datumformat. Beispiel: "2017-01-12T19:22:02+01:00" 
    '       @param       : Ordnerpfad
    '       @return      : Datumformat.
    '    ---------------------------------------------------------------------------------------
    Public Function GetCreatedAt() As String
        Return DateTime.Now.ToString(timeFilterSettings)
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: CreateFolder(modelPath As String)
    '       Zweck        : Erstellt ein UTF8WIthoutBOM File.
    '       @param       : Stringinhalt und Pfad.
    '       Quellort     : Von MainSpecIFSub(modelPaths() As String)
    '    ---------------------------------------------------------------------------------------
    Public Sub WriteStringToSpecIF(str As String, path As String)
        Dim utf8WithoutBom As New System.Text.UTF8Encoding(False)
        Using sw As IO.StreamWriter = New IO.StreamWriter(path, True, utf8WithoutBom)
            sw.Write(str)
            sw.Close()
        End Using
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Function     : MELHierarchieElementExist(elementID As String) As Boolean
    '       Zweck        : Überprüft ob ein Object in der Hierarchieliste schon eingetagen ist.
    '       @param       : Elementid
    '       @return      : True, wenn gefunden..
    '    ---------------------------------------------------------------------------------------
    Public Function MELHierarchieElementExist(elementID As String) As Boolean
        If InStr(ProjectData.MELHierarchieDS, """object"": """ & elementID) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : PrepareStringForExport(str As String) As String
    '       Zweck        : Zwei Funktionen. Den Aufruf von ConvHtml() und dann das Löschen von "##..##". 
    '       @param       : String
    '       @return      : String ohne ##
    '    ---------------------------------------------------------------------------------------
    Public Function PrepareStringForExport(str As String) As String
        'Html Entities Konventierung und Löschen von allen restlichen ##Dummies## first 2000439 200451
        Dim iPosFirst As Integer, iPosLast As Integer
        If str <> "" Then
            str = ConvHTML(str)
            iPosFirst = InStr(str, "##") 'Erste Position 
            Do While iPosFirst > 0
                iPosLast = InStr(iPosFirst + 2, str, "##") + 2 'Letzte Position
                If iPosLast - iPosFirst < 44 Then 'Differenz                   
                    If Mid(str, iPosFirst - 1, 1) = "," Then iPosFirst = iPosFirst - 1 'Falls ein Komma vor der ersten Position gefunden wird. ",##"  
                    str = Replace(str, Mid(str, iPosFirst, iPosLast - iPosFirst), "")
                End If
                iPosFirst = InStr(str, "##")
            Loop
            Return str
        Else
            SendError("Fehler in der Funktion PrepareStringForExport. String ist Leer", -1)
            Return ""
        End If
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : GetModelElementBaseName(ByVal modelElementName As String) As String
    '       Zweck        : Manche Modellelemente enthalten Zusatzpunkte, die entfernt werden müssen.
    '       @param       : String
    '       @return      : String ohne Punkt
    '    ---------------------------------------------------------------------------------------
    Public Function GetModelElementBaseName(ByVal modelElementName As String) As String
        If InStr(modelElementName, ".") > 0 Then 'Manche Elemente besitzen im Namen ".". 
            Return Left(modelElementName, InStr(modelElementName, ".") - 1)
        Else
            Return modelElementName
        End If
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: ResetCount()
    '       Zweck        : Setzt die Zähler der Diagrammtypen wieder auf Null
    '       @param       : -
    '       Quellort     : -
    '    ---------------------------------------------------------------------------------------
    Public Sub ResetCount()
        iState = 0
        iActivity = 0
        iSystem = 0
        iFunction = 0
        iXl = 0
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Function     : FindModelType(vsoDocument As Visio.Document) As String
    '       Zweck        : Ermittelt Diagrammtyp anhand von Modellelementen. Modellelementen werden in Settings.vb gespeichert.
    '       @param       : Geöffnetes Visio Dokument.
    '       @return      : String mit Digrammtyp Namen.
    '    ---------------------------------------------------------------------------------------
    Private iState As Integer = 0, iActivity As Integer = 0, iSystem As Integer = 0, iFunction As Integer = 0 'Überprüfung ob ein Diagrammtyp nicht mehr als 1x vorkommt
    Public Function FindModelType(vsoDocument As Visio.Document) As String
        Dim vsoPage As Visio.Page = vsoDocument.Pages(1)
        Dim boolstateElement As Boolean = True, boolactivityElement As Boolean = True
        Dim boolsystemstructurElement As Boolean = True, boolfunctionstructurElement As Boolean = True
        Dim elementText As String

        For Each element As Visio.Shape In vsoPage.Shapes
            elementText = GetModelElementBaseName(element.Master.NameU)
            'Es wird durch alle Modellelemente durchgegangen und daraus das Modell festgestellt.
            If Not StateElementTypes.Contains(elementText) And boolstateElement Then
                boolstateElement = False
            End If
            If Not ActivityElementTypes.Contains(elementText) And boolactivityElement Then
                boolactivityElement = False
            End If
            If Not SystemstructurElementTypes.Contains(elementText) And boolsystemstructurElement Then
                boolsystemstructurElement = False
            End If
            If Not FunctionstructurElementTypes.Contains(elementText) And boolfunctionstructurElement Then
                boolfunctionstructurElement = False
            End If
        Next
        'Fehlermeldung.
        If boolstateElement And boolsystemstructurElement And boolactivityElement And boolfunctionstructurElement Then
            If vsoPage.Shapes.Count = 0 Then
                Throw New System.Exception("Keine Modellelemente vorhanden.")
            End If
            Return "None"
        End If

        If KvDiagramm(boolstateElement, boolsystemstructurElement, boolactivityElement, boolfunctionstructurElement) Then
            If boolstateElement Then
                iState = iState + 1
                If iState > 1 Then Throw New System.Exception("Nur ein Zustandsdiagramm erlaubt.") Else Return "State"
            End If
            If boolsystemstructurElement Then
                iSystem = iSystem + 1
                If iSystem > 1 Then Throw New System.Exception("Nur ein Systemmodell erlaubt.") Else Return "System"
            End If
            If boolactivityElement Then
                If boolfunctionstructurElement Then
                    iFunction = iFunction + 1
                    If iFunction > 1 Then Throw New System.Exception("Nur eine Funktionshierarchie erlaubt.") Else Return "Function"
                Else
                    iActivity = iActivity + 1
                    If iActivity > 1 Then Throw New System.Exception("Nur ein Aktivitätsdiagramm erlaubt.") Else Return "Activity"
                End If
            End If
        Else
            Return "None"
        End If
        Return "None"
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : KvDiagramm(vsoDocument As Visio.Document) As String
    '       Zweck        : KV Diagramm.
    '       @param       : Zustände
    '       @return      : True oder False
    '       Quellaufruf  : FindModelType(vsoDocument As Visio.Document) As String
    '    ---------------------------------------------------------------------------------------
    Private Function KvDiagramm(a As Boolean, b As Boolean, c As Boolean, d As Boolean)
        If (a And Not b And Not c And Not d) Or (Not a And b And Not c And Not d) Or (Not a And Not b And c) Or (Not a And Not b And d) Then 'Kürzung
            Return True
        End If
        Return False
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: SendError
    '       Zweck        : Fehlermeldung geben.
    '       @param       : Inhalt der Nachricht und Errornumber.
    '    ---------------------------------------------------------------------------------------
    Public Sub SendError(msg As String, errnumber As Integer)
        Dim errMsg As String
        'Darstellung
        Form1.Label2.Visible = True
        Form1.Label2.ForeColor = Color.Red
        Form1.Timer1.Interval = 10000
        Form1.Timer1.Start()

        'StandardError.
        errMsg = "Es ist ein Fehler aufgetreten:" & msg

        'Einfach mit Case und Error Nummer: 
        If errnumber = -1 Then 'Fehler nur bei nicht erwarteten Meldungen. Diese werden mit MsgBox betont.
            errMsg = "Fehler entstanden. Siehe Meldung."
            Form1.Label2.Text = errMsg
            MsgBox(msg)
        End If
        If errnumber = 0 Then 'Wenn Textfehler in einer externen Anweisung geschrieben wird.
            'Nur Text
            If msg.Length < 41 Then
                errMsg = msg
            Else
                errMsg = "Fehler entstanden. Siehe Meldung."
                Form1.Label2.Text = errMsg
                MsgBox(msg)
            End If
        End If
        If errnumber = 1 Then
            errMsg = "Speicherung der Einstellungen fehlgeschlagen."
        End If
        If errnumber = 2 Then
            errMsg = "Keine Doppeleinträge."
        End If
        If errnumber = 3 Then
            errMsg = "Keine Schreibrechte."
        End If
        If errnumber = 32 Then
            errMsg = "Datei ist im Hintergrund offen. Bitte schließen."
        End If
        If errnumber = 33 Then
            errMsg = "Fehler entstanden. Siehe Meldung."
            Form1.Label2.Text = errMsg
            MsgBox(msg)
        End If

        Form1.Label2.Text = errMsg 'Doppelt, wegen MsgBox
        Form1.Refresh()
        Form1.Timer2.Start()
        SystemSounds.Asterisk.Play()
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: SendInfo(msg As String)
    '       Zweck        : Informationsaustausch zwischen Benutzer und Anwendung.
    '       @param       : Inhalt der Nachricht.
    '    ---------------------------------------------------------------------------------------
    Public Sub SendInfo(msg As String, time As Integer)
        Form1.Label2.Visible = True
        Form1.Label2.ForeColor = Color.White
        Form1.Label2.Text = msg
        Form1.Timer1.Interval = time * 1000 '1000 ist eine Sekunde
        Form1.Timer1.Start()
        Form1.Refresh()
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: ReleaseObject(ByVal obj As Object)
    '       Zweck        : Speicher Visio/Excel Datei freigeben.
    '       @param       : -
    '    ---------------------------------------------------------------------------------------
    Public Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: SvgExport(vsoDocument As Visio.Document)
    '       Zweck        : Bilderstellung von Visio Dokumenten.
    '       @param       : Geöffnetes Visio Dokument.
    '    ---------------------------------------------------------------------------------------
    Public Sub SvgExport(vsoDocument As Visio.Document)
        Dim vsoPage As Visio.Page = vsoDocument.Pages(1)
        Dim path As String = ProjectData.projectFullPath & filesandimagesSettings & "\" & ConvUmlaut(RemoveWhitespace(System.IO.Path.GetFileNameWithoutExtension(vsoDocument.FullName)) & ".svg")
        vsoPage.Export(path)
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Function     : GetGenID(cb As Integer) As String
    '       Zweck        : Generiert eine ID aus Buchstaben und Zahlen.
    '       @param       : Länger der ID
    '       @return      : Random ID
    '    ---------------------------------------------------------------------------------------
    Public Function GetGenID(cb As Integer) As String
        countElements = countElements + 1 'Erhöhung des Zählers aller erzeugten Elemente.
        Randomize()
        Dim rgch As String = "abcdefghijklmnopqrstuvwxyz"
        rgch = rgch & UCase(rgch) & "0123456789"
        GetGenID = ""

        Dim i As Long
        For i = 1 To cb
            GetGenID = GetGenID & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
        Next
        Return GetGenID
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : GetIDofMEDS(name As String, type As String) As String
    '       Zweck        : Erhalt der ID aus den Gesamtliste der Modellelementen
    '       @param       : Name des Modellelementes und der Diagrammtyp aus dem das Element entstand
    '       @return      : ID
    '    ---------------------------------------------------------------------------------------
    Public Function GetIDofMEDS(name As String, type As String) As String
        For Each Element As ModelElementsDataStructur In Modelelements
            If Element.Name = name And Element.Type = type Then
                Return Element.ID
            End If
        Next
        Return ""
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : ConvertStringToJson(s As String) As String
    '       Zweck        : String wird für JSON angepasst.
    '       @param       : String 
    '       @return      : Konvertierter String
    '    ---------------------------------------------------------------------------------------
    Public Function ConvertStringToJson(s As String) As String
        ConvertStringToJson = Replace(s, vbLf, "</p><p>")
        ConvertStringToJson = Replace(ConvertStringToJson, """", "\""")
        ConvertStringToJson = Replace(ConvertStringToJson, ChrW(8232), " ")
        ConvertStringToJson = Replace(ConvertStringToJson, ChrW(8220), "\""")
        ConvertStringToJson = Replace(ConvertStringToJson, ChrW(8221), "\""")
        ConvertStringToJson = Replace(ConvertStringToJson, ChrW(8222), "\""")
        ConvertStringToJson = Replace(ConvertStringToJson, ChrW(8223), "\""")
        ConvertStringToJson = Replace(ConvertStringToJson, ChrW(10), "")
        ConvertStringToJson = Replace(ConvertStringToJson, ChrW(32), " ")
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: ZipFileToSpecIF()
    '       Zweck        : Verpackt den erstellente Ordner in eine Zipdatei und bennt diesen in ".specifz" um.
    '                      Sehr anfällig für Bugs. Daher viele Fehlerabfragen.
    '       @param       : -
    '       Quellort     : Von MainSpecIFSub(modelPaths() As String)
    '    ---------------------------------------------------------------------------------------
    Public Sub ZipFileToSpecIF()
        Dim specIFZip As String = ProjectData.projectFullPath & ".zip"
        Dim folderName As String = ConvUmlaut(RemoveWhitespace(ProjectData.projectName))
        Dim specIFzFilename As String = folderName & ".specifz"
        Dim specIFzpath As String = ProjectData.projectFullPath & ".specifz"

        If File.Exists(specIFZip) Then File.Delete(specIFZip) 'Löscht die vorhandene Zip Datei falls vorhanden.
        'Überprüft ob das Schreiben schon fertig ist
        While (IsZipFIleLocked(New FileInfo(ProjectData.projectFullPath & "\" & folderName & ".specif")))
            Threading.Thread.Sleep(1000) 'Muss warten, bis Zip fertig ist.
        End While
        '----------------------------------------------------------------
        Try
            ZipFile.CreateFromDirectory(ProjectData.projectFullPath, specIFZip, CompressionLevel.Fastest, False)
            'Überprüft, ob Zip schon fertig ist.
            While (IsZipFIleLocked(New FileInfo(specIFZip)))
                Threading.Thread.Sleep(1000)
            End While
        Catch ex As Exception
            Dim errorCode As Integer = Marshal.GetHRForException(ex) And ((1 << 16) - 1)
            If errorCode <> 145 Then
                SendError(ex.Message, -1)
            Else
                Threading.Thread.Sleep(1000)
                If File.Exists(specIFZip) Then File.Delete(specIFZip)
                ZipFile.CreateFromDirectory(ProjectData.projectFullPath, specIFZip, CompressionLevel.Fastest, False)
            End If
        End Try
        '---------------------------------------------------------------------
        Try
            If System.IO.File.Exists(specIFZip) Then
                If File.Exists(specIFzpath) Then File.Delete(specIFzpath)
                My.Computer.FileSystem.RenameFile(specIFZip, specIFzFilename)

                If Directory.Exists(ProjectData.projectFullPath) Then
                    Directory.Delete(ProjectData.projectFullPath, True) 'Löscht den Ordner
                End If
            End If
        Catch ex As Exception
            'Manchmal entsteht hier ein Fehler, weil Ordner noch nicht gelöscht werden kann.
            Dim errorCode As Integer = Marshal.GetHRForException(ex) And ((1 << 16) - 1)
            If errorCode = 145 Or errorCode = 32 Then
                Threading.Thread.Sleep(1000)
                If Directory.Exists(ProjectData.projectFullPath) Then
                    Directory.Delete(ProjectData.projectFullPath, True) 'Löscht den Ordner
                End If
            Else
                SendError(ex.Message, -1)
            End If
        Finally
        End Try
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Function     : IsZipFIleLocked(file As FileInfo) As Boolean
    '       Zweck        : Wichtig herauszufinden, ob ZIP fertig ist oder nicht.
    '       @param       : Pfad der Datei 
    '       @return      : True locked.
    '       Quellaufruf  : Von ZipFileToSpecIF()
    '    ---------------------------------------------------------------------------------------
    Private Function IsZipFIleLocked(file As FileInfo) As Boolean
        Dim stream As FileStream = Nothing
        Try
            stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None)
        Catch ex As Exception
            Return True
        Finally
            If stream IsNot Nothing Then stream.Close()
        End Try
        Return False
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : IsNullOrBlank(ByVal str As String) As Boolean
    '       Zweck        : Überprüft ob ein Wert Null oder Leer ist oder aus Leerzeichen besteht.
    '       @param       : String 
    '       @return      : True wenn leer.
    '    ---------------------------------------------------------------------------------------
    Public Function IsNullOrBlank(ByVal str As String) As Boolean
        If String.IsNullOrEmpty(str) Then
            Return True
        End If
        For Each c In str
            If Not Char.IsWhiteSpace(c) Then
                Return False
            End If
        Next
        Return True
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : StripSpacesArrayComma(input As String) As String()
    '       Zweck        : Wird benutzt, wenn die Spalte Funktion mehr als einen Eintrag hat.
    '                      Beispiel "Text schreiben, Text lesen" wird zum Array {"Text schreiben", "Text lesen"}
    '       @param       : String 
    '       @return      : String Array
    '       Quellaufruf  : Requirements.vb
    '    ---------------------------------------------------------------------------------------
    Public Function StripSpacesArrayComma(input As String) As String()
        Dim output() As String
        Dim i As Integer = 0
        Dim symbol As Char = ""
        If input.Contains(",") Then symbol = ","
        If symbol = vbNullChar AndAlso input.Contains("/") Then symbol = "/"
        If symbol = vbNullChar AndAlso input.Contains(";") Then symbol = "/"

        If Not symbol = vbNullChar Then
            output = input.Split(symbol)
            For i = 0 To (output.Count - 1)
                output(i) = LTrim(String.Join(" ", output(i).Split(New Char() {}, StringSplitOptions.RemoveEmptyEntries)))
            Next
            Return output
        Else
            Return New String() {input}
        End If
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : IsFileLocked(exception As Exception) As Boolean
    '       Zweck        : Für Fehlermeldungen, wenn Datei locked ist
    '       @param       : Fehlermeldung
    '       @return      : Fehlermeldungscode 32 oder 33 als boolean
    '       Quellaufruf  : Von istBox1_DragDrop
    '       Quelle       : http://stackoverflow.com/questions/11287502/vb-net-checking-if-a-file-is-open-before-proceeding-with-a-read-write
    '    --------------------------------------------------------------------------------------- 
    Public Function IsFileLocked(exception As Exception) As Boolean
        Dim errorCode As Integer = Marshal.GetHRForException(exception) And ((1 << 16) - 1)
        Return errorCode = 32 OrElse errorCode = 33
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : IsFileLocked(exception As Exception) As Boolean
    '       Zweck        : Für Fehlermeldungen, wenn Datei locked ist
    '       @param       : -
    '       @return      : Fehlermeldungscode
    '    --------------------------------------------------------------------------------------- 
    Public Function GetErrorCode(exception As Exception) As Integer
        Dim errorCode As Integer = Marshal.GetHRForException(exception) And ((1 << 16) - 1)
        Return errorCode
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: WriteProgressbar(nbr As Integer, visible As Boolean)
    '       Zweck        : Stellt Progressbar dar.
    '       @param       : Number in Prozent und Visisble für die Darstellung.
    '       Quellort     : Von MainSpecIFSub(modelPaths() As String)
    '    ---------------------------------------------------------------------------------------
    Public Sub WriteProgressbar(nbr As Integer, visible As Boolean)
        If nbr = 100 Then
            Form1.Label6.Visible = True
        End If
        If visible = True Then
            Form1.ProgressBar1.Visible = True
            Form1.ProgressBar1.Value = nbr
            Form1.Refresh()
        Else
            Form1.ProgressBar1.Visible = False
            Form1.ProgressBar1.Value = 0
            Form1.Refresh()
            Form1.ProgressBar1.Refresh()
        End If
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Function     : IsPathWritable(ByVal strPath As String) As Boolean
    '       Zweck        : Test ob Schreibrechte vorhanden sind.
    '       @param       : Pfad
    '       @return      : True/False
    '    --------------------------------------------------------------------------------------- 
    Public Function IsPathWritable(ByVal strPath As String) As Boolean
        IsPathWritable = True
        If Not Directory.Exists(strPath) Then
            IsPathWritable = False
        Else
            Try
                Dim fs As FileStream = File.Create(strPath & "\WriteTest.txt", 4096, FileOptions.DeleteOnClose)
                fs.Close()
            Catch ex As IOException
                IsPathWritable = False
            End Try
        End If
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: RestartProject()
    '       Zweck        : Setzt alle Daten fürs Projekt zurück.
    '       Quellort     : Von MainSpecIFSub(modelPaths() As String)
    '    ---------------------------------------------------------------------------------------
    Public Sub RestartProject()
        Dim specIFZip As String = ProjectData.projectFullPath & ".zip"
        Dim folderName As String = ConvUmlaut(RemoveWhitespace(ProjectData.projectName))
        Dim specIFzFilename As String = folderName & ".specifz"
        Dim specIFzpath As String = ProjectData.projectFullPath & ".specifz"
        Try
            If File.Exists(specIFZip) Then File.Delete(specIFZip) 'Löscht die vorhandene Zip Datei falls vorhanden.
            If File.Exists(specIFzpath) Then File.Delete(specIFzpath) 'Löscht SpecifZ, falls vorhanden.
            If Directory.Exists(ProjectData.projectFullPath) Then Directory.Delete(ProjectData.projectFullPath, True) 'Löscht den Ordner

        Catch ex As Exception
            'Manchmal entsteht hier ein Fehler, weil Ordner noch nicht gelöscht werden kann.
            Dim errorCode As Integer = Marshal.GetHRForException(ex) And ((1 << 16) - 1)
            If errorCode = 145 Or errorCode = 32 Then 'Error 145 ist: Verzeichnis ist nicht leer.
                Threading.Thread.Sleep(1000) '1 Sekunde Pause
                If Directory.Exists(ProjectData.projectFullPath) Then
                    Directory.Delete(ProjectData.projectFullPath, True) 'Löscht den Ordner
                End If
            Else
                SendError(ex.Message, -1)
            End If
        Finally
            ProjectData.ReFreshData()
            ResetCount()
            WriteProgressbar(10, False)
        End Try
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: AddRelation_vsoPage_Overview(vsoPageID As String, OverviewID As String, targetID As String)
    '       Zweck        : Relationen hinzuzufügen für die "Übersicht" und Diagram OVerview  pro Modell auf reqif.net
    '                      Alle Modelle verfügen dann eine Relation z.B. Diagramm Overview zeigt auf ModellelementX
    '       @param       : -
    '    ---------------------------------------------------------------------------------------
    Public Sub AddRelation_vsoPage_Overview(vsoPageID As String, OverviewID As String, targetID As String)
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
            "        ""id"": ""RVis-" & OverviewID & "-" & targetID & """," & vbCrLf &
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
    '       Sub-Anweisung: AddObjectState(id As String, name As String, text As String)
    '       Zweck        : Fügt Zustände als Objekt hinzu.
    '       @param       : -
    '       Quellort     : Aus allen Modellen.
    '    ---------------------------------------------------------------------------------------
    Public Sub AddObjectState(id As String, name As String, text As String)
        If IsNullOrBlank(name) Then Throw New System.Exception("Es fehlt die Bezeichnung in einem Modellelement. Bitte ergänzen.")
        ProjectData.ObjekteDS = ProjectData.ObjekteDS & "{" & vbCrLf &
        "        ""id"": """ & id & """," & vbCrLf &
        "        ""title"": """ & name & """," & vbCrLf &
        "        ""attributes"": [{" & vbCrLf &
        "            ""title"": ""dcterms:title""," & vbCrLf &
        "            ""attributeType"": ""AT-Sta-Name""," & vbCrLf &
        "            ""value"": """ & name & """" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""dcterms:description""," & vbCrLf &
        "            ""attributeType"": ""AT-Sta-Text""," & vbCrLf &
        "            ""value"": ""<div><p>" & ConvertStringToJson(text) & "</p></div>""" & vbCrLf &
        "        }]," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""objectType"": ""OT-Sta""" & vbCrLf &
        "    },"
    End Sub
End Module
