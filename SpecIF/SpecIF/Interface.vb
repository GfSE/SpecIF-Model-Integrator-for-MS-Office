Imports System.IO
'    ---------------------------------------------------------------------------------------
'       Klasse  : Form1
'       Datum   : 29.04.2017
'       Author  : Philipp Mochine
'       Projekt : Die Vorstellung der Möglichkeiten von SpecIF (specif.de) mit SysML bzw. UML Notation und JSON als Struktur.
'                 Teilaufgabe der Masterarbeit: Entwicklung eines Integrationsmodells zur Überführung von Systemmodellen.
'                 Kommentare beziehen sich immer auf die darauffolgende Anweisung/Module oder Variablen.
'       Zweck   : Interfacemodul des Projektes. Hier befinden sich alle "Button" und deren Aktionen.                            
'       Quellort: -
'    ---------------------------------------------------------------------------------------
Public Class Form1
    Private openFilepath() As String 'Speichert alle Modellpfaden ab.
    ' Friend WithEvents ListB1 As New MyListBox
    Private varOnOff As Boolean
    Public lastSpecIFzPath As String = "" 'Pfad zum Ordner der generierten Specif Datei.
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: New()
    '       Zweck        : Initialisiert Variablen sowie Interface-Formen.      
    '       Quellaufruf  : -
    '    ---------------------------------------------------------------------------------------
    Public Sub New()
        InitializeComponent()
        ReDim openFilepath(0)
        Button2.Enabled = False
        ProgressBar1.Value = 10
        ProgressBar1.Visible = False
        Label2.Visible = False
        Label6.Visible = False
        varOnOff = True
        SettingsOnOff()
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    '       Zweck        : Öffnet die Hauptfunktion des MainSpecIF.vb Moduls.                       
    '       @param       : -     
    '       Quellaufruf  : Klick des Benutzers.
    '    ---------------------------------------------------------------------------------------    
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If varOnOff = False Then 'Frontinterface wird aktuell darsgestellt.
            Button2.Enabled = False
            SaveCurrentInfos()
            MainSpecIFSub(openFilepath)
            Button2.Enabled = True
            Refresh()
        Else
            'Speicherung der Daten.
            SaveCurrentInfos()
            SendInfo("Daten wurden gespeichert.", 2)
            SettingsOnOff()
        End If
    End Sub

    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: SaveCurrentInfos
    '       Zweck        : Speichert aktuelle Projektinformationen ab.                  
    '       @param       : -     
    '       Quellaufruf  : Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    '    --------------------------------------------------------------------------------------- 
    Private Sub SaveCurrentInfos()
        Try
            If ProjectData.projectinfoName Is Nothing Then 'Im Falle, wenn etwas nicht gespeichert wurde.
                If TextBox1.Text <> "" Then
                    ProjectData.projectinfoName = TextBox1.Text
                Else
                    ProjectData.projectinfoName = "Modelproject"
                End If
                If TextBox2.Text <> "" Then
                    ProjectData.projectinfoID = TextBox2.Text
                Else
                    ProjectData.projectinfoID = "P-" & GetGenID(27)
                End If
                If TextBox3.Text <> "" Then
                    ProjectData.projectInformation = TextBox3.Text
                Else
                    ProjectData.projectInformation = "Projekt Informationen"
                End If
            Else
                If varOnOff = True Then 'Nur Speichern, wenn wir uns in den Einstellungen befinden.
                    ProjectData.projectinfoName = TextBox1.Text
                    ProjectData.projectinfoID = TextBox2.Text
                    ProjectData.projectInformation = TextBox3.Text
                End If
            End If
        Catch ex As Exception
            SendError(ex.Message, 1)
        End Try
    End Sub

    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: ListBox1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragEnter
    '       Zweck        : Damit das Dragen angezeigt wird.                 
    '       @param       : -     
    '       Quellaufruf  : Drag einer Datei.
    '       Quelle       : https://support.microsoft.com/de-de/help/307966/how-to-provide-file-drag-and-drop-functionality-in-a-visual-c-application
    '    ---------------------------------------------------------------------------------------    
    Private Sub ListBox1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: ListBox1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragDrop
    '       Zweck        : Ermöglicht den Drap&Drop von Dateien. Nur Excel und Visio-Dateien sind erlaubt.
    '                      Fügt Modellpfade zu "openFilepath" hinzu.
    '       @param       : -     
    '       Quellaufruf  : Drag einer Datei.
    '    ---------------------------------------------------------------------------------------  
    Private Sub ListBox1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            Dim i As Integer
            Dim file As FileInfo
            Dim fileExtension As String
            Dim stream As FileStream = Nothing
            ' Daten als Array. 
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            ' Schleife durch alle Modellen
            Try
                For i = 0 To MyFiles.Length - 1 'Menge der gedroppten Files
                    fileExtension = System.IO.Path.GetExtension(MyFiles(i))
                    If ExcelExtension.Contains(fileExtension) Or VisioExtension.Contains(fileExtension) Then 'Nur Excel oder Visio Datein.
                        If Not FileExists(MyFiles(i)) Then 'Datei darf nicht doppelt existieren. (Nur ein Name).
                            If openFilepath.Count > 1 AndAlso openFilepath.Count <= 3 Then
                                SendInfo("Wegen Datengröße Firefox verwenden.", 3)
                            ElseIf openFilepath.Count > 3 Then
                                SendInfo("Datengröße kann für Reader zu groß sein.", 3)
                            End If
                            file = New FileInfo(MyFiles(i))
                            stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None) 'Ein Stream wird geöffnet, um zu Schauen, ob Datei lesbar ist.
                            stream.Close()
                            If openFilepath(0) = "" Then
                                openFilepath(0) = MyFiles(i)
                            Else
                                ReDim Preserve openFilepath(UBound(openFilepath) + 1)
                                openFilepath(UBound(openFilepath)) = MyFiles(i)
                            End If
                            ListBox1.Items.Add(System.IO.Path.GetFileNameWithoutExtension(MyFiles(i))) 'Darstellung in die Listbox.
                        Else
                            SendError("", 2)
                        End If
                    Else
                        SendError("Die Datei " & System.IO.Path.GetFileName(MyFiles(i)) & " hat keine gültige Dateiendung.", 0)
                    End If
                    If ListBox1.Items.Count <> 0 Then
                        Button2.Enabled = True
                    End If
                Next
            Catch ex As Exception
                If TypeOf ex Is IOException AndAlso IsFileLocked(ex) Then
                    SendError(ex.Message, -1)
                End If
            End Try
        End If
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Function     : fileExists(name As String) As Boolean
    '       Zweck        : Überprüft ob Datei schon in "openFilepath" Array vorhanden ist.
    '       @param       : Enthält Pfad 
    '       @return      : Boolean. True wenn Datei gefunden.
    '       Quellaufruf  : Von istBox1_DragDrop und ListBox1_MouseClick
    '    --------------------------------------------------------------------------------------- 
    Private Function FileExists(path As String) As Boolean
        If openFilepath(0) = "" Then
            Return False
        Else
            If openFilepath.Contains(path) Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: ListBox1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBox1.Click
    '       Zweck        : Fügt Modellpfade zu "openFilepath" hinzu.
    '       @param       : -
    '       Quellaufruf  : Klick des Benutzers.
    '    ---------------------------------------------------------------------------------------
    Private Sub ListBox1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBox1.Click
        If ListBox1.SelectedIndex = -1 Then
            Dim myStream As Stream = Nothing
            Dim openFileDialog1 As New OpenFileDialog() With {
                .InitialDirectory = modelsInitialDirectorySettings,
                .Filter = modelsFilterSettings,
                .FilterIndex = 2,
                .RestoreDirectory = True
            }
            'Öffnet das Fenster zum Suchen bzw.Auswählen der Datei
            If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Try
                    'Öffnet Datei als Stream
                    myStream = openFileDialog1.OpenFile()
                    If (myStream IsNot Nothing) Then
                        If Not FileExists(openFileDialog1.FileName) Then
                            If openFilepath.Count > 1 AndAlso openFilepath.Count <= 3 Then
                                SendInfo("Wegen Datengröße Firefox verwenden.", 3)
                            ElseIf openFilepath.Count > 3 Then
                                SendError("Datengröße kann für Reader zu groß sein.", 0)
                            End If
                            'Fügt den Pfad hinzu
                            If openFilepath(0) = "" Then
                                openFilepath(0) = openFileDialog1.FileName
                            Else
                                ReDim Preserve openFilepath(UBound(openFilepath) + 1)
                                openFilepath(UBound(openFilepath)) = openFileDialog1.FileName
                            End If
                            ListBox1.Items.Add(System.IO.Path.GetFileNameWithoutExtension(openFileDialog1.FileName))
                            If ListBox1.Items.Count <> 0 Then
                                Button2.Enabled = True
                            End If
                        Else
                            SendError("", 2)
                        End If
                    End If
                Catch Ex As Exception
                    SendError(Ex.Message, -1)
                Finally
                    If (myStream IsNot Nothing) Then
                        myStream.Close()
                    End If
                End Try
            End If
        End If
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
    '       Zweck        : Löscht Einträge beim Klicken.
    '       @param       : -
    '       Quellaufruf  : Klick des Benutzers.
    '    ---------------------------------------------------------------------------------------
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim fileName As String
        Dim i As Integer
        Try
            If ListBox1.SelectedIndex <> -1 Then 'Nur wenn Datei ausgewählt wurde
                fileName = ListBox1.SelectedItem
                For i = 0 To openFilepath.Length
                    If Path.GetFileNameWithoutExtension(openFilepath(i)) = fileName Then
                        fileName = openFilepath(i)
                        openFilepath = openFilepath.Where(Function(s) s <> fileName).ToArray
                        If openFilepath.Length = 0 Then
                            ReDim openFilepath(0)
                            Button2.Enabled = False
                        End If
                        Exit For
                    End If
                Next
                ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            End If
        Catch ex As Exception
            SendError(ex.Message, -1)
        End Try
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
    '       Zweck        : Einstellungsbutton auf dem Interface. Öffnet die Einstellungen.
    '       @param       : -
    '       Quellaufruf  : Klick des Benutzers.
    '    ---------------------------------------------------------------------------------------
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        SettingsOnOff()
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: SettingsOnOff()
    '       Zweck        : Wechselt zwischen Einstellungen und Hauptinterface.
    '       @param       : -
    '       Quellaufruf  : Initialisierung und PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
    '    ---------------------------------------------------------------------------------------
    Private Sub SettingsOnOff()
        Try
            If varOnOff = True Then 'Frontinterface darstellen.
                'Settings
                Label3.Visible = False
                Label4.Visible = False
                Label5.Visible = False
                TextBox1.Visible = False
                TextBox2.Visible = False
                TextBox3.Visible = False
                'FrontInterface
                ListBox1.Visible = True
                Label1.Text = "Drag && Drop your model"
                Button2.Text = "SpecIF"
                If openFilepath(0) = "" Then
                    Button2.Enabled = False
                End If
                varOnOff = False
            Else
                'FrontInterface
                ListBox1.Visible = False
                Label1.Text = "Project settings"
                'Settings
                Label3.Visible = True
                Label4.Visible = True
                Label5.Visible = True
                If ProjectData.projectinfoName <> "" Then
                    TextBox1.Text = ProjectData.projectinfoName
                Else
                    TextBox1.Text = "Modelproject"
                End If
                If ProjectData.projectinfoID Is Nothing Then
                    TextBox2.Text = "P-" & GetGenID(27)
                Else
                    TextBox2.Text = ProjectData.projectinfoID
                End If
                If ProjectData.projectInformation Is Nothing Then
                    TextBox3.Text = "Projekt Informationen"
                Else
                    TextBox3.Text = ProjectData.projectInformation
                End If

                TextBox1.Visible = True
                TextBox2.Visible = True
                TextBox3.Visible = True
                Button2.Enabled = True
                Button2.Text = "Speichern"
                SendInfo("Bitte Projektdaten eingeben.", 4)
                varOnOff = True
            End If
        Catch ex As Exception
            SendError(ex.Message, -1)
        End Try
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
    '       Zweck        : Wechselt zwischen Einstellungen und Hauptinterface.
    '       @param       : -
    '       Quellaufruf  : Initialisierung und PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
    '    ---------------------------------------------------------------------------------------
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label2.Visible = False
        Timer1.Stop()
        Label6.Visible = False
        lastSpecIFzPath = ""
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: Timer2_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
    '       Zweck        : Deaktiviert Progressbar
    '       @param       : -
    '       Quellaufruf  : MainSpecIFSub(modelPaths() As String)
    '    ---------------------------------------------------------------------------------------
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        WriteProgressbar(0, False)
        Timer2.Stop()
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        If lastSpecIFzPath <> "" Then
            Process.Start(lastSpecIFzPath)
        End If
    End Sub
End Class
