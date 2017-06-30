Imports Excel = Microsoft.Office.Interop.Excel
'    ---------------------------------------------------------------------------------------
'       Modul   : Requirements
'       Datum   : 29.04.2017
'       Author  : Philipp Mochine
'       Zweck   : Anforderungen aus einer Excelliste werden eingelesen und in das Integrationsmodell hinzugefügt.
'                 Beziehungen zwischen Anforderungen und Funktionshierarchie werden hier automatisch gesetzt.
'       Quellort: Von ExecuteModel(modelPath As String)
'    ---------------------------------------------------------------------------------------
Module Requirements
    Private xlWorkBook As Excel.Workbook
    Private xlWorkSheet As Excel.Worksheet
    Private Const modelType As String = "Requirement"
    Private reqRange As Excel.Range
    Private reqColumnTitle As Collection
    Private reqColumnAT As Collection
    Private ReadOnly Property XlTitle As String() = New String() {"dcterms:title", "reqif.name", "titel", "title"}
    Private ReadOnly Property XlDescription As String() = New String() {"dcterms:description", "ReqIF.Text", "Beschreibung", "description"}
    Private ReadOnly Property XlState As String() = New String() {"specIF:state", "reqif.foreignstate", "status", "state"}
    Private ReadOnly Property XlTypeTitle As String() = New String() {"type", "typ", "anforderungsart", "art", "ireb"}
    Private ReadOnly Property XlType As String() = New String() {"constraints", "quality", "function", "role"}
    Private ReadOnly Property XlFunction As String() = New String() {"function", "funktion"}
    Private Const roleType As String = "stakeholder"
    Private HierarchieList As New List(Of ModelElementsDataStructur) 'Eine Hierarchieliste wird geführt und anschließend diese zu sortieren.
    Private modelDescription As String = ConvertStringToJson("Eine Anforderung beschreibt eine oder mehrere Eigenschaften oder Verhaltensweise eines Systems. Diese Anforderungen müssen vom System erfüllt werden.")
    Private Const elementDescription As String = "ist ein Element des Aktivitätsmodells. Es führt Aktionen aus."
    '    ---------------------------------------------------------------------------------------
    '       Function     : RequirementsAnalyse(xlWorkBooktemp As Excel.Workbook) As Boolean
    '       Zweck        : Liest Exceldaten ein und erstellt die JSON Struktur. 
    '       @param       : Pointer auf das Excel Dokument.
    '       Quellort     : Von ExecuteModel(modelPath As String)
    '    ---------------------------------------------------------------------------------------
    Public Function RequirementsAnalyse(xlWorkBooktemp As Excel.Workbook) As Boolean
        RequirementsAnalyse = True
        xlWorkBook = xlWorkBooktemp
        xlWorkSheet = xlWorkBook.Worksheets(1)
        Try
            CreateRequirementsSpecIFBase()
            CreateObjectType()
            CreateRequirementsSpecIFCode()
            SetHierarchie()
        Catch ex As Exception
            RequirementsAnalyse = False
            SendError(ex.Message, -1)
        Finally
            ReleaseObject(xlWorkSheet)
            HierarchieList = New List(Of ModelElementsDataStructur)
        End Try
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: CreateRequirementsSpecIFCode()
    '       Zweck        : Liest Exceldaten ein und erstellt die JSON Struktur. 
    '       @param       : -
    '       Quellort     : Von RequirementsAnalyse(xlWorkBooktemp As Excel.Workbook)
    '    ---------------------------------------------------------------------------------------
    Private Sub CreateRequirementsSpecIFCode()
        Dim rngCell As Excel.Range
        Dim i As Integer = 1
        Dim id As String, title As String = "", type As String = ""
        Dim objecttemp As String = "", elementID As String
        Dim functionName() As String = New String() {}

        For Each rngCell In reqRange
            If reqColumnTitle(i) = XlTitle(0) Then title = rngCell.Value 'Titel wird gesetzt.
            If XlTypeTitle.Contains(reqColumnTitle(i).ToLower()) Then
                If XlType.Contains((rngCell.Value).ToLower()) Then
                    type = rngCell.Value
                Else
                    If IsNullOrBlank(rngCell.Value) Then
                        If CheckIfEmptyRow(rngCell) Then
                            Throw New System.Exception("Zeile " & rngCell.Row & " ist leer. Bitte ändern.")
                        End If
                        Throw New System.Exception("IREB-Bezeichnung fehlt in Zeile " & rngCell.Row)
                    End If
                    'Wenn eine Falsche Bezeichnung verwendet wurde.
                    Throw New System.Exception(rngCell.Value & " ist eine falsche Bezeichnung. Erlaubt sind nur diese Status: 'Constraints', 'Quality', 'Function', 'Role'")
                End If
            End If
            'Wenn Funktionsspalte erreicht wird, wird ein Array erstellt.
            If XlFunction.Contains(reqColumnTitle(i).ToLower()) And Not IsNullOrBlank(rngCell.Value) Then functionName = StripSpacesArrayComma(rngCell.Value)

            objecttemp = objecttemp & "{" & vbCrLf & 'Elemente der Zeile werden hier gesammelt
            "            ""title"": """ & reqColumnTitle(i) & """," & vbCrLf &
            "            ""attributeType"": """ & reqColumnAT(i) & """," & vbCrLf &
            "            ""value"": """ & ConvertStringToJson(rngCell.Value) & """" & vbCrLf & '??? Testen ob es geht
            "        },"
            i = i + 1

            If i > reqRange.Columns.Count Then 'Wurde die letzte Spalte abgearbeitet.
                i = 1
                id = "O-" & GetGenID(27)
                If title <> "" Then
                    ProjectData.ObjekteDS = ProjectData.ObjekteDS & "{" & vbCrLf & 'Elemente werden hier zusammengefasst als ein Objekt
                    "        ""id"": """ & id & """," & vbCrLf &
                    "        ""title"": """ & ConvertStringToJson(title) & """," & vbCrLf &
                    "        ""attributes"": [" & Left(objecttemp, Len(objecttemp) - 1) & "]," & vbCrLf &
                    "        ""revision"": 0," & vbCrLf &
                    "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
                    "        ""objectType"": ""OT-Req-Exl""" & vbCrLf &
                    "},"
                    If title.ToLower.Contains(roleType) Then type = "Role" 'Wenn Stakeholder vorhanden ist, wird der Typ zu "Role".

                    HierarchieList.Add(New ModelElementsDataStructur(title, id, type))
                    'Hier werden die Relationen zwischen Anforderungen und Funktion gesetzt
                    If functionName IsNot Nothing AndAlso functionName.Count > 0 Then
                        For Each element As String In functionName
                            If element <> "" Then
                                elementID = GetIDofMEDS(element, "Function")
                                If elementID = "" Then
                                    'Wenn Anforderung zuerst gelesen wird. 
                                    'Es wird die ID von der ganzen Zeile angegeben.
                                    Modelelements.Add(New ModelElementsDataStructur(element, id, modelType))
                                Else
                                    'Wenn FunctionStructur schon gelesen wurde Relation hinzufügen.
                                    AddRelation(id, elementID)
                                End If
                            End If
                        Next
                    End If
                    objecttemp = ""
                    ReDim functionName(-1)
                Else
                    objecttemp = ""
                    Throw New System.Exception("Titel darf nicht leer sein. Zeile: " & rngCell.Row)
                End If
            End If
        Next
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Function     : CheckIfEmptyRow() As Boolean
    '       Zweck        : 
    '       @param       : 
    '       @return      : True, wenn Zeile leer ist
    '    ---------------------------------------------------------------------------------------
    Public Function CheckIfEmptyRow(rngCell As Excel.Range) As Boolean
        Dim row As Integer
        Dim i As Integer
        Dim emptyCell = True 'True wenn leer

        row = rngCell.Row
        For i = 1 To reqRange.Columns.Count
            If Not IsNullOrBlank(reqRange.Cells(row, i).Value) Then
                emptyCell = False 'Es wurde was gefunden.
            End If
        Next
        Return emptyCell
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: AddRelation(reqID As String, funcID As String)
    '       Zweck        : Erstellt Relation zwischen Funktionselement und Anforderungselement
    '       @param       : ID von Anforderung und Funktion
    '       Quellort     : Von CreateRequirementsSpecIFCode()
    '    ---------------------------------------------------------------------------------------
    Private Sub AddRelation(reqID As String, funcID As String)
        ProjectData.RelationenDS = ProjectData.RelationenDS & "{" & vbCrLf &
    "        ""id"": ""RVis-" & reqID & "-" & funcID & """," & vbCrLf &
    "        ""title"": ""erfüllt""," & vbCrLf &
    "        ""revision"": 0," & vbCrLf &
    "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
    "        ""relationType"": ""RT-Satisfaction""," & vbCrLf &
    "        ""source"": {" & vbCrLf &
    "            ""id"": """ & funcID & """," & vbCrLf &
    "            ""revision"": 0" & vbCrLf &
    "        }," & vbCrLf &
    "        ""target"": {" & vbCrLf &
    "            ""id"": """ & reqID & """," & vbCrLf &
    "            ""revision"": 0" & vbCrLf &
    "        }" & vbCrLf &
    "    },"
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: CreateObjectType()
    '       Zweck        : Erstellt Objekttyp OT-Req-Exl und Objekte
    '       @param       : -
    '       Quellort     : Von RequirementsAnalyse(xlWorkBooktemp As Excel.Workbook)
    '    ---------------------------------------------------------------------------------------
    Private Sub CreateObjectType()
        Dim objecttemp As String = "", title As String, id As String
        Dim i As Integer = 1
        Dim titleexists As Boolean = False 'Muss True sein, damit Anforderung analysiert wird.
        Dim stateexists As Boolean = False 'Muss True sein, damit Anforderung analysiert wird.

        reqColumnTitle = New Collection 'Liste von allen Spalteninhalte
        reqColumnAT = New Collection 'Liste der Anforderungsid. Wird synchron mit reqColumnTitle gebildet.

        Do While xlWorkSheet.Cells(1, i).value <> "" 'Spalten müssen konsitent sein. Bei einer leeren Spalte wird hier abgebrochen.
            title = xlWorkSheet.Cells(1, i).value
            Dim comp As StringComparison = StringComparison.OrdinalIgnoreCase
            If XlTitle.Contains(title.ToLower()) Then 'Wenn Spalte z.B. Titel heißt, wird dies zu dcterms:title. Es ist Pflicht einen Titel anzugeben.
                titleexists = True
                title = XlTitle(0)
            End If
            If XlDescription.Contains(title.ToLower()) Then title = XlDescription(0) 'Beschreibung
            If XlState.Contains(title.ToLower()) Then title = "SpecIF:State" 'Status Titel
            If XlTypeTitle.Contains(title.ToLower()) Then stateexists = True

            id = "AT-" & GetGenID(27)
            objecttemp = objecttemp & "{" & vbCrLf &
            "            ""id"": """ & id & """," & vbCrLf &
            "            ""title"": """ & title & """," & vbCrLf &
            "            ""dataType"": ""XLS_Text""," & vbCrLf &
            "            ""revision"": 0," & vbCrLf &
            "            ""changedAt"": """ & GetCreatedAt() & """" & vbCrLf &
            "        },"
            i = i + 1
            reqColumnTitle.Add(title)
            reqColumnAT.Add(id)
        Loop
        'Fehler Möglichkeiten
        If i = 1 Then Throw New System.Exception("Keine Spalte gefunden.") 'Wenn es keine Spalte gab.
        If titleexists = False Then Throw New System.Exception("Spalte mit ""Titel"" nicht gefunden. Für SpecIF ist das eine Pflichtangabe.")
        If stateexists = False Then Throw New System.Exception("Spalte mit ""Typ"" nicht gefunden. Für SpecIF ist das eine Pflichtangabe. Klassifzierung der Anforderungsarten mit den Spaltentitel: ""Type"", ""Typ"", ""Anforderungsart"", ""Art"", ""IREB""")

        objecttemp = ",{" & vbCrLf &
            "		""id"":  ""OT-Req-Exl""," & vbCrLf &
            "		""title"": ""SpecIF:Requirement"", " & vbCrLf &
            "		""description"":  ""Eine &#8623; Anforderung dokumentiert einzelnes physisches oder funktionales Bedürfnis, das der betreffende Entwurf, das Produkt oder der Prozess erfüllen muss.""," & vbCrLf &
            "		""icon"": ""&#8623;""," & vbCrLf &
            "		""attributeTypes"": [" & Left(objecttemp, Len(objecttemp) - 1) & "]," & vbCrLf &
            "		""revision"": 0," & vbCrLf &
            "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
            "        ""creation"": [""manual""]" & vbCrLf &
            "	}"
        ProjectData.CompleteDS = Replace(ProjectData.CompleteDS, "##objectTypes##", objecttemp & "##objectTypes##") 'Wird direkt in das Integrationsmodell bei den Objekttypen gespeichert.
        Dim lastrow As Integer = xlWorkSheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row 'Letzte Reihe wird emitterlt.
        reqRange = xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(lastrow, i - 1)) 'Range wird gesetzt..

    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: CreateRequirementsSpecIFBase()
    '       Zweck        : Erstellt die JSON Basis-Struktur mit Objekten und Hierarchie
    '       @param       : -
    '       Quellort     : Von RequirementsAnalyse(xlWorkBooktemp As Excel.Workbook)
    '    ---------------------------------------------------------------------------------------
    Private Sub CreateRequirementsSpecIFBase()
        ProjectData.ObjekteDS = ProjectData.ObjekteDS & "{" & vbCrLf &
        "        ""id"": ""Fld-Anforderung""," & vbCrLf &
        "        ""title"": ""Anforderung""," & vbCrLf &
        "        ""attributes"": [{" & vbCrLf &
        "            ""title"": ""dcterms:title""," & vbCrLf &
        "            ""attributeType"": ""AT-Fld-Name""," & vbCrLf &
        "            ""value"": ""Anforderung""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""dcterms:description""," & vbCrLf &
        "            ""attributeType"": ""AT-Fld-Text""," & vbCrLf &
        "            ""value"": ""<div><p>" & modelDescription & "</p></div>""" & vbCrLf &
        "        }]," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""objectType"": ""OT-Fld""" & vbCrLf &
        "    },{" & vbCrLf &
        "        ""id"": ""Fld-Rollen""," & vbCrLf &
        "        ""title"": ""Rollen""," & vbCrLf &
        "        ""attributes"": [{" & vbCrLf &
        "            ""title"": ""dcterms:title""," & vbCrLf &
        "            ""attributeType"": ""AT-Fld-Name""," & vbCrLf &
        "            ""value"": ""Rollen""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""dcterms:description""," & vbCrLf &
        "            ""attributeType"": ""AT-Fld-Text""," & vbCrLf &
        "            ""value"": ""<div><p>Ist eine Randbedingung der Benutzer bzw. Stakeholder des Systems.</p></div>""" & vbCrLf &
        "        }]," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""objectType"": ""OT-Fld""" & vbCrLf &
        "    },{" & vbCrLf &
        "        ""id"": ""Fld-Quality""," & vbCrLf &
        "        ""title"": ""Qualität""," & vbCrLf &
        "        ""attributes"": [{" & vbCrLf &
        "            ""title"": ""dcterms:title""," & vbCrLf &
        "            ""attributeType"": ""AT-Fld-Name""," & vbCrLf &
        "            ""value"": ""Qualität""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""dcterms:description""," & vbCrLf &
        "            ""attributeType"": ""AT-Fld-Text""," & vbCrLf &
        "            ""value"": ""<div><p>Ist eine nichtfunktionale Anforderung. Sie beschreiben die Qualität des Systems, also wie gut die Leistung sein soll. </p></div>""" & vbCrLf &
        "        }]," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""objectType"": ""OT-Fld""" & vbCrLf &
        "    },{" & vbCrLf &
        "        ""id"": ""Fld-Function""," & vbCrLf &
        "        ""title"": ""Funktion""," & vbCrLf &
        "        ""attributes"": [{" & vbCrLf &
        "            ""title"": ""dcterms:title""," & vbCrLf &
        "            ""attributeType"": ""AT-Fld-Name""," & vbCrLf &
        "            ""value"": ""Funktion""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""dcterms:description""," & vbCrLf &
        "            ""attributeType"": ""AT-Fld-Text""," & vbCrLf &
        "            ""value"": ""<div><p>Ist eine funktionale Anforderung. Sie beschreiben was das System tun soll.</p></div>""" & vbCrLf &
        "        }]," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""objectType"": ""OT-Fld""" & vbCrLf &
        "    },{" & vbCrLf &
        "        ""id"": ""Fld-Constraints""," & vbCrLf &
        "        ""title"": ""Randbedingung""," & vbCrLf &
        "        ""attributes"": [{" & vbCrLf &
        "            ""title"": ""dcterms:title""," & vbCrLf &
        "            ""attributeType"": ""AT-Fld-Name""," & vbCrLf &
        "            ""value"": ""Randbedingung""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""title"": ""dcterms:description""," & vbCrLf &
        "            ""attributeType"": ""AT-Fld-Text""," & vbCrLf &
        "            ""value"": ""<div><p>Eine Randbedingung ist eine Anforderung, die die Art und Weise zum Erfüllen des Systems einschränkt. Die Klassifizierung erfolgt nach IREB. </p></div>""" & vbCrLf &
        "        }]," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""objectType"": ""OT-Fld""" & vbCrLf &
        "    },"

        ProjectData.HierarchieDS = ProjectData.HierarchieDS & "{" & vbCrLf &
        "    ""id"": ""SH-Fld-Anforderung""," & vbCrLf &
        "    ""object"": ""Fld-Anforderung""," & vbCrLf &
        "    ""revision"": 0," & vbCrLf &
        "    ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "    ""nodes"": [{" & vbCrLf &
        "        ""id"": ""SH-Fld-Rollen""," & vbCrLf &
        "        ""object"": ""Fld-Rollen""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""nodes"": [##Role##]" & vbCrLf &
        "     },{" & vbCrLf &
        "        ""id"": ""SH-Fld-Quality""," & vbCrLf &
        "        ""object"": ""Fld-Quality""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""nodes"": [##Quality##]" & vbCrLf &
        "     },{" & vbCrLf &
        "        ""id"": ""SH-Fld-Function""," & vbCrLf &
        "        ""object"": ""Fld-Function""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""nodes"": [##Function##]" & vbCrLf &
        "     },{" & vbCrLf &
        "        ""id"": ""SH-Fld-Constraints""," & vbCrLf &
        "        ""object"": ""Fld-Constraints""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""nodes"": [##Constraints##]" & vbCrLf &
        "     }]" & vbCrLf &
        "},"
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: SetHierarchie()
    '       Zweck        : Setzt Hierarchie und sortiert nach Alphabet.
    '       @param       : -
    '       Quellort     : Von RequirementsAnalyse(xlWorkBooktemp As Excel.Workbook)
    '    ---------------------------------------------------------------------------------------
    Private Sub SetHierarchie()
        Dim sortedList As New List(Of ModelElementsDataStructur) 'temporäre Liste
        Dim hrarStr As String

        sortedList = HierarchieList.OrderBy(Function(x) x.Type).ThenBy(Function(x) x.Name).ToList
        For Each Element As ModelElementsDataStructur In sortedList
            hrarStr = "{" & vbCrLf &
            "                   ""id"": ""SH-" & Element.ID & "-" & modelType & """," & vbCrLf &
            "                   ""object"": """ & Element.ID & """," & vbCrLf &
            "                   ""revision"": 0," & vbCrLf &
            "                   ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
            "                   ""nodes"": []" & vbCrLf &
            "               },##" & Element.Type & "##"
            ProjectData.HierarchieDS = Replace(ProjectData.HierarchieDS, "##" & Element.Type & "##", hrarStr)
        Next
    End Sub
End Module
