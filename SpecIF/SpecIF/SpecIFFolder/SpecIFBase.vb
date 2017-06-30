'    ---------------------------------------------------------------------------------------
'       Modul   : SpecIFBase
'       Datum   : 29.04.2017
'       Author  : Philipp Mochine
'       Zweck   : Erstellt die Basisstruktur der JSON bzw. SpecIF Datei                           
'       Quellort: Von MainSpecIFSub(modelPaths() As String)
'    ---------------------------------------------------------------------------------------
Module SpecIFBase
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: CreateSpecIFBase()
    '       Zweck        : Füllt das Integrationsmodell mit Basisstruktur.
    '       @param       : -
    '       Quellort     : Von MainSpecIFSub(modelPaths() As String)
    '    ---------------------------------------------------------------------------------------
    Public Sub CreateSpecIFBase()
        ProjectData.CompleteDS = GetCreatedFramework()
        ProjectData.CompleteDS = Replace(ProjectData.CompleteDS, "##dataTypes##", GetdataTypes())
        ProjectData.CompleteDS = Replace(ProjectData.CompleteDS, "##objectTypes##", GetobjectTypes() & "##objectTypes##")
        ProjectData.CompleteDS = Replace(ProjectData.CompleteDS, "##relationTypes##", GetrelationTypes())
        ProjectData.CompleteDS = Replace(ProjectData.CompleteDS, "##hierarchyTypes##", GethierarchyTypes())

        ProjectData.CompleteDS = Replace(ProjectData.CompleteDS, "##hierarchies##", Gethierarchies())
        ProjectData.CompleteDS = Replace(ProjectData.CompleteDS, "##objects##", Getobjects() & "##objects##")
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Function     : GetCreatedFramework() As String
    '       Zweck        : Erstellt Framework in einer JSON-Struktur.
    '       @param       : -
    '       @return      : String.
    '       Quellort     : Von CreateSpecIFBase()
    '    ---------------------------------------------------------------------------------------
    Private Function GetCreatedFramework() As String
        Dim framework As String
        framework = "{" & vbCrLf &
        "    ""id"": ""##projectinfoID##""," & vbCrLf &
        "    ""title"": ""##projectinfoName##""," & vbCrLf &
        "    ""specifVersion"": ""0.9.2""," & vbCrLf &
        "    ""tool"": ""Interactive-Spec""," & vbCrLf &
        "    ""toolVersion"": ""0.92.31""," & vbCrLf &
        "    ""rights"": {" & vbCrLf &
        "        ""title"": ""Creative Commons 4.0 CC BY-SA""," & vbCrLf &
        "        ""type"": ""dcterms:rights""," & vbCrLf &
        "        ""url"": ""https://creativecommons.org/licenses/by-sa/4.0/""" & vbCrLf &
        "    }," & vbCrLf &
        "    ""createdAt"": ""##createdAt##""," & vbCrLf &
        "    ""createdBy"": {" & vbCrLf &
        "        ""familyName"": ""##infoFamilyName##""," & vbCrLf &
        "        ""givenName"": ""##infoGivenName##""," & vbCrLf &
        "        ""org"": {" & vbCrLf &
        "            ""organizationName"": ""##infoOrganisationName##""" & vbCrLf &
        "        }," & vbCrLf &
        "        ""email"": {" & vbCrLf &
        "            ""type"": ""text/html""," & vbCrLf &
        "            ""value"": ""##infoEmail##""" & vbCrLf &
        "        }" & vbCrLf &
        "    }," & vbCrLf &
        "    ""dataTypes"": [##dataTypes##]," & vbCrLf &
        "    ""objectTypes"": [##objectTypes##]," & vbCrLf &
        "    ""relationTypes"": [##relationTypes##]," & vbCrLf &
        "    ""hierarchyTypes"": [##hierarchyTypes##]," & vbCrLf &
        "    ""objects"": [##objects##]," & vbCrLf &
        "    ""relations"": [##relations##]," & vbCrLf &
        "    ""hierarchies"": [##hierarchies##]," & vbCrLf &
        "    ""files"": [##files##]" & vbCrLf &
        "}"
        framework = Replace(framework, "##projectinfoID##", ProjectData.projectinfoID)
        framework = Replace(framework, "##projectinfoName##", ProjectData.projectinfoName)
        framework = Replace(framework, "##createdAt##", GetCreatedAt)
        framework = Replace(framework, "##infoFamilyName##", ProjectData.infoFamilyName)
        framework = Replace(framework, "##infoGivenName##", ProjectData.infoGivenName)
        framework = Replace(framework, "##infoOrganisationName##", ProjectData.infoOrganisationName)
        framework = Replace(framework, "##infoEmail##", ProjectData.infoEmail)

        ProjectData.MELHierarchieDS = "{" & vbCrLf &
        "        ""id"": ""SH-Fld-Modellelemente""," & vbCrLf &
        "        ""object"": ""Fld-Modellelemente""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "        ""nodes"": [{" & vbCrLf &
        "            ""id"": ""SH-Fld-Akteur""," & vbCrLf &
        "            ""object"": ""Fld-Akteur""," & vbCrLf &
        "            ""revision"": 0," & vbCrLf &
        "            ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "            ""nodes"": [##Akteur##]" & vbCrLf &
        "        },{" & vbCrLf &
        "            ""id"": ""SH-Fld-Zustand""," & vbCrLf &
        "            ""object"": ""Fld-Zustand""," & vbCrLf &
        "            ""revision"": 0," & vbCrLf &
        "            ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "            ""nodes"": [##Zustand##]" & vbCrLf &
        "        },{" & vbCrLf &
        "            ""id"": ""SH-Fld-Ereignis""," & vbCrLf &
        "            ""object"": ""Fld-Ereignis""," & vbCrLf &
        "            ""revision"": 0," & vbCrLf &
        "            ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
        "            ""nodes"": [##Ereignis##]" & vbCrLf &
        "        }]" & vbCrLf &
        "    }"
        Return framework
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: JoinMElHierarchie()
    '       Zweck        : Fügt Modellelemente in die Hierarchie zu.
    '       @param       : -
    '       Quellort     : Von MainSpecIFSub(modelPaths() As String)
    '    ---------------------------------------------------------------------------------------
    Public Sub JoinMElHierarchie()
        ProjectData.CompleteDS = Replace(ProjectData.CompleteDS, "##MELhierarchies##", ProjectData.MELHierarchieDS & "##MELhierarchies##")
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Sub-Anweisung: JoinSpecIFBaseWithProjectData()
    '       Zweck        : Füllt Daten aus ProjecData in das Integrationsmodell.
    '       @param       : -
    '       Quellort     : Von MainSpecIFSub(modelPaths() As String)
    '    ---------------------------------------------------------------------------------------
    Public Sub JoinSpecIFBaseWithProjectData()
        ProjectData.CompleteDS = Replace(ProjectData.CompleteDS, "##objects##", ProjectData.ObjekteDS & "##objects##")
        ProjectData.CompleteDS = Replace(ProjectData.CompleteDS, "##relations##", ProjectData.RelationenDS & "##relations##")
        'In Hierarchie können sich immer Reste befinden von den Zeichen ##. Erstmal alles entfernen.
        ProjectData.HierarchieDS = PrepareStringForExport(ProjectData.HierarchieDS)
        ProjectData.CompleteDS = Replace(ProjectData.CompleteDS, "##hierarchies##", ProjectData.HierarchieDS & "##hierarchies##")

        ProjectData.HierarchieDS = ""
        ProjectData.ObjekteDS = ""
        ProjectData.RelationenDS = ""
    End Sub
    '    ---------------------------------------------------------------------------------------
    '       Function     : GetdataTypes() As String
    '       Zweck        : Erstellt Datatype in einer JSON-Struktur.
    '       @param       : -
    '       @return      : String.
    '       Quellort     : Von CreateSpecIFBase()
    '    ---------------------------------------------------------------------------------------
    Private Function GetdataTypes() As String
        Dim dataTypes As String
        dataTypes = "{" & vbCrLf &
        "        ""id"": ""DT-ShortString""," & vbCrLf &
        "        ""title"": ""dcterms:title""," & vbCrLf &
        "        ""description"": ""Titel""," & vbCrLf &
        "        ""type"": ""xs:string""," & vbCrLf &
        "        ""maxLength"": 96," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""" & vbCrLf &
        "    }, {" & vbCrLf &
        "        ""id"": ""DT-FormattedText""," & vbCrLf &
        "        ""title"": ""dcterms:description""," & vbCrLf &
        "        ""description"": ""Descrition""," & vbCrLf &
        "        ""type"": ""xhtml""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""" & vbCrLf &
        "    }," & vbCrLf &
        "	{" & vbCrLf &
        "        ""id"": ""XLS_Text""," & vbCrLf &
        "        ""title"": ""XLS.Text""," & vbCrLf &
        "        ""description"": ""String with length 8192""," & vbCrLf &
        "        ""type"": ""xs:string""," & vbCrLf &
        "        ""maxLength"": 8192," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""" & vbCrLf &
        "    }"
        dataTypes = Replace(dataTypes, "##createdAt##", GetCreatedAt)
        Return dataTypes
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : GetobjectTypes() As String
    '       Zweck        : Erstellt ObjectType in einer JSON-Struktur.
    '       @param       : -
    '       @return      : String.
    '       Quellort     : Von CreateSpecIFBase()
    '    ---------------------------------------------------------------------------------------
    Private Function GetobjectTypes() As String
        Dim objectTypes As String
        objectTypes = "{" & vbCrLf &
        "        ""id"": ""OT-Fld""," & vbCrLf &
        "        ""title"": ""SpecIF:Heading""," & vbCrLf &
        "        ""description"": ""Kapitelüberschriften (oder Diagrammart)""," & vbCrLf &
        "        ""attributeTypes"": [{" & vbCrLf &
        "            ""id"": ""AT-Fld-Name""," & vbCrLf &
        "            ""title"": ""dcterms:title""," & vbCrLf &
        "            ""dataType"": ""DT-ShortString""," & vbCrLf &
        "            ""revision"": 0," & vbCrLf &
        "            ""changedAt"": ""##createdAt##""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""id"": ""AT-Fld-Text""," & vbCrLf &
        "            ""title"": ""dcterms:description""," & vbCrLf &
        "            ""dataType"": ""DT-FormattedText""," & vbCrLf &
        "            ""revision"": 0," & vbCrLf &
        "            ""changedAt"": ""##createdAt##""" & vbCrLf &
        "        }]," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""," & vbCrLf &
        "        ""creation"": [""manual""]" & vbCrLf &
        "    }, {" & vbCrLf &
        "        ""id"": ""OT-Pln""," & vbCrLf &
        "        ""title"": ""SpecIF:Diagram""," & vbCrLf &
        "        ""description"": ""Diagram""," & vbCrLf &
        "        ""icon"": ""&#9635;""," & vbCrLf &
        "        ""attributeTypes"": [{" & vbCrLf &
        "            ""id"": ""AT-Pln-Name""," & vbCrLf &
        "            ""title"": ""dcterms:title""," & vbCrLf &
        "            ""dataType"": ""DT-ShortString""," & vbCrLf &
        "            ""revision"": 0," & vbCrLf &
        "            ""changedAt"": ""##createdAt##""" & vbCrLf &
        "        }, {" & vbCrLf &
        "            ""id"": ""AT-Pln-Text""," & vbCrLf &
        "            ""title"": ""dcterms:description""," & vbCrLf &
        "            ""dataType"": ""DT-FormattedText""," & vbCrLf &
        "            ""revision"": 0," & vbCrLf &
        "            ""changedAt"": ""##createdAt##""" & vbCrLf &
        "        },{" & vbCrLf &
        "            ""id"": ""AT-Pln-Type""," & vbCrLf &
        "            ""title"": ""SpecIF:Type""," & vbCrLf &
        "            ""dataType"": ""DT-ShortString""," & vbCrLf &
        "            ""revision"": 0," & vbCrLf &
        "            ""changedAt"": ""##createdAt##""" & vbCrLf &
        "        }], " & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""," & vbCrLf &
        "        ""creation"": [""auto"",""manual""]" & vbCrLf &
        "    }, {" & vbCrLf &
        "		""id"": ""OT-Act""," & vbCrLf &
        "		""title"": ""FMC:Actor""," & vbCrLf &
        "		""description"": ""Ein &#9632; Akteur steht für ein aktives Objekt, z.B. eine Aktivität, ein Prozessschritt, eine Funktion, eine Systemkomponente oder eine Rolle.""," & vbCrLf &
        "		""icon"": ""&#9632;""," & vbCrLf &
        "		""attributeTypes"": [{" & vbCrLf &
        "			""id"": ""AT-Act-Name""," & vbCrLf &
        "			""title"": ""dcterms:title""," & vbCrLf &
        "			""dataType"": ""DT-ShortString""," & vbCrLf &
        "			""revision"": 0," & vbCrLf &
        "			""changedAt"": ""##createdAt##""" & vbCrLf &
        "		},{" & vbCrLf &
        "			""id"": ""AT-Act-Text""," & vbCrLf &
        "			""title"": ""dcterms:description""," & vbCrLf &
        "			""dataType"": ""DT-FormattedText""," & vbCrLf &
        "			""revision"": 0," & vbCrLf &
        "			""changedAt"": ""##createdAt##""" & vbCrLf &
        "		}]," & vbCrLf &
        "		""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""," & vbCrLf &
        "        ""creation"": [""auto"",""manual""]" & vbCrLf &
        "	},{" & vbCrLf &
        "		""id"": ""OT-Sta""," & vbCrLf &
        "		""title"": ""FMC:State""," & vbCrLf &
        "		""description"": ""Ein &#9679; Zustand repräsentiert ein passives Objekt, z.B. eine Form, einen Wert, eine Bedingung, einen Informationsspeicher oder eine physische Beschaffenheit.""," & vbCrLf &
        "		""icon"": ""&#9679;""," & vbCrLf &
        "		""attributeTypes"": [{" & vbCrLf &
        "			""id"": ""AT-Sta-Name""," & vbCrLf &
        "			""title"": ""dcterms:title""," & vbCrLf &
        "			""dataType"": ""DT-ShortString""," & vbCrLf &
        "			""revision"": 0," & vbCrLf &
        "			""changedAt"": ""##createdAt##""" & vbCrLf &
        "		},{" & vbCrLf &
        "			""id"": ""AT-Sta-Text""," & vbCrLf &
        "			""title"": ""dcterms:description""," & vbCrLf &
        "			""dataType"": ""DT-FormattedText""," & vbCrLf &
        "			""revision"": 0," & vbCrLf &
        "			""changedAt"": ""##createdAt##""" & vbCrLf &
        "		}]," & vbCrLf &
        "		""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""," & vbCrLf &
        "        ""creation"": [""auto"",""manual""]" & vbCrLf &
        "	},{" & vbCrLf &
        "		""id"": ""OT-Evt""," & vbCrLf &
        "		""title"": ""FMC:Event""," & vbCrLf &
        "		""description"": ""Ein &#9830; Ereignis bezeichnet eine zeitliche Referenz, eine Aenderung einer Bedingung bzw. eines Zustandes oder generell ein Signal zur Synchronisation.""," & vbCrLf &
        "		""icon"": ""&#9830;""," & vbCrLf &
        "		""attributeTypes"": [{" & vbCrLf &
        "			""id"": ""AT-Evt-Name""," & vbCrLf &
        "			""title"": ""dcterms:title""," & vbCrLf &
        "			""dataType"": ""DT-ShortString""," & vbCrLf &
        "			""revision"": 0," & vbCrLf &
        "			""changedAt"": ""##createdAt##""" & vbCrLf &
        "		},{" & vbCrLf &
        "			""id"": ""AT-Evt-Text""," & vbCrLf &
        "			""title"": ""dcterms:description""," & vbCrLf &
        "			""dataType"": ""DT-FormattedText""," & vbCrLf &
        "			""revision"": 0," & vbCrLf &
        "			""changedAt"": ""##createdAt##""" & vbCrLf &
        "		}]," & vbCrLf &
        "		""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""," & vbCrLf &
        "        ""creation"": [""auto"",""manual""]" & vbCrLf &
        "	},{" & vbCrLf &
        "		""id"": ""OT-Req""," & vbCrLf &
        "		""title"": ""SpecIF:Requirement""," & vbCrLf &
        "		""description"": ""Eine &#8623; Anforderung dokumentiert einzelnes physisches oder funktionales Bedürfnis, das der betreffende Entwurf, das Produkt oder der Prozess erfüllen muss.""," & vbCrLf &
        "		""icon"": ""&#8623;""," & vbCrLf &
        "		""attributeTypes"": [{" & vbCrLf &
        "			""id"": ""AT-Req-Name""," & vbCrLf &
        "			""title"": ""dcterms:title""," & vbCrLf &
        "			""dataType"": ""DT-ShortString""," & vbCrLf &
        "			""revision"": 0," & vbCrLf &
        "			""changedAt"": ""##createdAt##""" & vbCrLf &
        "		},{" & vbCrLf &
        "			""id"": ""AT-Req-Text""," & vbCrLf &
        "			""title"": ""dcterms:description""," & vbCrLf &
        "			""dataType"": ""DT-FormattedText""," & vbCrLf &
        "			""revision"": 0," & vbCrLf &
        "			""changedAt"": ""##createdAt##""" & vbCrLf &
        "		}]," & vbCrLf &
        "		""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""," & vbCrLf &
        "        ""creation"": [""manual""]" & vbCrLf &
        "	}"

        objectTypes = Replace(objectTypes, "##createdAt##", GetCreatedAt)
        Return objectTypes
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : GetrelationTypes() As String
    '       Zweck        : Erstellt RelationType in einer JSON-Struktur.
    '       @param       : -
    '       @return      : String.
    '       Quellort     : Von CreateSpecIFBase()
    '    ---------------------------------------------------------------------------------------
    Private Function GetrelationTypes() As String
        Dim relationTypes As String
        relationTypes = "{" & vbCrLf &
        "        ""id"": ""RT-FuehrtZu""," & vbCrLf &
        "        ""title"": ""führt zu""," & vbCrLf &
        "        ""description"": ""Relation: Zustand zu Transition oder von Akteur zu Akteur.""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""," & vbCrLf &
        "        ""creation"": [""auto"", ""manual""]," & vbCrLf &
        "        ""sourceTypes"": [""OT-Sta"", ""OT-Act""]," & vbCrLf &
        "        ""targetTypes"": [""OT-Act""]" & vbCrLf &
        "	},{" & vbCrLf &
        "        ""id"": ""RT-resultiertIn""," & vbCrLf &
        "        ""title"": ""resultiert in""," & vbCrLf &
        "        ""description"": ""Relation: Transition zu Zustand.""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""," & vbCrLf &
        "        ""creation"": [""auto"", ""manual""]," & vbCrLf &
        "        ""sourceTypes"": [""OT-Act""]," & vbCrLf &
        "        ""targetTypes"": [""OT-Sta""]" & vbCrLf &
        "	},{" & vbCrLf &
        "        ""id"": ""RT-Visibility""," & vbCrLf &
        "        ""title"": ""SpecIF:shows""," & vbCrLf &
        "        ""description"": ""Relation: Plan shows Model-Element.""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""," & vbCrLf &
        "        ""creation"": [""auto""]," & vbCrLf &
        "        ""sourceTypes"": [""OT-Pln""]," & vbCrLf &
        "        ""targetTypes"": [""OT-Act"", ""OT-Sta"", ""OT-Evt""]" & vbCrLf &
        "    },{" & vbCrLf &
        "	    ""id"": ""RT-loestAus""," & vbCrLf &
        "        ""title"": ""löst aus""," & vbCrLf &
        "        ""description"": ""Relation: Event zu Akteur.""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""," & vbCrLf &
        "        ""creation"": [""auto"", ""manual""]," & vbCrLf &
        "        ""sourceTypes"": [""OT-Evt""]," & vbCrLf &
        "        ""targetTypes"": [""OT-Act""]" & vbCrLf &
        "	},{" & vbCrLf &
        "	    ""id"": ""RT-ruftauf""," & vbCrLf &
        "        ""title"": ""ruft auf""," & vbCrLf &
        "        ""description"": ""Relation: Zustand zu Ereignis.""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""," & vbCrLf &
        "        ""creation"": [""auto"", ""manual""]," & vbCrLf &
        "        ""sourceTypes"": [""OT-Sta""]," & vbCrLf &
        "        ""targetTypes"": [""OT-Evt""]" & vbCrLf &
        "	},{" & vbCrLf &
        "	    ""id"": ""RT-Containment""," & vbCrLf &
        "        ""title"": ""SpecIF:contains""," & vbCrLf &
        "        ""description"": ""Relation: Element enthält ein Modellelement.""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""," & vbCrLf &
        "        ""creation"": [""auto"", ""manual""]," & vbCrLf &
        "        ""sourceTypes"": [""OT-Act"",""OT-Sta"",""OT-Evt""]," & vbCrLf &
        "        ""targetTypes"": [""OT-Act"",""OT-Sta"",""OT-Evt""]" & vbCrLf &
        "	},{" & vbCrLf &
        "	    ""id"": ""RT-Dependency""," & vbCrLf &
        "        ""title"": ""SpecIF:dependsOn""," & vbCrLf &
        "        ""description"": ""Relation: Anforderung hängt von Anforderung ab.""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""," & vbCrLf &
        "        ""creation"": [""auto"", ""manual""]," & vbCrLf &
        "        ""sourceTypes"": [""OT-Req""]," & vbCrLf &
        "        ""targetTypes"": [""OT-Req""]" & vbCrLf &
        "	},{" & vbCrLf &
        "	    ""id"": ""RT-Satisfaction""," & vbCrLf &
        "        ""title"": ""oslc_rm:satisfies""," & vbCrLf &
        "        ""description"": ""Relation: Modelelement erfüllt Anforderung. ""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""," & vbCrLf &
        "        ""creation"": [""auto"", ""manual""]," & vbCrLf &
        "        ""sourceTypes"": [""OT-Act"",""OT-Sta"",""OT-Evt""]," & vbCrLf &
        "        ""targetTypes"": [""OT-Req""]" & vbCrLf &
        "	}"
        relationTypes = Replace(relationTypes, "##createdAt##", GetCreatedAt)
        Return relationTypes
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : GethierarchyTypes() As String
    '       Zweck        : Erstellt HierarchyType in einer JSON-Struktur.
    '       @param       : -
    '       @return      : String.
    '       Quellort     : Von CreateSpecIFBase()
    '    ---------------------------------------------------------------------------------------
    Private Function GethierarchyTypes() As String
        Dim hierarchyTypes As String
        hierarchyTypes = "{" & vbCrLf &
        "        ""id"": ""HT-SpecIF_Outline""," & vbCrLf &
        "        ""title"": ""SpecIF:Outline""," & vbCrLf &
        "        ""description"": ""Hierarchy type for outlines""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""" & vbCrLf &
        "    }"
        hierarchyTypes = Replace(hierarchyTypes, "##createdAt##", GetCreatedAt)
        Return hierarchyTypes
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : Getobjects() As String
    '       Zweck        : Erstellt Objekte in einer JSON-Struktur.
    '       @param       : -
    '       @return      : String.
    '       Quellort     : Von CreateSpecIFBase()
    '    ---------------------------------------------------------------------------------------
    Private Function Getobjects() As String
        Dim objects As String
        objects = "{" & vbCrLf &
"        ""id"": ""Fld-Modellelemente""," & vbCrLf &
"        ""title"": ""Modellelemente""," & vbCrLf &
"        ""attributes"": [{" & vbCrLf &
"            ""title"": ""dcterms:title""," & vbCrLf &
"            ""attributeType"": ""AT-Fld-Name""," & vbCrLf &
"            ""value"": ""Modellelemente""" & vbCrLf &
"        }, {" & vbCrLf &
"            ""title"": ""dcterms:description""," & vbCrLf &
"            ""attributeType"": ""AT-Fld-Text""," & vbCrLf &
"            ""value"": ""<div>Es wird in drei Modellelement-Typen gegliedert: &#9632; Akteur, &#9679; Zustand und &#9830; Ereignis. Damit ist es möglich bei der Überführung von verschiedenen Modellen ein Modellkern mit wenigen Modellelement-Typen zu beschreiben. Dieses basiert auf der Arbeit des Prof. Dr. Siegfried Wendt. </div>""" & vbCrLf &
"        }]," & vbCrLf &
"        ""revision"": 0," & vbCrLf &
"        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
"        ""objectType"": ""OT-Fld""" & vbCrLf &
"    },{" & vbCrLf &
"        ""id"": ""Fld-Akteur""," & vbCrLf &
"        ""title"": ""Akteur""," & vbCrLf &
"        ""attributes"": [{" & vbCrLf &
"            ""title"": ""dcterms:title""," & vbCrLf &
"            ""attributeType"": ""AT-Fld-Name""," & vbCrLf &
"            ""value"": ""Akteur""" & vbCrLf &
"        }, {" & vbCrLf &
"            ""title"": ""dcterms:description""," & vbCrLf &
"            ""attributeType"": ""AT-Fld-Text""," & vbCrLf &
"            ""value"": ""<div><p>&#9632; Akteur: Ist ein aktives Element, wie eine Aktivität, Funktion, Prozess-Schritt, Systemkomponente oder eine Rolle.</p></div>""" & vbCrLf &
"        }]," & vbCrLf &
"        ""revision"": 0," & vbCrLf &
"        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
"        ""objectType"": ""OT-Fld""" & vbCrLf &
"    },{" & vbCrLf &
"        ""id"": ""Fld-Zustand""," & vbCrLf &
"        ""title"": ""Zustand""," & vbCrLf &
"        ""attributes"": [{" & vbCrLf &
"            ""title"": ""dcterms:title""," & vbCrLf &
"            ""attributeType"": ""AT-Fld-Name""," & vbCrLf &
"            ""value"": ""Zustand""" & vbCrLf &
"        }, {" & vbCrLf &
"            ""title"": ""dcterms:description""," & vbCrLf &
"            ""attributeType"": ""AT-Fld-Text""," & vbCrLf &
"            ""value"": ""<div><p>&#9679; Zustand: Repräsentiert ein passives Element, wie eine Form, Wert, Informationsspeicher, Bedingung, oder eine physische Beschaffenheit.</p></div>""" & vbCrLf &
"        }]," & vbCrLf &
"        ""revision"": 0," & vbCrLf &
"        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
"        ""objectType"": ""OT-Fld""" & vbCrLf &
"    },{" & vbCrLf &
"        ""id"": ""Fld-Ereignis""," & vbCrLf &
"        ""title"": ""Ereignis""," & vbCrLf &
"        ""attributes"": [{" & vbCrLf &
"            ""title"": ""dcterms:title""," & vbCrLf &
"            ""attributeType"": ""AT-Fld-Name""," & vbCrLf &
"            ""value"": ""Ereignis""" & vbCrLf &
"        }, {" & vbCrLf &
"            ""title"": ""dcterms:description""," & vbCrLf &
"            ""attributeType"": ""AT-Fld-Text""," & vbCrLf &
"            ""value"": ""<div><p>&#9830; Ereignis: Bezeichnet eine zeitliche Referenz. Das kann eine Änderung einer Bedingung bzw. eines Zustandes oder generell ein Signal zur Synchronisation sein.</p></div>""" & vbCrLf &
"        }]," & vbCrLf &
"        ""revision"": 0," & vbCrLf &
"        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
"        ""objectType"": ""OT-Fld""" & vbCrLf &
"    },{" & vbCrLf &
"        ""id"": ""Fld-ProjectInformation""," & vbCrLf &
"        ""title"": ""Projekt Information""," & vbCrLf &
"        ""attributes"": [{" & vbCrLf &
"            ""title"": ""dcterms:title""," & vbCrLf &
"            ""attributeType"": ""AT-Fld-Name""," & vbCrLf &
"            ""value"": ""Projekt Information""" & vbCrLf &
"        }, {" & vbCrLf &
"            ""title"": ""dcterms:description""," & vbCrLf &
"            ""attributeType"": ""AT-Fld-Text""," & vbCrLf &
"            ""value"": ""<div><p>" & GetProjectInformation() & "</p></div>""" & vbCrLf &
"        }]," & vbCrLf &
"        ""revision"": 0," & vbCrLf &
"        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
"        ""objectType"": ""OT-Fld""" & vbCrLf &
"    },{" & vbCrLf &
"        ""id"": ""Fld-Systemmodell""," & vbCrLf &
"        ""title"": ""Systemmodell""," & vbCrLf &
"        ""attributes"": [{" & vbCrLf &
"            ""title"": ""dcterms:title""," & vbCrLf &
"            ""attributeType"": ""AT-Fld-Name""," & vbCrLf &
"            ""value"": ""Systemmodell""" & vbCrLf &
"        }, {" & vbCrLf &
"            ""title"": ""dcterms:description""," & vbCrLf &
"            ""attributeType"": ""AT-Fld-Text""," & vbCrLf &
"            ""value"": ""<div><p>Ein Systemmodell ist das Abbild eines komplexen Systems. Durch die Reduktion der Komplexität und der Berücksichtigung von relevanten Attributen, die für den definierten Zweck von Bedeutung sind, können Vorgänge und Funktionen des Systems dargestellt und verstanden werden. </p></div>""" & vbCrLf &
"        }]," & vbCrLf &
"        ""revision"": 0," & vbCrLf &
"        ""changedAt"": """ & GetCreatedAt() & """," & vbCrLf &
"        ""objectType"": ""OT-Fld""" & vbCrLf &
"    },"
        Return objects
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : GetProjectInformation() As String
    '       Zweck        : Projektinformationen.
    '       @param       : -
    '       @return      : String.
    '       Quellort     : Von Getobjects()
    '    ---------------------------------------------------------------------------------------
    Private Function GetProjectInformation() As String
        Return ProjectData.projectInformation
    End Function
    '    ---------------------------------------------------------------------------------------
    '       Function     : Gethierarchies() As String
    '       Zweck        : Erstellt Hierarchie in einer JSON-Struktur..
    '       @param       : -
    '       @return      : String.
    '       Quellort     : Von CreateSpecIFBase()
    '    ---------------------------------------------------------------------------------------
    Private Function Gethierarchies() As String
        Dim hierarchies As String
        hierarchies = "{" & vbCrLf &
        "        ""id"": ""SP-Knotenpunkt""," & vbCrLf &
        "        ""title"": ""Systemmodell: ##projectName##""," & vbCrLf &
        "        ""description"": ""Entwicklung eines Integrationsmodells zur Überführung von Systemmodellen.""," & vbCrLf &
        "        ""revision"": 0," & vbCrLf &
        "        ""changedAt"": ""##createdAt##""," & vbCrLf &
        "        ""hierarchyType"": ""HT-SpecIF_Outline""," & vbCrLf &
        "        ""nodes"": [{" & vbCrLf &
        "           ""id"": ""SH-Fld-ProjectInformation""," & vbCrLf &
        "           ""object"": ""Fld-ProjectInformation""," & vbCrLf &
        "           ""revision"": 0," & vbCrLf &
        "           ""changedAt"": ""##createdAt##""," & vbCrLf &
        "           ""nodes"": []" & vbCrLf &
        "        },{" & vbCrLf &
        "           ""id"": ""SH-Fld-Systemmodell""," & vbCrLf &
        "           ""object"": ""Fld-Systemmodell""," & vbCrLf &
        "           ""revision"": 0," & vbCrLf &
        "           ""changedAt"": ""##createdAt##""," & vbCrLf &
        "           ""nodes"": [##hierarchies##]" & vbCrLf &
        "    },##MELhierarchies##]" & vbCrLf &
        "	}"
        hierarchies = Replace(hierarchies, "##createdAt##", GetCreatedAt)
        hierarchies = Replace(hierarchies, "##projectName##", ProjectData.projectName)
        Return hierarchies
    End Function
End Module
