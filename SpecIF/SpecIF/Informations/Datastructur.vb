'    ---------------------------------------------------------------------------------------
'       Class   : Datastructur
'       Datum   : 29.04.2017
'       Author  : Philipp Mochine
'       Zweck   : Diese Klasse dient als Datenstruktur des gesamten Projektes. Pro Projekt wird eine neue Klasse angelegt.
'                 Es werden alle möglichen Daten gespeichert. Beispiele sind: Ordner und Namensposition der Modelle,
'                 die Strings für Objekte, Relationen und Hierarchie, die später in richtigem Format als SpecIF abgespeichert werden                      
'       Quellort: Von MainSpecIFSub(modelPaths() As String)
'    ---------------------------------------------------------------------------------------
Public Class Datastructur
    'Informationen
    Public infoFamilyName As String
    Public infoGivenName As String
    Public infoOrganisationName As String
    Public infoEmail As String
    'Projekt Information
    Public projectinfoName As String 'Mit Foldername
    Public projectName As String
    Public projectinfoID As String
    Public projectFullPath As String
    Public projectInformation As String
    'Modell Informationen
    Public modelName() As String
    'specIF Objekte
    Private specIFComplete As String 'Das komplette Integrationsmodell befindet sich hier.
    Private specIFObjekte As String
    Private specIFRelationen As String
    Private specIFHierarchie As String
    Private specIFMELHierarchie As String 'Hierarchie der Modellelemente.

    Property CompleteDS() As String
        Get
            Return specIFComplete
        End Get
        Set(value As String)
            specIFComplete = value
        End Set
    End Property
    Property ObjekteDS() As String
        Get
            Return specIFObjekte
        End Get
        Set(value As String)
            specIFObjekte = value
        End Set
    End Property

    Property RelationenDS() As String
        Get
            Return specIFRelationen
        End Get
        Set(value As String)
            specIFRelationen = value
        End Set
    End Property

    Property HierarchieDS() As String
        Get
            Return specIFHierarchie
        End Get
        Set(value As String)
            specIFHierarchie = value
        End Set
    End Property
    Property MELHierarchieDS() As String
        Get
            Return specIFMELHierarchie
        End Get
        Set(value As String)
            specIFMELHierarchie = value
        End Set
    End Property

    Public Sub New()
        ReDim modelName(0)
    End Sub
    Public Sub ReFreshData()
        modelName = Nothing
        projectFullPath = Nothing
        specIFComplete = Nothing
        specIFObjekte = Nothing
        specIFRelationen = Nothing
        specIFHierarchie = Nothing
        specIFMELHierarchie = Nothing

        ReDim ProjectData.modelName(0)
    End Sub
End Class


