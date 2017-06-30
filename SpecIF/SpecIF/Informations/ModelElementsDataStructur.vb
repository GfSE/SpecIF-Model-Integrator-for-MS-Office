'    ---------------------------------------------------------------------------------------
'       Class   : ModelElementsDataStructur
'       Datum   : 29.04.2017
'       Author  : Philipp Mochine
'       Zweck   : Diese Klasse dient als Datenstruktur der gesamten Modellelementen. In wird in einer Liste geführt.
'                 Gespeichert werden Modellelementname, ID und (Diagramm-)typ.                      
'    ---------------------------------------------------------------------------------------
Public Class ModelElementsDataStructur
    Private melName As String
    Private melID As String
    Private melType As String

    Public Sub New(ByVal Name As String, ByVal ID As String, ByVal Type As String)
        melName = Name
        melID = ID
        melType = Type
    End Sub
    Property Name() As String
        Get
            Return melName
        End Get
        Set(value As String)
            melName = value
        End Set
    End Property
    Property ID() As String
        Get
            Return melID
        End Get
        Set(value As String)
            melID = value
        End Set
    End Property
    Property Type() As String
        Get
            Return melType
        End Get
        Set(value As String)
            melType = value
        End Set
    End Property
End Class