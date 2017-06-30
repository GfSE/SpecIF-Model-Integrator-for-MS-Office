'    ---------------------------------------------------------------------------------------
'       Modul   : Settings
'       Datum   : 29.04.2017
'       Author  : Philipp Mochine
'       Zweck   : Einstellungen für die Anwendung.                  
'    ---------------------------------------------------------------------------------------
Module Settings
    'Setting für Interface.vb openFileDialog
    Public modelsInitialDirectorySettings As String = AppDomain.CurrentDomain.BaseDirectory 'Ort der Anwendung
    Public Const modelsFilterSettings As String = "Requirements (*.xlsx)|*.xlsx|All Model Files (*.vsdx;*.vsdm;*.xlsx)|*.vsdx;*.vsdm;*.xlsx|Visio Files (*.vsdx;*.vsdm)|*.vsdx;*.vsdm"
    'Setting für MainSpecIF.vb für die Methode ExecuteModel
    Public ReadOnly Property ExcelExtension As String() = New String() {".xlsx", ".xlsm", ".xlsb", ".xls"}
    Public ReadOnly Property VisioExtension As String() = New String() {".vsd", ".vsdx", ".vsdm"}
    'Setting für Function.vb für die Methode FindModelType. Shapenamen aus NameU
    Public ReadOnly Property StateElementTypes As String() = New String() {"Diagram Overview", "Submachine state", "State with internal behavior", "State", "Dynamic connector", "Initial state", "Final state"}
    Public ReadOnly Property ActivityElementTypes As String() = New String() {"Diagram Overview", "Initial node", "Final node", "Action", "Dynamic connector", "Fork node"}
    Public ReadOnly Property SystemstructurElementTypes As String() = New String() {"Package (expanded)", "Interface", "Composition", "Dynamic connector"}
    Public ReadOnly Property FunctionstructurElementTypes As String() = New String() {"Diagram Overview", "Action", "Dynamic connector"}
    'Setting für Function.vb für die Methode GetCreatedAt.
    Public Const timeFilterSettings As String = "yyyy-MM-dd\Thh:mm:ssK"
    'Setting für Function.vb für die Methode CreateFolder.
    Public Const filesandimagesSettings As String = "\files_and_images"
    'Settings für ProjektData
    Public Const pdEmail As String = "sysML@specIF.de"
    Public Const pdFamilyName As String = "SpecIF"
    Public Const pdGivenName As String = "SysML"
    Public Const pdOrganisationName As String = "SysMLWithSpecIF"

End Module
