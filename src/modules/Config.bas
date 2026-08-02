Attribute VB_Name = "Config"
'--------------------------------------------------------
' Public constant for toolkit

'---- Toolkit informations ----
' release
Public Const TOOLKIT_VERSION As String = "v1.2.0" 'Version format v{majorUpdate}.{minorUpdate}.{Patch}
Public Const TOOLKIT_RELEASE_DATE As String = "03/08/2026" 'Date format DD/MM/YYYY
    
' Git informations
Public Const GIT_REPO As String = "https://github.com/BeeniGit/WordVBAToolkit"
Public Const GIT_ISSUES As String = "/issues"
Public Const GIT_HELP_TXT As String = "/wiki/Tool-n°1-:-Text-Formatting"
Public Const GIT_HELP_WORDS As String = "/wiki/Tool-n°2-:-Table-from-list"

Public Const GIT_REPO_UPDATE As String = "https://api.github.com/repos/BeeniGit/WordVBAToolkit/releases/latest"

Public Const GIT_MAIN_CONTRIBUTOR As String = "BeeniGit"
Public Const GIT_MAJOR_CONTRIBUTORS As String = ""

' Open source Licence
Public Const OPEN_SOURCE_LICENCE = "MIT"

'---- TableFromList ----
' Table definition
Public Const TABLE_COLROW_MIN  As Long = 1
Public Const TABLE_COLROW_MAX As Long = 30
Public Const TABLE_COL_DEF_VALUE  As Long = 3
Public Const TABLE_ROW_DEF_VALUE  As Long = 5


' Words positions
Public Const WORDS_COL_MIN As Long = 1
Public Const WORDS_COL_MAX As Long = 20
Public Const WORDS_COL_DEF_VALUE As Long = 1
