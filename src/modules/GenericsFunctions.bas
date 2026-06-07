Attribute VB_Name = "GenericsFunctions"
Public toolkitVersion As String
Public toolkitReleaseDate As String
Public openSourceLicence As String

Public gitRepo As String
Public gitIssues As String
Public gitHelpTxt As String
Public gitHelpCSV As String
Public gitMainContributor As String
Public gitMajorContributors As String
'--------------------------------------------------------
' Function      : InitializeToolkit
' Author        : BeeniGit
' Date          : 15/05/2025
' Version       : 1.0
' History       :
'
' Description :
'   Initialize the public variables
'
' Parameters :
'   N/A
'
' Output :
'   N/A
'
' Example :
'   N/A
'
' Notes :
'   All of this information must be verified and updated with each new version of the toolkit.
'--------------------------------------------------------
Public Sub InitializeToolkit(show As Boolean)
    'Toolkit
    toolkitVersion = "1.0.0" 'Version format {majorUpdate}.{minorUpdate}.{Patch}
    toolkitReleaseDate = "07/06/2026" 'Date format DD/MM/YYYY
    
    'Git informations
    gitRepo = "https://github.com/BeeniGit/WordVBAToolkit"
    gitIssues = "/issues"
    gitHelpTxt = "/wiki/Tool-n°1-:-Text-Formatter"
    gitHelpCSV = "/wiki/Tool-n°2-:-Table-from-list"
    
    gitMainContributor = "BeeniGit"
    gitMajorContributors = ""
    
    'Open sour Licence
    openSourceLicence = "MIT"
End Sub

'--------------------------------------------------------
' Function      : GetColorFromName
' Author        : BeeniGit
' Date          : 15/05/2025
' Version       : 1.0
' History       :
'
' Description :
'   Function that returns the RBG code based on the name of the color passed as a parameter
'
' Parameters :
'   colorName (String) : Name of the color to convert
'
' Output :
'   Long : RGB value of the color passed as a parameter
'
' Example :
'   GetColorFromName("Red") return the RGB value (255, 0, 0) in Long format
'
' Notes :
'   In case of unknown color, the default RGB value will be Black.
'--------------------------------------------------------
Function GetColorFromName(colorName As String) As Long
    Select Case LCase(Trim(colorName))
        Case "black":   GetColorFromName = RGB(0, 0, 0)
        Case "blue":    GetColorFromName = RGB(0, 0, 255)
        Case "red":     GetColorFromName = RGB(255, 0, 0)
        Case "green":   GetColorFromName = RGB(0, 128, 0)
        Case "yellow":  GetColorFromName = RGB(255, 215, 0)
        Case "orange":  GetColorFromName = RGB(255, 128, 0)
        Case "brown":   GetColorFromName = RGB(165, 42, 42)
        Case "purple":  GetColorFromName = RGB(128, 0, 128)
        Case "gray":    GetColorFromName = RGB(128, 128, 128)
        Case Else:      GetColorFromName = RGB(0, 0, 0)
    End Select
End Function

'--------------------------------------------------------
' Function      : GetLineWidth
' Author        : BeeniGit
' Date          : 15/05/2025
' Version       : 1.0
'
' Description :
'   Function that returns the Word line width based on the thickness passed as a parameter.
'
' Parameters :
'   thicknessName (String) : Selected border width name
'
' Output :
'   WdLineWidth : Word object that define the line width
'
' Example :
'   GetLineWidth("Medium (1.5 pt)") return the Word value wdLineWidth150pt
'
' Notes :
'   In case of unknown parameter, the default value will be 0,75pt
'--------------------------------------------------------
Function GetLineWidth(thicknessName As String) As WdLineWidth
    Select Case thicknessName
        Case "Thin (0.25 pt)":      GetLineWidth = wdLineWidth025pt
        Case "Normal (0.75 pt)":    GetLineWidth = wdLineWidth075pt
        Case "Medium (1.5 pt)":     GetLineWidth = wdLineWidth150pt
        Case "Bold (2.25 pt)":      GetLineWidth = wdLineWidth225pt
        Case "Very Bold (3 pt)":    GetLineWidth = wdLineWidth300pt
        Case Else:                  GetLineWidth = wdLineWidth075pt
    End Select
End Function

'--------------------------------------------------------
' Function      : GetBorderWidthInPoints
' Author        : BeeniGit
' Date          : 03/05/2026
' Version       : 1.0
'
' Description :
'   Function that returns the number of points  Word line width based on the thickness passed as a parameter.
'
' Parameters :
'   width (WdLineWidth) : Conversion needed for this linewidth
'
' Output :
'   Double : points value according to the line witdh parameter
'
' Example :
'   wdLineWidth025pt return the Word value 0.25
'
' Notes :
'   In case of unknown parameter, the default value will be 0,75pt
'--------------------------------------------------------
Function GetBorderWidthInPoints(width As WdLineWidth) As Double

    Select Case width
        Case wdLineWidth025pt: GetBorderWidthInPoints = 0.25
        Case wdLineWidth075pt: GetBorderWidthInPoints = 0.75
        Case wdLineWidth150pt: GetBorderWidthInPoints = 1.5
        Case wdLineWidth225pt: GetBorderWidthInPoints = 2.25
        Case wdLineWidth300pt: GetBorderWidthInPoints = 3
        Case Else: GetBorderWidthInPoints = 0.75 ' fallback
    End Select

End Function

'--------------------------------------------------------
' Function      : GetLineStyleFromName
' Author        : BeeniGit
' Date          : 15/05/2025
' Version       : 1.0
' History    :
'
' Description :
'   Function that returns the Word line style based on the parameter.
'
' Parameters :
'   styleName (String) : Line style
'
' Output :
'   WdLineStyle : Word object that define the line style
'
' Example :
'   GetLineStyleFromName("Double line") return the value wdLineStyleDouble
'
' Notes :
'   In case of unknown parameter, the default value will be single line
'--------------------------------------------------------
Function GetLineStyleFromName(styleName As String) As WdLineStyle
    Select Case styleName
        Case "Single line":         GetLineStyleFromName = wdLineStyleSingle
        Case "Double line":         GetLineStyleFromName = wdLineStyleDouble
        Case "Dotted line":         GetLineStyleFromName = wdLineStyleDot
        Case "Dash line":           GetLineStyleFromName = wdLineStyleDashLargeGap
        Case "Double-dash line":    GetLineStyleFromName = wdLineStyleDashDotDot
        Case Else:                  GetLineStyleFromName = wdLineStyleSingle
    End Select
End Function

'--------------------------------------------------------
' Function      : GetDelimiterFromName
' Author        : BeeniGit
' Date          : 04/06/2026
' Version       : 1.0
' History    :
'
' Description :
'   Function that returns the word delimiter based on the description name
'
' Parameters :
'   delimiterName (String) : Delimiter name.
'
' Output :
'   GetDelimiterFromName : Delimiter character based on the delimiter name
'
' Example :
'   NA
'
' Notes :
'   NA
'--------------------------------------------------------
Function GetDelimiterFromName(delimiterName As String) As String
    Select Case delimiterName
        Case "Tabulation":      GetDelimiterFromName = vbTab
        Case "Semicolons ;":    GetDelimiterFromName = ";"
        Case "Commas":          GetDelimiterFromName = ","
        Case Else:              GetDelimiterFromName = ";" ' Default fallback
    End Select
End Function

'--------------------------------------------------------
' Function      : ShuffleArray
' Author        : BeeniGit
' Date          : 05/07/2025
' Version       : 1.0
' History    :
'
' Description :
'   Suffle a array with the "Randomize" Algorithm
'
' Parameters :
'   arr (String) : Initial array containing several words or characters.
'
' Output :
'   arr : Mixed array using the Fisher-Yates algorithm
'
' Example :
'   NA
'
' Notes :
'   NA
'--------------------------------------------------------
Sub ShuffleArray(arr() As String)
    Dim i As Long, j As Long
    Dim temp As String
    Randomize
    For i = UBound(arr) To LBound(arr) + 1 Step -1
        j = Int((i - LBound(arr) + 1) * Rnd + LBound(arr))
        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp
    Next i
End Sub

'--------------------------------------------------------
' Function      : ErrorMessageDisplay
' Author        : BeeniGit
' Date          : 07/06/2026
' Version       : 1.0
' History    :
'
' Description :
'   Generic function that display error messages to the user
'
' Parameters :
'   errorID (String) : Error ID
'
' Output :
'   MsgBox
'
' Example :
'   NA
'
' Notes :
'   NA
'--------------------------------------------------------
Sub ErrorMessageDisplay(errorID As String)
    Select Case errorID
        Case "CSV":             MsgBox "Please select a CSV file.", vbExclamation
        Case "NoWordsCSV":      MsgBox "No words found in the CSV file.", vbExclamation
        Case "TableNumeric":    MsgBox "Please enter numeric values for the columns and rows.", vbExclamation
        Case "WordsNumeric":    MsgBox "Please enter numeric values for the words column.", vbExclamation
        Case "TableMin":        MsgBox "Please enter at least 1 columns or rows.", vbExclamation
        Case "TableMax":        MsgBox "Please enter fewer than 30 columns or rows.", vbExclamation
        Case "WordsMin":        MsgBox "Please enter at least 1 columns for the words positions.", vbExclamation
        Case "WordsMax":        MsgBox "Please enter fewer than 30 columns for the words positions.", vbExclamation
    End Select
End Sub

