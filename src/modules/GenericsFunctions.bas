Attribute VB_Name = "GenericsFunctions"
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
' Date          : 23/07/2026
' Version       : 1.1
' History       : 1.1 - 23/07/2026 => Add the error "NoConnection" from the AboutForm. Add optionnal value for Table limit.
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
Sub ErrorMessageDisplay(errorID As String, Optional optionnalValue As Variant)
    Select Case errorID
        Case "NoCharacter":      MsgBox "No characters in the field", vbExclamation, "Formatting error"
        Case "CSV":             MsgBox "Please select a CSV file.", vbExclamation, "File error"
        Case "NoWordsCSV":      MsgBox "No words found in the CSV file.", vbExclamation, "File error"
        Case "TableNumeric":    MsgBox "Please enter numeric values for the columns and rows.", vbExclamation, "Table error"
        Case "WordsNumeric":    MsgBox "Please enter numeric values for the words column.", vbExclamation, "Words position error"
        Case "TableMin":        MsgBox "Please enter at least " & optionnalValue & " columns or rows.", vbExclamation, "Table error"
        Case "TableMax":        MsgBox "Please enter fewer than " & optionnalValue & " columns or rows.", vbExclamation, "Table error"
        Case "WordsMin":        MsgBox "Please enter at least " & optionnalValue & " columns for the words positions.", vbExclamation, "Words position error"
        Case "WordsMax":        MsgBox "Please enter fewer than " & optionnalValue & " columns for the words positions.", vbExclamation, "Words position error"
        Case "NoConnection":    MsgBox "Can't open the URL, check your connexion or the URL", vbExclamation, "URL error"
    End Select
End Sub

'--------------------------------------------------------
' Function      : FinishMessageDisplay
' Author        : BeeniGit
' Date          : 23/07/2026
' Version       : 1.0
' History       :
'
' Description :
'   Generic function that display finish messages to the user
'
' Parameters :
'   finishID (String) : finish ID
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
Sub FinishMessageDisplay(finishID As String, Optional optionnalValue1 As Variant, Optional optionnalValue2 As Variant)
    Select Case finishID
        Case "TextFormatting":  MsgBox "Formatting done.", vbInformation
        Case "WordsInsertion":  MsgBox "Insertion complete: " & optionnalValue1 & " words across " & optionnalValue2 & " pages.", vbInformation
    End Select
End Sub
