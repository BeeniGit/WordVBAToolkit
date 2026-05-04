Attribute VB_Name = "TextFormatting"
'--------------------------------------------------------
' Function      : LaunchSearchWithFormatting
' Author        : BeeniGit
' Date          : 15/05/2025
' Version       : 1.0
' History    :
'
' Description :
'   Function to search for a string in the text and change its formatting.
'
' Parameters :
'   userInput           (String)    : Search string
'   fontColorName       (String)    : Font color
'   fontSizeStr         (String)    : Font size
'   isBold              (Boolean)   : Bold formatting
'   isItalic            (Boolean)   : Italic formatting
'
' Output :
'   Search for the string in the word document and apply the desired formatting.
'
' Example :
'   N/A
'
' Notes :
'   N/A
'--------------------------------------------------------
Sub LaunchSearchWithFormatting( _
    userInput As String, _
    fontColorName As String, _
    fontSizeStr As String, _
    isBold As Boolean, _
    isItalic As Boolean, _
    isUnderlined As Boolean, _
    matchWord As Boolean)
    
    ' ====== Variable declaration ======
    Dim searchTerms() As String
    Dim term As Variant
    Dim rng As Range
    Dim doc As Document
    Dim fontColor As Long
    
    Set doc = ActiveDocument
    fontColor = GetColorFromName(fontColorName)
    searchTerms = Split(userInput, ";")     ' Get the user character string

    ' Search loop to find the user character string
    For Each term In searchTerms
        term = Trim(term)                   ' Delete the space
        
        ' Check if the string is empty
        If term <> "" Then
            Set rng = doc.Content
            With rng.Find
                .ClearFormatting            ' Delete previous formmating
                .Text = term                ' Search string
                .MatchCase = False          ' Take in consideation the text case (uppercase or lowercase)
                .MatchWholeWord = matchWord ' Take in consideration the entire word
                .MatchWildcards = False     ' No use of wildcards
                .Wrap = wdFindStop          ' The search stops at the end of the document
                
                Do While .Execute
                
                    ' Applying the selected formatting
                    With rng.Font
                        .Bold = isBold
                        .Italic = isItalic
                        .Underline = IIf(isUnderline, wdUnderlineSingle, wdUnderlineNone)
                        .color = fontColor
                        .size = val(fontSizeStr)
                    End With
                    rng.Collapse wdCollapseEnd
                Loop
            End With
        End If
    Next term
    
    MsgBox "Formatting done."
    
    ' Close the form
    Unload Textform
    
End Sub

