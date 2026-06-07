Attribute VB_Name = "WordListToTable"
'--------------------------------------------------------
' Function      : FillCell
' Author        : BeeniGit
' Date          : 02/05/2026
' Version       : 1.0
' History       :
'
' Description :
'   Subroutine that fills a Word table cell with formatted text.
'
' Parameters :
'   rng             (Range)    : Target cell range
'   txt             (String)   : Text to insert in the cell
'   fontColor       (Long)     : Font color (RGB or WdColor)
'   fontSize        (Single)   : Font size
'   isBold          (Boolean)  : Bold formatting flag
'   isItalic        (Boolean)  : Italic formatting flag
'   isUnderlined    (Boolean)  : Underline formatting flag
'
' Output :
'   N/A
'
' Example :
'   FillCell cellRange, "Hello", vbBlack, 12, True, False, False
'
' Notes :
'
'--------------------------------------------------------
Sub FillCell( _
    rng As Range, _
    txt As String, _
    fontColor As Long, _
    fontSize As Single, _
    isBold As Boolean, _
    isItalic As Boolean, _
    isUnderlined As Boolean)

    With rng

        ' Insert the text value
        .Text = txt

        ' Remove trailing paragraph mark if present (Word default behavior)
        If Right(.Text, 1) = Chr(13) Then
            .Text = Left(.Text, Len(.Text) - 1)
        End If

        ' Font formatting
        .Font.color = fontColor
        .Font.size = fontSize
        .Font.Bold = isBold
        .Font.Italic = isItalic
        .Font.Underline = isUnderlined

        ' Cell alignment (horizontal + vertical)
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cells.VerticalAlignment = wdCellAlignVerticalCenter

        ' Remove extra spacing inside cell
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = wdLineSpaceSingle
        End With

    End With

End Sub

'--------------------------------------------------------
' Function      : CreateOnePageTable
' Author        : BeeniGit
' Date          : 02/05/2026
' Version       : 1.0
' History       :
'
' Description :
'   Creates a single Word table representing one page of content
'   and fills it with words in a left-to-right or right-to-left order.
'
' Parameters :
'   doc                 (Document)   : Active Word document
'   rng                 (Range)      : Insertion range for the table
'   Words()             (String)     : Array of words to insert
'   startIndex          (Long)       : Starting index in Words array
'   colCount            (Long)       : Number of columns
'   rowCount            (Long)       : Number of rows
'   borderColor         (Long)       : Table border color
'   borderWidth         (WdLineWidth): Border thickness
'   borderStyle         (WdLineStyle): Border style
'   fontColor           (Long)       : Text color
'   fontSize            (Single)     : Font size
'   isBold              (Boolean)    : Bold formatting flag
'   isItalic            (Boolean)    : Italic formatting flag
'   isUnderlined        (Boolean)    : Underline formatting flag
'   fillRightToLeft     (Boolean)    : Fill direction flag
'
' Output :
'   Table (Table) : Created and filled Word table
'
' Example :
'   Set tbl = CreateOnePageTable(doc, rng, words, 1, 5, 10, vbBlack, wdLineWidth050pt, wdLineStyleSingle, vbBlack, 12, False, False, False, False)
'
' Notes :
'   Table height is calculated to fit exactly one page.
'--------------------------------------------------------
Function CreateOnePageTable( _
    doc As Document, _
    rng As Range, _
    Words() As String, _
    startIndex As Long, _
    colCount As Long, _
    rowCount As Long, _
    borderColor As Long, _
    borderWidth As WdLineWidth, _
    BorderStyle As WdLineStyle, _
    fontColor As Long, _
    fontSize As Single, _
    isBold As Boolean, _
    isItalic As Boolean, _
    isUnderlined As Boolean, _
    fillRightToLeft As Boolean) As Table
    
    ' ====== Variable declaration ======
    Dim tbl As Table
    Dim row As Long, col As Long
    Dim i As Long
    Dim usableHeight As Single
    Dim rowHeight As Single

    i = startIndex

    ' Create table at insertion range
    Set tbl = doc.Tables.Add(rng, rowCount, colCount)

    With tbl

        ' Basic table configuration
        .Borders.Enable = True
        .AllowAutoFit = False
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100

        ' Remove internal padding for precise layout
        .topPadding = 0
        .bottomPadding = 0
        .LeftPadding = 0
        .RightPadding = 0

        ' Define equal column width distribution
        For col = 1 To colCount
            .Columns(col).PreferredWidth = 100 / colCount
        Next col
        
        ' Calculate usable page height (excluding margins)
        usableHeight = doc.PageSetup.PageHeight - _
                       doc.PageSetup.TopMargin - _
                       doc.PageSetup.BottomMargin

        ' Define fixed row height to force single-page layout
        ' Convert border width to points
        Dim borderPt As Double
        Dim totalBorderHeight As Double
        
        borderPt = GetBorderWidthInPoints(borderWidth)
        
        ' Total horizontal borders (top + bottom + inner lines)
        totalBorderHeight = borderPt * (rowCount + 1)
        
        ' Final row height calculation
        rowHeight = ((usableHeight - totalBorderHeight) * 0.96) / rowCount

        For row = 1 To rowCount
            .Rows(row).HeightRule = wdRowHeightExactly
            .Rows(row).Height = rowHeight
        Next row

        ' Configure table borders
        With .Borders
            .InsideColor = borderColor
            .OutsideColor = borderColor
            .InsideLineStyle = BorderStyle
            .OutsideLineStyle = BorderStyle
            .InsideLineWidth = borderWidth
            .OutsideLineWidth = borderWidth
        End With

    End With

    ' ====== Fill table cells with word data ======
    For row = 1 To rowCount

        If fillRightToLeft Then

            ' Right-to-left filling mode
            For col = colCount To 1 Step -1

                If i <= UBound(Words) Then
                    FillCell tbl.Cell(row, col).Range, Words(i), _
                        fontColor, fontSize, isBold, isItalic, isUnderlined
                    i = i + 1
                End If

            Next col

        Else

            ' Left-to-right filling mode
            For col = 1 To colCount

                If i <= UBound(Words) Then
                    FillCell tbl.Cell(row, col).Range, Words(i), _
                        fontColor, fontSize, isBold, isItalic, isUnderlined
                    i = i + 1
                End If

            Next col

        End If

    Next row

    Set CreateOnePageTable = tbl

End Function

'--------------------------------------------------------
' Function      : InsertWordsFromFile
' Author        : BeeniGit
' Date          : 02/05/2026
' Version       : 1.1
' History       :
'
' Description :
'   Reads a CSV file and generates Word tables filled with words.
'   Each table corresponds to one page of the document.
'
' Parameters :
'   filePath            (String)   : Path to the CSV file
'   wordsDelimiter      (String)   : Delimiter for the file containing the word list
'   wordsColumn         (Long)     : Position od the word list
'   pageOrientation     (String)   : Page orientation (Portrait/Landscape)
'   colCount            (Long)     : Number of columns per table
'   rowCount            (Long)     : Number of rows per table
'   borderColorName     (String)   : Border color name
'   borderThicknessName (String)   : Border thickness name
'   borderStyleName     (String)   : Border style name
'   textColorName       (String)   : Text color name
'   textSizeStr         (String)   : Font size as string
'   isBold              (Boolean)  : Bold formatting flag
'   isItalic            (Boolean)  : Italic formatting flag
'   isUnderlined        (Boolean)  : Underline formatting flag
'   isRandom            (Boolean)  : Shuffle words before insertion
'
' Output :
'   Word tables inserted into the active document
'
' Example :
'   InsertWordsFromCSV "C:\file.csv", "Semicolons ;" , "1","Portrait", 5, 10, "Black", "Thin", "Single", "Black", "12", False, False, False, False
'
' Notes :
'   One table is created per page. Words are distributed sequentially.
'--------------------------------------------------------
Sub InsertWordsFromFile( _
    filePath As String, _
    wordsDelimiter As String, _
    wordsColumn As Long, _
    pageOrientation As String, _
    colCount As Long, _
    rowCount As Long, _
    borderColorName As String, _
    borderThicknessName As String, _
    borderStyleName As String, _
    textColorName As String, _
    textSizeStr As String, _
    isBold As Boolean, _
    isItalic As Boolean, _
    isUnderlined As Boolean, _
    isRandom As Boolean)

    ' ====== Variable declaration ======
    Dim tempArray() As String
    Dim wordArray() As String
    Dim wordCount As Long
    Dim fields() As String
    Dim line As String
    Dim i As Long

    Dim fso As Object, ts As Object
    Dim rng As Range
    Dim tbl As Table

    Dim textSize As Single
    Dim wordsPerPage As Long
    Dim totalPages As Long
    Dim p As Long
    Dim wordIndex As Long

    ' ====== Parse font size safely ======
    If IsNumeric(textSizeStr) Then
        textSize = val(textSizeStr)
    Else
        textSize = 12
    End If

    ' ====== Open and read CSV file ======
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False, -2)

    ' Skip header row if present
    If Not ts.AtEndOfStream Then ts.ReadLine

    ' Read file line by line
    Do While Not ts.AtEndOfStream

        line = ts.ReadLine

        If Trim(line) <> "" Then
            fields = Split(line, GetDelimiterFromName(wordsDelimiter))

            wordCount = wordCount + 1
            ReDim Preserve tempArray(1 To wordCount)
            tempArray(wordCount) = Trim(fields(wordsColumn - 1))
        End If

    Loop

    ts.Close

    ' ====== Validate input data ======
    If wordCount = 0 Then
        Call ErrorMessageDisplay("NoWordsCSV")
        Exit Sub
    End If

    ' Optional shuffle
    If isRandom Then ShuffleArray tempArray

    wordArray = tempArray

    ' ====== Set page orientation ======
    With ActiveDocument.PageSetup
        If pageOrientation = "Portrait" Then
            .Orientation = wdOrientPortrait
        Else
            .Orientation = wdOrientLandscape
        End If
    End With

    ' ====== Compute pagination ======
    wordsPerPage = colCount * rowCount
    totalPages = Int((wordCount - 1) / wordsPerPage) + 1

    ' Set insertion point at document end
    Set rng = ActiveDocument.Content
    rng.Collapse wdCollapseEnd

    wordIndex = 1

    ' ====== Generate one table per page ======
    For p = 1 To totalPages

        Set tbl = CreateOnePageTable( _
            ActiveDocument, rng, wordArray, wordIndex, _
            colCount, rowCount, _
            GetColorFromName(borderColorName), _
            GetLineWidth(borderThicknessName), _
            GetLineStyleFromName(borderStyleName), _
            GetColorFromName(textColorName), _
            textSize, isBold, isItalic, isUnderlined, _
            False)

        wordIndex = wordIndex + wordsPerPage

        Set rng = tbl.Range
        rng.Collapse wdCollapseEnd
        
        ' recreate range
        Set rng = ActiveDocument.Range(Start:=rng.End, End:=rng.End)
        
        ' Page break (except last page)
        If p < totalPages Then
            ' Insert page break
            rng.InsertBreak Type:=wdPageBreak
    
            ' The break is inside a paragraph
            Dim pRange As Range
            Set pRange = rng.Duplicate
            pRange.MoveStart wdParagraph, -2    ' Start the selection 2 paragraph before
            pRange.MoveEnd wdParagraph, 1       ' End the selection after the page break
    
            With pRange.Font
                .size = 1
            End With
        End If
    Next p

    MsgBox "Insertion complete: " & wordCount & " words across " & totalPages & " pages.", vbInformation

End Sub
