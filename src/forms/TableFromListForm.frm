VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TableFromListForm 
   Caption         =   "Convert words list into table"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10965
   OleObjectBlob   =   "TableFromListForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TableFromListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ====== Global variables ======
Private fileCSV As String

'--------------------------------------------------------
' Sub           : UserForm_Initialize
' Author        : BeeniGit
' Date          : 23/07/2026
' Version       : 1.2
' History       : 1.2 - 23/07/2026 => Change min, max and default value for spinbox
'
' Description :
'   Initialization of the form
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
'   N/A
'--------------------------------------------------------
Private Sub UserForm_Initialize()

    ' Initializing the Font Color Dropdown Menu
    With cmbFontColor
        .AddItem "Black"
        .AddItem "Red"
        .AddItem "Blue"
        .AddItem "Green"
        .AddItem "Yellow"
        .AddItem "Orange"
        .AddItem "Brown"
        .AddItem "Purple"
        .AddItem "Gray"
        .ListIndex = 0
    End With

    ' Font size combo box initialization
    With cmbFontSize
        .AddItem "8"
        .AddItem "9"
        .AddItem "10"
        .AddItem "11"
        .AddItem "12"
        .AddItem "14"
        .AddItem "16"
        .AddItem "18"
        .AddItem "20"
        .AddItem "24"
        .AddItem "28"
        .AddItem "36"
        .AddItem "48"
        .AddItem "72"
        .ListIndex = 4
    End With
    
    ' Initializing the Color
    With cmbBorderColor
        .AddItem "Black"
        .AddItem "Red"
        .AddItem "Blue"
        .AddItem "Green"
        .AddItem "Yellow"
        .AddItem "Orange"
        .AddItem "Brown"
        .AddItem "Purple"
        .AddItem "Gray"
        .ListIndex = 0
    End With
    
    ' Table Border Initialization
    With cmbBorderSize
        .AddItem "Thin (0.25 pt)"
        .AddItem "Normal (0.75 pt)"
        .AddItem "Medium (1.5 pt)"
        .AddItem "Bold (2.25 pt)"
        .AddItem "Very Bold (3 pt)"
        .ListIndex = 1
    End With
    
     With cmbBorderStyle
        .AddItem "Single line"              ' wdLineStyleSingle
        .AddItem "Double line"              ' wdLineStyleDouble
        .AddItem "Dotted line"              ' wdLineStyleDot
        .AddItem "Dash line"                ' wdLineStyleDashLargeGap
        .AddItem "Double-dash line"         ' wdLineStyleDashDotDot
    End With
    cmbBorderStyle.ListIndex = 0
    
    ' Initialize the page orientation
    With cmbOrientation
        .AddItem "Landscape"
        .AddItem "Portrait"
        .ListIndex = 0
    End With
    
    With cmbDelimiter
        .AddItem "Semicolons ;"
        .AddItem "Commas ,"
        .AddItem "Tabulation"
        .ListIndex = 0
    End With
    
    txtColumns.Value = TABLE_COL_DEF_VALUE
    txtRows.Value = TABLE_ROW_DEF_VALUE

    ' Columns SpinButton
    With spnColumns
        .Min = TABLE_COLROW_MIN
        .Max = TABLE_COLROW_MAX
        .Value = TABLE_COL_DEF_VALUE
    End With

    ' Row SpinButton
    With spnRows
        .Min = TABLE_COLROW_MIN
        .Max = TABLE_COLROW_MAX
        .Value = TABLE_ROW_DEF_VALUE
    End With
    
    txtWordsColumn = WORDS_COL_DEF_VALUE
    ' Columns SpinButton
    With spnWordsColumn
        .Min = WORDS_COL_MIN
        .Max = WORDS_COL_MAX
        .Value = WORDS_COL_DEF_VALUE
    End With
End Sub

'--------------------------------------------------------
' Sub           : spnColumns_Change
' Author        : BeeniGit
' Date          : 15/05/2025
' Version       : 1.0
' History       :
'
' Description :
'   Update the text for the number of columns in the form when the user clicks on the spinner.
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
'   N/A
'--------------------------------------------------------
Private Sub spnColumns_Change()
    txtColumns.Value = spnColumns.Value
End Sub

'--------------------------------------------------------
' Sub           : txtColumns_Change
' Author        : BeeniGit
' Date          : 23/07/2026
' Version       : 1.1
' History       : 1.1 - 23/07/2026 => Check the limit using constant and not the spinbox limit.
'
' Description :
'   Update the column spinbutton with the user value
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
'   N/A
'--------------------------------------------------------
Private Sub txtColumns_Change()
    ' ====== Variable declaration ======
    Dim val As Long
    
    ' Check if the user input for rows contains only numeric value
    If Not IsNumeric(txtColumns.Value) Then
        Call ErrorMessageDisplay("TableNumeric")
        Exit Sub
    End If
    
    ' Minimum number of columns for tables
    If txtColumns.Value < TABLE_COLROW_MIN Then
        Call ErrorMessageDisplay("TableMin", TABLE_COLROW_MIN)
        Exit Sub
    End If
    
    ' Maximum number of columns for tables
    If txtColumns.Value > TABLE_COLROW_MAX Then
        Call ErrorMessageDisplay("TableMax", TABLE_COLROW_MAX)
        Exit Sub
    End If
     
    val = CLng(txtColumns.Value)
    
    ' Check if the user value is a number
    If IsNumeric(txtColumns.Value) Then
    
        ' Check the limitation of the spinbutton
        If val >= TABLE_COL_MIN And val <= TABLE_COL_MAX Then
            
            ' Update the spinbutton
            spnColumns.Value = val
        End If
    End If
End Sub

'--------------------------------------------------------
' Sub           : spnRows_Change
' Author        : BeeniGit
' Date          : 15/05/2025
' Version       : 1.0
' History       :
'
' Description :
'   Update the row spinbutton with the user value
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
'   N/A
'--------------------------------------------------------
Private Sub spnRows_Change()
    txtRows.Value = spnRows.Value
End Sub

'--------------------------------------------------------
' Sub           : txtRows_Change
' Author        : BeeniGit
' Date          : 23/07/2026
' Version       : 1.1
' History       : 1.1 - 23/07/2026 => Check the limit using constant and not the spinbox limit.
'
' Description :
'   Update the value of the top linked to the row according to the value entered by the user in the text area
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
'   N/A
'--------------------------------------------------------
Private Sub txtRows_Change()
    ' ====== Variable declaration ======
    Dim val As Long
    
    ' Check if the user input for rows contains only numeric value
    If Not IsNumeric(txtRows.Value) Then
        Call ErrorMessageDisplay("TableNumeric")
        Exit Sub
    End If

    ' Minimum number of rows for tables
    If txtRows.Value < TABLE_COLROW_MIN Then
        Call ErrorMessageDisplay("TableMin", TABLE_COLROW_MIN)
        Exit Sub
    End If
    
    ' Maximum number of rows for tables
    If txtRows.Value > TABLE_COLROW_MAX Then
        Call ErrorMessageDisplay("TableMax", TABLE_COLROW_MAX)
        Exit Sub
    End If
    
    val = CLng(txtRows.Value)
    ' Check if the user value is a number
    If IsNumeric(txtRows.Value) Then
    
        ' Check the limitation of the spinbutton
        If val >= TABLE_COLROW_MIN And val <= TABLE_COLROW_MAX Then
        
            ' Update the spinbutton
            spnRows.Value = val
        End If
    End If
End Sub

'--------------------------------------------------------
' Sub           : spnWordsColumn_Change
' Author        : BeeniGit
' Date          : 03/06/2026
' Version       : 1.0
' History       :
'
' Description :
'   Update the words column spinbutton with the user value
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
'   N/A
'--------------------------------------------------------
Private Sub spnWordsColumn_Change()
    txtWordsColumn.Value = spnWordsColumn.Value
End Sub

'--------------------------------------------------------
' Sub           : txtWordsColumn_Change
' Author        : BeeniGit
' Date          : 23/07/2026
' Version       : 1.1
' History       : 1.1 - 23/07/2026 => Check the limit using constant and not the spinbox limit.
'
' Description :
'   Update the value of the top linked to the words columns according to the value entered by the user in the text area
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
'   N/A
'--------------------------------------------------------
Private Sub txtWordsColumn_Change()
    ' ====== Variable declaration ======
    Dim val As Long
    
    ' Check if the user input for words columns contains only numeric value
    If Not IsNumeric(txtWordsColumn.Value) Then
        Call ErrorMessageDisplay("WordsNumeric")
        Exit Sub
    End If
    
    ' Minimum number for the words position
    If txtWordsColumn.Value < WORDS_COL_MIN Then
        Call ErrorMessageDisplay("WordsMin", WORDS_COL_MIN)
        Exit Sub
    End If
    
    ' Maximum number for the words position
    If txtWordsColumn.Value > WORDS_COL_MAX Then
        Call ErrorMessageDisplay("WordsMax", WORDS_COL_MAX)
        Exit Sub
    End If
    
    val = CLng(txtWordsColumn.Value)
    
    ' Check if the user value is a number
    If IsNumeric(txtWordsColumn.Value) Then
    
        ' Check the limitation of the spinbutton
        If val >= WORDS_COL_MIN And val <= WORDS_COL_MAX Then
        
            ' Update the spinbutton
            spnWordsColumn.Value = val
        End If
    End If
End Sub

'--------------------------------------------------------
' Sub           : btnChooseFile_Click
' Author        : BeeniGit
' Date          : 15/05/2025
' Version       : 1.0
' History       :
'
' Description :
'   Ouvre une boite de dialogue Windows avec l'obligation de donner un fichier CSV
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
'   N/A
'--------------------------------------------------------
Private Sub btnChooseFile_Click()
    ' ====== Variable declaration ======
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Initialization of the dialog box
    With dlg
        .Title = "Choose CSV file"
        .Filters.Clear
        .Filters.Add "CSV file", "*.csv"
        If .show = -1 Then
            fileCSV = .SelectedItems(1)
            lblFile.Caption = fileCSV
        End If
    End With
End Sub

'--------------------------------------------------------
' Sub      : btnOK_Click
' Author   : BeeniGit
' Date     : 15/05/2025
' Version  : 1.0
' History  :
'
' Description :
'   Check the user values and lanch the formatting text function
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
'   N/A
'--------------------------------------------------------
Private Sub btnOK_Click()
    ' Checking for the presence of a file
    If fileCSV = "" Then
        Call ErrorMessageDisplay("CSV")
        Exit Sub
    End If

    ' Call the macro with the entered parameters
    Call InsertWordsFromFile( _
        fileCSV, _
        cmbDelimiter.Value, _
        CLng(txtWordsColumn.Value), _
        cmbOrientation.Value, _
        CLng(txtColumns.Value), _
        CLng(txtRows.Value), _
        cmbBorderColor.Value, _
        cmbBorderSize.Value, _
        cmbBorderStyle.Value, _
        cmbFontColor.Value, _
        cmbFontSize.Value, _
        chkBold.Value, _
        chkItalic.Value, _
        chkUnderline.Value, _
        chkRandom.Value)

    ' Call the form
    Unload Me
End Sub

'--------------------------------------------------------
' Sub      : btnCancel_Click
' Author   : BeeniGit
' Date     : 15/05/2025
' Version  : 1.0
' History  :
'
' Description :
'   Close and cancel the formatting form
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
'   N/A
'--------------------------------------------------------
Private Sub btnCancel_Click()
    Unload Me
End Sub

'--------------------------------------------------------
' Sub      : btnAbout_Click
' Author   : BeeniGit
' Date     : 15/05/2025
' Version  : 1.0
' History  :
'
' Description :
'   Open the About form of the toolkit
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
'   N/A
'--------------------------------------------------------
Private Sub btnAbout_Click()
    AboutForm.show
End Sub

'--------------------------------------------------------
' Sub      : btnHelp_Click
' Author   : BeeniGit
' Date     : 24/07/2026
' Version  : 1.2
' History  : 1.1 - 23/07/2026 => Add the function ErrorMessageDisplay to display the network error. Use constant for connection.
'            1.2 - 24/07/2026 => Integrate the code lines into a generic function "openURL"
'
' Description :
'   Open the git URL in the wiki section for the tool Table from list
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
'   N/A
'--------------------------------------------------------
Private Sub btnHelp_Click()
    Call openURL(GIT_REPO & GIT_HELP_WORDS)
End Sub
