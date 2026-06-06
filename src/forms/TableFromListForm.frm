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
' Sub      : btnHelp_Click
' Author   : BeeniGit
' Date     : 28/05/2026
' Version  : 1.0
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
    Call InitializeToolkit(True)
    If gitRepo <> "" Then
        On Error GoTo ErrorHandler
        
        ThisDocument.FollowHyperlink Address:=gitRepo & gitHelpCSV
        
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Can't open the URL, check your connexion or the URL", vbExclamation, "URL error"
End Sub

'--------------------------------------------------------
' Sub           : UserForm_Initialize
' Author        : BeeniGit
' Date          : 03/06/2026
' Version       : 1.1
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
    
    txtColumns.Value = 3
    txtRows.Value = 5

    ' Columns SpinButton
    With spnColumns
        .Min = 1
        .Max = 20
        .Value = 3
    End With

    ' Row SpinButton
    With spnRows
        .Min = 1
        .Max = 30
        .Value = 5
    End With
    
    ' Columns SpinButton
    With spnWordsColumn
        .Min = 1
        .Max = 30
        .Value = 1
    End With
End Sub

'--------------------------------------------------------
' Sub           : spnColumns_Change
' Author        : BeeniGit
' Date          : 15/05/2025
' Version       : 1.0
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
' Date          : 15/05/2025
' Version       : 1.0
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
    Dim val As Long: val = CLng(txtColumns.Value)
    
    ' Check if the user value is a number
    If IsNumeric(txtColumns.Value) Then
    
        ' Check the limitation of the spinbutton
        If val >= spnColumns.Min And val <= spnColumns.Max Then
            
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
' Date          : 15/05/2025
' Version       : 1.0
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
    Dim val As Long: val = CLng(txtRows.Value)
    
    ' Check if the user value is a number
    If IsNumeric(txtRows.Value) Then
    
        ' Check the limitation of the spinbutton
        If val >= spnRows.Min And val <= spnRows.Max Then
        
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
' Date          : 03/06/2026
' Version       : 1.0
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
    Dim val As Long: val = CLng(txtWordsColumn.Value)
    
    ' Check if the user value is a number
    If IsNumeric(txtWordsColumn.Value) Then
    
        ' Check the limitation of the spinbutton
        If val >= spnWordsColumn.Min And val <= spnWordsColumn.Max Then
        
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
        MsgBox "Please select a CSV file.", vbExclamation
        Exit Sub
    End If
    
    ' Verification of the completed figure for rows and columns
    If Not IsNumeric(txtColumns.Value) Or Not IsNumeric(txtRows.Value) Then
        MsgBox "Please enter numeric values for the columns and rows.", vbExclamation
        Exit Sub
    End If
    
    ' Verification of the completed figure for rows and columns
    If Not IsNumeric(txtWordsColumn.Value) Then
        MsgBox "Please enter numeric values for the words column.", vbExclamation
        Exit Sub
    End If
    
    ' Minimum number of columns and rows for tables
    If txtColumns.Value < 1 Or txtRows.Value < 1 Then
        MsgBox "Please enter at least 1 columns or rows.", vbExclamation
        Exit Sub
    End If
    
    ' Maximum number of columns and rows for tables
    If txtColumns.Value > 30 Or txtRows.Value > 30 Then
        MsgBox "Please enter fewer than 30 columns or rows.", vbExclamation
        Exit Sub
    End If
    
    ' Minimum number of for the word position
    If txtWordsColumn.Value < 1 Then
        MsgBox "Please enter at least 1 columns for the words positions.", vbExclamation
        Exit Sub
    End If
    
    ' Maximum number of for the word position
    If txtWordsColumn.Value > 30 Then
        MsgBox "Please enter fewer than 30 columns for the words positions.", vbExclamation
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
