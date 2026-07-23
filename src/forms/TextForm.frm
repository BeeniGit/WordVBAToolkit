VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TextForm 
   Caption         =   "WordVBAToolkit : Formatting macro"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10965
   OleObjectBlob   =   "TextForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TextForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ====== Global variables ======
Public OKClicked As Boolean

'--------------------------------------------------------
' Sub      : UserForm_Initialize
' Author   : BeeniGit
' Date     : 15/05/2025
' Version  : 1.0
' History  :
'
' Description :
'   Formatting form initialization
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

    ' Color combo box initialization
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
    
    ' Call the sub function for update the preview
    UpdatePreview
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
    
    ' Check if the input string from the user is empty
    If Trim(txtStrings) = "" Then
        MsgBox "No characters in the field", vbExclamation
        Exit Sub
    End If
    
    ' Call the module function
    Call LaunchSearchWithFormatting( _
                txtStrings.Text, _
                cmbFontColor.Value, _
                cmbFontSize, _
                chkBold.Value, _
                chkItalic.Value, _
                chkUnderline.Value, _
                chkMatchWord.Value, _
                chkNoModifications)

    ' Close the form
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
' Date     : 23/07/2026
' Version  : 1.1
' History  : 23/07/2026 => Add the function ErrorMessageDisplay to display the network error. Use constant for connection.
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
    If GIT_REPO <> "" Then
        On Error GoTo ErrorHandler
        
        ThisDocument.FollowHyperlink Address:=GIT_REPO & GIT_HELP_TXT
        
    End If

    Exit Sub

ErrorHandler:
    Call ErrorMessageDisplay("NoConnection")
End Sub

'--------------------------------------------------------
' Sub      : chkBold_Click / chkItalic_Click / chkUnderline_Click / cmbFontColor_Change / cmbFontSize_Change
' Author   : BeeniGit
' Date     : 15/05/2025
' Version  : 1.0
' History  :
'
' Description :
'   Launch the update function for the form preview when the user click on options
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
Private Sub chkBold_Click(): UpdatePreview: End Sub

Private Sub chkItalic_Click(): UpdatePreview: End Sub

Private Sub chkUnderline_Click(): UpdatePreview: End Sub

Private Sub cmbFontColor_Change(): UpdatePreview: End Sub

Private Sub cmbFontSize_Change(): UpdatePreview: End Sub

'--------------------------------------------------------
' Sub      : UpdatePreview
' Author   : BeeniGit
' Date     : 15/05/2025
' Version  : 1.0
' History  :
'
' Description :
'   Update the preview label with the user configuration.
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
Private Sub UpdatePreview()
    Dim color As Long
    Dim size As Single
    
    ' Get the RGB value of the color choose by the user
    color = GetColorFromName(cmbFontColor.Value)
    
    ' Get the font size choose by the user
    If IsNumeric(cmbFontSize.Value) Then
        size = val(cmbFontSize.Value)
    Else
        size = 12
    End If

    ' Update of the preview
    With lblPreview
        .Caption = "Example"
        .Font.size = size
        .Font.Bold = chkBold.Value
        .Font.Italic = chkItalic.Value
        .Font.Underline = chkUnderline.Value
        .ForeColor = color
    End With
End Sub
