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
        Call ErrorMessageDisplay("NoCharacter")
        Exit Sub
    End If
    
    ' Call the module function
    Call LaunchSearchWithFormatting( _
                txtStrings.Text, _
                cmbFontColor.value, _
                cmbFontSize, _
                chkBold.value, _
                chkItalic.value, _
                chkUnderline.value, _
                chkMatchWord.value, _
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
' Date     : 24/07/2026
' Version  : 1.2
' History  : 1.1 - 23/07/2026 => Add the function ErrorMessageDisplay to display the network error. Use constant for connection.
'            1.2 - 24/07/2026 => Integrate the code lines into a generic function "openURL"
'
' Description :
'   Open the git URL in the wiki section for the formatting tool
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
    Call openURL(GIT_REPO & GIT_HELP_TXT)
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
    color = GetColorFromName(cmbFontColor.value)
    
    ' Get the font size choose by the user
    If IsNumeric(cmbFontSize.value) Then
        size = val(cmbFontSize.value)
    Else
        size = 12
    End If

    ' Update of the preview
    With lblPreview
        .Caption = "Example"
        .Font.size = size
        .Font.Bold = chkBold.value
        .Font.Italic = chkItalic.value
        .Font.Underline = chkUnderline.value
        .ForeColor = color
    End With
End Sub
