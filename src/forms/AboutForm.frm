VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AboutForm 
   Caption         =   "About WordVBAToolkit"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5565
   OleObjectBlob   =   "AboutForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--------------------------------------------------------
' Sub      : UserForm_Initialize
' Author   : BeeniGit
' Date     : 18/03/2026
' Version  : 1.0
'
' Description :
'   About form initialization
'
' Parameters :
'   N/A
'
' Output :
'   N/A
'
' Exemple :
'   N/A
'
' Notes :
'   N/A
'--------------------------------------------------------
Private Sub UserForm_Initialize()
    Call InitializeToolkit(True)
    
    ' Update toolkit informations
    lblVersion.Caption = "Toolkit version : " & toolkitVersion
    lblReleaseDate.Caption = "Release date : " & toolkitReleaseDate
    lblLicence.Caption = "Open source licence : " & openSourceLicence
    
    lblMainContributor.Caption = "Principal contributor : " & gitMainContributor
    lblMajorContributors.Caption = "Major contributors : " & gitMajorContributors

End Sub

'--------------------------------------------------------
' Sub      : btnGit_Click
' Author   : BeeniGit
' Date     : 18/03/2026
' Version  : 1.0
'
' Description :
'   Open the git URL
'
' Parameters :
'   N/A
'
' Output :
'   N/A
'
' Exemple :
'   N/A
'
' Notes :
'   N/A
'--------------------------------------------------------
Private Sub btnGit_Click()
    If gitRepo <> "" Then
        On Error GoTo ErrorHandler
        
        ThisDocument.FollowHyperlink Address:=gitRepo
        
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Can't open the URL, check your connexion or the URL", vbExclamation, "URL error"

End Sub

'--------------------------------------------------------
' Sub      : btnGitIssue_Click
' Author   : BeeniGit
' Date     : 18/03/2026
' Version  : 1.0
'
' Description :
'   Open the git URL in the issue section
'
' Parameters :
'   N/A
'
' Output :
'   N/A
'
' Exemple :
'   N/A
'
' Notes :
'   N/A
'--------------------------------------------------------
Private Sub btnGitIssue_Click()
    
    If gitRepo <> "" Then
        On Error GoTo ErrorHandler
        
        ThisDocument.FollowHyperlink Address:=gitRepo & gitIssues
        
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Can't open the URL, check your connexion or the URL", vbExclamation, "URL error"

End Sub

'--------------------------------------------------------
' Sub      : btnClose_Click
' Author   : BeeniGit
' Date     : 18/03/2026
' Version  : 1.0
'
' Description :
'   Close the about form
'
' Parameters :
'   N/A
'
' Output :
'   N/A
'
' Exemple :
'   N/A
'
' Notes :
'   N/A
'--------------------------------------------------------
Private Sub btnClose_Click()
    Unload Me
End Sub

