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
' Date     : 23/07/2026
' Version  : 1.1
' History  : 1.1 - 23/07/2026 => Use constant for initialization
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
' Example :
'   N/A
'
' Notes :
'   N/A
'--------------------------------------------------------
Private Sub UserForm_Initialize()
    
    ' Update toolkit informations
    lblVersion.Caption = "Toolkit version : " & TOOLKIT_VERSION
    lblReleaseDate.Caption = "Release date : " & TOOLKIT_RELEASE_DATE
    lblLicence.Caption = "Open source licence : " & OPEN_SOURCE_LICENCE
    
    lblMainContributor.Caption = "Principal contributor : " & GIT_MAIN_CONTRIBUTOR
    lblMajorContributors.Caption = "Major contributors : " & GIT_MAJOR_CONTRIBUTORS

End Sub

'--------------------------------------------------------
' Sub      : btnGit_Click
' Author   : BeeniGit
' Date     : 23/07/2026
' Version  : 1.1
' History  : 23/07/2026 => Add the function ErrorMessageDisplay to display the network error. Use constant for connection.
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
' Example :
'   N/A
'
' Notes :
'   N/A
'--------------------------------------------------------
Private Sub btnGit_Click()
    If GIT_REPO <> "" Then
        On Error GoTo ErrorHandler
        
        ThisDocument.FollowHyperlink Address:=GIT_REPO
        
    End If

    Exit Sub

ErrorHandler:
    Call ErrorMessageDisplay("NoConnection")

End Sub

'--------------------------------------------------------
' Sub      : btnGitIssue_Click
' Author   : BeeniGit
' Date     : 23/07/2026
' Version  : 1.1
' History  : 23/07/2026 => Add the function ErrorMessageDisplay to display the network error. Use constant for connection.
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
' Example :
'   N/A
'
' Notes :
'   N/A
'--------------------------------------------------------
Private Sub btnGitIssue_Click()
    If GIT_REPO <> "" Then
        On Error GoTo ErrorHandler
        
        ThisDocument.FollowHyperlink Address:=GIT_REPO & GIT_ISSUES
        
    End If

    Exit Sub

ErrorHandler:
    Call ErrorMessageDisplay("NoConnection")

End Sub

'--------------------------------------------------------
' Sub      : btnClose_Click
' Author   : BeeniGit
' Date     : 18/03/2026
' Version  : 1.0
' History  :
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
' Example :
'   N/A
'
' Notes :
'   N/A
'--------------------------------------------------------
Private Sub btnClose_Click()
    Unload Me
End Sub

