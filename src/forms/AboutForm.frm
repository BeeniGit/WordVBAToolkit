VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AboutForm 
   Caption         =   "About WordVBAToolkit"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5955
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
' Date     : 24/07/2026
' Version  : 1.2
' History  : 1.1 - 23/07/2026 => Use constant for initialization
'            1.2 - 24/07/2026 => Add the lblUpdate function to check the local version vs the repository version
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
    
    Call lblUpdate
    
End Sub

'--------------------------------------------------------
' Sub      : lblUpdate
' Author   : BeeniGit
' Date     : 24/07/2026
' Version  : 1.0
' History  :
'
' Description :
'   Update the label lblUpdateStatus in accordance to the output from the function CheckForUpdates
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
Private Sub lblUpdate()
    If CheckForUpdates() Then
        lblUpdateStatus.Caption = "New release available : A new version is available : " & g_LatestVersion
    Else
        lblUpdateStatus.Caption = "New release available : You are already up to date"
    End If

End Sub

'--------------------------------------------------------
' Sub      : btnGit_Click
' Author   : BeeniGit
' Date     : 24/07/2026
' Version  : 1.2
' History  : 1.1 - 23/07/2026 => Add the function ErrorMessageDisplay to display the network error. Use constant for connection.
'            1.2 - 24/07/2026 => Integrate the code lines into a generic function "openURL"
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
    Call openURL(GIT_REPO)
End Sub

'--------------------------------------------------------
' Sub      : btnGitIssue_Click
' Author   : BeeniGit
' Date     : 24/07/2026
' Version  : 1.2
' History  : 1.1 - 23/07/2026 => Add the function ErrorMessageDisplay to display the network error. Use constant for connection.
'            1.2 - 24/07/2026 => Integrate the code lines into a generic function "openURL"
'
' Description :
'   Open the git URL in the issues section
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
    Call openURL(GIT_REPO & GIT_ISSUES)
End Sub

'--------------------------------------------------------
' Sub      : btnManualCheck_Click
' Author   : BeeniGit
' Date     : 24/07/2026
' Version  : 1.0
' History  :
'
' Description :
'   Launch the manual check of the local version vs last version on the repository
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
Private Sub btnManualCheck_Click()
    Call lblUpdate
End Sub

'--------------------------------------------------------
' Sub      : btnReleaseInfo_Click
' Author   : BeeniGit
' Date     : 24/07/2026
' Version  : 1.0
' History  :
'
' Description :
'   Open the last release information page of the toolkit.
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
Private Sub btnReleaseInfo_Click()
    Call openURL(g_LatestReleaseURL)
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

