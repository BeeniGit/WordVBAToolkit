Attribute VB_Name = "Config_Ribbon"
Private firstUpdateCheck As Boolean

'--------------------------------------------------------
' Function      : GetUpdateVisible
' Author        : BeeniGit
' Date          : 28/06/2026
' Version       : 1.0
' History       :
'
' Description :
'   Function call by the ribbon and set the visibility of the update button in accordance to the update status
'
' Parameters :
'   NA
'
' Output :
'   NA
'
' Example :
'   NA
'
' Notes :
'   NA
'--------------------------------------------------------
Sub GetUpdateVisible(control As IRibbonControl, ByRef returnedVal)
    ' Check the update status only one time
    If Not firstUpdateCheck Then
        firstUpdateCheck = True
        Call LaunchCheckUpdates
    End If
    returnedVal = g_UpdateAvailable
End Sub

'--------------------------------------------------------
' Function      : GetUpdateLabel
' Author        : BeeniGit
' Date          : 26/06/2026
' Version       : 1.0
' History       :
'
' Description :
'   Function call by the ribbon and set the text of the update button in accordance to the update status
'
' Parameters :
'   NA
'
' Output :
'   NA
'
' Example :
'   NA
'
' Notes :
'   NA
'--------------------------------------------------------
Sub GetUpdateLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "A new version is available : " & g_LatestVersion
End Sub

'--------------------------------------------------------
' Function      : OnUpdateClick
' Author        : BeeniGit
' Date          : 26/06/2026
' Version       : 1.0
' History       :
'
' Description :
'   Open the last release page of the tool.
'
' Parameters :
'   NA
'
' Output :
'   NA
'
' Example :
'   NA
'
' Notes :
'   NA
'--------------------------------------------------------
Sub OnUpdateClick(control As IRibbonControl)
    Call openURL(g_LatestReleaseURL)
End Sub

