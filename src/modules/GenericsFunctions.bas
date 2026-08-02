Attribute VB_Name = "GenericsFunctions"
' Update variables
Public g_LatestVersion              As String
Public g_LatestReleaseURL           As String
Public g_UpdateAvailable            As Boolean
Public g_NetworkError               As Boolean

' Registry Constant
Private Const REGISTRY_APP          As String = "WordVBAToolkit"
Private Const REGISTRY_SECTION      As String = "Updates"
Private Const REGISTRY_KEY_NOTIF    As String = "NotifyDisabled"
Private Const REGISTRY_KEY_DATE     As String = "LastCheckDate"
Private Const REGISTRY_KEY_STATUS   As String = "Status"
Private Const REGISTRY_KEY_VERSION  As String = "Version"
Private Const REGISTRY_KEY_URL      As String = "URL"

'--------------------------------------------------------
' Function      : GetColorFromName
' Author        : BeeniGit
' Date          : 15/05/2025
' Version       : 1.0
' History       :
'
' Description :
'   Function that returns the RBG code based on the name of the color passed as a parameter
'
' Parameters :
'   colorName (String) : Name of the color to convert
'
' Output :
'   Long : RGB value of the color passed as a parameter
'
' Example :
'   GetColorFromName("Red") return the RGB value (255, 0, 0) in Long format
'
' Notes :
'   In case of unknown color, the default RGB value will be Black.
'--------------------------------------------------------
Function GetColorFromName(colorName As String) As Long
    Select Case LCase(Trim(colorName))
        Case "black":   GetColorFromName = RGB(0, 0, 0)
        Case "blue":    GetColorFromName = RGB(0, 0, 255)
        Case "red":     GetColorFromName = RGB(255, 0, 0)
        Case "green":   GetColorFromName = RGB(0, 128, 0)
        Case "yellow":  GetColorFromName = RGB(255, 215, 0)
        Case "orange":  GetColorFromName = RGB(255, 128, 0)
        Case "brown":   GetColorFromName = RGB(165, 42, 42)
        Case "purple":  GetColorFromName = RGB(128, 0, 128)
        Case "gray":    GetColorFromName = RGB(128, 128, 128)
        Case Else:      GetColorFromName = RGB(0, 0, 0)
    End Select
End Function

'--------------------------------------------------------
' Function      : GetLineWidth
' Author        : BeeniGit
' Date          : 15/05/2025
' Version       : 1.0
'
' Description :
'   Function that returns the Word line width based on the thickness passed as a parameter.
'
' Parameters :
'   thicknessName (String) : Selected border width name
'
' Output :
'   WdLineWidth : Word object that define the line width
'
' Example :
'   GetLineWidth("Medium (1.5 pt)") return the Word value wdLineWidth150pt
'
' Notes :
'   In case of unknown parameter, the default value will be 0,75pt
'--------------------------------------------------------
Function GetLineWidth(thicknessName As String) As WdLineWidth
    Select Case thicknessName
        Case "Thin (0.25 pt)":      GetLineWidth = wdLineWidth025pt
        Case "Normal (0.75 pt)":    GetLineWidth = wdLineWidth075pt
        Case "Medium (1.5 pt)":     GetLineWidth = wdLineWidth150pt
        Case "Bold (2.25 pt)":      GetLineWidth = wdLineWidth225pt
        Case "Very Bold (3 pt)":    GetLineWidth = wdLineWidth300pt
        Case Else:                  GetLineWidth = wdLineWidth075pt
    End Select
End Function

'--------------------------------------------------------
' Function      : GetBorderWidthInPoints
' Author        : BeeniGit
' Date          : 03/05/2026
' Version       : 1.0
'
' Description :
'   Function that returns the number of points  Word line width based on the thickness passed as a parameter.
'
' Parameters :
'   width (WdLineWidth) : Conversion needed for this linewidth
'
' Output :
'   Double : points value according to the line witdh parameter
'
' Example :
'   wdLineWidth025pt return the Word value 0.25
'
' Notes :
'   In case of unknown parameter, the default value will be 0,75pt
'--------------------------------------------------------
Function GetBorderWidthInPoints(width As WdLineWidth) As Double

    Select Case width
        Case wdLineWidth025pt: GetBorderWidthInPoints = 0.25
        Case wdLineWidth075pt: GetBorderWidthInPoints = 0.75
        Case wdLineWidth150pt: GetBorderWidthInPoints = 1.5
        Case wdLineWidth225pt: GetBorderWidthInPoints = 2.25
        Case wdLineWidth300pt: GetBorderWidthInPoints = 3
        Case Else: GetBorderWidthInPoints = 0.75 ' fallback
    End Select

End Function

'--------------------------------------------------------
' Function      : GetLineStyleFromName
' Author        : BeeniGit
' Date          : 15/05/2025
' Version       : 1.0
' History    :
'
' Description :
'   Function that returns the Word line style based on the parameter.
'
' Parameters :
'   styleName (String) : Line style
'
' Output :
'   WdLineStyle : Word object that define the line style
'
' Example :
'   GetLineStyleFromName("Double line") return the value wdLineStyleDouble
'
' Notes :
'   In case of unknown parameter, the default value will be single line
'--------------------------------------------------------
Function GetLineStyleFromName(styleName As String) As WdLineStyle
    Select Case styleName
        Case "Single line":         GetLineStyleFromName = wdLineStyleSingle
        Case "Double line":         GetLineStyleFromName = wdLineStyleDouble
        Case "Dotted line":         GetLineStyleFromName = wdLineStyleDot
        Case "Dash line":           GetLineStyleFromName = wdLineStyleDashLargeGap
        Case "Double-dash line":    GetLineStyleFromName = wdLineStyleDashDotDot
        Case Else:                  GetLineStyleFromName = wdLineStyleSingle
    End Select
End Function

'--------------------------------------------------------
' Function      : GetDelimiterFromName
' Author        : BeeniGit
' Date          : 04/06/2026
' Version       : 1.0
' History    :
'
' Description :
'   Function that returns the word delimiter based on the description name
'
' Parameters :
'   delimiterName (String) : Delimiter name.
'
' Output :
'   GetDelimiterFromName : Delimiter character based on the delimiter name
'
' Example :
'   NA
'
' Notes :
'   NA
'--------------------------------------------------------
Function GetDelimiterFromName(delimiterName As String) As String
    Select Case delimiterName
        Case "Tabulation":      GetDelimiterFromName = vbTab
        Case "Semicolons ;":    GetDelimiterFromName = ";"
        Case "Commas":          GetDelimiterFromName = ","
        Case Else:              GetDelimiterFromName = ";" ' Default fallback
    End Select
End Function

'--------------------------------------------------------
' Function      : ShuffleArray
' Author        : BeeniGit
' Date          : 05/07/2025
' Version       : 1.0
' History    :
'
' Description :
'   Suffle a array with the "Randomize" Algorithm
'
' Parameters :
'   arr (String) : Initial array containing several words or characters.
'
' Output :
'   arr : Mixed array using the Fisher-Yates algorithm
'
' Example :
'   NA
'
' Notes :
'   NA
'--------------------------------------------------------
Sub ShuffleArray(arr() As String)
    Dim i As Long, j As Long
    Dim temp As String
    Randomize
    For i = UBound(arr) To LBound(arr) + 1 Step -1
        j = Int((i - LBound(arr) + 1) * Rnd + LBound(arr))
        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp
    Next i
End Sub

'--------------------------------------------------------
' Function      : GetLatestReleaseJSON
' Author        : BeeniGit
' Date          : 24/07/2026
' Version       : 1.0
' History       :
'
' Description :
'   Get a JSON file from the project repository. This JSON file contains the last production version.
'
' Parameters :
'   NA
'
' Output :
'   GetLatestReleaseJSON    (String) :  JSON file from the project repo
'
' Example :
'   NA
'
' Notes :
'   NA
'--------------------------------------------------------
Private Function GetLatestReleaseJSON() As String
    Dim http As Object
    On Error GoTo ErrHandler

    ' Definitoon of the http objet
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Definition of the http request to the Github's API
    http.Open "GET", GIT_REPO_UPDATE, False
    http.SetRequestHeader "User-Agent", "WordVBAToolkit-UpdateChecker"
    http.SetRequestHeader "Accept", "application/vnd.github+json"
    http.setTimeouts 3000, 3000, 5000, 5000
    
    ' Send the http request
    http.Send

    ' Get status of the request
    If http.status = 200 Then
        ' Github answer with the JSON file
        GetLatestReleaseJSON = http.responseText
        g_NetworkError = False
    Else
        ' Time out or any other errors
        g_NetworkError = True
        GetLatestReleaseJSON = ""
        Call ErrorMessageDisplay("UpdateError")
    End If
    Exit Function

ErrHandler:
    g_NetworkError = True
    GetLatestReleaseJSON = ""
End Function

'--------------------------------------------------------
' Function      : ExtractJSONValue
' Author        : BeeniGit
' Date          : 24/07/2026
' Version       : 1.0
' History       :
'
' Description :
'   Get a JSON file from the project repository. This JSON file contains the last production version.
'
' Parameters :
'   json                (String)    : JSON file
'   key                 (String)    : Searched key into JSON file
'
' Output :
'   ExtractJSONValue    (String)    : Value corresponding to the requested key
'
' Example :
'   ExtractJSONValue(JSONfile, "tag_name") => value under the "tag_name" of the JSON contains the value "V1.1.0", the function return "v1.1.0"
'
' Notes :
'   NA
'--------------------------------------------------------
Private Function ExtractJSONValue(ByVal json As String, ByVal key As String) As String
    Dim searchKey As String
    Dim posStart As Long
    Dim posQuoteStart As Long
    Dim posQuoteEnd As Long

    ' Definition of the search string
    searchKey = """" & key & """:"""
    
    ' Search the first position of the search string in the JSON file.
    posStart = InStr(json, searchKey) ' If not found, default value : 0
    If posStart = 0 Then Exit Function

    ' Add offset to the first position to be able to detect the end of the value
    posQuoteStart = posStart + Len(searchKey)
    
    ' Search for quote marks to be able to detect the end of the value.
    posQuoteEnd = InStr(posQuoteStart, json, """")
    If posQuoteEnd = 0 Then Exit Function

    ExtractJSONValue = Mid(json, posQuoteStart, posQuoteEnd - posQuoteStart)
End Function

'--------------------------------------------------------
' Function      : CompareVersions
' Author        : BeeniGit
' Date          : 24/07/2026
' Version       : 1.0
' History       :
'
' Description :
'   Compare the version between the twon inputs.
'
' Parameters :
'   lovalVersion        (String)    : Local version get by the programm
'   lastVersion         (String)    : Last version stored into the git repository
'
' Output :
'   CompareVersions     (Integer)   : Output value of the comparaison. More informations about the output value avaliable in the Notes.
'
' Example :
'   CompareVersions("v1.0.0", "v1.1.0") return -1
'
' Notes :
' Output values of the function
'       1 : The local version is higher than the last version .
'       -1: The local version is lower than the latest version.
'       0 : The two versions are identical.
'--------------------------------------------------------
Function CompareVersions(ByVal lovalVersion As String, ByVal lastVersion As String) As Integer
    Dim localVersionParts() As String
    Dim lastVersionParts() As String
    Dim localVersionValue As Long
    Dim lastVersionValue As Long
    Dim i As Long
    Dim maxParts As Long

    ' Split the version values using the "." character. Remove the first character of the versions.
    localVersionParts = Split(Replace(LCase(lovalVersion), "v", ""), ".")
    lastVersionParts = Split(Replace(LCase(lastVersion), "v", ""), ".")
    
    ' Definition of the number of parts
    maxParts = IIf(UBound(localVersionParts) > UBound(lastVersionParts), UBound(localVersionParts), UBound(lastVersionParts))

    ' Version comparaison
    For i = 0 To maxParts
        localVersionValue = 0: lastVersionValue = 0
        If i <= UBound(localVersionParts) Then localVersionValue = val(localVersionParts(i))
        If i <= UBound(lastVersionParts) Then lastVersionValue = val(lastVersionParts(i))
        If localVersionValue <> lastVersionValue Then
            CompareVersions = IIf(localVersionValue > lastVersionValue, 1, -1)
            Exit Function
        End If
    Next i
    CompareVersions = 0
End Function

'--------------------------------------------------------
' Function      : CheckForUpdates
' Author        : BeeniGit
' Date          : 31/07/2026
' Version       : 1.0
' History       :
'
' Description :
'   Check the local version of the toolkit with the last version of the repository
'   This function send the request to the git repo to get the JSON file, check the tag into and compare the actual version "TOOLKIT_VERSION" with the version get into the JSON.
'
' Parameters :
'   NA
'
' Output :
'   CheckForUpdates     (Integer) : Update status of the check Update. For more information about the output value, check the notes.
'
' Example :
'   NA
'.
' Notes :
'   0 => Local version match with the last version on the git repo
'   1 => A new tool version is available.
'   15 => Network error.
'   16 => JSON file key error
'   17 => No check or notifications of the update.
'--------------------------------------------------------
Function CheckForUpdates(Optional ByVal forceCheck As Boolean = False) As Integer
    Dim updateStatus As Integer
    Dim json As String
    Dim latestTag As String
    Dim latestURL As String
    
    If Not IIf(GetRegistryValue(REGISTRY_KEY_NOTIF) = "1", True, False) Or forceCheck Then
        ' Get the JSON file from the last release
        json = GetLatestReleaseJSON()
        
        ' Check if empty file
        If json = "" And g_NetworkError Then
            updateStatus = 15
        Else:
            ' Search the key "tag_name" into the json file
            latestTag = ExtractJSONValue(json, "tag_name")
            
            ' Search the key "html_url" into the json file
            latestURL = ExtractJSONValue(json, "html_url")
            If latestTag = "" Or latestURL = "" Then
                updateStatus = 16
            Else
                g_LatestVersion = latestTag
                g_LatestReleaseURL = latestURL
                
                ' Compare the TOOLKIT_VERSION and the version of the git repo
                If CompareVersions(TOOLKIT_VERSION, latestTag) < 0 Then
                    updateStatus = 1
                Else
                    updateStatus = 0
                End If
            
                Call SetRegistryValue(REGISTRY_KEY_DATE, CStr(Now))
            End If
        End If
    Else
        updateStatus = 17
    End If
    
    Call SetRegistryValue(REGISTRY_KEY_STATUS, CStr(updateStatus))
    Call SetRegistryValue(REGISTRY_KEY_VERSION, CStr(g_LatestVersion))
    Call SetRegistryValue(REGISTRY_KEY_URL, CStr(g_LatestReleaseURL))
    CheckForUpdates = updateStatus
End Function

'--------------------------------------------------------
' Function      : LaunchCheckUpdates
' Author        : BeeniGit
' Date          : 31/07/2026
' Version       : 1.0
' History       :
'
' Description :
'   Check the conditions to launch the CheckForUpdates.
'   The conditions are the manual check, the first start-up with empty registry or the last check was made more than a day ago.
'   With this function, the check is made only one time per day to avoid network flowding or application freezing.
'
' Parameters :
'   NA
'
' Output :
'   LaunchCheckUpdates  (Integer) : Update status of the check Update. For more information about the output value, check the notes of the function "CheckForUpdates".
'
' Example :
'   NA
'
' Notes :
'   NA
'--------------------------------------------------------
Function LaunchCheckUpdates(Optional ByVal forceCheck As Boolean = False) As Integer
    Dim checkStatus As Integer
    Dim lastCheckDate As String
    
    ' Get the last check date into the registry.
    lastCheckDate = GetRegistryValue(REGISTRY_KEY_DATE)
    
    ' If first tool opening or manual check.
    If forceCheck Or lastCheckDate = "" Then
        checkStatus = CheckForUpdates(forceCheck)
    
    ' If lastCheckDate is corresponding to the date format.
    ElseIf IsDate(lastCheckDate) Then
        ' If lastCheckDate is lowwer
        If CDate(lastCheckDate) < Now - 1 Then
            checkStatus = CheckForUpdates()
        Else
            checkStatus = CInt(GetRegistryValue(REGISTRY_KEY_STATUS))
            g_LatestVersion = GetRegistryValue(REGISTRY_KEY_VERSION)
            g_LatestReleaseURL = GetRegistryValue(REGISTRY_KEY_URL)
        End If
    End If
    g_UpdateAvailable = checkStatus
    LaunchCheckUpdates = checkStatus
End Function

'--------------------------------------------------------
' Function      : SetRegistryValue
' Author        : BeeniGit
' Date          : 31/07/2026
' Version       : 1.0
' History       :
'
' Description :
'   Save into Windows registry the tool informations
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
Sub SetRegistryValue(registryKey As String, value As String)
    Call SaveSetting(REGISTRY_APP, REGISTRY_SECTION, registryKey, value)
End Sub

'--------------------------------------------------------
' Function      : GetRegistryValue
' Author        : BeeniGit
' Date          : 31/07/2026
' Version       : 1.0
' History       :
'
' Description :
'   Get from the Windows registry the tool information
'
' Parameters :
'   NA
'
' Output :
'   String : Value from the target registry
'
' Example :
'   NA
'
' Notes :
'   NA
'--------------------------------------------------------
Function GetRegistryValue(registryKey As String) As String
    GetRegistryValue = GetSetting(REGISTRY_APP, REGISTRY_SECTION, registryKey)
End Function

'--------------------------------------------------------
' Function      : SetUpdateNotificationDisabled
' Author        : BeeniGit
' Date          : 31/07/2026
' Version       : 1.0
' History       :
'
' Description :
'   Save into Windows registry the user preferences for the update notifications
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
'   Store the string character "0" if the parameter "disabled is False.
'   Store the string character "1" if the parameter "disabled is True.
'--------------------------------------------------------
Sub SetUpdateNotificationDisabled(ByVal disabled As Boolean)
    Call SetRegistryValue(REGISTRY_KEY_NOTIF, IIf(disabled, "1", "0"))
End Sub

'--------------------------------------------------------
' Function      : GetUpdateNotificationDisabled
' Author        : BeeniGit
' Date          : 31/07/2026
' Version       : 1.0
' History       :
'
' Description :
'   Get from the Windows registry the user preferences for the update notifications
'
' Parameters :
'   NA
'
' Output :
'   GetUpdateNotificationDisabled   (Boolean) : User preferences stored into the registry
'
' Example :
'   NA
'
' Notes :
'   NA
'--------------------------------------------------------
Function GetUpdateNotificationDisabled() As Boolean
    Dim regValue As Integer
    
    regValue = GetRegistryValue(REGISTRY_KEY_NOTIF)
    
    If regValue = 1 Then
        GetUpdateNotificationDisabled = True
    Else
        GetUpdateNotificationDisabled = False
    End If
End Function

'--------------------------------------------------------
' Function      : openURL
' Author        : BeeniGit
' Date          : 24/07/2026
' Version       : 1.0
' History       :
'
' Description :
'   Open a URL in the user’s primary browser. Consider network configuration errors.
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
Sub openURL(url As String)
    ' Check if the url is NULL
    If url <> "" Then
        ' Link the error to the ErrorHandler section
        On Error GoTo ErrorHandler
        
        ' Open the URL
        ThisDocument.FollowHyperlink Address:=url
        
    End If

    Exit Sub

ErrorHandler:
    Call ErrorMessageDisplay("NoConnection")
End Sub

'--------------------------------------------------------
' Function      : ErrorMessageDisplay
' Author        : BeeniGit
' Date          : 23/07/2026
' Version       : 1.1
' History       : 1.1 - 23/07/2026 => Add the error "NoConnection" from the AboutForm. Add optionnal value for Table limit.
'
' Description :
'   Generic function that display error messages to the user
'
' Parameters :
'   errorID (String) : Error ID
'
' Output :
'   MsgBox
'
' Example :
'   NA
'
' Notes :
'   NA
'--------------------------------------------------------
Sub ErrorMessageDisplay(errorID As String, Optional optionnalValue As Variant)
    Select Case errorID
        Case "NoCharacter":     MsgBox "No characters in the field", vbExclamation, "Formatting error"
        Case "CSV":             MsgBox "Please select a CSV file.", vbExclamation, "File error"
        Case "NoWordsCSV":      MsgBox "No words found in the CSV file.", vbExclamation, "File error"
        Case "TableNumeric":    MsgBox "Please enter numeric values for the columns and rows.", vbExclamation, "Table error"
        Case "WordsNumeric":    MsgBox "Please enter numeric values for the words column.", vbExclamation, "Words position error"
        Case "TableMin":        MsgBox "Please enter at least " & optionnalValue & " columns or rows.", vbExclamation, "Table error"
        Case "TableMax":        MsgBox "Please enter fewer than " & optionnalValue & " columns or rows.", vbExclamation, "Table error"
        Case "WordsMin":        MsgBox "Please enter at least " & optionnalValue & " columns for the words positions.", vbExclamation, "Words position error"
        Case "WordsMax":        MsgBox "Please enter fewer than " & optionnalValue & " columns for the words positions.", vbExclamation, "Words position error"
        Case "NoConnection":    MsgBox "Can't open the URL, check your connection or the URL", vbExclamation, "URL error"
        Case "UpdateError":     MsgBox "Error n:" & optionnalValue & " Can't access to the update status, check your connection", vbExclamation, "URL error"
    End Select
End Sub

'--------------------------------------------------------
' Function      : FinishMessageDisplay
' Author        : BeeniGit
' Date          : 23/07/2026
' Version       : 1.0
' History       :
'
' Description :
'   Generic function that display finish messages to the user
'
' Parameters :
'   finishID (String) : finish ID
'
' Output :
'   MsgBox
'
' Example :
'   NA
'
' Notes :
'   NA
'--------------------------------------------------------
Sub FinishMessageDisplay(finishID As String, Optional optionnalValue1 As Variant, Optional optionnalValue2 As Variant)
    Select Case finishID
        Case "TextFormatting":  MsgBox "Formatting done.", vbInformation
        Case "WordsInsertion":  MsgBox "Insertion complete: " & optionnalValue1 & " words across " & optionnalValue2 & " pages.", vbInformation
    End Select
End Sub
