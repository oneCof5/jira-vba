Attribute VB_Name = "RestHelper"
Option Explicit

Public sUser As String
Public sBasicAuth As String
Public sBaseUrl As String

Public Function SetBaseJiraUrl(ByVal sJiraRoot As String) As String
    
    Dim s As String

    ' Set the Jira base
    s = LCase("https://" & sJiraRoot & ".silverchair.com")
    
    SetBaseJiraUrl = s

End Function

' This function performs a REST API call. If the API returns data, that data is then returned by this function
Public Function JiraRestAPI(ByVal sRestApiUrl As String, ByVal sMethod As String, ByVal sBasicAuth As String, _
    Optional ByVal lStartAt As Long = 0, Optional ByVal lMaxResults As Long = 1000, _
    Optional ByVal sParams As String, Optional sJsonData As String) As String

    Dim oJiraService As New MSXML2.XMLHTTP60
    Dim sUrl As String, sLoginReason As String, sDeniedReason As String
        
    ' Confirm nothing's screwed with the baseUrl
    If Len(sBaseUrl) = 0 Then sBaseUrl = SetBaseJiraUrl(Range("sJiraRoot").Value)
            
    ' append the REST API to the base url
    sUrl = sBaseUrl & sRestApiUrl
    
    ' append any parameters as needed
    If sMethod = "GET" Then
        sUrl = sUrl & "?startAt=" & lStartAt _
            & "&maxResults=" & lMaxResults
    End If
     
    If Len(sParams) > 0 Then
        sUrl = sUrl & sParams
    End If
    
    ' Debug.Print sUrl
    
    ' Execute the Web Service
    With oJiraService
        ' Open the request
        .Open sMethod, sUrl
         
         ' Build the headers
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "User-Agent", "ThisIsADummyUserAgent"
        .setRequestHeader "Authorization", "Basic " & sBasicAuth
        
        ' do it
        If sMethod = "POST" Then
            .send (sJsonData)
        Else:
            .send
        End If
        
        ' Handle the response. First confirm we can log in
        sLoginReason = .getResponseHeader("X-Seraph-LoginReason")
        sDeniedReason = .getResponseHeader("X-Authentication-Denied-Reason")
        
        If sLoginReason <> "OK" Then
            MsgBox "Login failed. Response is: " & sLoginReason & vbNewLine _
                & "Message: " & sDeniedReason, vbCritical
                sBasicAuth = ""
                Exit Function
        Else:
            If .Status <> "200" Then
                MsgBox "HTTP Error :  " & .responseText, vbCritical, "HTTP Error Code " & .Status
                Exit Function
            Else
                JiraRestAPI = oJiraService.responseText
            End If
        End If
    
    End With

End Function

Public Function GetJiraUsername() As String
    
    Dim bError As Boolean, bValidUser As Boolean
    
    ' get the User Name
    sUser = InputBox("JIRA User Name", "Enter JIRA Credentials")
    If checkInputBoxEntry(sUser) Then bError = True

    ' Ensure we have something to encode
    If bError Then
        MsgBox "Error in entering user name. Either cancelled or no entry provided.", vbInformation
        Exit Function
    Else
        ' Double check that this is user is a valid user in our employee list
        bValidUser = ValidateJiraUsername(sUser)
        
        If bValidUser Then
            GetJiraUsername = sUser
        Else:
            ' this employee doesnt match. ensure that we continue
            If MsgBox("The entry of [" & sUser & "] does not match a current Employee on the Employees worksheet" & vbNewLine _
                & "Do you wish to proceed anyway?", _
                vbYesNo, "Confirm Username") = vbNo Then
                    MsgBox ("Cancelled per user request.")
                    sUser = ""
                    Exit Function
            Else:
                GetJiraUsername = sUser
            End If
        End If
    End If
    
End Function

Public Function GetJiraPassword() As String
    
    Dim bError As Boolean
    Dim sPass As String
    
    'get the Passord
    sPass = InputBoxDK("JIRA Password", "Enter JIRA Credentials")
    If checkInputBoxEntry(sPass) Then bError = True

    ' Ensure we have something to encode
    If bError Then
        MsgBox "Error in entering user name. Either cancelled or no entry provided.", vbInformation
        Exit Function
    Else
        GetJiraPassword = sPass
    End If
    
End Function


Public Function SetBasicAuth() As String

    Dim sPass As String

    If Len(sUser) = 0 Then
        sUser = GetJiraUsername()
    End If
    
    ' confim we actually have a user
    If Len(sUser) > 0 Then
        sPass = GetJiraPassword()
    End If
    
    ' Now Encode it
    If ((Len(sUser) > 0) And (Len(sPass) > 0)) Then
        SetBasicAuth = Base64Encode(sUser & ":" & sPass)
    Else:
        SetBasicAuth = ""
    End If

End Function


Function checkInputBoxEntry(sEntry As String) As Boolean
    Dim isError As Boolean
        
        ' https://stackoverflow.com/questions/26264814/how-to-detect-if-user-select-cancel-inputbox-vba-excel
        If StrPtr(sEntry) = 0 Then
            isError = True
        ElseIf sEntry = vbNullString Then
            isError = True
        Else
            isError = False
        End If
        
    checkInputBoxEntry = isError

End Function


Function AssembleJson(ByVal oJson As Dictionary, sThisUserName) As String

    Dim sIssueKey As String, sComment As String, sDate As String, _
        sEpicLink As String, sJson As String
    Dim lTimeSpentSeconds As Long
        
    ' Assign values
    If Len(sComment) < 0 Then
        sComment = "Working on Issue " & sIssueKey
    Else:
        sComment = oJson("comment")
    End If
    sIssueKey = oJson("issueKey")
    sEpicLink = oJson("epicLink")
    lTimeSpentSeconds = oJson("timeSpentSeconds")
    sDate = oJson("dateStarted")
    
    If Len(sEpicLink) > 0 Then
        sIssueKey = sEpicLink
    End If
    
    ' Build the JSON data
    sJson = "{" & vbNewLine
    sJson = sJson & "  ""issue"": {""key"":""" & sIssueKey & """}, " & vbNewLine
    sJson = sJson & "  ""author"": {""name"":""" & sThisUserName & """}," & vbNewLine
    sJson = sJson & "  ""comment"":""" & sComment & """," & vbNewLine
    sJson = sJson & "  ""dateStarted"":""" & sDate & """," & vbNewLine
    sJson = sJson & "  ""timeSpentSeconds"":" & lTimeSpentSeconds & vbNewLine
    sJson = sJson & "}"
    
    AssembleJson = sJson
    
    ' Debug.Print AssembleJson

End Function

Public Sub CanLogIntoJIRA()
    
    Dim sJson As String
    
    If Len(sBasicAuth) = 0 Then
        sBasicAuth = SetBasicAuth()
    End If
    
    ' Confirm login via GET of scrum board result. If we have a response, we're good
    sJson = JiraRestAPI("/rest/agile/1.0/board", "GET", sBasicAuth, 0, 1)

End Sub

Function ValidateJiraUsername(ByVal sUser As String) As Boolean

    Dim wks As Worksheet
    Dim lRow As Long, lLastRow As Long
    Dim bFound As Boolean
    
    Set wks = ThisWorkbook.Worksheets("Employees")
    
    With wks
    
        lLastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        
        For lRow = 2 To lLastRow
            If sUser = .Cells(lRow, 1).Value Then
                bFound = True
                Exit For
            End If
        
        Next lRow
        
    End With

    ValidateJiraUsername = bFound

End Function

' Connect to the Jira REST API and get details for this issue key
Public Function GetIssuesFromJira(ByVal lRow As Long, ByVal sKey As String, _
    Optional ByVal sParams As String) As Dictionary
    
    Dim sJql As String, sJson As String
    Dim oJson As Object
    
    ' Confirm our authentication is already handled
    If Len(sBasicAuth) = 0 Then
        sBasicAuth = RestHelper.SetBasicAuth()
    End If
    
    ' Validate this issue exists via GET
    sJql = "jql=key=" & sKey
    
    ' append to existing params (if any)
    sParams = sParams & "&" & sJql
    
    ' Go get the data
    sJson = RestHelper.JiraRestAPI("/rest/api/2/search", "GET", sBasicAuth, 0, 1, sParams)
        
    'parse the results
    If Len(sJson) > 0 Then
        Set oJson = JsonConverter.ParseJson(sJson)

        ' Check if the results are empty
        ' Assumes that resulting JSON will contain a key of "errorMessages" if there's an error
        If oJson.Exists("errorMessages") Then
            MsgBox "Error processing line " & lRow & vbNewLine _
                & oJson("errorMessages")(1) & vbNewLine _
                & "Correct this entry and reprocess."
            Set GetIssuesFromJira = Nothing
            Exit Function
        Else:
            Set GetIssuesFromJira = oJson
        End If
    End If


End Function

' Connect to the Jira REST API and get details for Employees
Public Function GetEmployeesFromJira(Optional ByVal sParams As String) As Dictionary
    
    Dim sJson As String, sRestApiUrl As String, sUrl As String
    Dim oJson As Object
    
    ' Confirm our authentication is already handled
    If Len(sBasicAuth) = 0 Then
        sBasicAuth = RestHelper.SetBasicAuth()
    End If
    
    ' Validate this issue exists via GET
    sRestApiUrl = "/rest/api/2/user/picker"
    sParams = sParams & "&query=@silverchair.com"
    
    ' Go get the data
    sJson = RestHelper.JiraRestAPI(sRestApiUrl, "GET", sBasicAuth, 0, 1000, sParams)
    sUrl = AssembleUrl(sRestApiUrl, "GET", 0, 1000, sParams)
    ' Debug.Print sUrl
        
    'parse the results
    If Len(sJson) > 0 Then
        Set oJson = JsonConverter.ParseJson(sJson)

        ' Check if the results are empty
        ' Assumes that resulting JSON will contain a key of "errorMessages" if there's an error
        If oJson.Exists("errorMessages") Then
            MsgBox "Empty result set returned. Check in a web browser via this URL:" & vbNewLine _
                & sUrl & vbNewLine _
                & "Error response from REST API:" & vbNewLine _
                & oJson("errorMessages")(1) & vbNewLine
            
            Set GetEmployeesFromJira = Nothing
            Exit Function
        Else:
            Set GetEmployeesFromJira = oJson
        End If
    End If

End Function

' Connect to the Jira REST API and get details for Employees
Public Function GetUserFromJira(ByVal UserName As String) As Dictionary
    
    Dim sJson As String, sUrl As String, sRestApiUrl As String, sParams As String
    Dim oJson As Object
    
    ' Confirm our authentication is already handled
    If Len(sBasicAuth) = 0 Then
        sBasicAuth = RestHelper.SetBasicAuth()
    End If
    
    ' Validate this user exists via GET
    sParams = "&username=" & UserName _
        & "&expand=groups"
    
    sRestApiUrl = "/rest/api/2/user"
    
    ' Go get the data
    sJson = RestHelper.JiraRestAPI(sRestApiUrl, "GET", sBasicAuth, 0, 1, sParams)
    sUrl = AssembleUrl(sRestApiUrl, "GET", 0, 1, sParams)
        
    'parse the results
    If Len(sJson) > 0 Then
        Set oJson = JsonConverter.ParseJson(sJson)

        ' Check if the results are empty
        ' Assumes that resulting JSON will contain a key of "errorMessages" if there's an error
        If oJson.Exists("errorMessages") Then
            MsgBox "Empty result set returned. Check in a web browser via this URL:" & vbNewLine _
                & sUrl & vbNewLine _
                & "Error response from REST API:" & vbNewLine _
                & oJson("errorMessages")(1) & vbNewLine
            
            Set GetUserFromJira = Nothing
            
            Exit Function
        Else:
            Set GetUserFromJira = oJson
        End If
    End If

End Function

Public Function AssembleUrl(ByVal sRestApiUrl As String, ByVal sMethod As String, _
    Optional ByVal lStartAt As Long = 0, Optional ByVal lMaxResults As Long = 1000, _
    Optional ByVal sParams As String) As String

    Dim sUrl As String

    ' Confirm nothing's screwed with the baseUrl
    If Len(sBaseUrl) = 0 Then sBaseUrl = SetBaseJiraUrl(Range("sJiraRoot").Value)
            
    ' append the REST API to the base url
    sUrl = sBaseUrl & sRestApiUrl
    
    ' append any parameters as needed
    If sMethod = "GET" Then
        sUrl = sUrl & "?startAt=" & lStartAt _
            & "&maxResults=" & lMaxResults
    End If
     
    If Len(sParams) > 0 Then
        sUrl = sUrl & sParams
    End If
    
    AssembleUrl = sUrl
    ' Debug.Print sUrl

End Function

