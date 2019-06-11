Attribute VB_Name = "TempoWorklogs"
Option Explicit

Public sBasicAuth As String, sUsername As String

Sub createWorklogs()
On Error GoTo err_createWorklogs

    Dim wb As Workbook
    Dim wksTimeLogs As Worksheet, wksSetup As Worksheet, wksTeam As Worksheet
    Dim lLastRow As Long, lRow As Long, lPeopleIdx As Long, lWorklogIdx As Long, _
        lTimeSpentInMinutes As Long
    Dim sJiraBaseUrl As String, sJson As String, sJsonBody As String, sJql As String
    Dim sThisUser As String, sThisUserName As String, sThisUserEmail As String, _
        sRequestorName As String, sRequestorEmail As String, _
        sWorklogAudit As String
    Dim bInclude As Boolean
    Dim dDate As Date
    Dim oIssues As Dictionary, oJsonIssue As Dictionary, oJsonWorklog As Dictionary, _
        oJson As Dictionary
    Dim vIssue As Variant
        
    Set wb = ThisWorkbook
    Set wksSetup = wb.Worksheets("Setup")
    Set wksTeam = wb.Worksheets("Team Members")
    
    ' Set the Jira base
    sJiraBaseUrl = "https://" & Range("sJiraRoot").Value & ".silverchair.com"
    
    ' Set the basic authentication for REST
    If Len(sBasicAuth) = 0 Then
        sBasicAuth = GetJiraCredentials()
    End If
    
    ' ensure we have creds
    If Len(sBasicAuth) = 0 Then
        GoTo exit_createWorklogs
    End If
    
    Application.StatusBar = "Validating Jira Credentials"
    
    ' Confirm login via GET of scrum board result. If we have a response, we're good
    sJson = JiraRestAPI(sBasicAuth, sJiraBaseUrl & "/rest/agile/1.0/board", "GET", 0, 1)
    
    If Len(sJson) = 0 Then
        ' Houston we have issues. Fix the username and password and quit.
        MsgBox "Issue with User Name or Password. Please try again.", vbExclamation, "Error"
        GoTo exit_createWorklogs
    Else:
        sJson = ""
    End If
    
    Application.StatusBar = "Validating Issues"
    
    ' Get the issues from the worksheet and drop in a dictionary
    Set oIssues = GetIssues()
           
    ' Validate the issues in the API; if there's a problem, quit.
    For Each vIssue In oIssues
    
        Application.StatusBar = "Validating Issues: " & oIssues(vIssue)("key") _
            & " (" & vIssue & " of " & oIssues.Count & ")"
        
        ' Validate this issue exists via GET
        sJql = "jql=key=" & oIssues(vIssue)("key")
        sJson = JiraRestAPI(sBasicAuth, sJiraBaseUrl & "/rest/api/2/search", "GET", 0, , sJql)
        
        'parse the results
        Set oJsonIssue = JsonConverter.ParseJson(sJson)

        ' Check if the results are empty
        ' Assumes that resulting JSON will contain a key of "errorMessages" if there's an error
        If oJsonIssue.Exists("errorMessages") Then
            ' Ooops!  Invalid spreadsheet entry; report and quit
            MsgBox "Error processing line " & oIssues(vIssue)("rowIdx") & vbNewLine _
                & oJsonIssue("errorMessages")(1) & vbNewLine _
                & "Correct this entry and reprocess.  Time Logs were NOT posted due to this error."
            GoTo exit_createWorklogs
        End If
    Next
    
    ' Now that we have the issues and they are valid, let's get the people we're logging time for
    With wksTeam
        
        Application.StatusBar = "Posting Time"

        lLastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        
        ' First, look up the requestor
        For lRow = 2 To lLastRow
            If sUsername = .Cells(lRow, 2).Value Then
                ' this is it!
                sRequestorName = .Cells(lRow, 3).Value
                sRequestorEmail = .Cells(lRow, 4).Value
                Exit For
            End If
        Next lRow
        
        For lRow = 2 To lLastRow
            
            bInclude = .Cells(lRow, 1).Value
            ' If the INCLUDE is True for this person, let's log some time
            If bInclude Then
            
                ' This is a valid user, so capture it
                sThisUserName = .Cells(lRow, 2).Value ' User name
                sThisUser = .Cells(lRow, 3).Value ' Display Name
                sThisUserEmail = .Cells(lRow, 4).Value ' Email
                
                Application.StatusBar = "Posting Time: " & sThisUser
                
                'Write the header for this user
                sWorklogAudit = "Work Log Audit: " & sRequestorName & " has posted time for " & sThisUser & " as follows: " & vbNewLine _
                    & "Worklog No., " _
                    & "Work Date, " _
                    & "Time Spent, " _
                    & "Issue Key (Summary), " _
                    & "Timesheet Comment" _
                    & vbNewLine
            
                'Iterate over the issues log for the work to record for this person
                For Each vIssue In oIssues
                
                    Application.StatusBar = "Posting Time: " & sThisUser _
                        & " (" & vIssue & " of " & oIssues.Count & ")"
                    
                    ' Assemble the JSON
                    sJsonBody = AssembleJson(oIssues(vIssue)("key"), sThisUserName, oIssues(vIssue)("date"), _
                        oIssues(vIssue)("timeInMinutes"), oIssues(vIssue)("comment"))
                    
                    ' Post it. If successful, this will return JSON with the worklog created.
                    sJson = JiraRestAPI(sBasicAuth, sJiraBaseUrl & "/rest/tempo-timesheets/3/worklogs", "POST", , , , sJsonBody)
                    
                    ' Process the returned JSON to report out
                    ' Parse the return into a Dictionary for reporting
                    Set oJson = JsonConverter.ParseJson(sJson)
                
                    ' Add a worklog audit line
                    sWorklogAudit = sWorklogAudit & reportWorklog(oJson)
                
                Next ' get next oIssue
                
                ' Finished with this user, so trigger the mail
                Application.StatusBar = "Posting Time: " & sThisUser _
                    & "(Sending Email)"
                Call sendEmail(sThisUserName, sThisUserEmail, sRequestorName, sRequestorEmail, sWorklogAudit)
                
            End If
            
        Next lRow
    End With
    
exit_createWorklogs:
    Application.StatusBar = "Done!"
    Exit Sub

err_createWorklogs:
    MsgBox Err.Description, vbCritical, "Error: " & Err.Number
    Resume exit_createWorklogs
    
End Sub

Sub test()
    Dim sJson As String, sTest As String
    Dim oJson As Dictionary, oPeople As Dictionary
    
    ' This sJson is a sample of data returned from a single worklog POST
    sJson = "{" & vbNewLine _
    & """timeSpentSeconds"": 60," & vbNewLine _
    & """dateStarted"": ""2019-06-05T00:00:00.000""," & vbNewLine _
    & """dateCreated"": ""2019-06-05T21:22:15.000""," & vbNewLine _
    & """dateUpdated"": ""2019-06-05T21:22:15.000""," & vbNewLine _
    & """comment"": ""User Story Refinement for one minute""," & vbNewLine _
    & """self"": ""https://newjirasandbox.silverchair.com/rest/api/2/tempo-timesheets/3/worklogs/894724""," & vbNewLine _
    & """id"": 894724," & vbNewLine _
    & """jiraWorklogId"": 894724," & vbNewLine _
    & """author"": {" & vbNewLine _
    & "    ""self"": ""https://newjirasandbox.silverchair.com/rest/api/2/user?username=cpearson""," & vbNewLine _
    & "    ""name"": ""cpearson""," & vbNewLine _
    & "    ""key"": ""cpearson""," & vbNewLine _
    & "    ""displayName"": ""Chris Pearson""," & vbNewLine _
    & "    ""avatar"": ""https://newjirasandbox.silverchair.com/secure/useravatar?size=small&ownerId=cpearson&avatarId=25375""" & vbNewLine _
    & "}," & vbNewLine
    
    sJson = sJson _
    & """issue"": {" & vbNewLine _
    & "    ""self"": ""https://newjirasandbox.silverchair.com/rest/api/2/issue/243185""," & vbNewLine _
    & "    ""id"": 243185," & vbNewLine _
    & "    ""projectId"": 13680," & vbNewLine _
    & "    ""key"": ""SCMP-11878""," & vbNewLine _
    & "    ""remainingEstimateSeconds"": 0," & vbNewLine _
    & "    ""issueType"": {" & vbNewLine _
    & "        ""name"": ""Epic""," & vbNewLine _
    & "        ""iconUrl"": ""https://newjirasandbox.silverchair.com/secure/viewavatar?size=xsmall&avatarId=17177&avatarType=issuetype""" & vbNewLine _
    & "    }," & vbNewLine _
    & "    ""summary"": ""Implementation Epic: PSI IP Intrusion and IP Registry Integration""" & vbNewLine _
    & "}," & vbNewLine _
    & """worklogAttributes"": []," & vbNewLine _
    & """workAttributeValues"": []" & vbNewLine _
    & "}"
    
    ' Parse the return into a Dictionary for reporting
    Set oJson = JsonConverter.ParseJson(sJson)
    Set oPeople = GetPeople()
    
    sTest = reportWorklog(oPeople, oJson)
    
End Sub


Function reportWorklog(oWork As Dictionary) As String
On Error GoTo Err_reportWorklog

    Dim sString As String, sNewString As String
    Dim dDate As Date
    Dim sDate As String
           
    sDate = oWork("dateStarted")
    sDate = Left(sDate, 10)
        
    sNewString = sNewString _
        & oWork("jiraWorklogId") & ", " _
        & sDate & ", " _
        & (oWork("timeSpentSeconds") / 60) & "m" & ", " _
        & oWork("issue")("key") & " [" _
        & oWork("issue")("summary") & "]" & ", " _
        & oWork("comment") _
        & vbNewLine
    
    sString = sString & sNewString
 '   Debug.Print sString
    
    reportWorklog = sString

Exit_reportWorklog:
    Exit Function

Err_reportWorklog:
    MsgBox Err.Description, , "Error: " & Err.Number
    Resume Exit_reportWorklog
    
End Function


Function AssembleJson(ByVal sKey As String, ByVal sThisUserName As String, ByVal sDate As String, _
    ByVal lTimeSpentInMinutes As Long, Optional ByVal sComment As String) As String
    
    If Len(sComment) < 0 Then sComment = "Working on Issue " & sKey
    
    AssembleJson = "{""issue"": {""key"":""" & sKey & """}, " _
        & """author"": {""name"":""" & sThisUserName & """}," _
        & """comment"":""" & sComment & """," _
        & """dateStarted"":""" & sDate & """," _
        & """timeSpentSeconds"":" & (lTimeSpentInMinutes * 60) & "}"
        
        ' Debug.Print AssembleJson

End Function

Function GetIssues() As Dictionary
On Error GoTo Err_GetIssues

' This will return a dictionary of the issues from the worksheet
' LONG_INTEGER (index), one per issue
'   "rowIndex" - what row on the worksheet this came from
'   "key" - the Jira key to look up
'   "timeInMinutes" - the time spent in minutes
'   "date" - the effective date of the work
'   "comment" - a comment for the worklog

    Dim wb As Workbook
    Dim wksTimeLogs As Worksheet
    Dim lLastRow As Long, lRow As Long, lIdx As Long, _
        lTimeSpentInMinutes As Long
    Dim sKey As String, sComment As String, sDate As String
    Dim dDate As Date
    Dim oDict As Dictionary, oDictI As Dictionary
        
    Set wb = ThisWorkbook
    Set wksTimeLogs = wb.Worksheets("Issues")
    Set oDict = CreateObject("Scripting.Dictionary")

    With wksTimeLogs
    
        lLastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        For lRow = 2 To lLastRow
        
            ' Jira Issue Key
            sKey = .Cells(lRow, 1).Value
            
            ' Time
            lTimeSpentInMinutes = .Cells(lRow, 2).Value
            
            ' Date
            dDate = Range("effectiveDate").Value
            
            ' If blank, use today's date (a blank date will be year 1899)
            If Year(dDate) > 2000 Then
                sDate = Format(dDate, "YYYY-MM-DD")
            Else
                sDate = Format(Now(), "YYYY-MM-DD")
            End If
            
            ' Comment
            sComment = .Cells(lRow, 3).Value

            ' add to dict
            Set oDictI = CreateObject("Scripting.Dictionary")
            With oDictI
                .item("rowIdx") = lRow
                .item("key") = sKey
                .item("timeInMinutes") = lTimeSpentInMinutes
                .item("date") = sDate
                .item("comment") = sComment
            End With
            
            ' inject the item into the parent object
            Set oDict(lIdx) = oDictI
            ' clear the issue dictionary
            Set oDictI = Nothing
            
            ' increment dict index
            lIdx = lIdx + 1
        
        Next lRow

    End With
    
    If oDict.Count > 0 Then
        Set GetIssues = oDict
    End If
    
Exit_GetIssues:
    Exit Function

Err_GetIssues:
    MsgBox Err.Description
    Resume Exit_GetIssues
    
End Function

Function GetJiraCredentials() As String

    Dim sUser As String, sPass As String
    Dim isError As Boolean: isError = False
        
    ' get the User Name
    sUser = InputBox("JIRA User Name", "Enter JIRA Credentials")
    If checkInputBoxEntry(sUser) Then isError = True

    'get the Passord
    sPass = InputBoxDK("JIRA Password", "Enter JIRA Credentials")
    If checkInputBoxEntry(sPass) Then isError = True
    
    ' Ensure we have something to encode
    If isError Then
        MsgBox "Error in entering user name / password. Either cancelled or no entry provided.", vbInformation
        Exit Function
    Else
        GetJiraCredentials = Base64Encode(sUser + ":" + sPass)
        'Update the username of the person requesting the post
        sUsername = sUser
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


Function JiraRestAPI(ByVal sUserPass As String, ByVal sRestApiUrl As String, ByVal sMethod As String, _
    Optional ByVal lStartAt As Long = 0, Optional ByVal lMaxResults As Long = 1000, _
    Optional ByVal sParams As String, Optional sBody As String) As String

    Dim oJiraService As New MSXML2.XMLHTTP60
    Dim sUrl As String
        
    ' construct the URL
    Select Case sMethod
        Case "GET"
            sUrl = sRestApiUrl _
                & "?startAt=" & lStartAt _
                & "&maxResults=" & lMaxResults
        Case Else
            sUrl = sRestApiUrl
    End Select
    
    
    If Len(sParams) > 0 Then sUrl = sUrl & "&" & sParams
    
    With oJiraService
         
         .Open sMethod, sUrl
         
         .setRequestHeader "Content-Type", "application/json"
         .setRequestHeader "Accept", "application/json"
         .setRequestHeader "User-Agent", "ThisIsADummyUserAgent"
         .setRequestHeader "Authorization", "Basic " & sUserPass
         .send (sBody)
        
         If .Status <> "200" Then
             MsgBox "Error :  " & .responseText, vbCritical, "HTTP Error Code " & .Status
             JiraRestAPI = ""
         Else
             JiraRestAPI = oJiraService.responseText
         End If
    End With
End Function

Sub sendEmail(ByVal sReceiverName As String, ByVal sReceiverEmail As String, _
    ByVal sSenderName As String, ByVal sSenderEmail As String, ByVal sMessageBody As String)
    
    On Error GoTo err_sendEmail
    
    Dim olApp As Outlook.Application
    Dim olMsg As Outlook.MailItem
    Dim sPath As String, sFile As String
    Dim blnCreated As Boolean
    
    sPath = ThisWorkbook.Path
    sFile = sPath & "\timesheet.jpeg"
    
    ' Set Outlook, and if not running start it
    Set olApp = Outlook.Application
    If olApp Is Nothing Then
        Set olApp = Outlook.Application
         blnCreated = True
        Err.Clear
    Else
        blnCreated = False
    End If
      
    Set olMsg = olApp.CreateItem(olMailItem)
    With olMsg
        .To = sReceiverEmail
        .CC = sSenderEmail
        .Subject = "Timesheet Entry Posted! " _
            & sSenderName & " posted time for " & sReceiverName
        ' .Attachments.Add sFile, 1, 0
        ' .HTMLBody = sMessageBody
        .Body = sMessageBody
        '.Display
        .send
    End With

exit_sendEmail:
    Set olMsg = Nothing
    Set olApp = Nothing
    Exit Sub
    
err_sendEmail:
    MsgBox (Err.Description), vbCritical, "Error: " & Err.Number
    Resume exit_sendEmail
End Sub

