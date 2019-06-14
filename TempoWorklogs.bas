Attribute VB_Name = "TempoWorklogs"
Option Explicit

Public sBasicAuth As String, sUsername As String

Sub createWorklogs()
On Error GoTo err_createWorklogs

    Dim wb As Workbook
    Dim wksTimeLogs As Worksheet, wksSetup As Worksheet, wksTeam As Worksheet, _
        wksEmail As Worksheet
    Dim lLastRow As Long, lRow As Long, lPeopleIdx As Long, lWorklogIdx As Long, _
        lTimeSpentInMinutes As Long
    Dim sJiraBaseUrl As String, sJson As String, sJsonBody As String, sJql As String
    Dim sThisUser As String, sThisUserName As String, sThisUserEmail As String, _
        sRequestorName As String, sRequestorEmail As String, _
        sWorklogAudit As String
    Dim bInclude As Boolean, bHTML As Boolean: bHTML = False
    Dim dDate As Date
    Dim oIssues As Dictionary, oJsonIssue As Dictionary, oJsonWorklog As Dictionary, _
        oJson As Dictionary
    Dim vIssue As Variant
            
    Set wb = ThisWorkbook
    Set wksSetup = wb.Worksheets("Setup")
    Set wksTeam = wb.Worksheets("Team Members")
    Set wksEmail = wb.Worksheets("Email")
    
    ' HTML?
    If wksEmail.Shapes("Check Box 4").OLEFormat.Object.Value = 1 Then
        '1 is checked so True
        bHTML = True
    Else:
        bHTML = False
    End If
        
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
    Set oIssues = GetWorksheetIssues()
           
    ' For each person, post the work
    With wksTeam
        
        Application.StatusBar = "Posting Time"

        lLastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        
        ' First, look up the requestor
        For lRow = 3 To lLastRow
            If sUsername = .Cells(lRow, 2).Value Then
                ' this is it!
                sRequestorName = .Cells(lRow, 3).Value
                sRequestorEmail = .Cells(lRow, 4).Value
                Exit For
            End If
        Next lRow
        
        For lRow = 3 To lLastRow
            
            bInclude = .Cells(lRow, 1).Value
            ' If the INCLUDE is True for this person, let's log some time
            If bInclude Then
            
                ' This is a valid user, so capture it
                sThisUserName = .Cells(lRow, 2).Value ' User name
                sThisUser = .Cells(lRow, 3).Value ' Display Name
                sThisUserEmail = .Cells(lRow, 4).Value ' Email
                
                Application.StatusBar = "Posting Time: " & sThisUser
                                
                'Iterate over the issues log for the work to record for this person
                For Each vIssue In oIssues
                
                    Application.StatusBar = "Posting Time: " & sThisUser _
                        & " (" & vIssue + 1 & " of " & oIssues.Count & ")"
                    
                    ' Assemble the JSON
                    sJsonBody = AssembleJson(oIssues(vIssue), sThisUserName)
                                        
                    ' Post it. If successful, this will return JSON with the worklog created.
                    sJson = JiraRestAPI(sBasicAuth, sJiraBaseUrl & "/rest/tempo-timesheets/3/worklogs", "POST", , , , sJsonBody)
                    
                    ' Process the returned JSON to report out
                    ' Parse the return into a Dictionary for reporting
                    Set oJson = JsonConverter.ParseJson(sJson)
                
                    ' Add a worklog audit line
                    sWorklogAudit = sWorklogAudit & reportWorklog(oJson, bHTML)
                
                Next ' get next oIssue
                
                ' Finished with this user, so trigger the mail
                
                ' Update status bar
                Application.StatusBar = "Posting Time: " & sThisUser _
                    & "(Sending Email)"
                                    
                'Assemble the mail message (wrap header and footer around the worklog)
                sWorklogAudit = assembleEmailMsgBody(sWorklogAudit, bHTML)
                
                ' Send / Display the mail
                Call sendEmail(sThisUserName, sThisUserEmail, sRequestorName, sRequestorEmail, sWorklogAudit, bHTML)
                
                ' Clean Up
                sWorklogAudit = ""
                
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

Function assembleEmailMsgBody(ByVal sBody As String, Optional ByVal bHTML As Boolean = False) As String
    
    Dim sBodyText As String, sAdmTime As String, sTotalTime As String
    Dim wks As Worksheet
    
    Set wks = ThisWorkbook.Worksheets("Issues")
    sAdmTime = Format(wks.Range("adminTime").Value, "#,##0")
    sTotalTime = Format(wks.Range("totalTime").Value, "#,##0")
                    
    'Write the header for this user
    If bHTML Then
    
        sBodyText = emailIntroSection(sTotalTime, bHTML, sAdmTime)
        
        sBodyText = sBodyText _
            & "<table class=MsoTable15Plain4 border=0 cellspacing=0 cellpadding=0 style='border-collapse:collapse'>" _
            & "<tr>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal><b>Worklog No.<o:p></o:p></b></p></td>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal><b>Work Date<o:p></o:p></b></p></td>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal><b>Time Spent<o:p></o:p></b></p></td>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal><b>Issue Key<o:p></o:p></b></p></td>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal><b>Issue Summary<o:p></o:p></b></p></td>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal><b>Timesheet Comment<o:p></o:p></b></p></td>" _
            & "</tr>"
        
    Else:
        sBodyText = "Worklog No., " _
            & "Work Date, " _
            & "Time Spent, " _
            & "Issue Key (Summary), " _
            & "Timesheet Comment" _
            & vbNewLine
    End If
    
    'Append the log entries
    sBodyText = sBodyText & sBody
    
    If bHTML Then
        sBodyText = sBodyText & "</table>" _
            & "<p><o:p></o:p></p>"
    End If
    
    assembleEmailMsgBody = sBodyText

End Function

Function emailIntroSection(sTotalTime As String, Optional ByVal bHTML As Boolean = False, _
    Optional ByVal sAdmTime As String) As String

    Dim sMsgBody As String
    Dim wks As Worksheet
    
    Set wks = ThisWorkbook.Worksheets("Email")

    ' inject the happy image and informational text
    ' Start the table and the table row
    sMsgBody = "<table class=MsoTable15Plain4 border=0 cellspacing=0 cellpadding=0 style='border-collapse:collapse'>" _
        & "<tr>"
    ' Image
    sMsgBody = sMsgBody _
        & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'>" _
        & "<p class=MsoNormal>" _
        & "<img src=""cid:timesheet.jpeg"" height=128><o:p></o:p>" _
        & "</p>" _
        & "</td>"
    ' Informational text from the text box as well as admin time.
    sMsgBody = sMsgBody _
        & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'>"
    
    sMsgBody = sMsgBody _
        & "<p class=MsoNormal>" _
        & "<b>" & sTotalTime & "m</b> of time has been logged on your behalf!<o:p></o:p>" _
        & "</p><p class=MsoNormal><o:p></o:p></p>"
    
    sMsgBody = sMsgBody _
        & "<p class=MsoNormal>" _
        & wks.Range("emailBody").Value & "<o:p></o:p>" _
        & "</p>"
    ' If there is ADMIN time, notify the user
    If Len(sAdmTime) > 0 Then
        sMsgBody = sMsgBody _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'>" _
            & "<p class=MsoNormal>" _
            & "<span style='font-size:16.0pt;color:red'>" _
            & "ACTION REQUIRED: The total time does <u>not</u> inlude your admin time. Please record " _
            & sAdmTime & " minutes to your personal admin code.  <o:p></o:p></span>" _
            & "</p>"
    End If
    
    sMsgBody = sMsgBody _
        & "</td>"
    ' close the table and table row
    sMsgBody = sMsgBody _
        & "</tr>" _
        & "</table>"
    ' add a line break
    sMsgBody = sMsgBody _
        & "<p><o:p></o:p></p>"
        
    emailIntroSection = sMsgBody

End Function


Sub test()
    Dim sJson As String, sTest As String, sMsgBody As String
    Dim oJson As Dictionary
    
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
    
    ' Assemble body
    ' First, the table row(s)
    sMsgBody = reportWorklog(oJson, True)
    
    ' Second, the table head and close
    sMsgBody = assembleEmailMsgBody(sMsgBody, True)

    
    Call sendEmail("Chris P.", "cpearson@silverchair.com", "Chris", "cpearson@silverchair.com", sMsgBody, True)
    
End Sub



Function reportWorklog(ByVal oWork As Dictionary, Optional ByVal bHTML As Boolean = False) As String
On Error GoTo Err_reportWorklog

    Dim sString As String, sNewString As String
    Dim dDate As Date
    Dim sDate As String
           
    sDate = oWork("dateStarted")
    sDate = Left(sDate, 10)
    
    If bHTML Then
        sNewString = sNewString _
            & "<tr>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal>" & oWork("jiraWorklogId") & "</p></td>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal>" & sDate & "</p></td>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal>" & (oWork("timeSpentSeconds") / 60) & "m" & "</p></td>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal>" & oWork("issue")("key") & "</p></td>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal>" & oWork("issue")("summary") & "</p></td>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal>" & oWork("comment") & "</p></td>" _
            & "</tr>"
    Else:
        sNewString = sNewString _
            & oWork("jiraWorklogId") & ", " _
            & sDate & ", " _
            & (oWork("timeSpentSeconds") / 60) & "m" & ", " _
            & oWork("issue")("key") & " [" _
            & oWork("issue")("summary") & "]" & ", " _
            & oWork("comment") _
            & vbNewLine
    End If
    
    sString = sString & sNewString
 '   Debug.Print sString
    
    reportWorklog = sString

Exit_reportWorklog:
    Exit Function

Err_reportWorklog:
    MsgBox Err.Description, , "Error: " & Err.Number
    Resume Exit_reportWorklog
    
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

Function GetWorksheetIssues() As Dictionary
On Error GoTo Err_GetWorksheetIssues

' This will return a dictionary of the issues from the worksheet
' LONG_INTEGER (index), one per issue
'   "rowIndex" - what row on the worksheet this came from
'   "key" - the Jira key to look up
'   "timeInMinutes" - the time spent in minutes
'   "date" - the effective date of the work
'   "comment" - a comment for the worklog
'
'{
'issue:
'    {
'       key:string
'    }
'       author:Ignored in PUT operations

'    comment:stringDescription of the worklog
'    dateStarted:stringYYYY-MM-ddT00:00:00.000+0000
'    timeSpentSeconds:numberTime worked in seconds
'}

    Dim wb As Workbook
    Dim wksTimeLogs As Worksheet
    Dim lLastRow As Long, lRow As Long, lIdx As Long
    Dim sKey As String, sDate As String, sComment As String, sIssueSummary As String, sType As String, sEpicLink As String
    Dim lTimeSpentSeconds As Long
    Dim dDate As Date
    Dim oDict As Dictionary, oDictI As Dictionary
        
    Set wb = ThisWorkbook
    Set wksTimeLogs = wb.Worksheets("Issues")
    Set oDict = CreateObject("Scripting.Dictionary")

    With wksTimeLogs
    
        lLastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        For lRow = 6 To lLastRow
        
            ' assign vars
            sKey = .Cells(lRow, 1).Value ' Jira Issue Key
            dDate = .Cells(lRow, 7).Value ' Date
            sDate = Format(dDate, "yyyy-mm-ddThh:mm:ss.000+0000")  ' Converted Date
            lTimeSpentSeconds = .Cells(lRow, 6).Value * 60 ' Converted Time in Minutes to Seconds
            sComment = .Cells(lRow, 4).Value ' Time Log Comment
            sIssueSummary = .Cells(lRow, 9).Value ' Issue Summary
            sType = .Cells(lRow, 8).Value ' Issue Type
            sEpicLink = .Cells(lRow, 10).Value ' Epic Link

            ' add to dict
            Set oDictI = CreateObject("Scripting.Dictionary")
            With oDictI
                .Item("rowIdx") = lRow
                .Item("issueKey") = sKey
                .Item("issueSummary") = sIssueSummary
                .Item("epicLink") = sEpicLink
                .Item("issueType") = sType
                .Item("timeSpentSeconds") = lTimeSpentSeconds
                .Item("dateStarted") = sDate
                .Item("comment") = sComment
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
        Set GetWorksheetIssues = oDict
    End If
    
Exit_GetWorksheetIssues:
    Exit Function

Err_GetWorksheetIssues:
    MsgBox Err.Description
    Resume Exit_GetWorksheetIssues
    
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
    
    
    If Len(sParams) > 0 Then sUrl = sUrl & sParams
    
    Debug.Print sUrl
    
    
    With oJiraService
         
         .Open sMethod, sUrl
         
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "User-Agent", "ThisIsADummyUserAgent"
        .setRequestHeader "Authorization", "Basic " & sUserPass
        If sMethod = "POST" Then
            .send (sBody)
        Else:
            .send
        End If
        
         If .Status <> "200" Then
             MsgBox "Error :  " & .responseText, vbCritical, "HTTP Error Code " & .Status
             JiraRestAPI = ""
         Else
             JiraRestAPI = oJiraService.responseText
         End If
    End With
End Function

Sub sendEmail(ByVal sReceiverName As String, ByVal sReceiverEmail As String, _
    ByVal sSenderName As String, ByVal sSenderEmail As String, ByVal sMessageBody As String, _
    Optional ByVal bHTML As Boolean = False)
        
    On Error GoTo err_sendEmail
    
    Dim olApp As Outlook.Application
    Dim olMsg As Outlook.MailItem
    Dim sPath As String, sFile As String, sBody As String, _
        sSubject As String, sMsg As String, sMsDiv As String
    Dim blnCreated As Boolean, bPreview As Boolean
    Dim wb As Workbook
    Dim wksEmail As Worksheet
    
    Set wb = ThisWorkbook
    Set wksEmail = wb.Worksheets("Email")
    
    ' Preview?
    If wksEmail.Shapes("Check Box 3").OLEFormat.Object.Value = 1 Then
        '1 is checked so True
        bPreview = True
    Else:
        bPreview = False
    End If
    
    sBody = Range("emailBody").Value
    sSubject = Range("subject").Value
    
    If sSubject = "" Then
        sSubject = "Timesheet Entry! " _
            & sSenderName & " has posted time for you"
    End If
    
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
    
    If bHTML Then
        With olMsg
            .Display
        End With
        sMsg = olMsg.HTMLBody
        sMsDiv = "<div class=WordSection1><p class=MsoNormal><o:p>"
        
    End If
    
    With olMsg
        .To = sReceiverEmail
        .CC = sSenderEmail
        .Subject = "Timesheet Entry Posted! " _
            & sSenderName & " posted time for " & sReceiverName
        .Attachments.Add sFile, 1, 0
        ' Check if plain text or HTML
        If bHTML Then
            .HTMLBody = Replace(sMsg, sMsDiv, sMsDiv & sMessageBody)
        Else:
            .Body = sMessageBody
        End If
        
        ' Check if preview or direct send
        If Not bPreview Then
            .send
        End If
    End With

exit_sendEmail:
    Set olMsg = Nothing
    Set olApp = Nothing
    Exit Sub
    
err_sendEmail:
    MsgBox (Err.Description), vbCritical, "Error: " & Err.Number
    Resume exit_sendEmail
End Sub

Sub UpdateIssues()
On Error GoTo Err_UpdateIssues
    
    Dim wb As Workbook
    Dim wksSetup As Worksheet, wksTimeLogs As Worksheet
    Dim r As Range
    Dim lLastRow As Long, lRow As Long, lOffset As Long
    Dim sJiraBaseUrl As String, sJson As String, sParams As String, _
        sThisUser As String, sKey As String
    Dim oJson As Dictionary
    Dim dDate As Date
    
    Application.ScreenUpdating = False
        
    Set wb = ThisWorkbook
    Set wksSetup = wb.Worksheets("Setup")
    Set wksTimeLogs = wb.Worksheets("Issues")
    
    ' Set the Jira base
    sJiraBaseUrl = "https://" & Range("sJiraRoot").Value & ".silverchair.com"
    
    ' Set the basic authentication for REST
    If Len(sBasicAuth) = 0 Then
        sBasicAuth = GetJiraCredentials()
    End If
    
    ' ensure we have creds
    If Len(sBasicAuth) = 0 Then
        GoTo Exit_UpdateIssues
    End If
    
    With wksTimeLogs
        
        lLastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        lOffset = 5
        
        ' fill Today's Date if Blank
        dDate = .Range("effectiveDate")
        If Year(dDate) < 2019 Then
            dDate = Date
            .Range("effectiveDate").Value = dDate
        End If
        
        ' Iterate over the issues
        For lRow = 6 To lLastRow
            
            sKey = .Cells(lRow, 1).Value
            
            ' Update Status Bar
            Application.StatusBar = "Validating Issues: " & sKey _
                & " (" & lRow - lOffset & " of " & lLastRow - lOffset & ")"
        
            ' Handle blank entries
            If sKey <> "" Then
            
                sParams = "&fields=summary,issuetype,customfield_11732"
            
                ' Validate this issue exists via GET
                Set oJson = GetIssues(lRow, sJiraBaseUrl, sKey, sParams)
                
                If Not IsNull(oJson) Then
                        
                    ' Calculated Duration
                    .Cells(lRow, 5).NumberFormat = "h:mm"
                    .Cells(lRow, 5).FormulaR1C1 = "=IF(RC[-2]-RC[-3]>0, RC[-2]-RC[-3],"""")"
                    
                    ' Time Spent in Minutes
                    .Cells(lRow, 6).NumberFormat = "#,##0"
                    ' =IF(ISERROR(E6*1440),"",E6*1440)
                    .Cells(lRow, 6).FormulaR1C1 = "=IF(ISERROR(RC[-1]*1440), """",RC[-1]*1440)"
                    
                    ' Calculated Start Time
                    .Cells(lRow, 7).NumberFormat = "m/d/yyyy h:mm"
                    .Cells(lRow, 7).Formula = "=$G$1 + B" & lRow
                
                    'type
                    .Cells(lRow, 8).Value = oJson("issues")(1)("fields")("issuetype")("name")
                    'summary
                    .Cells(lRow, 9).Value = oJson("issues")(1)("fields")("summary")
                    ' epic Link
                    .Cells(lRow, 10).Value = oJson("issues")(1)("fields")("customfield_11732")
                    ' Epic summary
                End If
            Else:
                Debug.Print "Row " & lRow & " is blank!"
            End If

        Next lRow
    
    End With
            
    Application.StatusBar = "Validating Issues: Done"
        
Exit_UpdateIssues:
    Application.ScreenUpdating = True
    Exit Sub

Err_UpdateIssues:
    MsgBox Err.Description, vbCritical, "Error: " & Err.Number
End Sub

Function GetIssues(ByVal lRow As Long, ByVal sJiraBaseUrl As String, _
    ByVal sKey As String, Optional ByVal sParams As String) As Dictionary
    
    Dim sJql As String, sJson As String
    Dim oJson As Object
    
    ' Validate this issue exists via GET
    sJql = "jql=key=" & sKey
    sParams = sParams & "&" & sJql
    sJson = JiraRestAPI(sBasicAuth, sJiraBaseUrl & "/rest/api/2/search", "GET", 0, 1, sParams)
        
    'parse the results
    Set oJson = JsonConverter.ParseJson(sJson)

    ' Check if the results are empty
    ' Assumes that resulting JSON will contain a key of "errorMessages" if there's an error
    If oJson.Exists("errorMessages") Then
        ' Ooops!  Invalid spreadsheet entry; report and quit
        MsgBox "Error processing line " & lRow & vbNewLine _
            & oJson("errorMessages")(1) & vbNewLine _
            & "Correct this entry and reprocess."
        Set GetIssues = Nothing
        Exit Function
    Else:
        Set GetIssues = oJson
    End If

End Function

Sub UpdateUsers()
On Error GoTo Err_UpdateUsers

    Dim wb As Workbook
    Dim wksSetup As Worksheet, wksTeam As Worksheet
    Dim lLastRow As Long, lRow As Long
    Dim sJiraBaseUrl As String, sJson As String, sParam As String, _
        sThisUser As String
    Dim oJson As Dictionary
    
    Application.ScreenUpdating = False
        
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
        GoTo Exit_UpdateUsers
    End If
    
    With wksTeam
        
        lLastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        
        ' Iterate over the people
        For lRow = 3 To lLastRow
            
            ' Update Status Bar
            Application.StatusBar = "Processing row " & lRow & " of " & lLastRow
            
            sThisUser = .Cells(lRow, 2).Value
            
            ' Check this row's user name
            sParam = "&username=" & sThisUser
            sJson = JiraRestAPI(sBasicAuth, sJiraBaseUrl & "/rest/api/2/user", "GET", 0, 1, sParam)
            
            If Len(sJson) = 0 Then
                ' Houston we have issues. Fix the username and password and quit.
                MsgBox "Issue with User Name or Password. Please try again.", vbExclamation, "Error"
                GoTo Exit_UpdateUsers
            Else:
                'parse the results
                Set oJson = JsonConverter.ParseJson(sJson)
                ' Check if the results are empty
                ' Assumes that resulting JSON will contain a key of "errorMessages" if there's an error
                If oJson.Exists("errorMessages") Then
                    ' Ooops!  Invalid spreadsheet entry; report and quit
                    MsgBox "Error processing row " & lRow & vbNewLine _
                        & oJson("errorMessages")(1) & vbNewLine _
                        & "Correct this entry and retry."
                    GoTo Exit_UpdateUsers
                Else:
                    'OK!
                    .Cells(lRow, 3).Value = oJson("displayName")
                    .Cells(lRow, 4).Value = oJson("emailAddress")
                    .Cells(lRow, 5).Value = oJson("avatarUrls")("48x48")
                    
                End If
            
            End If

        Next lRow
    
    End With
    
Exit_UpdateUsers:
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    Exit Sub

Err_UpdateUsers:
    MsgBox Err.Description, vbCritical, "Error: " & Err.Number
    Resume Exit_UpdateUsers
    
End Sub

