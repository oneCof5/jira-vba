Attribute VB_Name = "TempoWorklogs"
Option Explicit

Sub createWorklogs()
On Error GoTo err_createWorklogs

    Dim wksTeam As Worksheet, wksEmail As Worksheet
    Dim lLastRow As Long, lRow As Long, lIssueIdx As Long, lIssueCount As Long
    Dim sJson As String, sJsonBody As String, sJql As String, _
        sThisUser As String, sThisUserName As String, sThisUserEmail As String, _
        sRequestorName As String, sRequestorEmail As String, _
        sWorklogAudit As String, sInclude As String, sTempDate As String
    Dim dDate As Date, dWork As Date, dToday As Date
    Dim oIssues As Dictionary, oJsonIssue As Dictionary, oJsonWorklog As Dictionary, _
        oJson As Dictionary
                            
    Set wksTeam = ThisWorkbook.Worksheets("Team Members")
    Set wksEmail = ThisWorkbook.Worksheets("Email")
        
    Application.StatusBar = "Validating Issues"
    
    ' Ensure the issues are valid
    Call TempoWorklogs.UpdateIssues
        
    ' Get the issues from the worksheet and drop in a dictionary
    Set oIssues = GetWorksheetIssues()
    
    ' Handle the case where the date is not today and force confirmation
    dToday = Date
    sTempDate = Left(oIssues(0)("dateStarted"), 10)
    dWork = sTempDate
    
    If dWork < dToday Then
        If MsgBox("You're posting time for a date that is not today's date. Are you sure this is correct?", _
            vbYesNo, "Confirm Date other than Today") = vbNo Then
            MsgBox ("Cancelled per user request. Check the date of the time entry.")
            GoTo exit_createWorklogs:
        End If
    End If
           
    ' For each person, post the work
    With wksTeam
        
        Application.StatusBar = "Posting Time"

        lLastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        
        ' Validate data
        For lRow = 3 To lLastRow
            If .Cells(lRow, 2).Value = "" Then
                MsgBox "There is a data issue in row " & lRow & "." & vbNewLine _
                    & "Please validate there are no empty rows in the worksheet." & vbNewLine _
                    & "Processing stopped; No time logs have been posted."
                Exit For
            End If
        Next lRow
        
        ' First, look up the requestor
        For lRow = 3 To lLastRow
            If sUser = .Cells(lRow, 2).Value Then
                ' this is it!
                sRequestorName = .Cells(lRow, 3).Value
                sRequestorEmail = .Cells(lRow, 4).Value
                Exit For
            End If
        Next lRow
        
        For lRow = 3 To lLastRow
            
            sInclude = .Cells(lRow, 1).Value
            ' If the INCLUDE is True for this person, let's log some time
            If sInclude = "Y" Then
            
                ' This is a valid user, so capture it
                sThisUserName = .Cells(lRow, 2).Value ' User name
                sThisUser = .Cells(lRow, 3).Value ' Display Name
                sThisUserEmail = .Cells(lRow, 4).Value ' Email
                
                Application.StatusBar = "Posting Time: " & sThisUser
                
                'Iterate over the issues log for the work to record for this person
                lIssueCount = UBound(oIssues.Keys)
                
                For lIssueIdx = 0 To lIssueCount
                
                    Application.StatusBar = "Posting Time: " & sThisUser _
                        & " (" & lIssueIdx + 1 & " of " & lIssueCount + 1 & ")"
                                                                        
                    ' Assemble the JSON
                    sJsonBody = RestHelper.AssembleJson(oIssues(lIssueIdx), sThisUserName)
                                        
                    ' Post it. If successful, this will return JSON with the worklog created.
                    sJson = JiraRestAPI("/rest/tempo-timesheets/3/worklogs", "POST", sBasicAuth, , , , sJsonBody)
                    
                    ' Process the returned JSON to report out and parse the return into a Dictionary for reporting
                    Set oJson = JsonConverter.ParseJson(sJson)
                
                    ' Add a worklog audit entry to the log
                    sWorklogAudit = sWorklogAudit & reportWorklog(oJson)
                
                Next ' get next oIssue
                
                ' Finished with this user, so trigger the mail
                
                ' Update status bar
                Application.StatusBar = "Posting Time: " & sThisUser _
                    & "(Sending Email)"
                                    
                'Assemble the mail message (wrap header and footer around the worklog)
                sWorklogAudit = Email.assembleEmailMsgBody(sWorklogAudit)
                
                ' Send / Display the mail
                Call Email.sendEmail(sThisUser, sThisUserEmail, sRequestorName, sRequestorEmail, sWorklogAudit)
                
                ' Clean Up
                sWorklogAudit = ""
                
            End If
            
        Next lRow
    End With
    
exit_createWorklogs:
    wksEmail.Activate
    Application.StatusBar = "Done!"
    Exit Sub

err_createWorklogs:
    MsgBox Err.Description, vbCritical, "Error: " & Err.Number
    Resume exit_createWorklogs
    
End Sub

Function reportWorklog(ByVal oWork As Dictionary) As String
On Error GoTo Err_reportWorklog

    Dim sString As String
    Dim dDate As Date
    Dim sDate As String
           
    sDate = oWork("dateStarted")
    sDate = Left(sDate, 10)
    
    sString = "<tr>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal>" & oWork("jiraWorklogId") & "</p></td>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal>" & sDate & "</p></td>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal>" & (oWork("timeSpentSeconds") / 60) & "m" & "</p></td>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal>" & oWork("issue")("key") & "</p></td>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal>" & oWork("issue")("summary") & "</p></td>" _
            & "<td valign=top style='padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal>" & oWork("comment") & "</p></td>" _
            & "</tr>"
    
    ' Debug.Print sString
    
    reportWorklog = sString

Exit_reportWorklog:
    Exit Function

Err_reportWorklog:
    MsgBox Err.Description, , "Error: " & Err.Number
    Resume Exit_reportWorklog
    
End Function

Function GetWorksheetIssues() As Dictionary
On Error GoTo Err_GetWorksheetIssues

    Dim wksTimeLogs As Worksheet
    Dim lLastRow As Long, lRow As Long, lIdx As Long
    Dim sKey As String, sDate As String, sComment As String, sIssueSummary As String, sType As String, sEpicLink As String
    Dim lTimeSpentSeconds As Long
    Dim dDate As Date
    Dim oDict As Dictionary, oDictI As Dictionary
        
    Set wksTimeLogs = ThisWorkbook.Worksheets("Issues")
    Set oDict = CreateObject("Scripting.Dictionary")

    With wksTimeLogs
    
        lLastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        For lRow = 6 To lLastRow
        
            ' assign vars
            sKey = .Cells(lRow, 1).Value ' Jira Issue Key
            dDate = .Cells(lRow, 7).Value ' Date
            sDate = Format(dDate, "yyyy-mm-ddThh:mm:ss.000+0000")  ' Converted Date
            lTimeSpentSeconds = .Cells(lRow, 6).Value * 60 ' Converted Time in Minutes to Seconds
            If Len(.Cells(lRow, 4).Value) > 0 Then
                sComment = .Cells(lRow, 4).Value ' Time Log Comment
                sComment = sComment & " (Issue Key: " & sKey & ")"
            End If
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

Sub UpdateIssues()
On Error GoTo Err_UpdateIssues
    
    Dim wks As Worksheet
    Dim r As Range
    Dim lLastRow As Long, lRow As Long, lOffset As Long
    Dim sJson As String, sParams As String, _
        sThisUser As String, sKey As String
    Dim oJson As Dictionary
    Dim dDate As Date
    
    Application.ScreenUpdating = False
        
    Set wks = ThisWorkbook.Worksheets("Issues")
        
    With wks
        
        lLastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        lOffset = 5
        
        ' fill Today's Date if Blank
        dDate = .Range("effectiveDate")
        If Year(dDate) < 2019 Then
            dDate = Date
            .Range("effectiveDate").Value = dDate
        End If
        
        ' Validate the data
        For lRow = 6 To lLastRow
            If (Len(.Cells(lRow, 1).Value) = 0 Or IsEmpty(Trim$(.Cells(lRow, 1).Value))) Then
                MsgBox "Found a data issue in row " & lRow & "." & vbNewLine _
                & "Please correct.  Process stopped and no time was posted."
                GoTo Exit_UpdateIssues
            End If
        Next
        
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
                Set oJson = GetIssuesFromJira(lRow, sKey, sParams)
                
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


Sub GetEmployees()
On Error GoTo Err_GetEmployees

    Dim wksEmployee As Worksheet
    Dim lLastRow As Long, lRow As Long, lRowIdx As Long
    Dim sJson As String, sJsonUser As String, sParams As String
    Dim oJson As Dictionary, oJsonUser As Dictionary, oJsonUserGroups As Dictionary
    Dim bTimeSheetUser As Boolean
    
    Application.ScreenUpdating = False
            
    ' Confirm our authentication is already handled
    If Len(sBasicAuth) = 0 Then
        sBasicAuth = RestHelper.SetBasicAuth()
    End If

    ' Clear Users
    Call TempoWorklogs.clearWorksheetData("Employees", 2)
    
    Set wksEmployee = ThisWorkbook.Worksheets("Employees")
    
    With wksEmployee
        
        lLastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
            
        ' Update Status Bar
        Application.StatusBar = "Retrieving users from Jira"
                               
        ' Get the Employees and put into a dictionary
        Set oJson = RestHelper.GetEmployeesFromJira()
        
        ' Use lRow and lIdx because there will be Idx values we don't write to rows
        lRow = 2 ' start at the top
        For lRowIdx = 1 To oJson("users").Count
        
            Application.StatusBar = "Processing " & (lRowIdx) & " of " & oJson("users").Count
            
            bTimeSheetUser = False
            
            ' Get the JIRA data for this user
            Set oJsonUser = RestHelper.GetUserFromJira(oJson("users")(lRowIdx)("name"))
                                       
            ' Is this user able to log time?
            bTimeSheetUser = CanLogTime(oJsonUser)
            
            If bTimeSheetUser Then
                ' Write the info to the worksheet.
                .Cells(lRow, 1).Value = oJsonUser("name")
                .Cells(lRow, 2).Value = oJsonUser("displayName")
                .Cells(lRow, 3).Value = oJsonUser("emailAddress")
                
                ' We wrote something. Move to next Row
                lRow = lRow + 1
                
            End If
        
        Next lRowIdx
    
    End With
    
Exit_GetEmployees:
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    Exit Sub

Err_GetEmployees:
    MsgBox Err.Description, vbCritical, "Error: " & Err.Number
    Resume Exit_GetEmployees
    
End Sub


Sub clearWorksheetData(ByVal wksName As String, lFirstRow As Long)
    
    Dim wks As Worksheet
    Dim lLastRow As Long, lRow As Long, lIdx As Long
 
    Set wks = ThisWorkbook.Worksheets(wksName)
    
    With wks
        lLastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        If lLastRow >= lFirstRow Then
            Rows(lFirstRow & ":" & lLastRow).EntireRow.Delete
        End If
    End With
    
End Sub

Function CanLogTime(ByVal oJson As Dictionary) As Boolean

    Dim v As Variant
    Dim b As Boolean
    
    b = False
        
    For Each v In oJson("groups")("items")
        If v("name") = "jira_timesheets_users" Then
            b = True
            Exit For
        End If
    Next
    
    CanLogTime = b
            
End Function


