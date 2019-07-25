Attribute VB_Name = "Email"
Option Explicit

Function assembleEmailMsgBody(ByVal sBody As String) As String
    
    Dim sBodyText As String, sAdmTime As String, sTotalTime As String
    Dim wks As Worksheet
    
    Set wks = ThisWorkbook.Worksheets("Issues")
    sAdmTime = Format(wks.Range("adminTime").Value, "#,##0")
    sTotalTime = Format(wks.Range("totalTime").Value, "#,##0")
                    
    'Write the header for this user
    
    sBodyText = Email.emailIntroSection(sTotalTime, sAdmTime)
    
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
            
    'Append the log entries
    sBodyText = sBodyText & sBody
    sBodyText = sBodyText & "</table>" _
            & "<p><o:p></o:p></p>"
    
    assembleEmailMsgBody = sBodyText

End Function

Function emailIntroSection(sTotalTime As String, Optional ByVal sAdmTime As String) As String

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
        & "<img src=""cid:timesheet.jpg"" height=128><o:p></o:p>" _
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

Sub sendEmail(ByVal sReceiverName As String, ByVal sReceiverEmail As String, _
    ByVal sSenderName As String, ByVal sSenderEmail As String, ByVal sMessageBody As String)
    On Error GoTo err_sendEmail
    
    Dim olApp As Outlook.Application
    Dim olMsg As Outlook.MailItem
    Dim sPath As String, sFile As String, sBody As String, _
        sSubject As String, sMsg As String, sMsDiv As String
    Dim bCreated As Boolean, bPreview As Boolean
    Dim shp As Shape
    Dim wksEmail As Worksheet, wksImages As Worksheet
    
    Set wksEmail = ThisWorkbook.Worksheets("Email")
    Set wksImages = ThisWorkbook.Worksheets("Images")
    
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
    
    ' attach the image
    sPath = ThisWorkbook.Path
    sFile = sPath & "\timesheet.jpg"

    Call ExportMyPicture(wksEmail.Range("meme").Value)
        
    ' Create a new instance of outlook
    Set olApp = New Outlook.Application
    Set olMsg = olApp.CreateItem(olMailItem)
    
    With olMsg
        .Display
    
        sMsg = .HTMLBody
        sMsDiv = "<div class=WordSection1><p class=MsoNormal><o:p>"
    
        .To = sReceiverEmail
        .CC = sSenderEmail
        .Subject = "Timesheet Entry Posted! " _
            & sSenderName & " posted time for " & sReceiverName
        .Attachments.Add sFile, 1, 0

        .HTMLBody = Replace(sMsg, sMsDiv, sMsDiv & sMessageBody)
        
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


Sub test()
    Dim sJson As String, sWorklogAudit As String
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
    
    ' Add a worklog audit entry to the log
    sWorklogAudit = sWorklogAudit & TempoWorklogs.reportWorklog(oJson)
                    
    'Assemble the mail message (wrap header and footer around the worklog)
    sWorklogAudit = Email.assembleEmailMsgBody(sWorklogAudit)
    
    ' Send / Display the mail
    Call Email.sendEmail("Me", "cpearson@silverchair.com", "You", "cpearson.silverchair@gmail.com", sWorklogAudit)
    
End Sub


Sub ExportMyPicture(sName As String)
On Error GoTo Finish
    
    Dim MyChart As String, sFile As String
    Dim lPicWidth As Long, lPicHeight As Long
    Dim shpPicture As Shape
    Dim wks As Worksheet
        
    Application.ScreenUpdating = False
    
    
    Set wks = ThisWorkbook.Worksheets("Images")
    
    ' Define the output
    sFile = ThisWorkbook.Path & "\timesheet.jpg"
    
    ' MyPicture = Selection.Name
    Set shpPicture = wks.Shapes(sName)
    With shpPicture
        lPicHeight = .Height
        lPicWidth = .Width
    End With
    
    Charts.Add
    ActiveChart.Location Where:=xlLocationAsObject, Name:="Images"
    Selection.Border.LineStyle = 0
    MyChart = Selection.Name & " " & Split(ActiveChart.Name, " ")(2)
    
    With wks
        With .Shapes(MyChart)
            .Width = lPicWidth
            .Height = lPicHeight
        End With
        
        .Shapes(sName).Copy
        
        With ActiveChart
            .ChartArea.Select
            .Paste
        End With
        
        .ChartObjects(1).Chart.Export Filename:=sFile, FilterName:="jpg"
        .Shapes(MyChart).Cut
    End With
    
    Application.ScreenUpdating = True
    Exit Sub
    
Finish:
    MsgBox "You must select a picture"
End Sub

Public Sub updateMemeChoices()

    Dim wksEmail As Worksheet, wksImages As Worksheet
    Dim r As Range
    Dim s As String
    Dim shp As Shape
    
    Set wksEmail = ThisWorkbook.Worksheets("Email")
    Set wksImages = ThisWorkbook.Worksheets("Images")
    
    For Each shp In wksImages.Shapes
        If shp.Type = msoPicture Then
            If Len(s) = 0 Then
                s = shp.Name
            Else:
                s = s & "," & shp.Name
            End If
            ' Debug.Print s
        End If
    Next shp
    
    Set r = wksEmail.Range("meme")
    
    With r.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:=s
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub
