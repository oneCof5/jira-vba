Attribute VB_Name = "Users"
Option Explicit

Sub UpdateUsers()
On Error GoTo Err_UpdateUsers

    Dim wb As Workbook
    Dim wksSetup As Worksheet, wksTeam As Worksheet
    Dim lLastRow As Long, lRow As Long
    Dim sJiraBaseUrl As String, sJson As String, sParam As String, _
        sThisUser As String
    Dim oJson As Dictionary
        
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
        For lRow = 2 To lLastRow
            
            ' Update Status Bar
            Application.StatusBar = "Processing row " & lRow & " of " & lLastRow
            
            sThisUser = .Cells(lRow, 2).Value
            
            ' Check this row's user name
            sParam = "username=" & sThisUser
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
    Exit Sub

Err_UpdateUsers:
    MsgBox Err.Description, vbCritical, "Error: " & Err.Number
    Resume Exit_UpdateUsers
    
End Sub
