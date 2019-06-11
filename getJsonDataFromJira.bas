Attribute VB_Name = "getJsonDataFromJIRA"
Public usernamep As String

Public Function GetIssues(lngStartAt As Long, apiUrl As String, query As String) As String

    Dim oJiraService As New MSXML2.XMLHTTP60
    Dim oJson As Object
    Dim sUser, sPass, sUrl As String
        
    ' set the user / password
    If usernamep = "" Then
        ' get the user name
        sUser = InputBox("JIRA User Name", "Enter JIRA Credentials")
        sPass = InputBoxDK("JIRA Password", "Enter JIRA Credentials")
        If sUser + ":" + sPass = ":" Then
            usernamep = "cancel"
        Else
            usernamep = Base64Encode(sUser + ":" + sPass)
        End If
    End If
     
    If usernamep = "cancel" Then
        MsgBox ("User canceled login.")
    Else
        With oJiraService
            
            sUrl = apiUrl _
            & "?startAt=" & lngStartAt _
            & "&maxResults=1000"

             sUrl = sUrl + "&jql=" + query
             
             .Open "GET", sUrl
             
             .setRequestHeader "Content-Type", "application/json"
             .setRequestHeader "Accept", "application/json"
             .setRequestHeader "Authorization", "Basic " & usernamep
             .send
            
             If .Status = "401" Then
                 MsgBox "Something wrong with query in GetIssues, check your network connection, :  " + .responseText
                 GetIssues = ""
             Else
                 GetIssues = oJiraService.responseText
             End If
        End With
    End If
End Function

