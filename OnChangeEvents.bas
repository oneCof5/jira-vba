Attribute VB_Name = "OnChangeEvents"
Option Explicit

' Create the OnChange events for Worksheets
Private Sub CreateOnChange()
On Error GoTo Err_CreateOnChange

    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim LineNum As Long
    Dim sArray(0 To 3) As String, sProcName As String
    Dim i As Integer
    
    Set VBProj = ActiveWorkbook.VBProject
    
    sArray(0) = "Sheet1"
    sArray(1) = "Sheet3"
    sArray(2) = "Sheet4"
    sArray(3) = "Sheet5"
    
    For i = 0 To UBound(sArray)
        
        Set VBComp = VBProj.VBComponents(sArray(i))
        Set CodeMod = VBComp.CodeModule
        
        Select Case sArray(i)
            Case "Sheet1"
                sProcName = "TeamMembersOnChange(Target)"
            Case "Sheet3"
                sProcName = "IssuesOnChange(Target)"
            Case "Sheet4"
                sProcName = "SetupOnChange(Target)"
            Case "Sheet5"
                sProcName = "EmailOnChange(Target)"
            Case Else
                Exit Sub
        End Select
        
        With CodeMod
        
            ' Delete the existing
            .DeleteLines 1, (.CountOfDeclarationLines + .CountOfLines)
        
            LineNum = .CreateEventProc("Change", "Worksheet")
            LineNum = LineNum + 1
            .InsertLines LineNum, "     "
            LineNum = LineNum + 1
            .InsertLines LineNum, "     Call OnChangeEvents." _
                & sProcName
        End With
        
        sProcName = ""
        
    Next
    
Exit_CreateOnChange:
    Exit Sub

Err_CreateOnChange:
    MsgBox Err.Description
    Resume Exit_CreateOnChange

End Sub


Public Sub TeamMembersOnChange(Target As Range)

    Dim wks As Worksheet
    Dim lRow As Long, lCol As Long
    Dim rUser As Range, rInclude As Range
    
    Set wks = ThisWorkbook.Worksheets("Team Members")
    
    lRow = Target.Row
    lCol = Target.Column
    
    If (lRow > 2) And (lCol < 3) Then
    
        Application.EnableEvents = False
        
        Set rInclude = wks.Cells(lRow, 1)
        Set rUser = wks.Cells(lRow, 2)
    
        ' Include
        With rInclude.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Formula1:="Y,N"
        End With
                    
        ' User Name
        With rUser.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Formula1:="=Employees!$A$2:$A$300"
        End With
        
        With wks
            
            If Len(rUser.Value) > 0 Then
                rUser.Value = LCase(rUser.Value)
                .Cells(lRow, 3).Formula = "=VLOOKUP(B" & lRow & ",Employees!A:C,2,FALSE)"
                .Cells(lRow, 4).Formula = "=VLOOKUP(B" & lRow & ",Employees!A:C,3,FALSE)"
            End If
            
        End With
    
        Application.EnableEvents = True
    
    End If
    
End Sub

Public Sub IssuesOnChange(Target As Range)
On Error GoTo Err_IssuesOnChange

    Dim wks As Worksheet
    Dim lRow As Long, lCol As Long
    Dim rIssueKey As Range, rStartTime As Range, rEndTime As Range, rComment As Range
        
    ' If this is a delete of the entire row, skip it!
    If Target.Rows.count > 0 And Target.Columns.count = Columns.count Then
        GoTo Exit_IssuesOnChange
    End If
        
    lRow = Target.Row
    lCol = Target.Column
    
    If (lRow > 5) And (lCol = 1) Then
    
        Application.EnableEvents = False
        
        Set wks = ThisWorkbook.Worksheets("Issues")
        
        With wks
            
            Set rIssueKey = .Cells(lRow, 1)
            Set rStartTime = .Cells(lRow, 2)
            Set rEndTime = .Cells(lRow, 3)

            ' If we added an issue, ensure it's a project key in upper case
            If Len(rIssueKey.Value) > 0 Then
                rIssueKey.Value = UCase(rIssueKey.Value)
            'If the issue key is blank, skip the rest
            ElseIf Len(rIssueKey.Value) = 0 Then
                GoTo Exit_IssuesOnChange
            End If
            
            'If there's no time already, set to now
            If rStartTime.Value = 0 Then
                rStartTime.Value = Format(Time(), "HH:MM")
            End If
            
            'if the row above's end time is blank, set it too.
            If lRow > 6 Then
                If .Cells(lRow - 1, 3).Value = 0 Then
                    wks.Cells(lRow - 1, 3).Value = rStartTime.Value
                End If
            End If
                            
            ' Calculated Duration
            .Cells(lRow, 5).NumberFormat = "h:mm"
            .Cells(lRow, 5).FormulaR1C1 = "=IF(RC[-2]-RC[-3]>0, RC[-2]-RC[-3],"""")"
                
            ' Time Spent in Minutes
            .Cells(lRow, 6).NumberFormat = "#,##0"
            .Cells(lRow, 6).FormulaR1C1 = "=IF(ISERROR(RC[-1]*1440), """",RC[-1]*1440)"
        
        End With
        
    End If

Exit_IssuesOnChange:
    Application.EnableEvents = True
    Exit Sub
    
Err_IssuesOnChange:
    MsgBox Err.Description
    Resume Exit_IssuesOnChange

End Sub

Public Sub SetupOnChange(Target As Range)

    If (Target.Row = 1 And Target.Column = 2) Then
        sBaseUrl = RestHelper.SetBaseJiraUrl(Range("sJiraRoot").Value)
        ' Debug.Print sBaseUrl
    End If

End Sub

Public Sub EmailOnChange(Target)

    Dim sFile As String
    Dim Pic As Picture
    Dim rImageCell As Range
    Dim wks As Worksheet
    Dim shp As Shape
        
    If (Target.Row = 4 And Target.Column = 2) Then
        ' image was changed. update display.
        ' me.Range("memePreview").

        Set wks = ThisWorkbook.Worksheets("Email")
        
        ' attach the image
        sFile = ThisWorkbook.Path & "\timesheet.jpg"
        
        Call ExportMyPicture(wks.Range("meme").Value)
        
        For Each shp In wks.Shapes
            ' Debug.Print "Name=" & shp.Name & " | Type=" & shp.Type
            If (shp.Type = 11 Or shp.Type = 13) Then
                shp.Delete
            End If
        Next shp
        
        Set rImageCell = wks.Range("memePreview").MergeArea
    
        Set Pic = wks.Pictures.Insert(sFile)
        
        With Pic
            .ShapeRange.LockAspectRatio = msoTrue
            .Left = rImageCell.Left
            .Top = rImageCell.Top
'            .Width = rImageCell.Width
            .Height = rImageCell.Height
        End With
        
        wks.Activate

    End If

End Sub
    
Private Sub ThisWorkbookOnOpen()

    sBaseUrl = RestHelper.SetBaseJiraUrl(Range("sJiraRoot").Value)
    'Debug.Print sBaseUrl
    
    ' Update meme drop down
    Call Email.updateMemeChoices

End Sub



