VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton1_Click()
    
    Application.ScreenUpdating = False
    
    Call TempoWorklogs.UpdateIssues
    
    Application.ScreenUpdating = True


End Sub

Private Sub Worksheet_Change(ByVal Target As Excel.Range)

    Dim lRow As Long

    lRow = Target.Row
        
    If (Target.Column > 0 And Target.Column < 5) Then
        ' Calculated Duration
        Me.Cells(lRow, 5).NumberFormat = "h:mm"
        Me.Cells(lRow, 5).FormulaR1C1 = "=IF(RC[-2]-RC[-3]>0, RC[-2]-RC[-3],"""")"
            
        ' Time Spent in Minutes
        Me.Cells(lRow, 6).NumberFormat = "#,##0"
        Me.Cells(lRow, 6).FormulaR1C1 = "=IF(ISERROR(RC[-1]*1440), """",RC[-1]*1440)"
    End If

End Sub
