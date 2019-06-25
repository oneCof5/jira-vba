VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub cmdUpdateUser_Click()

    Application.ScreenUpdating = False
    
    Call TempoWorklogs.UpdateUsers
    
    Application.ScreenUpdating = True

End Sub

Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    
    Dim wks As Worksheet, wksDv As Worksheet
    Dim r As Range, rDv As Range
    Dim lRow As Long
    
    Set wksDv = ThisWorkbook.Worksheets("DV")
    Set wks = ThisWorkbook.Worksheets("Team Members")
    
    lRow = Target.Row
    
    Set r = wks.Cells(lRow, 1)
    Set rDv = wksDv.Range("A2:A3")
    
    If lRow > 2 Then
    
        With r.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Formula1:="='" & wksDv.Name & "'!" & rDv.Address
        End With
    End If
        
End Sub
