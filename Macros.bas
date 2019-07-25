Attribute VB_Name = "Macros"
Option Explicit

Sub clearIssues()
    ' This is a macro assigned to a form on the issues page
    Dim wks As Worksheet
    
    Set wks = ThisWorkbook.Worksheets("Issues")
    
    Range("effectiveDate").Value = ""
    Range("adminTime").Value = ""
            
    wks.Cells(1, 1).Select
    
    Call TempoWorklogs.clearWorksheetData("Issues", 6)
End Sub

Sub clearPeople()
    ' This is a macro assigned to a form on the people page
    Dim wks As Worksheet
    
    Set wks = ThisWorkbook.Worksheets("Team Members")
    wks.Cells(1, 1).Select
    Call TempoWorklogs.clearWorksheetData("Team Members", 3)
End Sub

Sub UpdateSilverchairEmployees()
    ' This is a macro assigned to a form on the people page
    Dim wks As Worksheet
    
    Set wks = ThisWorkbook.Worksheets("Employees")
    wks.Cells(1, 1).Select
    Call TempoWorklogs.GetEmployees
End Sub

Sub createAndPostWorklogs()
    ' This is a macro assigned to a form on the issues page
    
    Call TempoWorklogs.createWorklogs

End Sub
