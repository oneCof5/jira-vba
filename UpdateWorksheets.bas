Attribute VB_Name = "UpdateWorksheets"
Option Base 1

Public Function extractSprintNumber(theWholeEnchilada As String) As String
    Dim regEx As New RegExp
    Dim sprint As String
    
    With regEx
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "pi ?[0-9]{2,4} ?- ?[0-9]{2,4}|sp(rint )? ?[0-9]{2,4}"
    End With
            
    ' match "Sprint xxx" or "SpXXX" variations with between 2 and 4 digits for the sprint
    If regEx.Test(theWholeEnchilada) Then
        ' set the string to the matched value
        sprint = regEx.Execute(theWholeEnchilada)(0)
        ' now extract only the number
        regEx.Pattern = "[0-9]{2,4} ?- ?[0-9]{2,4}|[0-9]{2,4}"
        ' return only the matched number
        ' Debug.Print (regEx.Execute(sprint)(0))
        extractSprintNumber = regEx.Execute(sprint)(0)
    Else
        extractSprintNumber = "None"
    End If

End Function

Public Sub UpdateBurnUpData()
    On Error GoTo Err_UpdateBurnUpData
    
    Dim wksBurn As Worksheet, wksSetup As Worksheet
    Dim lastRowIdx As Long, sprRowIdx As Long
    Dim sJQL As String
    Dim rngActiveSheet As Range
    Dim bolFirstFutureSprint As Boolean, bolUsed As Boolean
    
    bolUsed = False
    bolFirstFutureSprint = False
           
    Set wksSetup = ThisWorkbook.Sheets("Setup")
    Set wksBurn = ThisWorkbook.Sheets("BurnUp")
    
    lastRowIdx = wksBurn.UsedRange.Rows(wksBurn.UsedRange.Rows.Count).Row
    
    ' CLEAR EXISTING VALUES
    Set rngActiveSheet = wksBurn.Range(Cells(2, 5), Cells(lastRowIdx, 22))
    rngActiveSheet.Select
    Selection.ClearContents
                        
    For sprRowIdx = 2 To lastRowIdx
        ' UPDATE THE STATUS BAR TEXT
        Application.StatusBar = "Retreiving records from JIRA (Sprint " _
            & wksBurn.Cells(sprRowIdx, 1).Value & ")"
            
        ' GET THE CONTEXT ROW END DATE
        Dim strSprintEndDate As String
        Dim strThisSprintNumber As String
        strSprintEndDate = wksBurn.Cells(sprRowIdx, 2).Value
        strThisSprintNumber = wksBurn.Cells(sprRowIdx, 1).Value
        
        Dim sprDate As Date, thisDate As Date
        sprDate = DateValue(strSprintEndDate)
        thisDate = Format(Now(), "yyyy/mm/dd")
        
        If bolFirstFutureSprint = False Then
            If sprDate > thisDate Then
                bolFirstFutureSprint = True
            End If
        End If
                
        ' BUILD SECTION
        sJQL = wksSetup.Range("sJQLSourceBuild").Value
        Call generateSectionData(sprRowIdx, 5, bolFirstFutureSprint, bolUsed, strSprintEndDate, sprDate, thisDate, sJQL, strThisSprintNumber)
            
        ' ACCESS SECTION
        'sJQL = wksSetup.Range("sJQLSourceAccess").Value
        'Call generateSectionData(sprRowIdx, 11, bolFirstFutureSprint, bolUsed, strSprintEndDate, sprDate, thisDate, sJQL, strThisSprintNumber)
                    
        ' Update the totals
        
        With wksBurn
            .Cells(sprRowIdx, 17).FormulaR1C1 = "=RC[-6]+RC[-12]" ' Sprint Planned Points
            .Cells(sprRowIdx, 18).FormulaR1C1 = "=RC[-6]+RC[-12]" ' Sprint Completed Points
            .Cells(sprRowIdx, 19).FormulaR1C1 = "=RC[-6]+RC[-12]" ' Sprint Surplus/Defecit
            .Cells(sprRowIdx, 20).FormulaR1C1 = "=RC[-6]+RC[-12]" ' Total Planned Points
            If sprDate < thisDate Then
                .Cells(sprRowIdx, 21).FormulaR1C1 = "=IF(RC[-6]+RC[-12],RC[-6]+RC[-12], """" )" ' Total Completed Points
                .Cells(sprRowIdx, 22).Value = "" ' Projected Total Completed Points
            Else
                If (bolFirstFutureSprint = True And bolUsed = False) Then
                    .Cells((sprRowIdx - 1), 22).FormulaR1C1 = "=RC[-6]+RC[-12]" ' Completed Points
                End If

                .Cells(sprRowIdx, 21).Value = "" ' Total Completed Points
                .Cells(sprRowIdx, 22).FormulaR1C1 = "=IF(RC[-6]+RC[-12],RC[-6]+RC[-12], """" )" ' Projected Total Completed Points
            End If
        End With
        
        ' Have we set the first future Sprint?
        If bolFirstFutureSprint Then
            bolUsed = True
        End If
                    
    Next sprRowIdx
        
    ' UPDATE THE STATUS BAR TEXT
    Application.StatusBar = ""

Exit_UpdateBurnUpData:
    Exit Sub
    
Err_UpdateBurnUpData:
    MsgBox Err.Description, vbExclamation, "Error"
    Resume Exit_UpdateBurnUpData

End Sub

Sub generateSectionData(sprRowIdx As Long, lngColStart As Long, _
    bolFirstFutureSprint As Boolean, bolUsed As Boolean, _
    strSprintEndDate As String, _
    sprDate As Date, thisDate As Date, sJQL As String, strThisSprintNumber As String)
    
    On Error GoTo Err_generateSectionData
        
    Dim totalPointsPlan As Long, pointsPlan As Long, pointsComplete As Long, totalPointsComplete As Long, totalProjectedPointsComplete As Long
        
    ' BUILD
    Dim obj As Scripting.Dictionary

    ' GET THE ITEMS FROM JIRA
    Set obj = getStuffFromJIRA(strSprintEndDate, sJQL)
    
    ' NOW WE KNOW WE'VE GOT EVERYTHING
    Dim calculatedStoryPoints As Collection
    Set calculatedStoryPoints = getCollection(obj, strThisSprintNumber)
                
    ' WRITE TO THE SHEET
    Call writeValuesToWorksheet(sprRowIdx, lngColStart, _
        bolFirstFutureSprint, bolUsed, _
        strSprintEndDate, _
        sprDate, thisDate, _
        calculatedStoryPoints.Item(1), _
        calculatedStoryPoints.Item(2), _
        calculatedStoryPoints.Item(3))
    
Exit_generateSectionData:
    Exit Sub
    
Err_generateSectionData:
    MsgBox Err.Description, vbExclamation, "Error"
    Resume Exit_generateSectionData

End Sub


Public Function getStuffFromJIRA(theDate As String, sJQLSource As String) As Dictionary
   
    Dim oJson As Scripting.Dictionary
    Dim objDict As Scripting.Dictionary
    Dim objDictI As Scripting.Dictionary
    Dim objDictII As Scripting.Dictionary
    Dim objDictIII As Scripting.Dictionary
        
    ' set the top level dictionary object
    Set objDict = CreateObject("Scripting.Dictionary")
    
    'How many issues as of this sprint?
    Dim lngStartAt As Long: lngStartAt = 1
    Dim lngIssueIdx As Long: lngIssueIdx = 0
    Dim lngTotalIssues As Long: lngTotalIssues = 0
    Dim bolGotEmAll As Boolean: bolGotEmAll = False
    Dim sFields As String, sApiUrl As String, sJQL As String, sResponse As String

    sApiUrl = "https://jira.silverchair.com/rest/api/2/search"
    sFields = "&fields=key,status,labels,customfield_10930,customfield_10013"
    
    ' Retrieve the JQL
    sJQL = sJQLSource _
        & " AND created <= '" & Format(theDate, "yyyy/mm/dd") & "'"
    
    ' Append the fields
    sJQL = sJQL & sFields & "&expand=changelog"
    
    ' Take the JIRA API data and store in an object that represents all items (not just first 1000)
    Do Until bolGotEmAll
        ' Call subroutine to get JIRA
        sResponse = GetIssues(lngStartAt, sApiUrl, sJQL)
        
        'parse the results
        Set oJson = JsonConverter.ParseJson(sResponse)
        lngTotalIssues = CLng(oJson("total"))
        
        ' Loop over the issues pulled in this query
        For Each JIRA_ISSUE In oJson("issues")
            ' set child of O dictionary object
            Set objDictI = CreateObject("Scripting.Dictionary")
            objDictI.Item("key") = JIRA_ISSUE("key")
            objDictI.Item("status") = JIRA_ISSUE("fields")("status")("name")
            objDictI.Item("storypoints") = JIRA_ISSUE("fields")("customfield_10013")
            
            'set "sprints" grandchild of O / child of objDictI object
            Set objDictII = CreateObject("Scripting.Dictionary")
            
            If Not IsNull(JIRA_ISSUE("fields")("customfield_10930")) Then
                Dim intSprint As Integer: intSprint = 0
                For Each sprint In JIRA_ISSUE("fields")("customfield_10930")
                    objDictII.Item(intSprint) = sprint
                    intSprint = intSprint + 1
                Next
            Else
                objDictII.Item(0) = "None"
            End If
            
            ' inject the sprint into the item object
            Set objDictI("sprints") = objDictII
            
            'set "history" grandchild of O / child of objDictI object
            Set objDictII = CreateObject("Scripting.Dictionary")
            
            Dim intHistory As Integer: intHistory = 0
            
            For Each HISTORY In JIRA_ISSUE("changelog")("histories")
                For Each HISTORY_ITEM In HISTORY("items")
                    If HISTORY_ITEM("field") = "status" Then
                        ' set new instance of objDictIII for these status changes on this date
                        Set objDictIII = CreateObject("Scripting.Dictionary")

                        objDictIII.Item("created") = HISTORY("created")
                        objDictIII.Item("fromString") = HISTORY_ITEM("fromString")
                        objDictIII.Item("toString") = HISTORY_ITEM("toString")
                        
                        ' inject the history items into the history object
                        Set objDictII(intHistory) = objDictIII
                        intHistory = intHistory + 1
                    End If
                Next
            Next
            
            ' inject the history into the item object
            Set objDictI("histories") = objDictII
                        
            ' inject the item into the parent object
            Set objDict(lngIssueIdx) = objDictI
            
            ' clean up any loose ends
            Set objDictII = Nothing
                        
            lngIssueIdx = lngIssueIdx + 1
        Next
        
        If lngIssueIdx + 1 < lngTotalIssues Then
            ' dont have them all, rinse and repeat
            lngStartAt = lngStartAt + 1000
        Else
            ' have them all
            bolGotEmAll = True
        End If
    Loop

    Set getStuffFromJIRA = objDict

End Function


Public Function getCollection(objDict As Dictionary, sThisSprintNumber As String) As Collection
' http://www.geeksengine.com/article/vba-function-multiple-values.html
    Dim var As Collection
    Set var = New Collection
    Dim sSprintNumber As String
    
    Dim totalPointsPlan As Long
    Dim pointsPlan As Long
    Dim pointsComplete As Long
    Dim i As Integer
    
    For i = 0 To (objDict.Count - 1)
        
        ' add to rolling total plan as of this sprint
        If Not IsNull(objDict(i)("storypoints")) Then
            If IsNumeric(objDict(i)("storypoints")) Then
                ' Total
                totalPointsPlan = totalPointsPlan + CLng(objDict(i)("storypoints"))
            End If
        End If
        
        ' Sprints
        If Not IsNull(objDict(i)("sprints")) Then
            ' Sprints Exist, so figure out start and end
            Dim z As Integer: z = objDict(i)("sprints").Count - 1
            Dim sprintIdx As Long
            For sprintIdx = 0 To z
                If Not IsNull(objDict(i)("sprints")(sprintIdx)) Then
                    sSprintNumber = objDict(i)("sprints")(sprintIdx)
                    sSprintNumber = extractSprintNumber(sSprintNumber)
                    ' If sSprintNumber = wksBurn.Cells(rowIdx, 1).Value Then
                    If sSprintNumber = sThisSprintNumber Then
                        ' this is the context sprint
                        If Not IsNull(objDict(i)("storypoints")) Then
                            If IsNumeric(objDict(i)("storypoints")) Then
                                pointsPlan = pointsPlan + CLng(objDict(i)("storypoints"))
                            End If
                        End If

                        ' is this the last sprint and the story is complete?
                        If sprintIdx = z Then
                            Select Case objDict(i)("status")
                                Case "Complete", "Ready for Staging Validation", "Ready for Post", "Ready for Staging Post", "Ready for QA Validation", _
                                    "Ready for QA Post", "Ready for Prod Validation", "Quick Closed", "Archived", "Ready for Release"
                                    If Not IsNull(objDict(i)("storypoints")) Then
                                        If IsNumeric(objDict(i)("storypoints")) Then
                                            pointsComplete = pointsComplete + CLng(objDict(i)("storypoints"))
                                        End If
                                    End If
                            End Select
                        End If
                    End If
                End If
            Next sprintIdx
        End If
                
    Next ' Get Next Issue
    
    ' Add items to the collection
    var.Add totalPointsPlan ' "John"
    var.Add pointsPlan ' "Star"
    var.Add pointsComplete
    
    Set getCollection = var
    
End Function


Sub writeValuesToWorksheet(rowIdx As Long, lngColStart As Long, _
    bolFirstFutureSprint As Boolean, bolUsed As Boolean, _
    strSprintEndDate As String, _
    sprDate As Date, thisDate As Date, _
    totalPointsPlan As Long, pointsPlan As Long, pointsComplete As Long)
    
On Error GoTo Err_writeValuesToWorksheet
    
    Dim wksBurn As Worksheet
    
    Set wksBurn = ThisWorkbook.Sheets("BurnUp")
       
    With wksBurn
        
        .Cells(rowIdx, lngColStart).Value = pointsPlan ' Planned Points
        .Cells(rowIdx, lngColStart + 3).Value = totalPointsPlan ' Total Planned Points
    
        
        If sprDate < thisDate Then
            .Cells(rowIdx, lngColStart + 1).Value = pointsComplete ' Completed Points
            .Cells(rowIdx, lngColStart + 2).FormulaR1C1 = "=RC[-1]-RC[-2]" ' Surplus or Defecit
            If rowIdx = 2 Then
                .Cells(rowIdx, lngColStart + 4).FormulaR1C1 = "=RC[-3]" ' Total Completed Points
            Else
                .Cells(rowIdx, lngColStart + 4).FormulaR1C1 = "=R[-1]C+RC[-3]" ' Total Completed Points
            End If
            
            .Cells(rowIdx, lngColStart + 5).Value = "" ' Projected Total Completed Points
        Else
            If (bolFirstFutureSprint = True And bolUsed = False) Then
                .Cells((rowIdx - 1), lngColStart + 5).FormulaR1C1 = "=RC[-1]" ' Completed Points
            End If
            .Cells(rowIdx, lngColStart + 1).Value = "" ' Completed Points
            .Cells(rowIdx, lngColStart + 2).Value = "" ' Surplus or Defecit
            .Cells(rowIdx, lngColStart + 4).Value = "" ' Total Completed Points
            .Cells(rowIdx, lngColStart + 5).FormulaR1C1 = "=R[-1]C+RC[-5]" ' Projected Total Completed Points
        End If
    End With
        

Exit_writeValuesToWorksheet:
    Exit Sub

Err_writeValuesToWorksheet:
    MsgBox Err.Description, vbExclamation, "Error"
    Resume Exit_writeValuesToWorksheet

End Sub

Sub SprintBoards()

    Dim sApiUrl As String
    Dim sBoardID As String
    Dim sJQL As String: sJQL = ""
    Dim lngStartAt As Long: lngStartAt = 1
    
    sBoardID = 369
    
    sApiUrl = "https://jira.silverchair.com/rest/agile/1.0/board/" _
        & sBoardID & "/sprint"
    ' sApiUrl = "https://jira.silverchair.com/rest/api/2/search"

    sResponse = GetIssues(lngStartAt, sApiUrl, sJQL)
    'Debug.Print sResponse

End Sub

