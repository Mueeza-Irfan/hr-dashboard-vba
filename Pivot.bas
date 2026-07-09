Attribute VB_Name = "Pivot"
'  STEP 2 - PivotTables side by side in "Pivot" sheet
'  Dataset: Employee HR data (Sheet "Data", columns A:W)
'
'  Columns: A Employee ID | B Employee Name | C Gender | D Age |
'           E Department | F Designation | G Joining Date |
'           H Employment Status | I Work Mode | J Education |
'           K Hire Source | L City | M Manager | N Salary | O Bonus |
'           P Attendance % | Q Leave Days | R Overtime Hours |
'           S Performance Rating | T Training Hours |
'           U Satisfaction Score | V Promotion | W Attrition Reason
'
'  Layout (all on ROW 2, titles on ROW 1, 1-col gap between blocks):
'
'  PT_Department        -> A2   (3 cols: Department | Headcount | Avg Salary)
'  PT_Gender            -> E2   (2 cols: Gender | Headcount)
'  PT_JoiningTrend      -> H2   (2 cols: Year | Headcount)
'  PT_EmploymentStatus  -> K2   (2 cols: Status | Headcount)
'  PT_WorkMode          -> N2   (2 cols: Work Mode | Headcount)
'  PT_Designation       -> Q2   (3 cols: Designation | Avg Salary | Avg Bonus)
'  PT_City              -> U2   (2 cols: City | Headcount)
'  PT_Education         -> X2   (2 cols: Education | Headcount)
'  PT_AttritionReason   -> AA2  (2 cols: Attrition Reason | Headcount)
'  PT_Promotion         -> AD2  (2 cols: Promotion | Headcount)
'  PT_Performance       -> AG2  (4 cols: Department | Avg Perf | Avg Satisfaction | Avg Attendance)
'  PT_TopSalary         -> AL2  (2 cols: Employee Name | Salary - Top 5 only)
'-------------------------------------------------------------
Sub CreatePivotSheets()
    Dim wsData  As Worksheet
    Dim wsPvt   As Worksheet
    Dim pc      As PivotCache
    Dim pt      As PivotTable
    Dim lastRow As Long
    Dim sField  As String

    ' --- Validate Data sheet ---
    Set wsData = Nothing
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("Data")
    On Error GoTo 0
    If wsData Is Nothing Then
        MsgBox "Sheet named 'Data' not found!", vbCritical
        Exit Sub
    End If

    ' --- Get existing "Pivot" sheet ---
    Set wsPvt = Nothing
    On Error Resume Next
    Set wsPvt = ThisWorkbook.Sheets("Pivot")
    On Error GoTo 0
    If wsPvt Is Nothing Then
        MsgBox "Sheet named 'Pivot' not found!" & vbNewLine & _
               "Please create a sheet named 'Pivot' first.", vbCritical
        Exit Sub
    End If

    ' --- Clear all existing content + PivotTables (clean re-run) ---
    ' FIX: Cells.Clear alone removes all pivot table definitions too -
    ' no need for a separate loop over wsPvt.PivotTables, which was
    ' causing a 1004 error because clearing a pivot mid-loop shifts
    ' the collection's indices out from under a forward For Each.
    wsPvt.Cells.Clear

    ' --- Build shared PivotCache ---
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:="Data!$A$1:$W$" & lastRow)

    ' =========================================================
    ' PT_Department -> A2  (2 cols: Department | Headcount )
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("A2"), _
        TableName:="PT_Department")
    With pt
        .PivotFields("Department").Orientation = xlRowField
        .PivotFields("Department").Position = 1
        With .PivotFields("Employee ID")
            .Orientation = xlDataField
            .Function = xlCount
            .NumberFormat = "#,##0"
            .Name = "Headcount"
        End With
         
    End With

    ' =========================================================
    ' PT_Gender -> E2  (2 cols: Gender | Headcount)
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("E2"), _
        TableName:="PT_Gender")
    With pt
        .PivotFields("Gender").Orientation = xlRowField
        .PivotFields("Gender").Position = 1
        With .PivotFields("Employee ID")
            .Orientation = xlDataField
            .Function = xlCount
            .NumberFormat = "#,##0"
            .Name = "Headcount"
        End With
    End With

    ' =========================================================
    ' PT_JoiningTrend -> H2  (2 cols: Year | Headcount)
    ' =========================================================
    ' =========================================================
    ' PT_JoiningTrend -> H2  (3 cols: Year | Quarter | Headcount)
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("H2"), _
        TableName:="PT_JoiningTrend")
    With pt
        .PivotFields("Joining Date").Orientation = xlRowField
        .PivotFields("Joining Date").Position = 1
        With .PivotFields("Employee ID")
            .Orientation = xlDataField
            .Function = xlCount
            .NumberFormat = "#,##0"
            .Name = "Headcount"
        End With
    End With
    ' Group Joining Date by YEAR and QUARTER separately
    ' (so 2024-Qtr1 and 2025-Qtr1 stay distinct, not merged)
    On Error Resume Next
    pt.PivotFields("Joining Date").Group Periods:=Array(False, False, False, False, False, True, True)
    On Error GoTo 0

    ' =========================================================
    ' PT_EmploymentStatus -> K2  (2 cols: Employment Status | Headcount)
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("K2"), _
        TableName:="PT_EmploymentStatus")
    With pt
        .PivotFields("Employment Status").Orientation = xlRowField
        .PivotFields("Employment Status").Position = 1
        With .PivotFields("Employee ID")
            .Orientation = xlDataField
            .Function = xlCount
            .NumberFormat = "#,##0"
            .Name = "Headcount"
        End With
    End With

    ' =========================================================
    ' PT_WorkMode -> N2  (2 cols: Work Mode | Headcount)
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("N2"), _
        TableName:="PT_WorkMode")
    With pt
        .PivotFields("Work Mode").Orientation = xlRowField
        .PivotFields("Work Mode").Position = 1
        With .PivotFields("Employee ID")
            .Orientation = xlDataField
            .Function = xlCount
            .NumberFormat = "#,##0"
            .Name = "Headcount"
        End With
    End With

    ' =========================================================
    ' PT_Designation -> Q2  (3 cols: Designation | Avg Salary | Avg Bonus)
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("Q2"), _
        TableName:="PT_Designation")
    With pt
        .PivotFields("Designation").Orientation = xlRowField
        .PivotFields("Designation").Position = 1
        With .PivotFields("Salary")
            .Orientation = xlDataField
            .Function = xlAverage
            .NumberFormat = "#,##0.00"
            .Name = "Avg Salary"
        End With
        With .PivotFields("Bonus")
            .Orientation = xlDataField
            .Function = xlAverage
            .NumberFormat = "#,##0.00"
            .Name = "Avg Bonus"
        End With
    End With

    ' =========================================================
    ' PT_City -> U2  (2 cols: City | Headcount)
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("U2"), _
        TableName:="PT_City")
    With pt
        .PivotFields("City").Orientation = xlRowField
        .PivotFields("City").Position = 1
        With .PivotFields("Employee ID")
            .Orientation = xlDataField
            .Function = xlCount
            .NumberFormat = "#,##0"
            .Name = "Headcount"
        End With
    End With

    ' =========================================================
    ' PT_Education -> X2  (2 cols: Education | Headcount)
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("X2"), _
        TableName:="PT_Education")
    With pt
        .PivotFields("Education").Orientation = xlRowField
        .PivotFields("Education").Position = 1
        With .PivotFields("Employee ID")
            .Orientation = xlDataField
            .Function = xlCount
            .NumberFormat = "#,##0"
            .Name = "Headcount"
        End With
    End With

    ' =========================================================
    ' PT_AttritionReason -> AA2  (2 cols: Attrition Reason | Headcount)
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("AA2"), _
        TableName:="PT_AttritionReason")
    With pt
        .PivotFields("Attrition Reason").Orientation = xlRowField
        .PivotFields("Attrition Reason").Position = 1
        With .PivotFields("Employee ID")
            .Orientation = xlDataField
            .Function = xlCount
            .NumberFormat = "#,##0"
            .Name = "Headcount"
        End With
    End With

    ' =========================================================
    ' PT_Promotion -> AD2  (2 cols: Promotion | Headcount)
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("AD2"), _
        TableName:="PT_Promotion")
    With pt
        .PivotFields("Promotion").Orientation = xlRowField
        .PivotFields("Promotion").Position = 1
        With .PivotFields("Employee ID")
            .Orientation = xlDataField
            .Function = xlCount
            .NumberFormat = "#,##0"
            .Name = "Headcount"
        End With
    End With

    ' =========================================================
    ' PT_Performance -> AG2  (4 cols: Department | Avg Performance | Avg Satisfaction | Avg Attendance)
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("AG2"), _
        TableName:="PT_Performance")
    With pt
        .PivotFields("Department").Orientation = xlRowField
        .PivotFields("Department").Position = 1
        With .PivotFields("Performance Rating")
            .Orientation = xlDataField
            .Function = xlAverage
            .NumberFormat = "0.0"
            .Name = "Avg Performance"
        End With
        With .PivotFields("Satisfaction Score")
            .Orientation = xlDataField
            .Function = xlAverage
            .NumberFormat = "0.0"
            .Name = "Avg Satisfaction"
        End With
        With .PivotFields("Attendance %")
            .Orientation = xlDataField
            .Function = xlAverage
            .NumberFormat = "0.0%"
            .Name = "Avg Attendance"
        End With
    End With

    ' =========================================================
    ' PT_TopSalary -> AL2  (2 cols: Employee Name | Salary) - Top 5 only
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("AL2"), _
        TableName:="PT_TopSalary")
    With pt
        .PivotFields("Employee Name").Orientation = xlRowField
        .PivotFields("Employee Name").Position = 1
        With .PivotFields("Salary")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
            .Name = "Total Salary"
        End With
    End With
    sField = pt.DataFields(1).Name
    pt.PivotFields("Employee Name").AutoSort xlDescending, sField
    On Error Resume Next
    pt.PivotFields("Employee Name").PivotFilters.Add2 Type:=xlTop10Items, Value1:=5
    On Error GoTo 0

    ' --- Write titles above each PivotTable in row 1 ---
    Call WritePivotTitles(wsPvt)

    ' --- Auto-fit all used columns ---
    wsPvt.Columns("A:AM").AutoFit

    ' --- Freeze top row so headers stay visible when scrolling ---
    wsPvt.Activate
    wsPvt.Range("A3").Select
    ActiveWindow.FreezePanes = True

End Sub


'-------------------------------------------------------------
'  HELPER - Write bold green title above each PivotTable
'-------------------------------------------------------------
Sub WritePivotTitles(wsPvt As Worksheet)
    Dim darkGreen  As Long: darkGreen = RGB(30, 110, 60)
    Dim lightGreen As Long: lightGreen = RGB(226, 245, 228)
    Dim white      As Long: white = RGB(255, 255, 255)

    Dim titles(11, 2) As String
    titles(0, 0) = "Headcount by Department": titles(0, 1) = "A": titles(0, 2) = "B"
    titles(1, 0) = "Headcount by Gender": titles(1, 1) = "E": titles(1, 2) = "F"
    titles(2, 0) = "Joining Trend by Year": titles(2, 1) = "H": titles(2, 2) = "I"
    titles(3, 0) = "Headcount by Employment Status": titles(3, 1) = "K": titles(3, 2) = "L"
    titles(4, 0) = "Headcount by Work Mode": titles(4, 1) = "N": titles(4, 2) = "O"
    titles(5, 0) = "Avg Salary & Bonus by Designation": titles(5, 1) = "Q": titles(5, 2) = "S"
    titles(6, 0) = "Headcount by City": titles(6, 1) = "U": titles(6, 2) = "V"
    titles(7, 0) = "Headcount by Education": titles(7, 1) = "X": titles(7, 2) = "Y"
    titles(8, 0) = "Attrition Reason Breakdown": titles(8, 1) = "AA": titles(8, 2) = "AB"
    titles(9, 0) = "Promotion Status": titles(9, 1) = "AD": titles(9, 2) = "AE"
    titles(10, 0) = "Performance & Satisfaction by Dept": titles(10, 1) = "AG": titles(10, 2) = "AJ"
    titles(11, 0) = "Top 5 Employees by Salary": titles(11, 1) = "AL": titles(11, 2) = "AM"

    Dim i As Integer
    For i = 0 To 11
        Dim titleRange As String
        titleRange = titles(i, 1) & "1:" & titles(i, 2) & "1"

        With wsPvt.Range(titleRange)
            .Merge
            .Value = titles(i, 0)
            .Font.Bold = True
            .Font.Size = 14
            .Font.Color = white
            .Font.Name = "Calibri"
            .Interior.Color = darkGreen
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = 20
        End With

        wsPvt.Range(titleRange).Borders(xlEdgeBottom).LineStyle = xlContinuous
        wsPvt.Range(titleRange).Borders(xlEdgeBottom).Color = lightGreen
        wsPvt.Range(titleRange).Borders(xlEdgeBottom).Weight = xlThin
    Next i

    wsPvt.Rows(1).RowHeight = 20
End Sub


'=============================================================
'  UTILITY: Refresh all PivotTables (run after data changes)
'=============================================================
Sub RefreshAllPivots()
    Dim wsPvt As Worksheet
    Dim pt    As PivotTable

    On Error Resume Next
    Set wsPvt = ThisWorkbook.Sheets("Pivot")
    On Error GoTo 0

    If wsPvt Is Nothing Then
        MsgBox "Sheet named 'Pivot' not found!", vbCritical
        Exit Sub
    End If

    For Each pt In wsPvt.PivotTables
        pt.RefreshTable
    Next pt

    MsgBox "All PivotTables refreshed successfully!", vbInformation, "Done"
End Sub




