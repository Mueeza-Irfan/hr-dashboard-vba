Attribute VB_Name = "Dashboard"
'=============================================================
'  HR WORKFORCE DASHBOARD - GREEN THEME (IMPROVED)
'  Built from: Sheet "Data" (employee records) +
'              Sheet "Pivot" (12 PivotTables: PT_Department,
'              PT_Gender, PT_JoiningTrend, PT_EmploymentStatus,
'              PT_WorkMode, PT_Designation, PT_City, PT_Education,
'              PT_AttritionReason, PT_Promotion, PT_Performance,
'              PT_TopSalary)
'
'  NOTE: Adjust the two flagged formulas in BuildKPICards
'        ("Active" / "Yes") if your Employment Status or
'        Promotion column uses different text values.
'
'  CHANGES vs original:
'   1. GroupJoiningTrendByYear  -> fixes the unreadable
'      "barcode" joining trend chart by grouping raw dates
'      into actual Years.
'   2. CleanAttritionReasonPivot -> hides the "(blank)" item
'      (Active employees have no attrition reason) so the
'      chart isn't dominated by a meaningless bar.
'   3. Chart 5 (Salary & Bonus) -> Bonus is now a line on a
'      secondary axis instead of a nearly-invisible column.
'   4. Category axis labels rotated + font unified for
'      Department / Designation / Attrition charts.
'   5. Data label number formats added where missing.
'-------------------------------------------------------------
'  MAIN ENTRY - for first-time build
'-------------------------------------------------------------
Sub BuildDashboard()
    Call DashboardMain
    MsgBox "Dashboard created successfully!", vbInformation, "Done"
End Sub

'-------------------------------------------------------------
'  Button re-run entry point
'-------------------------------------------------------------
Sub RefreshDashboard()
    Call DashboardMain
    MsgBox "Dashboard refreshed successfully!", vbInformation, "Done"
End Sub


'-------------------------------------------------------------
'  CORE LOGIC
'-------------------------------------------------------------
Private Sub DashboardMain()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Dim wsPvt As Worksheet
    Set wsPvt = Nothing
    On Error Resume Next
    Set wsPvt = ThisWorkbook.Sheets("Pivot")
    On Error GoTo 0
    If wsPvt Is Nothing Then
        MsgBox "Sheet 'Pivot' not found!", vbCritical
        GoTo CleanUp
    End If

    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Dashboard")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = "Dashboard"
        ws.Tab.Color = RGB(30, 110, 60)
    End If

    ws.Cells.Clear
    ws.Cells.Interior.ColorIndex = xlNone

    Dim chtObj As ChartObject
    For Each chtObj In ws.ChartObjects
        chtObj.Delete
    Next chtObj

    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ' Turn off Grand Total rows on the pivots we chart
    Call PrepPivotsForCharts(wsPvt)

    ' NEW: fix the two data-quality issues at the pivot level
    Call GroupJoiningTrendByYear(wsPvt)
    Call CleanAttritionReasonPivot(wsPvt)

    Call SetupDashboardLayout(ws)
    Call BuildHeader(ws)
    Call BuildKPICards(ws, wsPvt)
    Call CreateAllCharts(ws, wsPvt)
    Call AddSlicersAndTimeline(ws, wsPvt)
    Call AddMacroButton(ws)

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub


'-------------------------------------------------------------
'  Disable row grand totals on the PivotTables that feed charts
'-------------------------------------------------------------
Sub PrepPivotsForCharts(wsPvt As Worksheet)
    Dim names As Variant
    Dim i     As Integer
    names = Array("PT_Department", "PT_Gender", "PT_JoiningTrend", _
                  "PT_EmploymentStatus", "PT_Designation", "PT_AttritionReason")

    For i = 0 To UBound(names)
        On Error Resume Next
        wsPvt.PivotTables(CStr(names(i))).RowGrand = False
        On Error GoTo 0
    Next i
End Sub


'-------------------------------------------------------------
'  NEW: Group the Joining Date field on PT_JoiningTrend into
'  Years only, so the chart shows ~5-8 bars/points instead of
'  a smear of 2,000 individual dates.
'  ADJUST "Joining Date" below if your field is named differently.
'-------------------------------------------------------------
Sub GroupJoiningTrendByYear(wsPvt As Worksheet)
    Dim pt As PivotTable
    Dim pf As PivotField

    On Error Resume Next
    Set pt = wsPvt.PivotTables("PT_JoiningTrend")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    On Error Resume Next
    Set pf = pt.PivotFields("Joining Date")
    On Error GoTo 0
    If pf Is Nothing Then Exit Sub

    On Error Resume Next
    ' Ungroup first in case it was previously grouped by Month/Quarter
    pf.LabelRange.Ungroup
    ' Periods array = Seconds, Minutes, Hours, Days, Months, Quarters, Years
    pf.LabelRange.Group Start:=True, End:=True, _
        Periods:=Array(False, False, False, False, False, False, True)
    On Error GoTo 0
End Sub


'-------------------------------------------------------------
'  NEW: Hide the "(blank)" item on PT_AttritionReason so the
'  chart isn't dominated by Active employees (who have no
'  attrition reason). Only real resignation reasons remain.
'  ADJUST the row field index below if Attrition Reason isn't
'  the first row field in that pivot.
'-------------------------------------------------------------
Sub CleanAttritionReasonPivot(wsPvt As Worksheet)
    Dim pt As PivotTable
    Dim pf As PivotField

    On Error Resume Next
    Set pt = wsPvt.PivotTables("PT_AttritionReason")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    On Error Resume Next
    Set pf = pt.RowFields(1)
    pf.PivotItems("(blank)").Visible = False
    On Error GoTo 0
End Sub


'-------------------------------------------------------------
'  Layout: column widths + row heights
'-------------------------------------------------------------
Sub SetupDashboardLayout(ws As Worksheet)
    Dim c As Integer
    For c = 1 To 17
        ws.Columns(c).ColumnWidth = 9.5
    Next c

    ws.Rows("1:2").RowHeight = 20
    ws.Rows("3:3").RowHeight = 5
    ws.Rows("4:6").RowHeight = 16
    ws.Rows("7:7").RowHeight = 5

    Dim r As Integer
    For r = 8 To 31
        ws.Rows(r).RowHeight = 15
    Next r

    ws.Rows("32:32").RowHeight = 20
    For r = 33 To 46
        ws.Rows(r).RowHeight = 15
    Next r

    ws.Range("A1:Q46").Interior.Color = RGB(240, 250, 242)
End Sub


'-------------------------------------------------------------
'  Header banner
'-------------------------------------------------------------
Sub BuildHeader(ws As Worksheet)
    ws.Range("A1:Q2").Merge
    With ws.Range("A1")
        .Value = "HR WORKFORCE DASHBOARD"
        .Font.Name = "Calibri"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(30, 110, 60)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub


'-------------------------------------------------------------
'  KPI Cards
'-------------------------------------------------------------
Sub BuildKPICards(ws As Worksheet, wsPvt As Worksheet)
    Dim green      As Long: green = RGB(30, 110, 60)
    Dim greenDeep  As Long: greenDeep = RGB(15, 75, 40)
    Dim bgWhite    As Long: bgWhite = RGB(255, 255, 255)
    Dim shadow     As Long: shadow = RGB(205, 230, 210)
    Dim textDark   As Long: textDark = RGB(30, 30, 30)
    Dim redAlert   As Long: redAlert = RGB(180, 40, 40)

    Dim kpiLabels(5) As String
    Dim kpiFmlas(5)  As String
    Dim kpiFmts(5)   As String
    Dim kpiCols(5)   As Integer

    kpiLabels(0) = "Total Employees"
    kpiLabels(1) = "Avg Salary"
    kpiLabels(2) = "Avg Performance"
    kpiLabels(3) = "Avg Satisfaction"
    kpiLabels(4) = "Promotion Rate"
    kpiLabels(5) = "Attrition Rate"

    kpiFmlas(0) = "=COUNTA(Data!A2:A100000)"
    kpiFmlas(1) = "=AVERAGE(Data!N2:N100000)"
    kpiFmlas(2) = "=AVERAGE(Data!S2:S100000)"
    kpiFmlas(3) = "=AVERAGE(Data!U2:U100000)"
    kpiFmlas(4) = "=COUNTIF(Data!V2:V100000,""Yes"")/COUNTA(Data!A2:A100000)"
    kpiFmlas(5) = "=1-(COUNTIF(Data!H2:H100000,""Active"")/COUNTA(Data!A2:A100000))"

    kpiFmts(0) = "#,##0"
    kpiFmts(1) = "#,##0"
    kpiFmts(2) = "0.0"
    kpiFmts(3) = "0.0"
    kpiFmts(4) = "0.0%"
    kpiFmts(5) = "0.0%"

    kpiCols(0) = 2: kpiCols(1) = 4: kpiCols(2) = 6
    kpiCols(3) = 8: kpiCols(4) = 10: kpiCols(5) = 12

    Dim k As Integer
    For k = 0 To 5
        Dim c As Integer: c = kpiCols(k)

        With ws.Range(ws.Cells(4, c), ws.Cells(4, c + 1))
            .Merge
            .Value = kpiLabels(k)
            .Interior.Color = green
            .Font.Bold = True
            .Font.Size = 9
            .Font.Name = "Calibri"
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).Color = greenDeep
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeLeft).Color = greenDeep
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeRight).Color = greenDeep
        End With

        With ws.Range(ws.Cells(5, c), ws.Cells(5, c + 1))
            .Merge
            .Formula = kpiFmlas(k)
            .NumberFormat = kpiFmts(k)
            .Interior.Color = bgWhite
            .Font.Bold = True
            .Font.Size = 13
            .Font.Name = "Calibri"
            .Font.Color = textDark
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeLeft).Color = greenDeep
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeRight).Color = greenDeep
        End With

        ' NEW: flag Attrition Rate in red if it's above 10% - a
        ' small professional touch so the KPI itself signals risk
        If k = 5 Then
            With ws.Cells(5, c)
                If .Value > 0.1 Then .Font.Color = redAlert
            End With
        End If

        With ws.Range(ws.Cells(6, c), ws.Cells(6, c + 1))
            .Merge
            .Interior.Color = shadow
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).Color = greenDeep
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlEdgeBottom).Color = greenDeep
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeLeft).Color = greenDeep
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeRight).Color = greenDeep
        End With
    Next k
End Sub


'-------------------------------------------------------------
'  CREATE ALL CHARTS
'-------------------------------------------------------------
Sub CreateAllCharts(ws As Worksheet, wsPvt As Worksheet)
    Dim cGreen  As Long: cGreen = RGB(30, 110, 60)
    Dim cMint   As Long: cMint = RGB(103, 190, 130)
    Dim cLime   As Long: cLime = RGB(168, 213, 90)
    Dim cAmber  As Long: cAmber = RGB(230, 168, 68)
    Dim cTeal   As Long: cTeal = RGB(45, 140, 140)
    Dim cWhite  As Long: cWhite = RGB(255, 255, 255)

    Dim L1  As Double: L1 = ws.Cells(8, 1).Left
    Dim L2  As Double: L2 = ws.Cells(8, 6).Left
    Dim L3  As Double: L3 = ws.Cells(8, 12).Left
    Dim T1  As Double: T1 = ws.Cells(8, 1).Top
    Dim T2  As Double: T2 = ws.Cells(19, 1).Top
    Dim CW1 As Double: CW1 = ws.Cells(8, 6).Left - ws.Cells(8, 1).Left - 3
    Dim CW2 As Double: CW2 = ws.Cells(8, 12).Left - ws.Cells(8, 6).Left - 3
    Dim CW3 As Double: CW3 = ws.Cells(8, 17).Left - ws.Cells(8, 12).Left - 3
    Dim CH  As Double: CH = ws.Cells(19, 1).Top - ws.Cells(8, 1).Top - 3

    Dim cht As ChartObject
    Dim p   As Integer

    ' CHART 1 - Headcount by Gender (Doughnut)
    Set cht = ws.ChartObjects.Add(Left:=L1, Top:=T1, Width:=CW1, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.PivotTables("PT_Gender").TableRange1
        .ChartType = xlDoughnut
        .HasTitle = True
        .ChartTitle.Text = "Headcount by Gender"
        Call ApplyTitleStyle(.ChartTitle, cGreen)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        .SeriesCollection(1).Points(1).Interior.Color = cGreen
        .SeriesCollection(1).Points(2).Interior.Color = cAmber
        .SeriesCollection(1).Points(3).Interior.Color = cTeal
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.ShowPercentage = True
        .SeriesCollection(1).DataLabels.ShowValue = False
        .SeriesCollection(1).DataLabels.Font.Size = 9
        .SeriesCollection(1).DataLabels.Font.Bold = True
        .SeriesCollection(1).DataLabels.Font.Color = cWhite
        On Error GoTo 0
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        .Legend.Font.Size = 8
        .ShowAllFieldButtons = False
    End With

    ' CHART 2 - Joining Trend by Year (now genuinely grouped by year)
    Set cht = ws.ChartObjects.Add(Left:=L2, Top:=T1, Width:=CW2, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.PivotTables("PT_JoiningTrend").TableRange1
        .ChartType = xlColumnClustered   ' clearer than a smeared line for ~5-8 years
        .HasTitle = True
        .ChartTitle.Text = "Joining Trend by Year & Quarter"
        Call ApplyTitleStyle(.ChartTitle, cGreen)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        With .SeriesCollection(1)
            .Name = "Headcount"
            .Interior.Color = cGreen
            .HasDataLabels = True
            .DataLabels.ShowValue = True
            .DataLabels.NumberFormat = "#,##0"
            .DataLabels.Font.Size = 8
        End With
        On Error GoTo 0
        .HasLegend = False
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 9
        .Axes(xlCategory).TickLabels.Font.Name = "Calibri"
        Call RemoveGridlines(cht.Chart)
        .ShowAllFieldButtons = False
    End With

    ' CHART 3 - Headcount by Department (Column)
    Set cht = ws.ChartObjects.Add(Left:=L3, Top:=T1, Width:=CW3, Height:=CH)
    With cht.Chart
        Dim rngDept As Range
        Set rngDept = wsPvt.PivotTables("PT_Department").TableRange1
        Set rngDept = rngDept.Resize(, 2)
        .SetSourceData Source:=rngDept
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Headcount by Department"
        Call ApplyTitleStyle(.ChartTitle, cGreen)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        .SeriesCollection(1).Interior.Color = cGreen
        .SeriesCollection(1).Name = "Headcount"
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.ShowValue = True
        .SeriesCollection(1).DataLabels.Font.Size = 8
        On Error GoTo 0
        .HasLegend = False
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue).TickLabels.Font.Size = 8
        ' NEW: rotate + resize category labels so department names
        ' don't get squished to unreadable 7pt text
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Name = "Calibri"
        .Axes(xlCategory).TickLabels.Orientation = -25
        Call RemoveGridlines(cht.Chart)
        .ShowAllFieldButtons = False
    End With

    ' CHART 4 - Headcount by Employment Status (Bar)
    Set cht = ws.ChartObjects.Add(Left:=L1, Top:=T2, Width:=CW1, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.PivotTables("PT_EmploymentStatus").TableRange1
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Employment Status"
        Call ApplyTitleStyle(.ChartTitle, cGreen)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        Dim statusCols(3) As Long
        statusCols(0) = cGreen
        statusCols(1) = cAmber
        statusCols(2) = cTeal
        statusCols(3) = cMint
        For p = 1 To .SeriesCollection(1).Points.Count
            If p <= 4 Then .SeriesCollection(1).Points(p).Interior.Color = statusCols(p - 1)
        Next p
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.ShowValue = True
        .SeriesCollection(1).DataLabels.NumberFormat = "#,##0"
        .SeriesCollection(1).DataLabels.Font.Size = 8
        On Error GoTo 0
        .HasLegend = False
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 8
        Call RemoveGridlines(cht.Chart)
        .ShowAllFieldButtons = False
    End With

    ' CHART 5 - Avg Salary & Bonus by Designation
    ' NEW: Bonus is now a line on a secondary axis instead of a
    ' near-invisible column dwarfed by Salary.
    Set cht = ws.ChartObjects.Add(Left:=L2, Top:=T2, Width:=CW2, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.PivotTables("PT_Designation").TableRange1
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Avg Salary & Bonus by Designation"
        Call ApplyTitleStyle(.ChartTitle, cGreen)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        .SeriesCollection(1).Interior.Color = cGreen
        .SeriesCollection(1).Name = "Avg Salary"

        .SeriesCollection(2).ChartType = xlLineMarkers
        .SeriesCollection(2).AxisGroup = xlSecondary
        .SeriesCollection(2).Name = "Avg Bonus"
        .SeriesCollection(2).Format.Line.ForeColor.RGB = cAmber
        .SeriesCollection(2).Format.Line.Weight = 2.25
        .SeriesCollection(2).MarkerStyle = xlMarkerStyleCircle
        .SeriesCollection(2).MarkerSize = 6
        .SeriesCollection(2).Format.MarkerFill.ForeColor.RGB = cAmber
        .SeriesCollection(2).Format.MarkerLine.ForeColor.RGB = cAmber

        .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue, xlPrimary).TickLabels.Font.Size = 8
        .Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue, xlSecondary).TickLabels.Font.Size = 8
        On Error GoTo 0
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Name = "Calibri"
        .Axes(xlCategory).TickLabels.Orientation = -25
        Call RemoveGridlines(cht.Chart)
        .ShowAllFieldButtons = False
    End With

    ' CHART 6 - Attrition Reason Breakdown (Bar)
    ' "(blank)" item already hidden by CleanAttritionReasonPivot
    Set cht = ws.ChartObjects.Add(Left:=L3, Top:=T2, Width:=CW3, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.PivotTables("PT_AttritionReason").TableRange1
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Attrition Reason Breakdown"
        Call ApplyTitleStyle(.ChartTitle, cGreen)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        .SeriesCollection(1).Interior.Color = cTeal
        .SeriesCollection(1).Name = "Headcount"
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.ShowValue = True
        .SeriesCollection(1).DataLabels.NumberFormat = "#,##0"
        .SeriesCollection(1).DataLabels.Font.Size = 8
        On Error GoTo 0
        .HasLegend = False
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Name = "Calibri"
        Call RemoveGridlines(cht.Chart)
        .ShowAllFieldButtons = False
    End With
End Sub


'-------------------------------------------------------------
'  SLICERS + TIMELINE
'-------------------------------------------------------------
Sub AddSlicersAndTimeline(ws As Worksheet, wsPvt As Worksheet)
    Dim sc  As SlicerCache
    Dim tlc As SlicerCache
    Dim sl  As Slicer
    Dim tl  As Object
    Dim pt  As PivotTable
    Dim i   As Integer

    Dim fieldNames() As String
    Dim baseTables()  As String
    fieldNames = Split("Department,Gender,Employment Status,Work Mode,Designation", ",")
    baseTables = Split("PT_Department,PT_Gender,PT_EmploymentStatus,PT_WorkMode,PT_Designation", ",")

    Dim cleanNames As Variant
    cleanNames = Array("Slicer_Department", "Slicer_Gender", "Slicer_Employment_Status", _
                        "Slicer_Work_Mode", "Slicer_Designation", "Timeline_Joining_Date")
    For i = 0 To UBound(cleanNames)
        On Error Resume Next
        ActiveWorkbook.SlicerCaches(CStr(cleanNames(i))).Delete
        On Error GoTo 0
    Next i

    ws.Range("A32:Q32").Merge
    With ws.Range("A32")
        .Value = "FILTERS"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Name = "Calibri"
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(30, 110, 60)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Dim leftPos     As Double: leftPos = ws.Cells(33, 1).Left
    Dim topPos      As Double: topPos = ws.Cells(33, 1).Top
    Dim slicerW     As Double: slicerW = 140
    Dim slicerH     As Double: slicerH = 110
    Dim gap         As Double: gap = 10

    For i = 0 To UBound(fieldNames)
        On Error Resume Next
        Set sc = ActiveWorkbook.SlicerCaches.Add2( _
                    wsPvt.PivotTables(baseTables(i)), fieldNames(i))
        On Error GoTo 0

        If Not sc Is Nothing Then
            sc.Name = "Slicer_" & Replace(fieldNames(i), " ", "_")

            For Each pt In wsPvt.PivotTables
                On Error Resume Next
                sc.PivotTables.AddPivotTable pt
                On Error GoTo 0
            Next pt

            Set sl = sc.Slicers.Add(ws, , sc.Name, fieldNames(i), _
                                     topPos, leftPos, slicerW, slicerH)
            On Error Resume Next
            sl.Style = "SlicerStyle6"
            On Error GoTo 0

            leftPos = leftPos + slicerW + gap
        End If
        Set sc = Nothing
    Next i

    On Error Resume Next
    Set tlc = ActiveWorkbook.SlicerCaches.Add2( _
                wsPvt.PivotTables("PT_JoiningTrend"), "Joining Date", , xlTimeline)
    On Error GoTo 0

    If Not tlc Is Nothing Then
        tlc.Name = "Timeline_Joining_Date"

        For Each pt In wsPvt.PivotTables
            On Error Resume Next
            tlc.PivotTables.AddPivotTable pt
            On Error GoTo 0
        Next pt

        On Error Resume Next
        Set tl = tlc.Timelines.Add(ws, topPos, leftPos, 260, slicerH)
        If Not tl Is Nothing Then
            tl.Caption = "Joining Date"
        End If
        On Error GoTo 0
    End If
End Sub


'-------------------------------------------------------------
'  MACRO BUTTON
'-------------------------------------------------------------
Sub AddMacroButton(ws As Worksheet)
    Dim btnLeft   As Double: btnLeft = ws.Cells(1, 15).Left + 4
    Dim btnTop    As Double: btnTop = ws.Cells(1, 1).Top + 4
    Dim btnWidth  As Double: btnWidth = ws.Cells(1, 17).Left + ws.Columns(17).Width - ws.Cells(1, 15).Left - 8
    Dim btnHeight As Double: btnHeight = ws.Cells(3, 1).Top - ws.Cells(1, 1).Top - 8

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                  btnLeft, btnTop, btnWidth, btnHeight)
    With shp
        .Name = "btnRefreshDashboard"
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.Solid
        .Line.ForeColor.RGB = RGB(30, 110, 60)
        .Line.Weight = 1.5
        .Line.Visible = msoTrue
        .TextFrame2.TextRange.Text = "Refresh Dashboard"
        With .TextFrame2.TextRange.Font
            .Bold = msoTrue
            .Size = 9
            .Fill.ForeColor.RGB = RGB(30, 110, 60)
        End With
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .OnAction = "RefreshDashboard"
    End With
End Sub


'-------------------------------------------------------------
'  HELPER: Title styling
'-------------------------------------------------------------
Sub ApplyTitleStyle(ct As ChartTitle, titleColor As Long)
    ct.Font.Bold = True
    ct.Font.Color = titleColor
    ct.Font.Size = 11
    ct.Font.Name = "Calibri"
End Sub

'-------------------------------------------------------------
'  HELPER: Chart area styling
'-------------------------------------------------------------
Sub ApplyChartAreaStyle(cht As Chart)
    cht.ChartArea.Interior.Color = RGB(238, 248, 240)
    With cht.ChartArea.Border
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .Color = RGB(30, 110, 60)
    End With
    cht.PlotArea.Interior.ColorIndex = xlNone
    On Error Resume Next
    cht.PlotArea.Border.LineStyle = xlNone
    On Error GoTo 0
End Sub

'-------------------------------------------------------------
'  HELPER: Remove all gridlines
'-------------------------------------------------------------
Sub RemoveGridlines(cht As Chart)
    On Error Resume Next
    cht.Axes(xlValue).HasMajorGridlines = False
    cht.Axes(xlValue).HasMinorGridlines = False
    cht.Axes(xlCategory).HasMajorGridlines = False
    cht.Axes(xlCategory).HasMinorGridlines = False
    On Error GoTo 0
End Sub




