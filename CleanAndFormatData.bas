Attribute VB_Name = "CleanAndFormatData"
'  DATA CLEANING & FORMATTING MACRO - GREEN THEME
'  Sheet: "Data"
'  Columns: Employee ID | Employee Name | Gender | Age | Department |
'           Designation | Joining Date | Employment Status | Work Mode |
'           Education | Hire Source | City | Manager | Salary | Bonus |
'           Attendance % | Leave Days | Overtime Hours | Performance Rating |
'           Training Hours | Satisfaction Score | Promotion | Attrition Reason
'
'  BUTTON: Assign this macro -> CleanAndFormatData
'  STYLE:  Green theme (TableStyleMedium7)
'          Header    = Dark green  RGB(30, 110, 60)
'          Odd rows  = Light green RGB(226, 245, 228)
'          Even rows = White       RGB(255, 255, 255)
'=============================================================

Option Explicit

'-------------------------------------------------------------
'  MAIN - Assign this to your macro button
'-------------------------------------------------------------
Sub CleanAndFormatData()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws      As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long

    Set ws = ThisWorkbook.Sheets("Data")

    ' Find actual data boundaries
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Run each cleaning & formatting step
    Call StepRemoveBlankRowsCols(ws, lastRow, lastCol)
    Call StepTrimWhitespace(ws, lastRow, lastCol)
    Call StepStandardiseText(ws, lastRow)
    Call StepFormatNumbers(ws, lastRow)
    Call StepCentreAlignment(ws, lastRow, lastCol)
    Call StepApplyGreenTheme(ws, lastRow, lastCol)   ' Green theme step
    Call StepFormatHeader(ws, lastCol)
    Call StepSetColumnWidths(ws)
    Call StepFreezeHeader(ws)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Green Theme Applied Successfully!" & vbNewLine & _
           lastRow - 1 & " records processed.", vbInformation, "Done"
End Sub


'-------------------------------------------------------------
'  STEP 1 - Remove entirely blank rows and columns
'-------------------------------------------------------------
Sub StepRemoveBlankRowsCols(ws As Worksheet, lastRow As Long, lastCol As Long)
    Dim r As Long
    Dim c As Long

    ' Delete blank rows (bottom-up to avoid row shift issues)
    For r = lastRow To 2 Step -1
        If WorksheetFunction.CountA(ws.Rows(r)) = 0 Then
            ws.Rows(r).Delete
        End If
    Next r

    ' Delete blank columns (right-to-left)
    For c = lastCol To 1 Step -1
        If WorksheetFunction.CountA(ws.Columns(c)) = 0 Then
            ws.Columns(c).Delete
        End If
    Next c
End Sub


'-------------------------------------------------------------
'  STEP 2 - Trim leading/trailing spaces from all text cells
'-------------------------------------------------------------
Sub StepTrimWhitespace(ws As Worksheet, lastRow As Long, lastCol As Long)
    Dim r    As Long
    Dim c    As Long
    Dim cell As Range

    For r = 1 To lastRow
        For c = 1 To lastCol
            Set cell = ws.Cells(r, c)
            If cell.Value <> "" And VarType(cell.Value) = vbString Then
                cell.Value = Trim(cell.Value)
            End If
        Next c
    Next r
End Sub


'-------------------------------------------------------------
'  STEP 3 - Standardise text columns (Proper Case)
'  Applies to: Employee Name, Gender, Department, Designation,
'              Employment Status, Work Mode, Education, Hire Source,
'              City, Manager, Promotion, Attrition Reason
'-------------------------------------------------------------
Sub StepStandardiseText(ws As Worksheet, lastRow As Long)
    Dim r      As Long
    Dim col    As Variant
    Dim colNum As Integer
    Dim headers As Variant

    headers = Array("Employee Name", "Gender", "Department", "Designation", _
                     "Employment Status", "Work Mode", "Education", "Hire Source", _
                     "City", "Manager", "Promotion", "Attrition Reason")

    For Each col In headers
        colNum = FindColumn(ws, CStr(col))
        If colNum > 0 Then
            For r = 2 To lastRow
                If ws.Cells(r, colNum).Value <> "" Then
                    ws.Cells(r, colNum).Value = _
                        WorksheetFunction.Proper(ws.Cells(r, colNum).Value)
                End If
            Next r
        End If
    Next col
End Sub


'-------------------------------------------------------------
'  STEP 4 - Format number columns with correct formats
'  Age                 -> 0
'  Salary, Bonus       -> #,##0.00
'  Attendance %        -> 0%   (stored as 0-100, shown as percent)
'  Leave Days          -> 0
'  Overtime Hours      -> 0.0
'  Performance Rating  -> 0.0
'  Training Hours      -> 0
'  Satisfaction Score  -> 0.0
'  Joining Date        -> yyyy-mm-dd (convert serials first)
'-------------------------------------------------------------
Sub StepFormatNumbers(ws As Worksheet, lastRow As Long)
    Dim colAge     As Integer: colAge = FindColumn(ws, "Age")
    Dim colSalary  As Integer: colSalary = FindColumn(ws, "Salary")
    Dim colBonus   As Integer: colBonus = FindColumn(ws, "Bonus")
    Dim colAttend  As Integer: colAttend = FindColumn(ws, "Attendance %")
    Dim colLeave   As Integer: colLeave = FindColumn(ws, "Leave Days")
    Dim colOT      As Integer: colOT = FindColumn(ws, "Overtime Hours")
    Dim colPerf    As Integer: colPerf = FindColumn(ws, "Performance Rating")
    Dim colTrain   As Integer: colTrain = FindColumn(ws, "Training Hours")
    Dim colSat     As Integer: colSat = FindColumn(ws, "Satisfaction Score")
    Dim colDate    As Integer: colDate = FindColumn(ws, "Joining Date")

    If colAge > 0 Then ws.Range(ws.Cells(2, colAge), ws.Cells(lastRow, colAge)).NumberFormat = "0"
    If colSalary > 0 Then ws.Range(ws.Cells(2, colSalary), ws.Cells(lastRow, colSalary)).NumberFormat = "#,##0.00"
    If colBonus > 0 Then ws.Range(ws.Cells(2, colBonus), ws.Cells(lastRow, colBonus)).NumberFormat = "#,##0.00"
    If colAttend > 0 Then ws.Range(ws.Cells(2, colAttend), ws.Cells(lastRow, colAttend)).NumberFormat = "0.0%"
    If colLeave > 0 Then ws.Range(ws.Cells(2, colLeave), ws.Cells(lastRow, colLeave)).NumberFormat = "0"
    If colOT > 0 Then ws.Range(ws.Cells(2, colOT), ws.Cells(lastRow, colOT)).NumberFormat = "0.0"
    If colPerf > 0 Then ws.Range(ws.Cells(2, colPerf), ws.Cells(lastRow, colPerf)).NumberFormat = "0.0"
    If colTrain > 0 Then ws.Range(ws.Cells(2, colTrain), ws.Cells(lastRow, colTrain)).NumberFormat = "0"
    If colSat > 0 Then ws.Range(ws.Cells(2, colSat), ws.Cells(lastRow, colSat)).NumberFormat = "0.0"

    ' Joining Date: convert Excel serial numbers to readable dates
    If colDate > 0 Then
        Dim dateCell As Range
        Dim r        As Long
        For r = 2 To lastRow
            Set dateCell = ws.Cells(r, colDate)
            If IsNumeric(dateCell.Value) And dateCell.Value > 1 Then
                dateCell.Value = CDate(dateCell.Value)
            End If
        Next r
        ws.Range(ws.Cells(2, colDate), ws.Cells(lastRow, colDate)).NumberFormat = "yyyy-mm-dd"
    End If
End Sub


'-------------------------------------------------------------
'  STEP 5 - Centre-align ALL cells (header + data)
'-------------------------------------------------------------
Sub StepCentreAlignment(ws As Worksheet, lastRow As Long, lastCol As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Uniform row height
    ws.Rows("1:" & lastRow).RowHeight = 18
End Sub


'-------------------------------------------------------------
'  STEP 6 - Apply Green Theme (TableStyleMedium7)
'-------------------------------------------------------------
Sub StepApplyGreenTheme(ws As Worksheet, lastRow As Long, lastCol As Long)
    Dim tbl As ListObject

    ' Detect data range automatically
    Dim rng As Range
    Set rng = ws.Range("A1").CurrentRegion

    ' Unlist any existing tables (removes table format, KEEPS data intact)
    On Error Resume Next
    For Each tbl In ws.ListObjects
        tbl.Unlist   ' Unlist preserves data; Delete would erase it
    Next tbl
    On Error GoTo 0

    ' Create table
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "tblEmployeeData"

    ' Apply Green Table Style
    tbl.TableStyle = "TableStyleMedium7"

    ' Show filter dropdowns
    tbl.ShowAutoFilter = True

    ' ---- Override row banding with exact green shades ----
    ' Odd rows  -> light green tint
    ' Even rows -> white
    Dim r       As Long
    Dim oddBg   As Long: oddBg = RGB(226, 245, 228)    ' Light green
    Dim evenBg  As Long: evenBg = RGB(255, 255, 255)   ' White

    For r = 2 To lastRow
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior
            If (r Mod 2) = 1 Then   ' Odd row  -> light green
                .Color = oddBg
            Else                     ' Even row -> white
                .Color = evenBg
            End If
        End With
    Next r

    ' Autofit columns
    rng.Columns.AutoFit
End Sub


'-------------------------------------------------------------
'  STEP 7 - Format header row: green theme styling
'           Dark green background, white bold text, centred
'-------------------------------------------------------------
Sub StepFormatHeader(ws As Worksheet, lastCol As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)            ' White text
        .Font.Name = "Calibri"
        .Interior.Color = RGB(30, 110, 60)          ' Dark green header
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 22
    End With
End Sub


'-------------------------------------------------------------
'  STEP 8 - Set column widths tailored to each column
'-------------------------------------------------------------
Sub StepSetColumnWidths(ws As Worksheet)
    Dim colMap As Variant
    Dim i      As Integer
    Dim colNum As Integer

    colMap = Array( _
        "Employee ID", 12, _
        "Employee Name", 20, _
        "Gender", 9, _
        "Age", 7, _
        "Department", 16, _
        "Designation", 18, _
        "Joining Date", 14, _
        "Employment Status", 16, _
        "Work Mode", 12, _
        "Education", 14, _
        "Hire Source", 14, _
        "City", 12, _
        "Manager", 16, _
        "Salary", 12, _
        "Bonus", 12, _
        "Attendance %", 13, _
        "Leave Days", 11, _
        "Overtime Hours", 14, _
        "Performance Rating", 16, _
        "Training Hours", 13, _
        "Satisfaction Score", 16, _
        "Promotion", 11, _
        "Attrition Reason", 18)

    For i = 0 To UBound(colMap) - 1 Step 2
        colNum = FindColumn(ws, CStr(colMap(i)))
        If colNum > 0 Then
            ws.Columns(colNum).ColumnWidth = CDbl(colMap(i + 1))
        End If
    Next i
End Sub


'-------------------------------------------------------------
'  STEP 9 - Freeze the header row
'-------------------------------------------------------------
Sub StepFreezeHeader(ws As Worksheet)
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
End Sub


'-------------------------------------------------------------
'  HELPER - Find column number by header name (case-insensitive)
'-------------------------------------------------------------
Function FindColumn(ws As Worksheet, headerName As String) As Integer
    Dim c       As Integer
    Dim lastCol As Integer
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If LCase(Trim(ws.Cells(1, c).Value)) = LCase(Trim(headerName)) Then
            FindColumn = c
            Exit Function
        End If
    Next c

    FindColumn = 0  ' Not found
End Function


'=============================================================
'  STANDALONE: Run this alone to apply only the green theme
'=============================================================
Sub ApplyGreenTheme()
    Dim ws  As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Set ws = ActiveSheet

    ' Detect data range automatically
    Set rng = ws.Range("A1").CurrentRegion

    ' Unlist existing table if already exists (keeps data, removes table format)
    On Error Resume Next
    For Each tbl In ws.ListObjects
        tbl.Unlist   ' Unlist preserves data; Delete would erase it
    Next tbl
    On Error GoTo 0

    ' Create table
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)

    ' Apply Green Table Style
    tbl.TableStyle = "TableStyleMedium7"   ' Green theme

    ' Optional: Center align
    rng.HorizontalAlignment = xlCenter

    ' Autofit columns
    rng.Columns.AutoFit

    MsgBox "Data Clean&Format Successfully!"
End Sub


