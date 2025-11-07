Attribute VB_Name = "modReports"
Option Explicit
Public PageName As String, RowName As String, ColumnName As String, DataName As String

Sub ShowForm()

    frmGenerateReports.Show
    
End Sub
Sub CreatePivot()

    Dim Destination As Range
    Dim RangeData As Range
    Dim ReportBook As Workbook
    Dim ReportSheet As Worksheet
    Dim Cache As PivotCache

    Set ReportBook = Workbooks("Reports.xlsx")
    Set ReportSheet = ReportBook.Worksheets("Reports")

    If Application.WorksheetFunction.CountA(ReportSheet.Cells) = 0 Then
        MsgBox "No data was copied to the Reports sheet. PivotTable generation cancelled.", vbExclamation
        Exit Sub
    End If

    Set RangeData = ReportSheet.Range("A1").CurrentRegion

    If Application.WorksheetFunction.CountA(RangeData.Rows(1)) < RangeData.Columns.Count Then
        MsgBox "The data used to build the PivotTable is missing one or more headers.", vbExclamation
        Exit Sub
    End If

    Set Destination = ReportSheet.Range("L1")

    Set Cache = ReportBook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=RangeData.Address(True, True, xlR1C1, True))

    Cache.CreatePivotTable TableDestination:=Destination, _
        TableName:="SalesPivot"

End Sub

Sub SetFields()
    Dim Table As PivotTable
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = Workbooks("Reports.xlsx")
    Set ws = wb.Worksheets("Reports")
    Set Table = ws.PivotTables("SalesPivot")

    ' Safety: only set fields that exist
    SafeSetField Table, PageName, xlPageField
    SafeSetField Table, RowName, xlRowField
    SafeSetField Table, ColumnName, xlColumnField
    SafeSetField Table, DataName, xlDataField

    wb.Activate
    ws.Activate
    If frmGenerateReports.Sales <> "Model" Then
        Table.PivotSelect "", xlDataOnly
        Selection.NumberFormat = "$#,##0"
    End If
    ws.Range("E1").Select
End Sub

Private Sub SafeSetField(ByVal pt As PivotTable, ByVal fldName As String, ByVal orient As XlPivotFieldOrientation)
    Dim exists As Boolean: exists = False
    Dim pf As PivotField

    If Len(fldName) = 0 Then Exit Sub

    On Error Resume Next
    Set pf = pt.PivotFields(fldName)
    exists = Not pf Is Nothing
    On Error GoTo 0

    If exists Then
        pf.Orientation = orient
    Else
        ' Optional: tell the user which field was missing
        ' Debug.Print "Missing Pivot field: " & fldName
    End If

End Sub


Sub ConsolidateData(ThisMonth)
Dim BeenThere As Boolean
Dim Compare As String
Dim Sheet As Worksheet

BeenThere = False
Sheets.Add
ActiveSheet.Name = "Reports"
        If frmGenerateReports.Period = "Month" Then
                Select Case ThisMonth
                        Case "Jan"
                                ThisMonth = 1
                        Case "Feb"
                                ThisMonth = 2
                        Case "Mar"
                                ThisMonth = 3
                        Case "Apr"
                                ThisMonth = 4
                        Case "May"
                                ThisMonth = 5
                        Case "Jun"
                                ThisMonth = 6
                        Case "Jul"
                                ThisMonth = 7
                        Case "Aug"
                                ThisMonth = 8
                        Case "Sep"
                                ThisMonth = 9
                        Case "Oct"
                                ThisMonth = 10
                        Case "Nov"
                                ThisMonth = 11
                        Case "Dec"
                                ThisMonth = 12
                End Select
        End If

                Select Case frmGenerateReports.Period
                        Case "Month"
                                For Each Sheet In Worksheets
                                        Sheet.Select
                                        Compare = ActiveSheet.Name
                                On Error Resume Next
                                Compare = Month(CDate(Compare))
                                        If Compare = ThisMonth Then
                                                frmGenerateReports.StartDate = Sheet.Name
                                                If BeenThere = False Then
                                                        GrabCells 1
                                                        BeenThere = True
                                                Else
                                                        GrabCells 2
                                                        DoEvents
                                                End If
                                        End If
                                Next
                        Case "All"
                                GrabCells 1
                                Do
                                        frmGenerateReports.StartDate = frmGenerateReports.StartDate + 1
                                        GrabCells 2
                                        DoEvents
                                Loop Until frmGenerateReports.StartDate = frmGenerateReports.EndDate
                        Columns("A:J").EntireColumn.AutoFit
                        Range("A1").Select
                        Case "Other"
                        GrabCells 1
                        Do
                                frmGenerateReports.StartDate = frmGenerateReports.StartDate + 1
                                GrabCells 2
                                DoEvents
                        Loop Until frmGenerateReports.StartDate = frmGenerateReports.EndDate
                        Columns("A:J").EntireColumn.AutoFit
                        Range("A1").Select
  
        End Select

End Sub

Sub GrabCells(StartingCell As Long)
    Dim callDate As String
    Dim src As Worksheet, dst As Worksheet
    Dim lastRow As Long

    callDate = Format(frmGenerateReports.StartDate, "d-mmm-yy")
    Set src = Sheets(callDate)
    Set dst = Sheets("Reports")

    ' 1) Ensure headers are written once at A1:J1
    If StartingCell = 1 Then
        dst.Range("A1:J1").Value = src.Range("B1:J1").Value
    End If

    ' 2) Append data rows if any
    lastRow = src.Cells(src.Rows.Count, "B").End(xlUp).Row
    If lastRow >= 2 Then
        src.Range("B2:J" & lastRow).Copy
        dst.Cells(dst.Rows.Count, "A").End(xlUp).Offset(1).PasteSpecial xlPasteValues
    End If
    
End Sub

Sub FinishReport()
    Dim Sheet As Worksheet
    Workbooks.Open Filename:=ThisWorkbook.path & "\Reports.xlsx"

        For Each Sheet In Worksheets
                Sheet.Select
                Cells.Select
                Selection.Clear
                Range("A1").Select
        Next
    Workbooks("Sales - Fiscal Year.xlsm").Activate
    Sheets("Reports").Select
    Cells.Select
    Selection.Copy
    Workbooks("Reports.xlsx").Activate
    ActiveSheet.Paste
    Range("A1").Select
    ' Ensure the target sheet in Reports.xlsx is named consistently
    If ActiveSheet.Name <> "Reports" Then
        On Error Resume Next
        ActiveSheet.Name = "Reports"
    On Error GoTo 0
    End If

' Quick header sanity check before we try to build a Pivot
    With Workbooks("Reports.xlsx").Worksheets("Reports")
        If Application.WorksheetFunction.CountA(.Cells) = 0 Then
            MsgBox "Reports sheet is empty. Aborting pivot creation.", vbExclamation
        Exit Sub  ' exits FinishReport safely
    End If

    Dim src As Range
    Set src = .Range("A1").CurrentRegion

    ' Require a fully populated header row (no blank field names)
    If Application.WorksheetFunction.CountA(src.Rows(1)) < src.Columns.Count Then
        MsgBox "Reports sheet has missing/blank headers. Aborting pivot creation.", vbExclamation
        Exit Sub
    End If
End With

    Windows("Sales - Fiscal Year.xlsm").Activate
    Application.CutCopyMode = False
    ActiveWindow.SelectedSheets.Delete
    Workbooks("Reports.xlsx").Activate
    CreatePivot
    SetFields

End Sub
