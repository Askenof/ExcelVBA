Attribute VB_Name = "modReports"
Option Explicit
Public PageName, RowName, ColumnName, DataName As String


Sub ShowForm()

    frmGenerateReports.Show
    
End Sub

Sub CreatePivot()

Dim Destination As Range, RangeData As Range

    Set Destination = Worksheets("Reports").Range("A1")
    Set RangeData = Range("A1", Range("I1").End(xlDown))
    
    ActiveWorkbook.PivotCaches.Create _
        (SourceType:=xlDatabase, SourceData:=RangeData) _
        .CreatePivotTable TableDestination:=Destination, _
        TableName:="SalesPivot"
        
End Sub

Sub SetFields()

Dim Table As PivotTable

    Set Table = Worksheets("Reports").PivotTables("SalesPivot")
    
    With Table
        .PivotFields(PageName).Orientation = xlPageField
        .PivotFields(RowName).Orientation = xlRowField
        .PivotFields(ColumnName).Orientation = xlColumnField
        .PivotFields(DataName).Orientation = xlDataField
    End With
    ActiveWorkbook.Sheets("Reports").Activate
        If frmGenerateReports.Sales <> "Model" Then
            Table.PivotSelect "", xlDataOnly
            Selection.NumberFormat = "$#,##0"
        End If
        Range("E1").Select

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

Sub GrabCells(StartingCell)
    Dim CallDate, Where, Where2 As String
    
    
    CallDate = Format(frmGenerateReports.StartDate, "d-mmm-yy")
    Sheets(CallDate).Select
        If Range("A2").Value <> Empty Then
                Where = "B" & StartingCell
                Where2 = "J" & StartingCell
                Range(Where, Range(Where2).End(xlDown)).Select
                Application.CutCopyMode = False
                Selection.Copy
                Range("A1").Select
                Sheets("Reports").Select
                ActiveSheet.Paste
                Range("A1").End(xlDown).Offset(1).Select
        Else
                If StartingCell = 1 Then
                        Application.CutCopyMode = False
                        Range("B1:J1").Select
                        Selection.Copy
                        Sheets("Reports").Select
                        ActiveSheet.Paste
                        Range("A2").Select
                End If
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
    Windows("Sales - Fiscal Year.xlsm").Activate
    Application.CutCopyMode = False
    ActiveWindow.SelectedSheets.Delete
    Workbooks("Reports.xlsx").Activate
    CreatePivot
    SetFields

End Sub
