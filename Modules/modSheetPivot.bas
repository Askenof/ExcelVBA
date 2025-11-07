Attribute VB_Name = "modSheetPivot"
Option Explicit

Sub CreatePivotTable()
Attribute CreatePivotTable.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CreatePivotTable Macro
'

'
    Application.CutCopyMode = False
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:="'26-Aug-16'!R1C1:R13C10", _
        Version:=8).CreatePivotTable _
        TableDestination:="Sheet1!R3C1", _
        TableName:="PivotTable1", _
        DefaultVersion:=8

    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
End Sub
Sub SetFields()
'
' SetFields Macro
'

'
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Year")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Classification")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Make")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Selling Price"), "Sum of Selling Price", xlSum
End Sub
