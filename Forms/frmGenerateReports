VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGenerateReports 
   Caption         =   "Generate Sales Report"
   ClientHeight    =   5400
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7584
   OleObjectBlob   =   "frmGenerateReports.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGenerateReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sales, Period, Month As String
Public StartDate, EndDate As Date

Private Sub cmdDisplay_Click()

    Me.Hide
    GetDates
    GetGroupedSales
    ConsolidateData Month
    FinishReport
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()

    With cboMonth
        .AddItem "Jan"
        .AddItem "Feb"
        .AddItem "Mar"
        .AddItem "Apr"
        .AddItem "May"
        .AddItem "Jun"
        .AddItem "Jul"
        .AddItem "Aug"
        .AddItem "Sep"
        .AddItem "Oct"
        .AddItem "Nov"
        .AddItem "Dec"
    End With
    cboMonth.Visible = False
    lblMonth.Visible = False
    
End Sub

Private Sub cmdCancel_Click()

        Unload Me

End Sub

Private Sub optAll_Click()

        Period = "All"

End Sub

Private Sub optClassification_Click()

        Sales = "Classification"

End Sub

Private Sub optModel_Click()

        Sales = "Model"

End Sub

Private Sub optMonth_Change()

        If optMonth = True Then
                cboMonth.Visible = True
                lblMonth.Visible = True
        Else
                cboMonth.Visible = False
                lblMonth.Visible = False
        End If

End Sub

Private Sub optMonth_Click()

        Period = "Month"

End Sub

Private Sub optOther_Change()

        If optOther = True Then
                txtStartDate.Enabled = True
                txtEndDate.Enabled = True
        Else
                txtStartDate.Enabled = False
                txtEndDate.Enabled = False
        End If

End Sub

Private Sub optOther_Click()

        Period = "Other"

End Sub

Private Sub optSalesPerson_Click()

        Sales = "Salesperson"

End Sub
Function GetDates()

    Select Case Period
        Case "Month"
            Month = cboMonth
        Case "All"
            StartDate = GetFirstDate
            EndDate = GetLastDate
        Case "Other"
            StartDate = CDate(txtStartDate)
            EndDate = CDate(txtEndDate)
        End Select

End Function


Function GetLastDate()
    Dim Sheet As Worksheet
    Dim LastDate, TestDate As Date

    
    LastDate = 0
    For Each Sheet In Worksheets
        TestDate = CDate(Sheet.Name)
        If TestDate > LastDate Then
        LastDate = TestDate
        End If
    Next Sheet
    GetLastDate = LastDate
    
End Function

Function GetFirstDate()
    Dim Sheet As Worksheet
    Dim FirstDate, TestDate As Date

    FirstDate = 99999
    For Each Sheet In Worksheets
                TestDate = CDate(Sheet.Name)
        If TestDate < FirstDate Then
        FirstDate = TestDate
        End If
    Next Sheet
    GetFirstDate = FirstDate
    
End Function

Sub GetGroupedSales()
    Select Case Sales
        Case "Salesperson"
            PageName = "Salesperson"
            RowName = "Year"
            ColumnName = "Make"
            DataName = "Selling Price"
        Case "Model"
            PageName = "Model"
            RowName = "Year"
            ColumnName = "Color"
            DataName = "Color"
        Case "Classification"
            PageName = "Classification"
            RowName = "Make"
            ColumnName = "Year"
            DataName = "Selling Price"
    End Select

End Sub

