Attribute VB_Name = "ModTotals"
Option Explicit

Sub AddTotals()

' Will sum the Selling Price in Column I and place the result two rows below
' the last selling price in that column
' To see each step of the code in action:
' Macros-Select-Step into-and move throughout the code with (Fn)F8

Dim LastCell As Range
Dim TotalFormula As String

    If Range("I2").Value <> 0 Then
        Set LastCell = Range("I2").End(xlDown)
        LastCell.Select
        ActiveCell.Offset(2).Select
        TotalFormula = "=Sum(I2:" & LastCell.Address(False, False) & ")"
        ActiveCell.Formula = TotalFormula
    Else
        Range("A3").Value = "No Sales"
    End If
    
End Sub

'Sub AllTotals()
'
'' This is how you type a For-to-Next Loop, i is just used as variable
'
'Dim i As Integer
'
'    For i = 1 To Worksheets.Count
'        Worksheets(i).Select
'        AddTotals
'    Next i
'
'End Sub

Sub AllTotals()

' This is how you type a For-each-next Loop

Dim Sheet As Worksheet

    For Each Sheet In Worksheets
        Sheet.Select
        AddTotals
    Next Sheet

End Sub
