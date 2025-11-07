Attribute VB_Name = "Module1"
Option Explicit

' Requires:
'  - Trust Center: "Trust access to the VBA project object model"
'  - Reference: Microsoft Visual Basic for Applications Extensibility 5.3

Private Function ExtFor(ByVal vbComp As VBIDE.VBComponent) As String
    Select Case vbComp.Type
        Case vbext_ct_StdModule:          ExtFor = ".bas"
        Case vbext_ct_ClassModule:        ExtFor = ".cls"
        Case vbext_ct_Document:           ExtFor = ".cls" ' ThisWorkbook / Sheets
        Case vbext_ct_MSForm:             ExtFor = ".frm" ' Will also emit a matching .frx
        Case Else:                        ExtFor = ".txt"
    End Select
End Function

Private Sub EnsurePath(ByVal path As String)
    Dim parts() As String, p As String, i As Long
    parts = Split(path, Application.PathSeparator)
    p = parts(0)
    For i = 1 To UBound(parts)
        p = p & Application.PathSeparator & parts(i)
        If Len(Dir$(p, vbDirectory)) = 0 Then MkDir p
    Next i
End Sub

Public Sub ExportAllVba(Optional ByVal targetFolder As String = "")
    Dim vbComp As VBIDE.VBComponent
    Dim fn As String

    If targetFolder = "" Then targetFolder = ThisWorkbook.path & Application.PathSeparator & "vba-src"
    EnsurePath targetFolder

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        fn = targetFolder & Application.PathSeparator & vbComp.Name & ExtFor(vbComp)
        If Len(Dir$(fn)) > 0 Then Kill fn              ' overwrite old export
        vbComp.Export fn
    Next vbComp

    MsgBox "Exported VBA to: " & targetFolder, vbInformation
End Sub

Public Sub Run_ExportAllVba()
    ExportAllVba  ' calls your parameterized routine
End Sub

