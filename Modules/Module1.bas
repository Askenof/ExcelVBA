Attribute VB_Name = "module1"
Option Explicit
' Requires:
'  - Trust Center: "Trust access to the VBA project object model"
'  - Reference: Microsoft Visual Basic for Applications Extensibility 5.3


Private Sub EnsureFolder(ByVal path As String)

    If Len(Dir$(path, vbDirectory)) = 0 Then MkDir path

End Sub

Private Function ExtFor(ByVal vbComp As VBIDE.VBComponent) As String
    Select Case vbComp.Type
        Case vbext_ct_StdModule:   ExtFor = ".bas"
        Case vbext_ct_ClassModule: ExtFor = ".cls"
        Case vbext_ct_Document:    ExtFor = ".cls"  ' ThisWorkbook/Sheets
        Case vbext_ct_MSForm:      ExtFor = ".frm"  ' A matching .frx will be emitted too
        Case Else:                 ExtFor = ".txt"
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

Private Sub CleanTargetFolders(ByVal base As String)

    On Error Resume Next
    Kill base & "\Modules\*.bas"
    Kill base & "\Sheets\*.cls"
    Kill base & "\Forms\*.frm"
    Kill base & "\Forms\*.frx"
    
    On Error GoTo 0
    
End Sub
Private Function TargetFor(ByVal vbComp As VBIDE.VBComponent) As String

    Dim base As String: base = ThisWorkbook.path
    Select Case vbComp.Type
    
        Case vbext_ct_StdModule
            TargetFor = base & "\Modules\"
        Case vbext_ct_ClassModule, vbext_ct_Document   ' sheets + ThisWorkbook
            TargetFor = base & "\Sheets\"
        Case vbext_ct_MSForm
            TargetFor = base & "\Forms\"
        Case Else
            TargetFor = base & "\"   ' fallback
            
    End Select
    
End Function

Public Sub ExportAllVba(Optional ByVal _
ignored As String = "")
    Dim vbComp As VBIDE.VBComponent
    Dim fn As String, p As String
    Dim base As String: base = ThisWorkbook.path

    ' make sure the three folders exist
    EnsureFolder base & "\Modules"
    EnsureFolder base & "\Sheets"
    EnsureFolder base & "\Forms"

    ' clean old exports so deletes/renames are reflected in Git
    CleanTargetFolders base

    ' export each component to its folder
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        p = TargetFor(vbComp)
        EnsureFolder p
        fn = p & vbComp.Name & ExtFor(vbComp)
        If Len(Dir$(fn)) > 0 Then Kill fn
        vbComp.Export fn
    Next vbComp

    MsgBox "Exported VBA to:" & vbCrLf & _
           base & "\Modules" & vbCrLf & _
           base & "\Sheets" & vbCrLf & _
           base & "\Forms", vbInformation
End Sub

Public Sub Run_ExportAllVba()
    ExportAllVba
End Sub

