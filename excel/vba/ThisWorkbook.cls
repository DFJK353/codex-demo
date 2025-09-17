VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = Me.Worksheets("科目余额表")
    On Error GoTo 0

    If Not ws Is Nothing Then
        EnsureAccountColumnsFormat ws
    End If
End Sub
