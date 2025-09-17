VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet科目余额表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    EnsureAccountColumnsFormat Me
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Columns("A:B")) Is Nothing Then Exit Sub
    EnsureAccountColumnsFormat Me
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    HandleAccountDoubleClick Target, Cancel, Me
End Sub
