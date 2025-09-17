Attribute VB_Name = "modLedgerNavigation"
Option Explicit

Private Const LEDGER_SHEET_NAME As String = "序时账"
Private Const ACCOUNT_COLUMNS As String = "A:B"

Public Sub EnsureAccountColumnsFormat(ByVal ws As Worksheet)
    On Error GoTo CleanExit
    Application.EnableEvents = False
    ws.Columns(ACCOUNT_COLUMNS).NumberFormat = "@"
CleanExit:
    Application.EnableEvents = True
End Sub

Public Sub HandleAccountDoubleClick(ByVal targetCell As Range, ByRef Cancel As Boolean, ByVal sourceSheet As Worksheet)
    Dim ledgerSheet As Worksheet
    Dim accountKey As String
    Dim candidates As Collection
    Dim destination As Range
    Dim candidate As Variant
    Dim success As Boolean

    If targetCell Is Nothing Then Exit Sub
    If targetCell.CountLarge > 1 Then Exit Sub
    If Intersect(targetCell, sourceSheet.Columns(ACCOUNT_COLUMNS)) Is Nothing Then Exit Sub

    EnsureAccountColumnsFormat sourceSheet

    accountKey = Trim$(CStr(targetCell.Value))
    If Len(accountKey) = 0 Then Exit Sub

    Set ledgerSheet = GetLedgerSheet()
    If ledgerSheet Is Nothing Then
        MsgBox "未找到名为 '" & LEDGER_SHEET_NAME & "' 的序时账工作表。", vbExclamation
        Exit Sub
    End If

    Set candidates = BuildAccountKeyCandidates(targetCell)
    For Each candidate In candidates
        accountKey = CStr(candidate)
        Set destination = FindAccountInLedger(accountKey, ledgerSheet)
        If Not destination Is Nothing Then
            success = True
            Exit For
        End If
    Next candidate

    If success Then
        Cancel = True
        Application.Goto destination, True
    Else
        MsgBox "在序时账中未找到与 '" & targetCell.Text & "' 匹配的分录。", vbInformation
    End If
End Sub

Private Function GetLedgerSheet() As Worksheet
    On Error Resume Next
    Set GetLedgerSheet = ThisWorkbook.Worksheets(LEDGER_SHEET_NAME)
    On Error GoTo 0
End Function

Private Function BuildAccountKeyCandidates(ByVal targetCell As Range) As Collection
    Dim candidates As New Collection

    AddCandidate candidates, NormalizeAccountKey(targetCell.Text)
    AddCandidate candidates, NormalizeAccountKey(CStr(targetCell.Value))
    AddCandidate candidates, targetCell.Text
    AddCandidate candidates, CStr(targetCell.Value)

    Set BuildAccountKeyCandidates = candidates
End Function

Private Sub AddCandidate(ByRef candidates As Collection, ByVal candidate As String)
    On Error GoTo CleanExit
    candidate = Trim$(candidate)
    If Len(candidate) = 0 Then Exit Sub
    candidates.Add candidate, candidate
CleanExit:
    On Error GoTo 0
End Sub

Private Function NormalizeAccountKey(ByVal rawValue As String) As String
    Dim sanitized As String
    Dim dotPos As Long
    Dim decimalPart As String

    sanitized = Replace(Trim$(rawValue), ",", "")
    dotPos = InStr(sanitized, ".")
    If dotPos > 0 Then
        decimalPart = Mid$(sanitized, dotPos + 1)
        If Len(decimalPart) = 0 Or decimalPart = String(Len(decimalPart), "0") Then
            sanitized = Left$(sanitized, dotPos - 1)
        End If
    End If

    NormalizeAccountKey = sanitized
End Function

Private Function FindAccountInLedger(ByVal accountKey As String, ByVal ledgerSheet As Worksheet) As Range
    Dim searchRange As Range

    If ledgerSheet Is Nothing Then Exit Function

    On Error Resume Next
    Set searchRange = ledgerSheet.UsedRange
    On Error GoTo 0

    If searchRange Is Nothing Then Exit Function

    With searchRange
        Set FindAccountInLedger = .Find(What:=accountKey, After:=searchRange.Cells(searchRange.Cells.Count), _
            LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End With
End Function
