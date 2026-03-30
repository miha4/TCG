Attribute VB_Name = "modUndo"

Option Explicit
' =============================================================================
' modUndo  generalni "Undo Last Action" (single-level snapshot)
' -----------------------------------------------------------------------------
' Uporaba:
'   modUndo.BeginSnapshot someRange, "OFFICE"
'   ... spremembe v someRange ...
'   modUndo.Undo
'
' Snapshot shrani:
'   - workbook fullname, sheet name, address
'   - Value2 (vedno 2D)
'   - FormulaR1C1 (vedno 2D)
'
' Ne shranjuje formatov (barve, conditional formatting, komentarji, itd.)
' =============================================================================

Private mReady As Boolean
Private mActionName As String

Private mWbFullName As String
Private mWsName As String
Private mAddressA1 As String

Private mVal2 As Variant          ' 2D variant array
Private mFormulaR1C1 As Variant   ' 2D variant array

' =============================================================================
'  PUBLIC API
' =============================================================================

Public Sub BeginSnapshot(ByVal rng As Range, Optional ByVal actionName As String = "")
    On Error GoTo EH

    If rng Is Nothing Then Err.Raise vbObjectError + 3100, "modUndo.BeginSnapshot", "Range is Nothing."

    Dim ws As Worksheet: Set ws = rng.Worksheet
    Dim wb As Workbook:  Set wb = ws.Parent

    mActionName = actionName
    mWbFullName = wb.fullName
    mWsName = ws.Name
    mAddressA1 = rng.Address(True, True)

    ' preberi in normaliziraj na 2D
    mVal2 = To2D(rng.Value2, rng.Rows.Count, rng.Columns.Count)
    mFormulaR1C1 = To2D(rng.FormulaR1C1, rng.Rows.Count, rng.Columns.Count)

    mReady = True
    Debug.Print "UNDO SNAPSHOT READY | action=" & mActionName & " | ws=" & mWsName & " | rng=" & mAddressA1
    Exit Sub

EH:
    mReady = False
    Err.Raise Err.Number, "modUndo.BeginSnapshot", Err.Description
End Sub

Public Sub Undo()
    If Not mReady Then
        Err.Raise vbObjectError + 3101, "modUndo.Undo", "Undo ni na voljo (ni snapshot-a)."
    End If

    Dim wb As Workbook: Set wb = FindWorkbookByFullName(mWbFullName)
    If wb Is Nothing Then
        Err.Raise vbObjectError + 3102, "modUndo.Undo", "Workbook za snapshot ni odprt: " & mWbFullName
    End If

    Dim ws As Worksheet: Set ws = GetWorksheetSafe(wb, mWsName)
    If ws Is Nothing Then
        Err.Raise vbObjectError + 3103, "modUndo.Undo", "Worksheet za snapshot ne obstaja: " & mWsName
    End If

    Dim rng As Range: Set rng = ws.Range(mAddressA1)

    ' shape check
    If rng.Rows.Count <> SafeUB(mVal2, 1) Or rng.Columns.Count <> SafeUB(mVal2, 2) Then
        Err.Raise vbObjectError + 3104, "modUndo.Undo", _
            "Range size se ne ujema s snapshotom. Range=" & rng.Address(False, False)
    End If

    Dim su As Boolean: su = Application.ScreenUpdating
    On Error GoTo EH_RESTORE
    Application.ScreenUpdating = False

    ' 1) povrni vrednosti
    rng.Value2 = mVal2

    ' 2) povrni formule (to ohrani formule, kjer so bile)
    rng.FormulaR1C1 = mFormulaR1C1

    Application.ScreenUpdating = su
    Debug.Print "UNDO APPLIED | action=" & mActionName & " | ws=" & mWsName & " | rng=" & mAddressA1
    Exit Sub

EH_RESTORE:
    Application.ScreenUpdating = su
    Err.Raise Err.Number, "modUndo.Undo", Err.Description
End Sub

Public Function HasSnapshot() As Boolean
    HasSnapshot = mReady
End Function

Public Function SnapshotActionName() As String
    SnapshotActionName = mActionName
End Function

Public Function SnapshotInfo() As String
    If Not mReady Then
        SnapshotInfo = "NO SNAPSHOT"
    Else
        SnapshotInfo = "action=" & mActionName & " | ws=" & mWsName & " | rng=" & mAddressA1
    End If
End Function

Public Sub ClearSnapshot()
    mReady = False
    mActionName = ""
    mWbFullName = ""
    mWsName = ""
    mAddressA1 = ""
    mVal2 = Empty
    mFormulaR1C1 = Empty
End Sub

' =============================================================================
'  PRIVATE HELPERS
' =============================================================================

Private Function FindWorkbookByFullName(ByVal fullName As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase$(wb.fullName) = LCase$(fullName) Then
            Set FindWorkbookByFullName = wb
            Exit Function
        End If
    Next wb
    Set FindWorkbookByFullName = Nothing
End Function

Private Function GetWorksheetSafe(ByVal wb As Workbook, ByVal SheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetSafe = wb.Worksheets(SheetName)
    On Error GoTo 0
End Function

' Pretvori scalar ali 2D v 2D (1..r, 1..c)
Private Function To2D(ByVal v As Variant, ByVal rCount As Long, ByVal cCount As Long) As Variant
    Dim a As Variant

    If rCount = 1 And cCount = 1 Then
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = v
        To2D = a
        Exit Function
    End If

    ' e je e array, ga vrni
    If IsArray(v) Then
        To2D = v
        Exit Function
    End If

    ' fallback: scalar v veji range (ne bi se smelo zgodit, ampak za ziher)
    ReDim a(1 To rCount, 1 To cCount)
    a(1, 1) = v
    To2D = a
End Function

' varno dobi UBound dim; e ni array, vrne 1
Private Function SafeUB(ByVal v As Variant, ByVal dimN As Long) As Long
    On Error GoTo EH
    If IsArray(v) Then
        SafeUB = UBound(v, dimN)
    Else
        SafeUB = 1
    End If
    Exit Function
EH:
    SafeUB = 1
End Function




