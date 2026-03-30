Attribute VB_Name = "DodajDrugeIzmene"
Option Explicit

Public Sub DodajDrugeIzmene_Run()
    DodajDrugeIzmene
End Sub

Public Sub DodajDrugeIzmene( _
        Optional ByVal wsName As String = "PREDOGLED", _
        Optional ByVal firstDataRow As Long = 3, _
        Optional ByVal firstSchedCol As Long = 5, _
        Optional ByVal codeCol As Long = 1, _
        Optional ByVal teamCol As Long = 4, _
        Optional ByVal lastRowCol As Long = 2) ' <-- B stolpec

    Dim ws As Worksheet, wsSet As Worksheet
    Set ws = ThisWorkbook.Worksheets(wsName)
    Set wsSet = ThisWorkbook.Worksheets("NASTAVITVE")

    ' lastRow iz stolpca B
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, lastRowCol).End(xlUp).Row
    If lastRow < firstDataRow Then Exit Sub

    ' lastCol iz key-based nastavitve "KONÈNI DATUM"
    Dim endDate As Date, m As Variant, lastCol As Long
    On Error GoTo CleanFail
    endDate = DateValue(modSettings.GetDateRequired(wsSet, "KONÈNI DATUM"))

    m = Application.Match(CDbl(endDate), ws.Range(ws.Cells(1, firstSchedCol), ws.Cells(1, ws.Columns.Count)), 0)
    If IsError(m) Then
        MsgBox "Konènega datuma (" & Format(endDate, "d.m.yyyy") & ") ne najdem v vrstici 1.", vbCritical
        Exit Sub
    End If
    lastCol = firstSchedCol + CLng(m) - 1
    If lastCol < firstSchedCol Then Exit Sub

    ' pospešitev
    Dim calc0 As XlCalculation
    calc0 = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    ' bulk read
    Dim arrCode As Variant, arrTeam As Variant, arrSched As Variant
    arrCode = ws.Range(ws.Cells(firstDataRow, codeCol), ws.Cells(lastRow, codeCol)).Value2
    arrTeam = ws.Range(ws.Cells(firstDataRow, teamCol), ws.Cells(lastRow, teamCol)).Value2
    arrSched = ws.Range(ws.Cells(firstDataRow, firstSchedCol), ws.Cells(lastRow, lastCol)).Value2

    Dim nRows As Long, nCols As Long
    nRows = UBound(arrSched, 1)
    nCols = UBound(arrSched, 2)

    ' 1) template po timih
    Dim tmpl As Object: Set tmpl = CreateObject("Scripting.Dictionary")
    Dim i As Long, j As Long, t As String, v As String
    Dim hasCycle As Boolean

    For i = 1 To nRows
        t = Trim$(CStr(arrTeam(i, 1)))
        If Len(t) > 0 Then
            hasCycle = False
            For j = 1 To nCols
                v = UCase$(Trim$(CStr(arrSched(i, j))))
                If v = "X1" Or v = "X2" Or v = "X3" Or v = "O" Then
                    hasCycle = True
                    Exit For
                End If
            Next j

            If hasCycle Then
                If Not tmpl.Exists(t) Then
                    Dim tmpRow() As String
                    ReDim tmpRow(1 To nCols)
                    For j = 1 To nCols
                        tmpRow(j) = UCase$(Trim$(CStr(arrSched(i, j))))
                    Next j
                    tmpl.Add t, tmpRow
                End If
            End If
        End If
    Next i

    If tmpl.Count = 0 Then
        MsgBox "Ni template ciklusa (X1/X2/X3/O).", vbExclamation
        GoTo CleanExit
    End If

    ' 2) polni samo vrstice, kjer je A != ""
    Dim filledPeople As Long, filledCells As Long
    Dim code As String
    Dim tmp As Variant

    For i = 1 To nRows
        code = Trim$(CStr(arrCode(i, 1)))
        If Len(code) > 0 Then
            t = Trim$(CStr(arrTeam(i, 1)))
            If Len(t) > 0 And tmpl.Exists(t) Then
                tmp = tmpl(t)
                
                If code = "-" Then
                ' --- RESET: izbriši vse izmene po template ciklu ---
                For j = 1 To nCols
                    v = tmp(j)
                    If v = "X1" Or v = "X2" Or v = "X3" Or v = "O" Then
                        If Len(Trim$(CStr(arrSched(i, j)))) > 0 Then
                            arrSched(i, j) = ""
                            filledCells = filledCells + 1
                        End If
                    End If
                Next j
                filledPeople = filledPeople + 1
                
                Else
                ' napolni celo vrstico po template ciklu
                For j = 1 To nCols
                    v = tmp(j)
                    If v = "X1" Or v = "X2" Or v = "X3" Or v = "O" Then
                        If Len(Trim$(CStr(arrSched(i, j)))) = 0 Then
                        arrSched(i, j) = code
                        filledCells = filledCells + 1
                        End If
                    End If
                Next j
                filledPeople = filledPeople + 1
                End If
            End If
        End If
    Next i

    ' bulk write
    ws.Range(ws.Cells(firstDataRow, firstSchedCol), ws.Cells(lastRow, lastCol)).Value2 = arrSched

    MsgBox "Konèano." & vbCrLf & _
           "Dopolnjenih oseb: " & filledPeople & vbCrLf & _
           "Vpisanih celic: " & filledCells, vbInformation

CleanExit:
    Application.Calculation = calc0
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.Calculation = calc0
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Napaka: " & Err.Description, vbCritical
End Sub




