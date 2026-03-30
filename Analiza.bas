Attribute VB_Name = "Analiza"

Option Explicit

' ============================================================
'  POROÈILO – BLOCK (3 stolpci na enoto)
'  - BEST DAYS (top tabela)
'  - LOG
'  - ANALIZA po osebah + % uspešnosti + skupni %
'
'  Opomba:
'   - hasPerson parameter v podpisu uporabljaš kot "hasUnit" filter (Boolean array firstRow..lastRow)
'   - uporablja modOffice.IsAnyShiftForQuota in modOffice_Logic.CanonicalShift
' ============================================================

Public Sub UstvariPorocilo_Block( _
        ByVal wsG As Worksheet, _
        ByVal firstRow As Long, ByVal lastRow As Long, _
        ByVal COL_NAME As Long, ByVal COL_PCT As Long, _
        ByRef shArr() As Variant, ByRef hasPerson() As Boolean, _
        ByVal daysWidth As Long, _
        ByVal FIRST_DATE_COL_G As Long, _
        ByVal ROW_DDGAMA As Long, _
        ByVal ROW_DATES_G As Long, _
        ByRef cntArr() As Double, _
        ByRef logArr() As String, _
        ByVal logCount As Long, _
        ByVal unitKey As String, _
        ByVal startCol As Long, _
        ByVal ScoreThreshold As Double)

    Dim wsR As Worksheet
    On Error Resume Next
    Set wsR = ThisWorkbook.Worksheets("POROÈILO")
    On Error GoTo 0
    If wsR Is Nothing Then
        Set wsR = ThisWorkbook.Worksheets.Add
        wsR.Name = "POROÈILO"
    End If

    ' --- poèisti samo ta blok (3 stolpci) ---
    wsR.Range(wsR.Cells(1, startCol), wsR.Cells(4000, startCol + 2)).ClearContents

    ' =========================
    ' 1) NASLOV
    ' =========================
    wsR.Cells(1, startCol).Value = "POROÈILO – OFFICE (" & unitKey & ")"
    wsR.Cells(1, startCol).Font.Bold = True
    wsR.Cells(1, startCol).Font.Size = 14

    ' =========================
    ' 2) BEST DAYS (top tabela)
    ' =========================
    wsR.Cells(3, startCol).Value = "BEST DAYS:"
    wsR.Cells(3, startCol).Font.Bold = True
    wsR.Cells(4, startCol).Value = "Opomba: OFFICE se dodeli le, ko je Count (RAM) STROGO > praga (" & CStr(ScoreThreshold) & ")."

    Dim score() As Double, idx() As Long
    Dim j As Long
    ReDim score(1 To daysWidth)
    ReDim idx(1 To daysWidth)

    For j = 1 To daysWidth
        score(j) = cntArr(j)
        idx(j) = j
    Next j

    QuickSortDesc score, idx, 1, daysWidth

    Dim arrTop() As Variant
    ReDim arrTop(1 To daysWidth, 1 To 3)
    
    Dim topN As Long
    Dim dd As String
    topN = 0

    For j = 1 To daysWidth
        dd = UCase$(Trim$(CStr(wsG.Cells(ROW_DDGAMA, FIRST_DATE_COL_G + idx(j) - 1).Value)))

        ' BEST DAYS: izloci vikende in praznike
        If dd <> "SO" And dd <> "NE" And dd <> "PR" Then
            topN = topN + 1
            arrTop(topN, 1) = idx(j)
            arrTop(topN, 2) = Format$(wsG.Cells(ROW_DATES_G, FIRST_DATE_COL_G + idx(j) - 1).Value, "d.m.")
            arrTop(topN, 3) = score(j)
        End If
    Next j

    wsR.Cells(5, startCol).Resize(1, 3).Value = Array("Št. dneva", "Datum", "Count (RAM)")
    wsR.Cells(5, startCol).Resize(1, 3).Font.Bold = True
    If topN > 0 Then
        wsR.Cells(6, startCol).Resize(topN, 3).Value = arrTop
    Else
        wsR.Cells(6, startCol).Value = "Ni delovnih dni za prikaz."
    End If

    ' =========================
    ' 3) LOG
    ' =========================
    Dim startRowLog As Long
    startRowLog = 7 + daysWidth + 4

    wsR.Cells(startRowLog - 2, startCol).Value = "DODELJEVANJE OFFICE (log):"
    wsR.Cells(startRowLog - 2, startCol).Font.Bold = True

    If logCount = 0 Then
        wsR.Cells(startRowLog, startCol).Value = "Ni bilo mogoèe dodeliti OFFICE (ni dneva z Count (RAM) STROGO > praga " & CStr(ScoreThreshold) & ", ali ni veljavnih izmen/dni)."
    Else
        Dim arrLog() As Variant
        ReDim arrLog(1 To logCount, 1 To 1)
        For j = 1 To logCount
            arrLog(j, 1) = logArr(j)
        Next j
        wsR.Cells(startRowLog, startCol).Resize(logCount, 1).Value = arrLog
    End If

    ' =========================
    ' 4) ANALIZA po osebah + %
    ' =========================
    ' Postavi analizo POD log, da ne prepiše BEST DAYS / log
    Dim rowA As Long, r As Long
    Dim stDel As Long, stOff As Long, ideal As Long
    Dim pct As Double
    Dim w As Long
    Dim totalIdeal As Long, totalOff As Long

    totalIdeal = 0
    totalOff = 0

    rowA = startRowLog + IIf(logCount = 0, 2, logCount + 2) + 2

    wsR.Cells(rowA, startCol).Value = "ANALIZA – dosežen odstotek OFFICE:"
    wsR.Cells(rowA, startCol).Font.Bold = True

    rowA = rowA + 2

    ' 3-stolpèni kompakt:
    ' A: Ime
    ' B: OFF/CILJ
    ' C: % uspešnosti + status
    wsR.Cells(rowA, startCol).Resize(1, 3).Value = Array("Ime", "OFF/CILJ", "% uspešnosti")
    wsR.Cells(rowA, startCol).Resize(1, 3).Font.Bold = True

    w = rowA + 1

    For r = firstRow To lastRow
        If hasPerson(r) Then

            stDel = 0
            stOff = 0

            For j = 1 To daysWidth
                Dim sKey As String
                sKey = modOffice_Logic.CanonicalShift(shArr(r, j))

                If modOffice.IsAnyShiftForQuota(sKey) Then stDel = stDel + 1
                If sKey = "O" Then stOff = stOff + 1
            Next j

            pct = wsG.Cells(r, COL_PCT).Value
            If Not IsNumeric(pct) Then pct = 1

            ideal = Round((1 - pct) * stDel, 0)
            If ideal < 0 Then ideal = 0

            totalIdeal = totalIdeal + ideal
            totalOff = totalOff + stOff

            wsR.Cells(w, startCol).Value = wsG.Cells(r, COL_NAME).Value
            wsR.Cells(w, startCol + 1).Value = CStr(stOff) & "od" & CStr(ideal)

            If ideal = 0 Then
                wsR.Cells(w, startCol + 2).Value = "—"
            Else
                wsR.Cells(w, startCol + 2).Value = Format$(stOff / ideal, "0%") & IIf(stOff >= ideal, "  DA", "  NE")
            End If

            w = w + 1
        End If
    Next r

    ' =========================
    ' 5) Skupni %
    ' =========================
    w = w + 1
    wsR.Cells(w, startCol).Value = "Skupni % uspešnosti enote:"
    wsR.Cells(w, startCol).Font.Bold = True

    If totalIdeal = 0 Then
        wsR.Cells(w, startCol + 1).Value = "—"
    Else
        wsR.Cells(w, startCol + 1).Value = Format$(totalOff / totalIdeal, "0.0%")
    End If

    wsR.Columns(startCol).Resize(, 3).AutoFit
End Sub


' =======================================================
'   QUICK SORT – DESCENDING
' =======================================================
Private Sub QuickSortDesc(ByRef s() As Double, _
                          ByRef idx() As Long, _
                          ByVal L As Long, _
                          ByVal r As Long)

    Dim i As Long, j As Long
    Dim x As Double
    Dim tmp As Double, tmpI As Long

    i = L: j = r
    x = s((L + r) \ 2)

    Do While i <= j
        Do While s(i) > x: i = i + 1: Loop
        Do While s(j) < x: j = j - 1: Loop

        If i <= j Then
            tmp = s(i): s(i) = s(j): s(j) = tmp
            tmpI = idx(i): idx(i) = idx(j): idx(j) = tmpI
            i = i + 1: j = j - 1
        End If
    Loop

    If L < j Then QuickSortDesc s, idx, L, j
    If i < r Then QuickSortDesc s, idx, i, r
End Sub


' ============================================================
'  ANALIZA – veè grafov, zloženi vertikalno (po enotah)
' ============================================================
Public Sub NarediGraf_Analiza_Block( _
        ByRef prvotni() As Double, _
        ByRef koncni() As Double, _
        ByVal daysWidth As Long, _
        ByVal wsG As Worksheet, _
        ByVal ROW_DATES_G As Long, _
        ByVal FIRST_DATE_COL_G As Long, _
        ByVal unitKey As String, _
        ByVal unitIdx As Long)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ANALIZA")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "ANALIZA"
    End If

    ' ============================
    ' 1) PODATKI – loèen blok na enoto
    ' ============================
    Dim startCol As Long
    startCol = 1 ' vedno zaèni v A, grafi bodo pod sabo

    Dim startRow As Long
    startRow = 1 + (unitIdx - 1) * (daysWidth + 6)

    Dim arr() As Variant
    Dim i As Long
    ReDim arr(1 To daysWidth, 1 To 4)

    For i = 1 To daysWidth
        arr(i, 1) = i
        arr(i, 2) = prvotni(i)
        arr(i, 3) = koncni(i)
        arr(i, 4) = wsG.Cells(ROW_DATES_G, FIRST_DATE_COL_G + i - 1).Value
    Next i

    ws.Cells(startRow, startCol).Resize(1, 4).Value = _
        Array("Dan", "Prvotni count", "Konèni count", "Datum")

    ws.Cells(startRow + 1, startCol).Resize(daysWidth, 4).Value = arr

    ws.Columns(startCol).Resize(, 4).AutoFit

    ' ============================
    ' 2) GRAF – po enoti
    ' ============================
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim chartName As String

    chartName = "GrafAnaliza_" & unitKey

    On Error Resume Next
    Set chartObj = ws.ChartObjects(chartName)
    On Error GoTo 0

    Dim chartTop As Double
    chartTop = ws.Cells(startRow + daysWidth + 2, startCol).Top

    If chartObj Is Nothing Then
        Set chartObj = ws.ChartObjects.Add( _
            Left:=50, _
            Top:=chartTop, _
            Width:=900, _
            Height:=350)
        chartObj.Name = chartName
    Else
        chartObj.Top = chartTop
    End If

    Set cht = chartObj.Chart
    cht.ChartType = xlColumnClustered

    ' poèisti samo serije
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    cht.SeriesCollection.NewSeries
    cht.SeriesCollection(1).Name = "Prvotni count"
    cht.SeriesCollection(1).Values = _
        ws.Range(ws.Cells(startRow + 1, startCol + 1), _
                 ws.Cells(startRow + daysWidth, startCol + 1))
    cht.SeriesCollection(1).XValues = _
        ws.Range(ws.Cells(startRow + 1, startCol + 3), _
                 ws.Cells(startRow + daysWidth, startCol + 3))

    cht.SeriesCollection.NewSeries
    cht.SeriesCollection(2).Name = "Konèni count"
    cht.SeriesCollection(2).Values = _
        ws.Range(ws.Cells(startRow + 1, startCol + 2), _
                 ws.Cells(startRow + daysWidth, startCol + 2))
    cht.SeriesCollection(2).XValues = _
        ws.Range(ws.Cells(startRow + 1, startCol + 3), _
                 ws.Cells(startRow + daysWidth, startCol + 3))

    cht.HasTitle = True
    cht.ChartTitle.Text = "Primerjava count – " & unitKey
    cht.Axes(xlCategory).TickLabels.NumberFormat = "d.m."
End Sub



