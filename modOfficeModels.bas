Attribute VB_Name = "modOfficeModels"
Option Explicit

' ============================================================
'  modOfficeModels
'  - modeli za dodeljevanje OFFICE ("O") v RAM (shArr)
'  - spreadMode: èe je kapacitete dovolj, razteguj O po obdobju
' ============================================================

' ---------------------------
' Day rules
' ---------------------------
Private Function IsBlockedDay(ByVal dayName As String) As Boolean
    ' dayArr je že UCase
    IsBlockedDay = (dayName = "SO" Or dayName = "NE" Or dayName = "PR")
End Function

' ---------------------------
' Candidate selection (best score)
' ---------------------------
Private Function FindBestDayForPerson( _
    ByVal r As Long, _
    ByVal daysWidth As Long, _
    ByRef cntArr() As Double, _
    ByRef shArr() As Variant, _
    ByRef dayArr() As String, _
    ByVal ScoreThreshold As Double, _
    ByRef anyDayAboveThreshold As Boolean) As Long

    Dim j As Long
    Dim bestJ As Long: bestJ = 0
    Dim bestScore As Double: bestScore = -1E+99

    Dim s As String, dn As String

    For j = 1 To daysWidth
        s = modOffice_Logic.CanonicalShift(shArr(r, j))

        If modOffice.IsOverwritableByOffice(s) Then
            dn = dayArr(j)

            If Not IsBlockedDay(dn) Then
                If cntArr(j) > ScoreThreshold Then
                    anyDayAboveThreshold = True
                    If cntArr(j) > bestScore Then
                        bestScore = cntArr(j)
                        bestJ = j
                    End If
                End If
            End If
        End If
    Next j

    FindBestDayForPerson = bestJ
End Function

' ---------------------------
' Candidate selection (next good day after lastJ, wrap-around)
' ---------------------------
Private Function FindNextGoodDayForPerson( _
    ByVal r As Long, _
    ByVal daysWidth As Long, _
    ByRef cntArr() As Double, _
    ByRef shArr() As Variant, _
    ByRef dayArr() As String, _
    ByVal ScoreThreshold As Double, _
    ByRef anyDayAboveThreshold As Boolean, _
    ByVal lastJ As Long) As Long

    Dim j As Long
    Dim s As String, dn As String
    Dim bestJ As Long: bestJ = 0
    Dim bestScore As Double: bestScore = -1E+99

    If lastJ < 0 Then lastJ = 0
    If lastJ > daysWidth Then lastJ = daysWidth

    ' 1) po zadnjem dodeljenem dnevu
    For j = lastJ + 1 To daysWidth
        s = modOffice_Logic.CanonicalShift(shArr(r, j))
        If modOffice.IsOverwritableByOffice(s) Then
            dn = dayArr(j)
            If Not IsBlockedDay(dn) Then
                If cntArr(j) > ScoreThreshold Then
                    anyDayAboveThreshold = True
                    If cntArr(j) > bestScore Then
                        bestScore = cntArr(j)
                        bestJ = j
                    End If
                End If
            End If
        End If
    Next j

    ' 2) wrap-around (od zaèetka do lastJ)
    For j = 1 To lastJ
        s = modOffice_Logic.CanonicalShift(shArr(r, j))
        If modOffice.IsOverwritableByOffice(s) Then
            dn = dayArr(j)
            If Not IsBlockedDay(dn) Then
                If cntArr(j) > ScoreThreshold Then
                    anyDayAboveThreshold = True
                    If cntArr(j) > bestScore Then
                        bestScore = cntArr(j)
                        bestJ = j
                    End If
                End If
            End If
        End If
    Next j

    FindNextGoodDayForPerson = bestJ
End Function

' ---------------------------
' Apply one assignment (RAM)
' ---------------------------
Private Sub ApplyOneOffice( _
    ByVal wsG As Worksheet, _
    ByVal COL_NAME As Long, _
    ByVal r As Long, _
    ByVal j As Long, _
    ByRef cntArr() As Double, _
    ByRef shArr() As Variant, _
    ByRef officeNeed() As Long, _
    ByRef totalNeed As Long, _
    ByRef addedOffice As Long, _
    ByRef dayArr() As String, _
    ByRef dateArr() As Date, _
    ByVal ScoreThreshold As Double)

    Dim cntBefore As Double: cntBefore = cntArr(j)

    shArr(r, j) = "O"
    addedOffice = addedOffice + 1

    cntArr(j) = cntArr(j) - 1
    If cntArr(j) < ScoreThreshold Then cntArr(j) = ScoreThreshold

    officeNeed(r) = officeNeed(r) - 1
    If officeNeed(r) < 0 Then officeNeed(r) = 0

    totalNeed = totalNeed - 1
    If totalNeed < 0 Then totalNeed = 0

    ' ---- log ----
    Dim who As String
    who = Trim$(CStr(wsG.Cells(r, COL_NAME).Value))
    If Len(who) = 0 Then who = "?"

    modOffice.AppendOfficeLog _
        "ASSIGN | who=" & who & _
        " | r=" & r & " | j=" & j & _
        " | date=" & Format$(dateArr(j), "dd.mm.yyyy") & _
        " | day=" & dayArr(j) & _
        " | scoreBefore=" & cntBefore & _
        " | scoreAfter=" & cntArr(j) & _
        " | needLeft=" & officeNeed(r) & _
        " | totalLeft=" & totalNeed
End Sub

Private Sub StopPersonNeed(ByVal r As Long, ByRef officeNeed() As Long, ByRef totalNeed As Long)
    modOffice.AppendOfficeLog "NO_CAND | r=" & r & " | need=" & officeNeed(r)
    totalNeed = totalNeed - officeNeed(r)
    If totalNeed < 0 Then totalNeed = 0
    officeNeed(r) = 0
End Sub

Private Sub SwapCand( _
    ByRef cand_r() As Long, ByRef cand_j() As Long, _
    ByRef cand_score() As Double, ByRef cand_need() As Long, _
    ByVal i As Long, ByVal k As Long)

    Dim tr As Long, tj As Long, tn As Long
    Dim ts As Double

    tr = cand_r(i): cand_r(i) = cand_r(k): cand_r(k) = tr
    tj = cand_j(i): cand_j(i) = cand_j(k): cand_j(k) = tj
    ts = cand_score(i): cand_score(i) = cand_score(k): cand_score(k) = ts
    tn = cand_need(i): cand_need(i) = cand_need(k): cand_need(k) = tn
End Sub

Private Function WeightedCandidateScore( _
    ByVal wsG As Worksheet, _
    ByVal COL_PCT As Long, _
    ByVal r As Long, _
    ByVal j As Long, _
    ByVal daysWidth As Long, _
    ByRef cntArr() As Double, _
    ByRef shArr() As Variant, _
    ByRef dateArr() As Date, _
    ByVal wSurplus As Double, _
    ByVal wMonthlyPct As Double) As Double

    Dim ws As Double, wM As Double
    ws = wSurplus
    wM = wMonthlyPct
    If (ws + wM) <= 0 Then ws = 1

    Dim scoreSurplus As Double
    scoreSurplus = cntArr(j)

    Dim pctOper As Double
    pctOper = modOffice.ParsePct(wsG.Cells(r, COL_PCT).Value)

    Dim y As Long, m As Long
    y = Year(dateArr(j))
    m = Month(dateArr(j))

    Dim d As Long
    Dim quotaM As Long, offM As Long
    Dim sCan As String

    quotaM = 0
    offM = 0
    For d = 1 To daysWidth
        If Year(dateArr(d)) = y And Month(dateArr(d)) = m Then
            sCan = modOffice_Logic.CanonicalShift(shArr(r, d))
            If modOffice.IsAnyShiftForQuota(sCan) Then quotaM = quotaM + 1
            If sCan = "O" Then offM = offM + 1
        End If
    Next d

    Dim targetM As Long, deficitM As Long
    targetM = WorksheetFunction.Round((1 - pctOper) * quotaM, 0)
    If targetM < 0 Then targetM = 0

    deficitM = targetM - offM
    If deficitM < 0 Then deficitM = 0

    WeightedCandidateScore = (ws * scoreSurplus) + (wM * deficitM)
End Function


' ============================================================
'  MODEL 1: FAIR PER ROUND
' ============================================================
Public Sub AssignOffice_FairPerRound( _
    ByVal wsG As Worksheet, _
    ByVal COL_NAME As Long, _
    ByVal COL_PCT As Long, _
    ByVal firstRow As Long, ByVal lastRow As Long, _
    ByVal daysWidth As Long, _
    ByRef cntArr() As Double, _
    ByRef shArr() As Variant, _
    ByRef hasPerson() As Boolean, _
    ByRef officeNeed() As Long, _
    ByRef dayArr() As String, _
    ByVal ScoreThreshold As Double, _
    ByRef totalNeed As Long, _
    ByRef addedOffice As Long, _
    ByRef dateArr() As Date, _
    ByVal spreadMode As Boolean, _
    ByVal WeightSurplus As Double, _
    ByVal WeightMonthlyPct As Double)

    Dim preventInfLoop As Long: preventInfLoop = 0

    Dim r As Long, j As Long
    Dim changed As Boolean
    Dim anyDayAboveThreshold As Boolean

    ' kandidatni seznami
    Dim maxCand As Long: maxCand = (lastRow - firstRow + 1)
    Dim cand_r() As Long, cand_j() As Long, cand_need() As Long
    Dim cand_score() As Double
    ReDim cand_r(1 To maxCand)
    ReDim cand_j(1 To maxCand)
    ReDim cand_need(1 To maxCand)
    ReDim cand_score(1 To maxCand)

    ' pointerji za spread
    Dim assigned() As Long
    ReDim assigned(firstRow To lastRow)

    ' original need (za gap)
    Dim needOrig() As Long
    ReDim needOrig(firstRow To lastRow)
    For r = firstRow To lastRow
        needOrig(r) = officeNeed(r)
    Next r

    Dim startAfter As Long, gap As Long, targetAfter As Long
    Dim candCount As Long, i As Long, k As Long

    Do While totalNeed > 0

        changed = False
        anyDayAboveThreshold = False

        ' 1) zberi kandidate (max 1/osebo)
        candCount = 0

        For r = firstRow To lastRow
            If hasPerson(r) And officeNeed(r) > 0 Then

                If spreadMode Then
                    ' baseline pointer (zadnji dodeljeni j)
                    startAfter = assigned(r)
                    If startAfter <= 0 Then
                        startAfter = ((r - firstRow) Mod daysWidth) + 1
                    End If

                    ' gap za razteg po obdobju
                    gap = 0
                    If needOrig(r) > 0 Then
                        gap = daysWidth \ (needOrig(r) + 1)
                        If gap < 1 Then gap = 1
                    End If

                    targetAfter = startAfter + gap
                    If targetAfter > daysWidth Then targetAfter = targetAfter - daysWidth

                    j = FindNextGoodDayForPerson(r, daysWidth, cntArr, shArr, dayArr, _
                                                ScoreThreshold, anyDayAboveThreshold, targetAfter)
                Else
                    j = FindBestDayForPerson(r, daysWidth, cntArr, shArr, dayArr, _
                                             ScoreThreshold, anyDayAboveThreshold)
                End If

                If j > 0 Then
                    candCount = candCount + 1
                    cand_r(candCount) = r
                    cand_j(candCount) = j
                    cand_score(candCount) = WeightedCandidateScore( _
                        wsG, COL_PCT, r, j, daysWidth, cntArr, shArr, dateArr, _
                        WeightSurplus, WeightMonthlyPct)
                    cand_need(candCount) = officeNeed(r)
                Else
                    StopPersonNeed r, officeNeed, totalNeed
                End If

            End If
        Next r

        If candCount = 0 Then Exit Do
        If Not anyDayAboveThreshold Then Exit Do

        ' 2) sort: need DESC, score DESC
        For i = 1 To candCount - 1
            For k = i + 1 To candCount
                If (cand_need(k) > cand_need(i)) Or _
                   (cand_need(k) = cand_need(i) And cand_score(k) > cand_score(i)) Then
                    SwapCand cand_r, cand_j, cand_score, cand_need, i, k
                End If
            Next k
        Next i

        ' 3) dodeljevanje (max 1/osebo na krog)
        Dim assignedThisRound() As Boolean
        ReDim assignedThisRound(firstRow To lastRow)

        For i = 1 To candCount

            If totalNeed <= 0 Then Exit For

            r = cand_r(i)
            j = cand_j(i)

            If assignedThisRound(r) Then GoTo NextCandidate
            If officeNeed(r) <= 0 Then GoTo NextCandidate

            ' re-check
            If cntArr(j) <= ScoreThreshold Then
                If spreadMode Then assigned(r) = j
                GoTo NextCandidate
            End If

            If IsBlockedDay(dayArr(j)) Then
                If spreadMode Then assigned(r) = j
                GoTo NextCandidate
            End If

            If Not modOffice.IsOverwritableByOffice(modOffice_Logic.CanonicalShift(shArr(r, j))) Then
                If spreadMode Then assigned(r) = j
                GoTo NextCandidate
            End If

            ApplyOneOffice wsG, COL_NAME, r, j, cntArr, shArr, _
                           officeNeed, totalNeed, addedOffice, _
                           dayArr, dateArr, ScoreThreshold

            assigned(r) = j
            assignedThisRound(r) = True
            changed = True

NextCandidate:
        Next i

        preventInfLoop = preventInfLoop + 1
        If preventInfLoop > 1000 Then Exit Do
        If Not changed Then Exit Do

    Loop
End Sub


' ============================================================
'  MODEL 2: GREEDY SEQUENTIAL
' ============================================================
Public Sub AssignOffice_GreedySequential( _
    ByVal wsG As Worksheet, _
    ByVal COL_NAME As Long, _
    ByVal firstRow As Long, ByVal lastRow As Long, _
    ByVal daysWidth As Long, _
    ByRef cntArr() As Double, _
    ByRef shArr() As Variant, _
    ByRef hasPerson() As Boolean, _
    ByRef officeNeed() As Long, _
    ByRef dayArr() As String, _
    ByVal ScoreThreshold As Double, _
    ByRef totalNeed As Long, _
    ByRef addedOffice As Long, _
    ByRef dateArr() As Date, _
    ByVal spreadMode As Boolean)

    Dim preventInfLoop As Long: preventInfLoop = 0

    Dim r As Long, bestJ As Long
    Dim changed As Boolean
    Dim anyDayAboveThreshold As Boolean

    Dim lastAssignedJ() As Long
    ReDim lastAssignedJ(firstRow To lastRow)

    ' original need (za gap)
    Dim needOrig() As Long
    ReDim needOrig(firstRow To lastRow)
    For r = firstRow To lastRow
        needOrig(r) = officeNeed(r)
    Next r

    Dim tries As Long
    Dim j As Long
    Dim startAfter As Long
    Dim gap As Long
    Dim targetAfter As Long

    Do While totalNeed > 0

        changed = False
        anyDayAboveThreshold = False

        For r = firstRow To lastRow
            If hasPerson(r) And officeNeed(r) > 0 Then

                If spreadMode Then
                    ' baseline pointer
                    startAfter = lastAssignedJ(r)
                    If startAfter <= 0 Then
                        startAfter = ((r - firstRow) Mod daysWidth) + 1
                    End If

                    ' gap za razteg po obdobju
                    gap = 0
                    If needOrig(r) > 0 Then
                        gap = daysWidth \ (needOrig(r) + 1)
                        If gap < 1 Then gap = 1
                    End If

                    targetAfter = startAfter + gap
                    If targetAfter > daysWidth Then targetAfter = targetAfter - daysWidth

                    ' poskusi najti veljaven dan, èe se kandidat vmes "pokvari", premakni naprej
                    For tries = 1 To daysWidth

                        j = FindNextGoodDayForPerson(r, daysWidth, cntArr, shArr, dayArr, _
                                                    ScoreThreshold, anyDayAboveThreshold, targetAfter)

                        If j = 0 Then
                            StopPersonNeed r, officeNeed, totalNeed
                            Exit For
                        End If

                        ' re-check (paranoja)
                        If cntArr(j) > ScoreThreshold _
                           And Not IsBlockedDay(dayArr(j)) _
                           And modOffice.IsOverwritableByOffice(modOffice_Logic.CanonicalShift(shArr(r, j))) Then

                            ApplyOneOffice wsG, COL_NAME, r, j, cntArr, shArr, _
                                           officeNeed, totalNeed, addedOffice, _
                                           dayArr, dateArr, ScoreThreshold

                            lastAssignedJ(r) = j
                            changed = True
                            Exit For
                        Else
                            ' kandidat ni šel skozi -> premakni iskanje naprej
                            targetAfter = j
                            lastAssignedJ(r) = j
                        End If

                    Next tries

                Else
                    bestJ = FindBestDayForPerson(r, daysWidth, cntArr, shArr, dayArr, _
                                                 ScoreThreshold, anyDayAboveThreshold)

                    If bestJ = 0 Then
                        StopPersonNeed r, officeNeed, totalNeed
                    ElseIf cntArr(bestJ) > ScoreThreshold Then
                        ApplyOneOffice wsG, COL_NAME, r, bestJ, cntArr, shArr, _
                                       officeNeed, totalNeed, addedOffice, _
                                       dayArr, dateArr, ScoreThreshold
                        changed = True
                    Else
                        StopPersonNeed r, officeNeed, totalNeed
                    End If
                End If

            End If
        Next r

        preventInfLoop = preventInfLoop + 1
        If preventInfLoop > 1000 Then Exit Do
        If Not changed Then Exit Do
        If Not anyDayAboveThreshold Then Exit Do

    Loop
End Sub



