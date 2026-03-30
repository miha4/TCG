Attribute VB_Name = "modOffice_Logic"
Option Explicit
' =============================================================================
' modOffice_Logic
' -----------------------------------------------------------------------------
' Namen:
'   - poslovna logika + “canonicalizacija” šifer izmen
'   - helperji, ki jih uporabljata modOffice (run/orchestracija)
'     in modOfficeModels (model flow)
'
' Filozofija:
'   - CleanShiftText: oèisti tekst (NBSP, trim), ne spreminja case
'   - CanonicalShift: normalizira samo znane kode (X*, O, …), ostalo pusti
'   - IsBlockedDay: SO/NE/PR
'   - FindBestDayForPerson: vrne najboljšega j za osebo glede na cntArr
'   - ApplyOneOffice: izvede 1 dodelitev O v RAM + znižanje cnt + log
'   - ComputeOfficeNeed: izraèuna potrebo po O (officeNeed) za vsako osebo
'
' Opomba:
'   - Za overwritable rule uporabljaš modOffice.IsOverwritableByOffice(...)
'     (ker ta je settings-driven in trenutno živi v modOffice).
'   - Za logiranje uporabljaš modOffice.AppendOfficeLog(...)
' =============================================================================


' ---------------------------
' Shift text normalization
' ---------------------------

' Oèisti NBSP + Trim. Ne spreminja case, da ohraniš "Za", "Ze", ipd.
Public Function CleanShiftText(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, ChrW(160), " ")
    CleanShiftText = Trim$(s)
End Function

' Kanonizira samo znane “sistemske” kode.
' Ostalo pusti v originalni obliki (da "Za" ostane "Za", ne "ZA").
Public Function CanonicalShift(ByVal v As Variant) As String
    Dim raw As String: raw = CleanShiftText(v)
    Dim u As String: u = UCase$(raw)

    ' Poznane šifre: razširi po potrebi
    Select Case True
        Case Len(u) = 0
            CanonicalShift = ""
        Case u = "O"
            CanonicalShift = "O"
        Case Left$(u, 1) = "X"
            ' X1, X2, X3, ... (tudi X10)
            CanonicalShift = u
        Case Else
            ' èudne / custom izmene ohrani
            CanonicalShift = raw
    End Select
End Function

' Èe rabiš striktno primerjanje na "X1"/"X2"/"O" ne glede na case,
' uporabi to (ne bo spreminjalo originala v prikazu, samo za check).
Public Function ShiftKey(ByVal v As Variant) As String
    ShiftKey = UCase$(CleanShiftText(v))
End Function

' ------------------------------------------------------------------------------
' --------------------- RAZDELI DELAVCE MED ENOTE GLEDE NA IME TIMOV -----------
' -------------------------------------------------------------------------------

Public Function UnitFromCycleTag(ByVal tag As String) As String
    UnitFromCycleTag = modSettings.MapUnitFromTeamTag(tag)
End Function


' --------------------------
' Compute officeNeed
' ---------------------------

' Izraèuna officeNeed za vse osebe v range-u.
' Pravila se opirajo na:
'   - pctOper (0..1)
'   - quotaCells: koliko celic šteje v kvoto (odvisno od modOffice.IsAnyShiftForQuota)
'   - existingO
'   - overwritableCells (odvisno od modOffice.IsOverwritableByOffice)
'
' Vrne:
'   - officeNeed(firstRow..lastRow)
'   - totalNeed (ByRef)
Public Sub ComputeOfficeNeed( _
    ByVal wsG As Worksheet, _
    ByVal firstRow As Long, _
    ByVal lastRow As Long, _
    ByVal daysWidth As Long, _
    ByVal COL_PCT As Long, _
    ByRef shArr() As Variant, _
    ByRef hasPerson() As Boolean, _
    ByRef officeNeed() As Long, _
    ByRef totalNeed As Long, _
    ByRef dateArr() As Date)

    Dim r As Long, j As Long
    Dim sCan As String
    Dim pctOper As Double
    Dim quotaCells As Long, overwritableCells As Long, existingO As Long
    Dim quotaByMonth As Object, existingByMonth As Object
    Dim mk As Variant, qM As Long, eM As Long

    totalNeed = 0

    For r = firstRow To lastRow
        If hasPerson(r) Then

            pctOper = modOffice.ParsePct(wsG.Cells(r, COL_PCT).Value)

            quotaCells = 0
            overwritableCells = 0
            existingO = 0
            Set quotaByMonth = CreateObject("Scripting.Dictionary")
            Set existingByMonth = CreateObject("Scripting.Dictionary")

            For j = 1 To daysWidth
                sCan = CanonicalShift(shArr(r, j))

                If modOffice.IsAnyShiftForQuota(sCan) Then quotaCells = quotaCells + 1
                If sCan = "O" Then existingO = existingO + 1
                If modOffice.IsOverwritableByOffice(sCan) Then overwritableCells = overwritableCells + 1

                mk = Format$(dateArr(j), "yyyymm")
                If modOffice.IsAnyShiftForQuota(sCan) Then
                    If Not quotaByMonth.Exists(mk) Then quotaByMonth.Add mk, 0
                    quotaByMonth(mk) = CLng(quotaByMonth(mk)) + 1
                End If
                If sCan = "O" Then
                    If Not existingByMonth.Exists(mk) Then existingByMonth.Add mk, 0
                    existingByMonth(mk) = CLng(existingByMonth(mk)) + 1
                End If
            Next j

            officeNeed(r) = 0

            If overwritableCells > 0 And pctOper < 1 And quotaCells > 0 Then
                Dim totalDesired As Long
                Dim needMonth As Long
                totalDesired = 0

                For Each mk In quotaByMonth.Keys
                    qM = CLng(quotaByMonth(mk))
                    eM = 0
                    If existingByMonth.Exists(mk) Then eM = CLng(existingByMonth(mk))

                    needMonth = WorksheetFunction.Round((1 - pctOper) * qM, 0) - eM
                    If needMonth < 0 Then needMonth = 0
                    totalDesired = totalDesired + needMonth
                Next mk

                officeNeed(r) = totalDesired
                If officeNeed(r) < 0 Then officeNeed(r) = 0
                If officeNeed(r) > overwritableCells Then officeNeed(r) = overwritableCells
            End If

            totalNeed = totalNeed + officeNeed(r)
        End If
    Next r
End Sub


Sub DeleteAllComments()
'
' DeleteAllComments Macro
'
    ActiveWindow.SmallScroll Down:=48
    ActiveWindow.SmallScroll ToRight:=219
    ActiveWindow.SmallScroll Down:=27
    Cells.Select
    ActiveWindow.SmallScroll ToRight:=-46
    ActiveWindow.ScrollColumn = 218
    ActiveWindow.ScrollColumn = 212
    ActiveWindow.ScrollColumn = 210
    ActiveWindow.ScrollColumn = 148
    ActiveWindow.ScrollColumn = 125
    ActiveWindow.ScrollColumn = 90
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.SmallScroll Down:=-121
    Range("E3:MY177").Select
    Selection.ClearComments
End Sub


