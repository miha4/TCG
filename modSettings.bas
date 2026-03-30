Attribute VB_Name = "modSettings"
Option Explicit
' --- Logger state ---
Public gStage As String
' =============================================================================
' modSettings (refactor - key-based)
' -----------------------------------------------------------------------------
' Branje nastavitev iz lista NASTAVITVE:
'   - Kljui v stolpcu C, vrednosti v stolpcu B (kot na tvojem screenshotu)
'   - Cache indeks: key -> row (hitro, isto)
'   - Tipizirani getterji: Text/Long/Date/Bool
'   - LoadMainSettings / LoadOfficeSettings / LoadUnitConfigs
'   - Wrapperji za nazaj, da lahko ostale module migrira postopno
' =============================================================================
' ----------------------------
'  PUBLIC TYPES (izvozni tip)
' namesto loenih spremenljivk jih tukaj spravimo v katlo
' Public Type je paket povezanih spremenljivk
' ----------------------------
Public Type TMainSettings
    SheetName As String
    PathGama As String
    GamaSheetName As String
    firstDataRowG As Long
    lastDataRowG As Long
    colIdG As Long
    colNameG As Long
    ColOJT As Long
    ColTeam As Long
    ColOperativePct As Long
    colLicenseG As Long
    startDate As Date
    endDate As Date
    ' datumi na GAMA listu (glava)
    firstDateRowG As Long
    firstDateColG As Long
    RowDDGama As Long
    daysWidth As Long
    GamaStartDateCol As Long
    ' splono
    selectedUnitsText As String
    excludedTypesCsv As String
    ' Preview / PREDOGLED
    PrevFirstDateCol As Long   ' prvi stolpec datumov v PREDOGLEDU
    PrevFirstDataRow As Long   ' prva vrstica zaposlenih v PREDOGLEDU
End Type


Public Type TOfficeSettings
    PrevFirstDataRow As Long
    PrevColID As Long
    PrevFirstDateCol As Long
    CounterGreaterThan As Long
    OverwriteShiftsCsv As String  ' npr "X1,X2"
    AllowCopyEmptyAsDelete As Boolean ' Èe bo kdaj rabil
    CountAllShiftsForWorkday As Boolean  ' "Prepiši izmene" DA/NE
    DebugUI As Boolean
    OfficeModelMode As String      ' "GLOBAL" / "SEQUENTIAL"
    WeightSurplus As Double         ' UTEZI
    WeightMonthlyPct As Double

    '--- Periodini sestanki ---
    MeetActive As Boolean
    MeetName As String
    MeetDOW As String
    MeetNthMonth As String
    MeetIDsSpec As String
    MeetNote As String
End Type

Public Type TUnitConfig
    unitKey As String
    NoNightFL_IdsCsv As String
    Overwrite As Boolean
    allow3NonFL As Boolean
    CoreShiftsText As String
End Type


' ----------------------------
'  INTERNAL CACHE
' ----------------------------

' Slovar: klju (stolpec C) -> tevilka vrstice
' Uporablja se za hitro iskanje nastavitev (namesto ponovnega iskanja po listu).
Private mKeyIndex As Object  ' Scripting.Dictionary (late-bound)
Private mDuplicateKeys As Object ' key -> "r1,r2,..."

' Zapomni si, za kateri list je bil indeks zgrajen
' e se list zamenja, se indeks ponovno zgradi.
Private mKeyIndexSheetName As String


' Stolpec, kjer so kljui nastavitev (C)
Private Const KEY_COL As Long = 3

' Stolpec, kjer so vrednosti nastavitev (B)
' Uporablja se v vseh Get* funkcijah za branje vrednosti.
Private Const VAL_COL As Long = 2
Private Const MAX_UNITCFG_SCAN_ROWS As Long = 2000

' =============================================================================
'  ENTRY POINTS
' =============================================================================

Public Function LoadMainSettings(ByVal wsSet As Worksheet) As TMainSettings
    ' spravimo spremenljivke v TMainSettings; ko to zapiem, dobim v pomnilnik vse spremenljivke
    Dim s As TMainSettings ' ustvarimo prazen settings objekt
    s.SheetName = wsSet.Name ' shrani ime lista

    BuildKeyIndex wsSet 'Zgradimo slovar klju -> vrstica
    ' --- BRANJE OSNOVNIH NASTAVITEV
    LogStep "READ SETTINGS", "Required keys"

    LogStep "Get PathG":              s.PathGama = NormalizeSharePointPath(GetTextRequiredAny(wsSet, "COMMON.PATH_GAMA", "Pot do datoteke GAMA", "pot do datoteke GAMA"))
    LogStep "Get SheetG":             s.GamaSheetName = GetTextRequiredAny(wsSet, "COMMON.GAMA_SHEET", "Ime lista v GAMA")
    LogStep "Get firstRow":           s.firstDataRowG = GetLongRequiredAny(wsSet, 1, "GAMA.ROW_FIRST_EMP", "Prva vrstica zaposlenih v GAMA")
    LogStep "Get lastRow":            s.lastDataRowG = GetLongRequiredAny(wsSet, s.firstDataRowG, "GAMA.ROW_LAST_EMP", "Zadnja vrstica zaposlenih v GAMA")
    LogStep "Get COL_ID":             s.colIdG = GetLongRequiredAny(wsSet, 1, "GAMA.COL_ID", "Stolpec ID zaposlenega v GAMA (npr. 9)", "Stolpec ID zaposlenega v GAMA")
    LogStep "Get COL_NAME":           s.colNameG = GetLongRequiredAny(wsSet, 1, "GAMA.COL_NAME", "Stolpec ime in priimek v GAMA (npr. 10)", "Stolpec ime in priimek v GAMA")
    LogStep "Get COL_TYPE":           s.ColOJT = GetLongRequiredAny(wsSet, 1, "GAMA.COL_OJT", "Stolpec OJT v GAMA (npr. 3)", "Stolpec OJT v GAMA")
    LogStep "Get COL_CYCLE":          s.ColTeam = GetLongRequiredAny(wsSet, 1, "GAMA.COL_TEAM", "Stolpec tima v GAMA (npr. 8)", "Stolpec tima v GAMA")
    LogStep "Get COL_PCT":            s.ColOperativePct = GetLongRequiredAny(wsSet, 1, "GAMA.COL_OPERATIVE_PCT", "Odstotek operative v GAMA")
    LogStep "Get ROW_DATES_G":        s.firstDateRowG = GetLongRequiredAny(wsSet, 3, "GAMA.ROW_DATES", "Prva vrstica datumov v GAMA")
    LogStep "Get COL_DATES_G":        s.firstDateColG = GetLongRequiredAny(wsSet, 11, "GAMA.COL_FIRST_DATE", "Prvi stolpec datumov v GAMA")
    LogStep "Get ROW_DDGAMA":         s.RowDDGama = GetLongRequiredAny(wsSet, 188, "GAMA.ROW_DDGAMA", "vrstica DDGAMA (prazniki)")
    LogStep "Get DaysWidth":          s.daysWidth = GetLongRequiredAny(wsSet, 1, "COMMON.DAYS_WIDTH", "DaysWidth")
    LogStep "Get COL_LIC":            s.colLicenseG = GetLongOptionalAny(wsSet, 0, "GAMA.COL_LICENSE", "Stolpec licence v GAMA")
    LogStep "Get PrevFirstDateCol":   s.PrevFirstDateCol = GetLongRequiredAny(wsSet, 1, "PREVIEW.COL_FIRST_DATE", "Prvi stolpec datumov v PREDOGLED", "prvi stolpec datumov v PREDOGLED")
    LogStep "Get PrevFirstDataRow":   s.PrevFirstDataRow = GetLongRequiredAny(wsSet, 1, "PREVIEW.ROW_FIRST_EMP", "Prva vrstica zaposlenih v PREDOGLED")

    LogStep "Get GamaStartDateCol":   s.GamaStartDateCol = GetLongRequiredAny(wsSet, -999, "GAMA.COL_START_DATE", "Prvi stolpec datuma v GAMA (SE PRERAÈUNA AVTOMATSKO)")
    LogStep "Get selectedUnits":      s.selectedUnitsText = GetTextOptionalAny(wsSet, "", "FACTORY.SELECTED_UNITS", "Planirane enote (OKZP, FIS, FDT, BRN, MBX, POW, CEK)")
    LogStep "Get excludedTypes":      s.excludedTypesCsv = GetTextOptionalAny(wsSet, "", "FACTORY.EXCLUDED_OJT", "Izloèeni glede na OJT filter")


    LoadMainSettings = s
End Function

Public Function LoadOfficeSettings(ByVal wsSet As Worksheet) As TOfficeSettings
    Dim o As TOfficeSettings
    BuildKeyIndex wsSet

    LogStep "Get PrevFirstDataRow":     o.PrevFirstDataRow = GetLongRequiredAny(wsSet, 1, "PREVIEW.ROW_FIRST_EMP", "Prva vrstica zaposlenih v PREDOGLED")
    LogStep "Get PrevColID":            o.PrevColID = GetLongRequiredAny(wsSet, 2, "PREVIEW.COL_ID", "Stolpec ID v PREDOGLED")
    LogStep "Get PrevFirstDateCol":     o.PrevFirstDateCol = GetLongRequiredAny(wsSet, 1, "PREVIEW.COL_FIRST_DATE", "Prvi stolpec datumov v PREDOGLED", "prvi stolpec datumov v PREDOGLED")
    LogStep "Get CounterGreaterThan":   o.CounterGreaterThan = GetLongOptionalAny(wsSet, 0, "OFFICE.COUNTER_GT", "Števec veèji od", "Stevec veèji od")
    LogStep "Get OverwriteShiftsCsv":    o.OverwriteShiftsCsv = GetTextOptionalAny(wsSet, "", "OFFICE.OVERWRITE_SHIFTS", "Prepiši izmene")
    LogStep "Get CountAllShiftsForWorkday":   o.CountAllShiftsForWorkday = ParseBool(GetTextOptionalAny(wsSet, "DA", "OFFICE.COUNT_ALL_SHIFTS", "Štej vse izmene (DA/NE)"))
    LogStep "Get DebugUI":              o.DebugUI = ParseBool(GetTextOptionalAny(wsSet, "NE", "OFFICE.DEBUG_UI", "DEBUG UI (DA/NE)"))
    LogStep "Get OfficeModelMode":       o.OfficeModelMode = UCase$(Trim$(GetTextOptionalAny(wsSet, "GLOBAL", "OFFICE.MODEL", "Office model (GLOBAL / SEQUENTIAL):")))
    LogStep "Get WeightSurplus":        o.WeightSurplus = GetDoubleOptionalAny(wsSet, 0.5, "OFFICE.W_SURPLUS", "Utež viški")
    LogStep "Get WeightMonthlyPct":     o.WeightMonthlyPct = GetDoubleOptionalAny(wsSet, 0.5, "OFFICE.W_MONTHLY_PCT", "Utež meseèni odstotek")

        ' --- Periodièni sestanki ---
        ' ta je še ostala, iz navodila smo že dali
    LogStep "Get MeetActive":    o.MeetActive = ParseBool(GetTextOptionalAny(wsSet, "DA", "MEETING.ACTIVE", "AKTIVACIJA MAKROJA (DA / NE)", "Aktivacija makroja (DA / NE)"))
    LogStep "Get MeetName":      o.MeetName = Trim$(GetTextOptionalAny(wsSet, "Sestanek", "MEETING.NAME", "Ime sestanka za log in komentar"))
    LogStep "Get MeetDOW":       o.MeetDOW = Trim$(GetTextOptionalAny(wsSet, "", "MEETING.DOW", "Dan v tednu za sestanek"))
    LogStep "Get MeetNth":       o.MeetNthMonth = Trim$(GetTextOptionalAny(wsSet, "", "MEETING.NTH_MONTH", "Kateri dan v mesecu"))
    LogStep "Get MeetIDsSpec":   o.MeetIDsSpec = Trim$(GetTextOptionalAny(wsSet, "", "MEETING.IDS", "ID-ji zaposlenih"))
    LogStep "Get MeetNote":      o.MeetNote = Trim$(GetTextOptionalAny(wsSet, "", "MEETING.NOTE", "Opomba (periodièni sestanek)"))



    LoadOfficeSettings = o
End Function

' =============================================================
' Funkcija ki pregleda drugo tabelo kjer so podatki o enotah
' =============================================================

Public Function LoadUnitConfigs(ByVal wsSet As Worksheet, _
                                Optional ByVal headerText As String = "NASTAVITVE PO ENOTAH") As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    LogMsg "DBG", "LoadUnitConfigs: start (sheet=" & wsSet.Name & ")"
    BuildKeyIndex wsSet
    LogMsg "DBG", "LoadUnitConfigs: key index pripravljen"

    Dim hdrRow As Long
    hdrRow = FindTextRow(wsSet, headerText, KEY_COL, 1, 500)
    If hdrRow = 0 Then hdrRow = FindTextAnywhereRow(wsSet, headerText, 1, 500)
    LogMsg "DBG", "LoadUnitConfigs: hdrRow=" & hdrRow

    If hdrRow = 0 Then
        LogMsg "DBG", "LoadUnitConfigs: header '" & headerText & "' ni najden. Vrnem prazen config."
        Set LoadUnitConfigs = dict
        Exit Function
    End If

    Dim colUnit As Long, colNoNight As Long, colOverwrite As Long, colAllow3 As Long, colCore As Long
    ResolveUnitConfigColumns wsSet, hdrRow, colUnit, colNoNight, colOverwrite, colAllow3, colCore

    If colUnit = 0 Then colUnit = 2
    If colNoNight = 0 Then colNoNight = 3
    If colOverwrite = 0 Then colOverwrite = 4
    If colAllow3 = 0 Then colAllow3 = 5
    If colCore = 0 Then colCore = 6

    Dim r As Long: r = hdrRow + 2
    Dim lastR As Long
    lastR = LastUsedRowAcrossCols(wsSet, hdrRow + 2, colUnit, colNoNight, colOverwrite, colAllow3, colCore)
    If lastR < r Then lastR = r
    If (lastR - r + 1) > MAX_UNITCFG_SCAN_ROWS Then
        LogMsg "WARN", "LoadUnitConfigs: scan obseg je velik (" & (lastR - r + 1) & " vrstic). Omejim na " & MAX_UNITCFG_SCAN_ROWS & "."
        lastR = r + MAX_UNITCFG_SCAN_ROWS - 1
    End If

    LogMsg "DBG", "LoadUnitConfigs: hdrRow=" & hdrRow & " | firstDataRow=" & r & " | lastR=" & lastR

    Dim unitKey As String
    Do While r <= lastR
        LogMsg "DBG", "LoadUnitConfigs: row=" & r
    
        LogMsg "DBG", "LoadUnitConfigs: reading UNIT at row " & r
        unitKey = UCase$(Trim$(wsSet.Cells(r, colUnit).Value & ""))
        LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | unitKey='" & unitKey & "'"
    
        If Len(unitKey) = 0 Then
            LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | unitKey prazno"
            If IsUnitConfigRowEmpty(wsSet, r, colUnit, colNoNight, colOverwrite, colAllow3, colCore) Then
                LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | vrstica prazna -> Exit Do"
                Exit Do
            End If
            LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | vrstica ni prazna -> skip"
            GoTo NextUnitRow
        End If
    
        LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | raw NoNight='" & CStr(wsSet.Cells(r, colNoNight).Text) & "'"
        LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | raw Overwrite='" & CStr(wsSet.Cells(r, colOverwrite).Text) & "'"
        LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | raw Allow3NonFL='" & CStr(wsSet.Cells(r, colAllow3).Text) & "'"
        LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | raw CoreShifts='" & CStr(wsSet.Cells(r, colCore).Text) & "'"
    
        Dim cfg As Object
        Set cfg = CreateObject("Scripting.Dictionary")
        LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | cfg created"
    
        cfg("NoNightCsv") = Trim$(wsSet.Cells(r, colNoNight).Value & "")
        LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | NoNightCsv='" & cfg("NoNightCsv") & "'"
    
        LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | parse Overwrite"
        cfg("Overwrite") = ParseBoolSafe(wsSet.Cells(r, colOverwrite).Value, False, "Overwrite", unitKey, r)
        LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | parsed Overwrite=" & CStr(cfg("Overwrite"))
    
        LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | parse Allow3NonFL"
        cfg("Allow3NonFL") = ParseBoolSafe(wsSet.Cells(r, colAllow3).Value, False, "Allow3NonFL", unitKey, r)
        LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | parsed Allow3NonFL=" & CStr(cfg("Allow3NonFL"))
    
        cfg("CoreShiftsCsv") = Trim$(wsSet.Cells(r, colCore).Value & "")
        LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | CoreShiftsCsv='" & cfg("CoreShiftsCsv") & "'"
    
        Set dict(unitKey) = cfg
        LogMsg "DBG", "LoadUnitConfigs: row=" & r & " | saved unit '" & unitKey & "'"
    
NextUnitRow:
        r = r + 1
    Loop

    Set LoadUnitConfigs = dict
End Function

Private Function ParseBoolSafe(ByVal raw As Variant, ByVal defaultValue As Boolean, _
                               ByVal fieldName As String, ByVal unitKey As String, ByVal rowNo As Long) As Boolean
    Dim rawText As String
    rawText = VariantToLogText(raw)

    On Error GoTo PARSE_ERR
    ParseBoolSafe = ParseBool(rawText)
    Exit Function

PARSE_ERR:
    LogMsg "WARN", "LoadUnitConfigs: neveljaven bool '" & fieldName & "' za unit=" & unitKey & _
                   " v vrstici " & rowNo & " (vrednost='" & rawText & "'). Uporabim default=" & CStr(defaultValue)
    ParseBoolSafe = defaultValue
    Err.Clear
End Function

Private Function VariantToLogText(ByVal v As Variant) As String
    On Error GoTo FALLBACK
    If IsError(v) Then
        VariantToLogText = "#ERROR"
    ElseIf IsEmpty(v) Then
        VariantToLogText = ""
    Else
        VariantToLogText = CStr(v)
    End If
    Exit Function

FALLBACK:
    VariantToLogText = "#UNPRINTABLE"
    Err.Clear
End Function

Private Function LastUsedRowAcrossCols(ByVal ws As Worksheet, ByVal firstDataRow As Long, _
                                       ByVal colUnit As Long, ByVal colNoNight As Long, _
                                       ByVal colOverwrite As Long, ByVal colAllow3 As Long, _
                                       ByVal colCore As Long) As Long
    Dim c As Variant
    Dim cols As Variant
    cols = Array(colUnit, colNoNight, colOverwrite, colAllow3, colCore)

    Dim lastR As Long
    lastR = firstDataRow - 1

    For Each c In cols
        If CLng(c) > 0 Then
            lastR = WorksheetFunction.Max(lastR, LastUsedRowFast(ws, CLng(c)))
        End If
    Next c

    LastUsedRowAcrossCols = lastR
End Function

Private Function FindTextAnywhereRow(ByVal ws As Worksheet, ByVal needle As String, _
                                     ByVal rMin As Long, ByVal rMax As Long) As Long
    Dim rng As Range
    Dim found As Range

    On Error Resume Next
    Set rng = ws.Range(ws.Cells(rMin, 1), ws.Cells(rMax, ws.UsedRange.Columns.Count))
    On Error GoTo 0

    If rng Is Nothing Then
        FindTextAnywhereRow = 0
        Exit Function
    End If

    Set found = rng.Find(What:=needle, After:=rng.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, _
                         SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If found Is Nothing Then
        FindTextAnywhereRow = 0
    Else
        FindTextAnywhereRow = found.Row
    End If
End Function
' ---------------------------------------------------------------------------------------------------------------------
' -------- Funkcija, ki najde kjer je stolpec enot, stolpec NoNight, stolpec AllowX3.... v tabeli nastavitve po enotah
' -------------------------------------------------------------------------------------------------------------------

Private Sub ResolveUnitConfigColumns(ByVal wsSet As Worksheet, ByVal hdrRow As Long, _
                                     ByRef colUnit As Long, ByRef colNoNight As Long, _
                                     ByRef colOverwrite As Long, ByRef colAllow3 As Long, _
                                     ByRef colCore As Long)

    colUnit = 0
    colNoNight = 0
    colOverwrite = 0
    colAllow3 = 0
    colCore = 0

    Dim scanRow As Long, c As Long, lastCol As Long
    Dim rawKey As String, key As String
    Dim usedCols As Long

    LogMsg "DBG", "ResolveUnitConfigColumns: start | hdrRow=" & hdrRow

    On Error Resume Next
    usedCols = wsSet.UsedRange.Columns.Count
    On Error GoTo 0

    LogMsg "DBG", "ResolveUnitConfigColumns: UsedRange.Columns.Count=" & usedCols

    For scanRow = hdrRow + 1 To hdrRow + 4

        lastCol = wsSet.Cells(scanRow, wsSet.Columns.Count).End(xlToLeft).Column

        If usedCols > 0 Then
            lastCol = WorksheetFunction.Min(lastCol, usedCols)
        End If

        If lastCol < 1 Then lastCol = 1

        LogMsg "DBG", "ResolveUnitConfigColumns: scanRow=" & scanRow & " | lastCol=" & lastCol

        For c = 1 To lastCol
            rawKey = CStr(wsSet.Cells(scanRow, c).Value & "")
            key = UCase$(Trim$(rawKey))
            key = Replace$(key, "_", "")
            key = Replace$(key, " ", "")

            If Len(Trim$(rawKey)) > 0 Then
                LogMsg "DBG", "ResolveUnitConfigColumns: row=" & scanRow & _
                              " col=" & c & _
                              " | raw='" & Replace(rawKey, "'", "''") & "'" & _
                              " | norm='" & key & "'"
            End If

            If key = "UNIT" Then
                colUnit = c
                LogMsg "DBG", "ResolveUnitConfigColumns: FOUND UNIT at col " & c & " (row " & scanRow & ")"
            End If

            If key = "NONIGHTFLIDS" Or key = "NONIGHTIDS" Then
                colNoNight = c
                LogMsg "DBG", "ResolveUnitConfigColumns: FOUND NoNight at col " & c & " (row " & scanRow & ")"
            End If

            If key = "OVERWRITE" Then
                colOverwrite = c
                LogMsg "DBG", "ResolveUnitConfigColumns: FOUND Overwrite at col " & c & " (row " & scanRow & ")"
            End If

            If key = "ALLOW3NONFL" Then
                colAllow3 = c
                LogMsg "DBG", "ResolveUnitConfigColumns: FOUND Allow3NonFL at col " & c & " (row " & scanRow & ")"
            End If

            If key = "CORESHIFTS" Or key = "CORESHIFTSCSV" Then
                colCore = c
                LogMsg "DBG", "ResolveUnitConfigColumns: FOUND CoreShifts at col " & c & " (row " & scanRow & ")"
            End If
        Next c

        LogMsg "DBG", "ResolveUnitConfigColumns: after scanRow=" & scanRow & _
                      " | colUnit=" & colUnit & _
                      " | colNoNight=" & colNoNight & _
                      " | colOverwrite=" & colOverwrite & _
                      " | colAllow3=" & colAllow3 & _
                      " | colCore=" & colCore

        If colUnit > 0 And colNoNight > 0 Then
            LogMsg "DBG", "ResolveUnitConfigColumns: enough columns found, exit"
            Exit Sub
        End If
    Next scanRow

    LogMsg "DBG", "ResolveUnitConfigColumns: finished | colUnit=" & colUnit & _
                  " | colNoNight=" & colNoNight & _
                  " | colOverwrite=" & colOverwrite & _
                  " | colAllow3=" & colAllow3 & _
                  " | colCore=" & colCore
End Sub

Private Function LastUsedRowFast(ByVal ws As Worksheet, ByVal col As Long) As Long
    Dim lastCell As Range
    Set lastCell = ws.Columns(col).Find(What:="*", After:=ws.Cells(1, col), LookIn:=xlFormulas, _
                                        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If lastCell Is Nothing Then
        LastUsedRowFast = 0
    Else
        LastUsedRowFast = lastCell.Row
    End If
End Function

Private Function IsUnitConfigRowEmpty(ByVal ws As Worksheet, ByVal r As Long, _
                                      ByVal colUnit As Long, ByVal colNoNight As Long, _
                                      ByVal colOverwrite As Long, ByVal colAllow3 As Long, _
                                      ByVal colCore As Long) As Boolean
    IsUnitConfigRowEmpty = (Len(Trim$(ws.Cells(r, colUnit).Value & "")) = 0 And _
                            Len(Trim$(ws.Cells(r, colNoNight).Value & "")) = 0 And _
                            Len(Trim$(ws.Cells(r, colOverwrite).Value & "")) = 0 And _
                            Len(Trim$(ws.Cells(r, colAllow3).Value & "")) = 0 And _
                            Len(Trim$(ws.Cells(r, colCore).Value & "")) = 0)
End Function

Private Sub BuildKeyIndex(ByVal wsSet As Worksheet)
    ' rebuild only if:
    ' - not built yet, or
    ' - different sheet
    If Not mKeyIndex Is Nothing Then
        If mKeyIndexSheetName = wsSet.Name Then Exit Sub
    End If

    Set mKeyIndex = CreateObject("Scripting.Dictionary")
    mKeyIndex.CompareMode = vbTextCompare
    Set mDuplicateKeys = CreateObject("Scripting.Dictionary")
    mDuplicateKeys.CompareMode = vbTextCompare
    mKeyIndexSheetName = wsSet.Name

    Dim lastRow As Long
    lastRow = LastUsedRowFast(wsSet, KEY_COL)
    If lastRow < 1 Then lastRow = 1

    Dim r As Long, k As String
    For r = 1 To lastRow
        k = Trim$(wsSet.Cells(r, KEY_COL).Value & "")
        If Len(k) > 0 Then
            If Not mKeyIndex.Exists(k) Then
                mKeyIndex(k) = r
            Else
                If mDuplicateKeys.Exists(k) Then
                    mDuplicateKeys(k) = CStr(mDuplicateKeys(k)) & "," & CStr(r)
                Else
                    mDuplicateKeys.Add k, CStr(mKeyIndex(k)) & "," & CStr(r)
                End If
            End If
        End If
    Next r
End Sub

Public Function GetDuplicateKeys(ByVal wsSet As Worksheet) As Object
    BuildKeyIndex wsSet

    Dim out As Object: Set out = CreateObject("Scripting.Dictionary")
    out.CompareMode = vbTextCompare

    Dim k As Variant
    For Each k In mDuplicateKeys.Keys
        out.Add CStr(k), CStr(mDuplicateKeys(k))
    Next k

    Set GetDuplicateKeys = out
End Function

' =============================================================================
'   LOOKUP FUNKCIJA
'   Vrne tevilko vrstice, kjer se nahaja doloen klju v NASTAVITVE
'   To je funkcija, ki iz cache slovarja vrne vrstico za doloen klju  ali 0, e klju ne obstaja.
' ==============================================================================
Private Function KeyRow(ByVal wsSet As Worksheet, ByVal keyText As String) As Long
    ' e e ni zgrajen, zgradi slovar, e je e, se ne zgodi ni
    BuildKeyIndex wsSet
    ' tukaj pridobimo vrstico od kljua
    If mKeyIndex.Exists(keyText) Then
        KeyRow = CLng(mKeyIndex(keyText))
    Else
        KeyRow = 0
    End If
End Function

' =============================================================================
'   GET TEXT OPTIONAL - Dobi text za doloeno nastavitev
'   e klju manjka, uporabi default, zato optional
' =============================================================================
Public Function GetTextOptionalAny(ByVal wsSet As Worksheet, ByVal defaultValue As String, ParamArray keyCandidates()) As String
    Dim i As Long, k As String, r As Long
    For i = LBound(keyCandidates) To UBound(keyCandidates)
        k = CStr(keyCandidates(i))
        r = KeyRow(wsSet, k)
        If r > 0 Then
            GetTextOptionalAny = Trim$(wsSet.Cells(r, VAL_COL).Value & "")
            Exit Function
        End If
    Next i
    GetTextOptionalAny = defaultValue
End Function

' =============================================================================
'   GET TEXT REQUIRED - Dobi text za doloeno nastavitev
'   Preveri, e klju obstaja; e ne, da napako
'   Prebere vrednost, odstrani morebitne presledke; e je prazna, vre napako
' =============================================================================
Public Function GetTextRequiredAny(ByVal wsSet As Worksheet, ParamArray keyCandidates()) As String
    Dim i As Long, k As String, r As Long, v As String
    For i = LBound(keyCandidates) To UBound(keyCandidates)
        k = CStr(keyCandidates(i))
        r = KeyRow(wsSet, k)
        If r > 0 Then
            v = Trim$(wsSet.Cells(r, VAL_COL).Value & "")
            If Len(v) = 0 Then Err.Raise vbObjectError + 1002, "modSettings", "Prazna vrednost za kljuc: '" & k & "'"
            GetTextRequiredAny = v
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 1001, "modSettings", "Manjka kljuc (aliases): '" & JoinAliases(keyCandidates) & "'"
End Function

' =============================================================================
'   GET LONG OPTIONAL - Vrne tevilko iz nastavitev, e ni nastavitve, vrne default
' =============================================================================
Public Function GetLongRequiredAny(ByVal wsSet As Worksheet, ByVal minValue As Long, ParamArray keyCandidates()) As Long
    Dim arr As Variant
    If IsArray(keyCandidates(0)) Then
        arr = keyCandidates(0)
    Else
        arr = keyCandidates
    End If

    Dim t As String
    t = FindRequiredAnyText(wsSet, arr)
    If Not IsNumeric(t) Then Err.Raise vbObjectError + 1011, "modSettings", "Vrednost ni tevilka (aliases): '" & JoinAliases(arr) & "'"
    Dim v As Long: v = CLng(t)
    If v < minValue Then Err.Raise vbObjectError + 1012, "modSettings", "Vrednost mora biti >= " & minValue & " (je " & v & ")"
    GetLongRequiredAny = v
End Function

Public Function GetLongOptionalAny(ByVal wsSet As Worksheet, ByVal defaultValue As Long, ParamArray keyCandidates()) As Long
    Dim arr As Variant
    If UBound(keyCandidates) >= LBound(keyCandidates) Then
        If IsArray(keyCandidates(0)) Then
            arr = keyCandidates(0)
        Else
            arr = keyCandidates
        End If
    Else
        GetLongOptionalAny = defaultValue
        Exit Function
    End If

    Dim t As String
    t = FindOptionalAnyText(wsSet, arr, vbNullString)
    If Len(Trim$(t)) = 0 Then
        GetLongOptionalAny = defaultValue
        Exit Function
    End If

    If Not IsNumeric(t) Then
        Err.Raise vbObjectError + 1014, "modSettings", "Vrednost ni številka (aliases): '" & JoinAliases(arr) & "'"
    End If
    GetLongOptionalAny = CLng(t)
End Function
' =============================================================================
'   GET DOUBLE OPTIONAL - Vrne decimalno številko iz nastavitev, sicer default
' =============================================================================
Public Function GetDoubleOptionalAny(ByVal wsSet As Worksheet, ByVal defaultValue As Double, ParamArray keyCandidates()) As Double
    Dim i As Long, k As String, r As Long, t As String
    For i = LBound(keyCandidates) To UBound(keyCandidates)
        k = CStr(keyCandidates(i))
        r = KeyRow(wsSet, k)
        If r > 0 Then
            t = Trim$(wsSet.Cells(r, VAL_COL).Value & "")
            If Len(t) = 0 Then
                GetDoubleOptionalAny = defaultValue
                Exit Function
            End If
            t = Replace(t, ",", ".")
            If IsNumeric(t) Then
                GetDoubleOptionalAny = CDbl(t)
            Else
                Err.Raise vbObjectError + 1013, "modSettings", "Vrednost ni številka za kljuè: '" & k & "' (vrednost='" & t & "')"
            End If
            Exit Function
        End If
    Next i
    GetDoubleOptionalAny = defaultValue
End Function


' =============================================================================
'   GET LONG REQUIRED - Vrne tevilko iz nastavitev, e ni nastavitve, vrne napako
' =============================================================================
Public Function GetLongRequired(ByVal wsSet As Worksheet, ByVal keyText As String, Optional ByVal minValue As Long = -6666) As Long
    Dim t As String: t = GetTextRequired(wsSet, keyText)
    If Not IsNumeric(t) Then Err.Raise vbObjectError + 1011, "modSettings", "Vrednost ni tevilka za klju: '" & keyText & "' (vrednost='" & t & "')"

    Dim v As Long: v = CLng(t)
    If v < minValue Then Err.Raise vbObjectError + 1012, "modSettings", "Vrednost za '" & keyText & "' mora biti >= " & minValue & " (je " & v & ")"
    GetLongRequired = v
End Function

Public Function GetDateRequiredAny(ByVal wsSet As Worksheet, ParamArray keyCandidates()) As Date
    Dim arr As Variant
    If IsArray(keyCandidates(0)) Then
        arr = keyCandidates(0)
    Else
        arr = keyCandidates
    End If

    Dim t As String
    t = FindRequiredAnyText(wsSet, arr)
    If IsDate(t) Then
        GetDateRequiredAny = CDate(t)
    Else
        Err.Raise vbObjectError + 1022, "modSettings", "Neveljaven datum (aliases): '" & JoinAliases(arr) & "' (vrednost='" & t & "')"
    End If
End Function



' =============================================================================
'   GET DATE REQUIRED - Preveri ali klju obstaja in nato prebere datum (tudi e je string ali date...)
' =============================================================================
Public Function GetDateRequired(ByVal wsSet As Worksheet, ByVal keyText As String) As Date
    Dim r As Long: r = KeyRow(wsSet, keyText)
    If r = 0 Then Err.Raise vbObjectError + 1020, "modSettings", "Manjka datum klju: '" & keyText & "'"

    Dim v As Variant: v = wsSet.Cells(r, VAL_COL).Value
    If IsDate(v) Then
        GetDateRequired = CDate(v)
    Else
        ' poskusi tudi preko CStr
        If IsDate(CStr(v)) Then
            GetDateRequired = CDate(CStr(v))
        Else
            Err.Raise vbObjectError + 1021, "modSettings", "Neveljaven datum za '" & keyText & "' (vrednost='" & (v & "") & "')"
        End If
    End If
End Function


Public Function GetTextRequired(ByVal wsSet As Worksheet, ByVal keyText As String) As String
    GetTextRequired = GetTextRequiredAny(wsSet, keyText)
End Function

Private Function FindRequiredAnyText(ByVal wsSet As Worksheet, ByVal keyCandidates As Variant) As String
    Dim t As String
    t = FindOptionalAnyText(wsSet, keyCandidates, vbNullString)
    If Len(Trim$(t)) = 0 Then
        Err.Raise vbObjectError + 1001, "modSettings", "Manjka kljuc (aliases): '" & JoinAliases(keyCandidates) & "'"
    End If
    FindRequiredAnyText = t
End Function

Private Function FindOptionalAnyText(ByVal wsSet As Worksheet, ByVal keyCandidates As Variant, ByVal defaultValue As String) As String
    Dim i As Long, k As String, r As Long
    For i = LBound(keyCandidates) To UBound(keyCandidates)
        k = CStr(keyCandidates(i))
        r = KeyRow(wsSet, k)
        If r > 0 Then
            FindOptionalAnyText = Trim$(wsSet.Cells(r, VAL_COL).Value & "")
            Exit Function
        End If
    Next i
    FindOptionalAnyText = defaultValue
End Function

Private Function JoinAliases(ByVal keyCandidates As Variant) As String
    Dim i As Long, s As String
    For i = LBound(keyCandidates) To UBound(keyCandidates)
        If Len(s) > 0 Then s = s & "', '"
        s = s & CStr(keyCandidates(i))
    Next i
    JoinAliases = s
End Function


' =============================================================================
'   PARSE BOOL: Pretvori DA -> 1 NE -> 0
' =============================================================================
Public Function ParseBool(ByVal raw As String) As Boolean
    Dim t As String: t = UCase$(Trim$(raw & ""))

    Select Case t
        Case "DA", "D", "YES", "Y", "TRUE", "1", "ON"
            ParseBool = True
        Case "NE", "N", "NO", "FALSE", "0", "OFF", ""
            ParseBool = False
        Case Else
            ' e je tevilka: >0 true
            If IsNumeric(t) Then
                ParseBool = (Val(t) <> 0)
            Else
                Err.Raise vbObjectError + 1030, "modSettings", "Ne razumem bool vrednosti: '" & raw & "' (priakujem DA/NE)"
            End If
    End Select
End Function


' =============================================================================
'  HELPERS (text search + path normalize + UI)
' =============================================================================

Private Function FindTextRow(ByVal ws As Worksheet, ByVal needle As String, _
                             ByVal col As Long, ByVal rMin As Long, ByVal rMax As Long) As Long
    Dim r As Long
    For r = rMin To rMax
        If Trim$(ws.Cells(r, col).Value & "") = needle Then
            FindTextRow = r
            Exit Function
        End If
    Next r
    FindTextRow = 0
End Function

Public Function NormalizeSharePointPath(ByVal p As String) As String
    ' Minimalna normalizacija. e ima specifino logiko za portal/sharepoint, jo daj sem.
    Dim t As String: t = Trim$(p)
    If Len(t) = 0 Then
        NormalizeSharePointPath = ""
        Exit Function
    End If

    ' odstrani narekovaje
    If Left$(t, 1) = """" And Right$(t, 1) = """" Then
        t = Mid$(t, 2, Len(t) - 2)
    End If

    NormalizeSharePointPath = t
End Function

Public Function OpenGamaWorkbook( _
    ByVal fullPath As String, _
    Optional ByVal readOnly As Boolean = False, _
    Optional ByVal showMessages As Boolean = True) As Workbook

    On Error GoTo FAIL

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase$(wb.fullName) = LCase$(fullPath) Then
            Set OpenGamaWorkbook = wb
            Exit Function
        End If
    Next wb

    If Len(fullPath) = 0 Then
        If showMessages Then MsgBox "Pot do GAMA je prazna.", vbCritical
        Set OpenGamaWorkbook = Nothing
        Exit Function
    End If

    If Dir(fullPath) = "" Then
        If showMessages Then MsgBox "GAMA datoteka ne obstaja:" & vbCrLf & fullPath, vbCritical
        Set OpenGamaWorkbook = Nothing
        Exit Function
    End If

    If readOnly Then
        Application.DisplayAlerts = False
        Set OpenGamaWorkbook = Workbooks.Open(Filename:=fullPath, UpdateLinks:=0, readOnly:=True, Notify:=False, AddToMru:=False)
        Application.DisplayAlerts = True
    Else
        Set OpenGamaWorkbook = Workbooks.Open(fullPath)
    End If

    Exit Function

FAIL:
    If showMessages Then
        MsgBox "Napaka pri odpiranju GAMA:" & vbCrLf & fullPath & vbCrLf & _
               "Detajl: " & Err.Description, vbCritical
    End If
    Set OpenGamaWorkbook = Nothing
End Function

Public Function GetWorksheetSafe(ByVal wb As Workbook, ByVal SheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetSafe = wb.Worksheets(SheetName)
    On Error GoTo 0
End Function

' ===========================================================================================================================
' ---- ZGRADI SLOVAR: Ko filtriram plan po enotah OKZP, FIS,... Vzame celico, kjer vpiem imena enot in iz njih zgradi slovar
Public Function NormalizeID(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace$(s, ChrW(160), " ")
    s = Trim$(s)
    NormalizeID = UCase$(s)
End Function

Public Function CsvToIdSet(ByVal raw As Variant) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim s As String
    s = CStr(raw & "")
    s = Replace$(s, vbCr, ",")
    s = Replace$(s, vbLf, ",")
    s = Replace$(s, vbTab, ",")
    s = Replace$(s, ";", ",")
    s = Replace$(s, ChrW(160), " ")
    s = Replace$(s, " ", ",")

    Do While InStr(s, ",,") > 0
        s = Replace$(s, ",,", ",")
    Loop
    If Left$(s, 1) = "," Then s = Mid$(s, 2)
    If Right$(s, 1) = "," Then s = Left$(s, Len(s) - 1)

    If Len(Trim$(s)) = 0 Then
        Set CsvToIdSet = d
        Exit Function
    End If

    Dim parts() As String, i As Long, p As String
    parts = Split(s, ",")
    For i = LBound(parts) To UBound(parts)
        p = NormalizeID(parts(i))
        If Len(p) > 0 Then
            If Not d.Exists(p) Then d.Add p, True
        End If
    Next i

    Set CsvToIdSet = d
End Function

'===========================================================================================================================

Public Function BuildUnitDict(ByVal s As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    s = UCase$(Trim$(s))
    If s = "" Then Set BuildUnitDict = d: Exit Function
    If s = "ALL" Or s = "VSE" Then d("ALL") = True: Set BuildUnitDict = d: Exit Function

    s = Replace(s, ";", ",")
    s = Replace(s, " ", "")
    Dim parts() As String: parts = Split(s, ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If parts(i) <> "" Then d(parts(i)) = True
    Next i

    Set BuildUnitDict = d
End Function

Public Function MapUnitFromTeamTag(ByVal tag As String) As String
    tag = UCase$(Trim$(tag))
    tag = Replace$(tag, " ", "")

    If tag Like "C#" Or tag Like "C##" Then
        MapUnitFromTeamTag = "OKZP": Exit Function
    ElseIf tag Like "BC#" Or tag Like "BC##" Then
        MapUnitFromTeamTag = "BRN": Exit Function
    ElseIf tag Like "MC#" Or tag Like "MC##" Then
        MapUnitFromTeamTag = "MBX": Exit Function
    ElseIf tag Like "PC#" Or tag Like "PC##" Then
        MapUnitFromTeamTag = "POW": Exit Function
    ElseIf tag Like "CC#" Or tag Like "CC##" Then
        MapUnitFromTeamTag = "CEK": Exit Function
    ElseIf tag Like "FDC#" Or tag Like "FDC##" Then
        MapUnitFromTeamTag = "FDT": Exit Function
    ElseIf tag Like "FIC#" Or tag Like "FIC##" Then
        MapUnitFromTeamTag = "FIS": Exit Function
    End If

    MapUnitFromTeamTag = ""
End Function

Public Sub ValidateSettings()
    Dim wsSet As Worksheet
    Set wsSet = ThisWorkbook.Worksheets("NASTAVITVE")

    Dim required As Variant
    required = Array( _
        "Pot do datoteke GAMA", "Ime lista v GAMA", "Prva vrstica zaposlenih v GAMA", _
        "Zadnja vrstica zaposlenih v GAMA", "Stolpec ID zaposlenega v GAMA (npr. 9)", _
        "Stolpec tima v GAMA (npr. 8)", "Prva vrstica datumov v GAMA", _
        "DaysWidth", "ZAÈETNI DATUM", "KONÈNI DATUM", _
        "Prvi stolpec datuma v GAMA (SE PRERAÈUNA AVTOMATSKO)")

    Dim errs As Collection: Set errs = New Collection
    Dim i As Long, k As String

    BuildKeyIndex wsSet
    For i = LBound(required) To UBound(required)
        k = CStr(required(i))
        If KeyRow(wsSet, k) = 0 Then errs.Add "Manjka kljuè: " & k
    Next i

    Dim dups As Object: Set dups = GetDuplicateKeys(wsSet)
    Dim dk As Variant
    For Each dk In dups.Keys
        errs.Add "Duplicate kljuè '" & CStr(dk) & "' v vrsticah: " & CStr(dups(dk))
    Next dk

    Dim s As TMainSettings
    On Error Resume Next
    s = LoadMainSettings(wsSet)
    If Err.Number <> 0 Then
        errs.Add "LoadMainSettings napaka: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    If s.endDate < s.startDate Then
        errs.Add "KONÈNI DATUM je pred ZAÈETNI DATUM."
    End If
    If CLng(s.endDate - s.startDate) + 1 <> s.daysWidth Then
        errs.Add "DaysWidth ni poravnan z datumoma (prièakovano " & (CLng(s.endDate - s.startDate) + 1) & ")."
    End If

    Dim wbG As Workbook, wsG As Worksheet
    Set wbG = OpenGamaWorkbook(s.PathGama, True, False)
    If wbG Is Nothing Then
        errs.Add "GAMA datoteke ni mogoèe odpreti: " & s.PathGama
    Else
        Set wsG = GetWorksheetSafe(wbG, s.GamaSheetName)
        If wsG Is Nothing Then
            errs.Add "GAMA list ne obstaja: " & s.GamaSheetName
        Else
            If Not IsDate(wsG.Cells(s.firstDateRowG, s.GamaStartDateCol).Value) Then
                errs.Add "GAMA start celica ni datum (" & s.firstDateRowG & "," & s.GamaStartDateCol & ")."
            ElseIf DateValue(wsG.Cells(s.firstDateRowG, s.GamaStartDateCol).Value) <> DateValue(s.startDate) Then
                errs.Add "Datum v GAMA start stolpcu ni poravnan z ZAÈETNI DATUM."
            End If

            Dim units As Object: Set units = BuildUnitDict(s.selectedUnitsText)
            If units.Count = 0 Or units.Exists("ALL") Then
                Set units = LoadUnitConfigs(wsSet)
            End If

            Dim uk As Variant, rngCheck As Range
            For Each uk In units.Keys
                If UCase$(CStr(uk)) <> "ALL" Then
                    Set rngCheck = Nothing
                    On Error Resume Next
                    Set rngCheck = wsG.Range("CountTbl" & CStr(uk))
                    On Error GoTo 0
                    If rngCheck Is Nothing Then errs.Add "Manjka named range CountTbl" & CStr(uk)
                End If
            Next uk
        End If
    End If

    If errs.Count > 0 Then
        Dim msg As String: msg = "ValidateSettings: najdene napake:" & vbCrLf
        For i = 1 To errs.Count
            msg = msg & " - " & CStr(errs(i)) & vbCrLf
        Next i
        MsgBox msg, vbCritical
    Else
        MsgBox "ValidateSettings: vse preveritve OK.", vbInformation
    End If
End Sub

' ===========================================================================================================================
' PREVERNJANJE ALI DOLOENA ENOTA SPADA MED IZBRANE ENOTE ZA PLANIRANJE
'===========================================================================================================================
Public Function IsUnitSelected(ByVal selectedUnits As Object, ByVal unitKey As String) As Boolean
    If selectedUnits Is Nothing Then IsUnitSelected = True: Exit Function
    If selectedUnits.Exists("ALL") Then IsUnitSelected = True: Exit Function
    IsUnitSelected = selectedUnits.Exists(UCase$(Trim$(unitKey)))
End Function





'===================================
' LOGGING
'===================================


Public Function GetStage() As String
    GetStage = gStage
End Function

Public Sub LogMsg(ByVal level As String, ByVal msg As String)
    Dim line As String
    line = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [" & level & "] " & msg
    Debug.Print line
End Sub

' Posodobi status v NASTAVITVE!G7 (+ Immediate)
Public Sub SetStatus(ByVal txt As String)
    On Error Resume Next
    With ThisWorkbook.Worksheets("NASTAVITVE").Range("H7")
        .Value = Format$(Now, "hh:nn:ss") & "  " & txt
    End With
    On Error GoTo 0
    DoEvents
End Sub

' Enotna toka za korake (stage)
Public Sub LogStep(ByVal stepName As String, Optional ByVal details As String = vbNullString)
    gStage = stepName
    If Len(details) > 0 Then
        LogMsg "STEP", stepName & " | " & details
        SetStatus stepName & " | " & details
    Else
        LogMsg "STEP", stepName
        SetStatus stepName
    End If
End Sub


