Attribute VB_Name = "modOffice"
Option Explicit

' =============================================================================
' modOffice (prenovljen)
' -----------------------------------------------------------------------------
' Vloga:
'   - RUN / orchestracija: prebere nastavitve, odpre GAMA, naloi RAM,
'     poklie model (GLOBAL/SEQUENTIAL), naredi poroilo + graf,
'     prepie nazaj v PREDOGLED z UNDO snapshotom.
'   - SETTINGS-driven helperji (ostanejo tukaj):
'       * IsAnyShiftForQuota
'       * IsOverwritableByOffice
'       * BuildAllowedShiftDict
'       * ParseScore / ParsePct
'       * DebugBlock
'       * AppendOfficeLog
'   - Office poslovni helperji (IsBlockedDay, FindBestDayForPerson, )
'     so v modOffice_Logic.
'
' Opomba:
'   - Za "weird shifts" (Za/Ze/...) NE delamo globalnega UCase.
'     Za primerjave uporabljamo modOffice_Logic.CanonicalShift / ShiftKey.
'
' Glavna sprememba v tej verziji:
'   - NI ve MsgBox vpraanj ("Obdrim spremembe?", "Undo sestanki?")
'   - Snapshot (modUndo.BeginSnapshot) OSTANE aktiven po akciji,
'     da lahko uporabnik rono sproi UndoLastAction kadarkoli.
' =============================================================================


' --- log buffer (RAM) ---
Private logArr() As String
Private logCount As Long

' ------------------- DEBUG SWITCHES (privzete) -------------------
Private Const DEBUG_PRINT As Boolean = True
Private Const DEBUG_SAMPLE_LIMIT As Long = 8
Private Const DEBUG_STEP_BY_STEP As Boolean = False

' runtime config (iz Settings)
Private gDEBUG_UI As Boolean
Private gCountAllShiftsForWorkday As Boolean
Private gOverwriteAllowed As Object   ' Scripting.Dictionary (keys = dovoljene izmene za prepis)

' =============================================================================
'  A) GLAVNI MAKRO: OFFICE na PREDOGLED (rotacija)
' =============================================================================

Public Sub ApplyOfficeToPreview()

    Dim errNum As Long, eDesc As String
    Dim prevScreenUpdating As Boolean, prevEnableEvents As Boolean
    Dim prevCalc As XlCalculation

    Dim wbThis As Workbook
    Dim wsSet As Worksheet, wsP As Worksheet
    Dim wbG As Workbook, wsG As Worksheet

    ' --- settings structs ---
    Dim s As modSettings.TMainSettings
    Dim os As modSettings.TOfficeSettings

    ' --- extracted settings (za berljivost) ---
    Dim pathG As String, sheetG As String
    Dim firstRow As Long, lastRow As Long
    Dim COL_ID As Long, COL_NAME_G As Long, COL_PCT As Long
    Dim ROW_DATES_G As Long, ROW_DDGAMA As Long
    Dim daysWidth As Long
    Dim startD As Date
    Dim START_DATE_COL_G As Long
    Dim COL_CYCLE As Long
    Dim FIRST_DATE_COL_G As Long

    ' --- preview layout ---
    Dim PREV_FIRST_ROW As Long, PREV_ID_COL As Long, PREV_FIRST_DATE_COL As Long

    ' --- office config ---
    Dim OverwriteShiftsCsv As String
    Dim CountAllShiftsForWorkday As Boolean
    Dim cfgDebugUI As Boolean
    Dim OfficeModelMode As String
    Dim ScoreThreshold As Double
    Dim WeightSurplus As Double
    Dim WeightMonthlyPct As Double

    ' --- runtime ---
    Dim dateArr() As Date

    Set wbThis = ThisWorkbook
    Set wsSet = wbThis.Worksheets("NASTAVITVE")
    Set wsP = wbThis.Worksheets("PREDOGLED")

    modSettings.LogStep "--OFFICE--", "ApplyOfficeToPreview START"

    ' ---- app state backup ----
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    On Error GoTo CLEANUP_ERR
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' ============================================================
    ' 1) LOAD SETTINGS
    ' ============================================================
    On Error GoTo SETTINGS_ERR
    s = modSettings.LoadMainSettings(wsSet)
    os = modSettings.LoadOfficeSettings(wsSet)
    On Error GoTo CLEANUP_ERR

    ' --- main (GAMA) ---
    pathG = s.PathGama
    sheetG = s.GamaSheetName
    firstRow = s.firstDataRowG
    lastRow = s.lastDataRowG
    COL_ID = s.colIdG
    COL_NAME_G = s.colNameG
    COL_PCT = s.ColOperativePct
    ROW_DATES_G = s.firstDateRowG
    ROW_DDGAMA = s.RowDDGama
    daysWidth = s.daysWidth
    startD = s.startDate
    START_DATE_COL_G = s.GamaStartDateCol
    COL_CYCLE = s.ColTeam
    ' KLJUNI FIX: ta stolpec je start-date v GAMI (preraèunan)
    FIRST_DATE_COL_G = s.GamaStartDateCol

    ' --- preview layout (office settings) ---
    PREV_FIRST_ROW = os.PrevFirstDataRow
    PREV_ID_COL = os.PrevColID
    PREV_FIRST_DATE_COL = os.PrevFirstDateCol

    ' --- office config ---
    OverwriteShiftsCsv = os.OverwriteShiftsCsv
    CountAllShiftsForWorkday = os.CountAllShiftsForWorkday
    cfgDebugUI = os.DebugUI
    OfficeModelMode = os.OfficeModelMode
    ScoreThreshold = os.CounterGreaterThan
    WeightSurplus = os.WeightSurplus
    WeightMonthlyPct = os.WeightMonthlyPct

    modSettings.LogStep "--OFFICE--", _
        "Loaded (daysWidth=" & daysWidth & ", model=" & OfficeModelMode & ")"

    ' ============================================================
    ' 2) APPLY RUNTIME CONFIG (module globals)
    ' ============================================================
    gDEBUG_UI = cfgDebugUI
    gCountAllShiftsForWorkday = CountAllShiftsForWorkday
    Set gOverwriteAllowed = BuildAllowedShiftDict(OverwriteShiftsCsv)

    DebugBlock "SETUP", _
        "PREDOGLED: firstRow=" & PREV_FIRST_ROW & vbCrLf & _
        "ID col=" & PREV_ID_COL & ", firstDateCol=" & PREV_FIRST_DATE_COL & vbCrLf & _
        "Overwrite=" & OverwriteShiftsCsv & vbCrLf & _
        "Kvota non-empty=" & IIf(gCountAllShiftsForWorkday, "DA", "NE (samo X*)") & vbCrLf & _
        "DEBUG_UI=" & IIf(gDEBUG_UI, "DA", "NE") & vbCrLf & _
        "daysWidth=" & daysWidth & ", threshold(os.CounterGreaterThan)=" & ScoreThreshold & vbCrLf & _
        "Model=" & OfficeModelMode & vbCrLf & _
        "Uteži: viški=" & WeightSurplus & " | meseèni%=" & WeightMonthlyPct

    ' ============================================================
    ' 3) OPEN GAMA
    ' ============================================================
    modSettings.LogStep "--OFFICE--", "Open workbook"
    Set wbG = modSettings.OpenGamaWorkbook(pathG, False, True)
    If wbG Is Nothing Then GoTo CLEANUP_OK

    Set wsG = modSettings.GetWorksheetSafe(wbG, sheetG)
    If wsG Is Nothing Then
        MsgBox "List '" & sheetG & "' ne obstaja v GAMA.", vbCritical
        GoTo CLEANUP_OK
    End If

    ' ============================================================
    ' 4) SANITY CHECK START DATE (PREDOGLED + GAMA)
    ' ============================================================
    Dim pStart As Variant, gStart As Variant

    pStart = wsP.Cells(1, PREV_FIRST_DATE_COL).Value
    If Not IsDate(pStart) Then
        MsgBox "PREDOGLED nima veljavnega datuma v vrstici 1." & vbCrLf & _
               "Celica: " & wsP.Name & "!" & wsP.Cells(1, PREV_FIRST_DATE_COL).Address(0, 0), vbCritical
        GoTo CLEANUP_OK
    End If

    If DateValue(pStart) <> DateValue(startD) Then
        MsgBox "Start date ni poravnan med NASTAVITVE in PREDOGLED!" & vbCrLf & _
               "START_DATE: " & Format$(startD, "dd.mm.yyyy") & vbCrLf & _
               "PREDOGLED: " & Format$(DateValue(pStart), "dd.mm.yyyy"), vbCritical
        GoTo CLEANUP_OK
    End If

    If START_DATE_COL_G <= 0 Then
        MsgBox "GamaStartDateCol ni nastavljen (<=0).", vbCritical
        GoTo CLEANUP_OK
    End If

    gStart = wsG.Cells(ROW_DATES_G, START_DATE_COL_G).Value
    If Not IsDate(gStart) Then
        MsgBox "GAMA datum v START stolpcu ni veljaven." & vbCrLf & _
               "Celica: " & wsG.Name & "!" & wsG.Cells(ROW_DATES_G, START_DATE_COL_G).Address(0, 0) & vbCrLf & _
               "Text: " & wsG.Cells(ROW_DATES_G, START_DATE_COL_G).Text, vbCritical
        GoTo CLEANUP_OK
    End If

    If DateValue(gStart) <> DateValue(startD) Then
        MsgBox "Start date ni poravnan med NASTAVITVE in GAMA!" & vbCrLf & _
               "START_DATE: " & Format$(startD, "dd.mm.yyyy") & vbCrLf & _
               "GAMA: " & Format$(DateValue(gStart), "dd.mm.yyyy") & _
               " (col=" & START_DATE_COL_G & ")", vbCritical
        GoTo CLEANUP_OK
    End If

    modSettings.LogStep "--OFFICE-- SANITY", "Dates OK (" & Format$(startD, "dd.mm.yyyy") & ")"

    ' --- dateArr iz GAMA (1..daysWidth) ---
    Dim j As Long
    ReDim dateArr(1 To daysWidth)
    For j = 1 To daysWidth
        dateArr(j) = DateValue(wsG.Cells(ROW_DATES_G, START_DATE_COL_G + j - 1).Value)
    Next j

    ' ============================================================
    ' 5) READ PREDOGLED -> RAM (shArr + hasPerson + rowMap)
    ' ============================================================
    modSettings.LogStep "--OFFICE-- MAP", "Load preview into RAM"

    Dim shArr() As Variant, hasPerson() As Boolean, rowMap() As Long
    Dim PREV_LAST_ROW As Long
    PREV_LAST_ROW = wsP.Cells(wsP.Rows.Count, PREV_ID_COL).End(xlUp).Row

    If PREV_LAST_ROW < PREV_FIRST_ROW Then
        MsgBox "PREDOGLED je prazen. Najprej zaeni BuildPreviewFromCycles.", vbCritical
        GoTo CLEANUP_OK
    End If

    Dim unitOfRow() As String
    Dim tag As String

    ReDim shArr(firstRow To lastRow, 1 To daysWidth)
    ReDim hasPerson(firstRow To lastRow)
    ReDim rowMap(firstRow To lastRow)
    ReDim unitOfRow(firstRow To lastRow)

    Dim idToRow As Object: Set idToRow = CreateObject("Scripting.Dictionary")
    idToRow.CompareMode = vbTextCompare

    Dim r As Long, key As String
    For r = firstRow To lastRow
        key = NormID(wsG.Cells(r, COL_ID))
        If Len(key) > 0 Then idToRow(key) = r
    Next r

    Dim pr As Long, id As String, gRow As Long
    For pr = PREV_FIRST_ROW To PREV_LAST_ROW
        id = NormID(wsP.Cells(pr, PREV_ID_COL))
        If Len(id) > 0 Then
            If idToRow.Exists(id) Then
                gRow = CLng(idToRow(id))
                hasPerson(gRow) = True
                rowMap(gRow) = pr
                tag = Trim$(CStr(wsG.Cells(gRow, COL_CYCLE).Value))
                unitOfRow(gRow) = UnitFromCycleTag(tag)
                For j = 1 To daysWidth
                    shArr(gRow, j) = modOffice_Logic.CleanShiftText(wsP.Cells(pr, PREV_FIRST_DATE_COL + j - 1).Value)
                Next j
            End If
        End If
    Next pr

    Dim mapped As Long: mapped = 0
    For r = firstRow To lastRow
        If hasPerson(r) Then mapped = mapped + 1
    Next r

    DebugBlock "MAP", _
        "PREDOGLED rows=" & (PREV_LAST_ROW - PREV_FIRST_ROW + 1) & vbCrLf & _
        "Mapped persons=" & mapped

    ' ============================================================
    ' 6) BUILD UNIT LIST (multi-unit support)
    ' ============================================================

    Dim units As Object
    Set units = CreateObject("Scripting.Dictionary")
    units.CompareMode = vbTextCompare

    Dim x As Long

    For x = firstRow To lastRow
        If hasPerson(x) Then
            If Len(unitOfRow(x)) > 0 Then
                If Not units.Exists(unitOfRow(x)) Then
                    units.Add unitOfRow(x), True
                End If
            End If
        End If
    Next x

    DebugBlock "UNITS", "Detected units: " & Join(units.Keys, ", ")


    ' ============================================================
    ' 7) DAY ARRAYS + sanity
    ' ============================================================
    modSettings.LogStep "--OFFICE-- LOAD", "Count + day arrays"


    Dim dayArr() As String
    ReDim dayArr(1 To daysWidth)

    Const PREV_DAY_ROW As Long = 2
    Const PREV_FIRST_DAY_COL As Long = 5 ' E

    Dim prevDay As String, gamaDay As String
    Dim mismatchCount As Long: mismatchCount = 0
    Dim firstMismatch As String: firstMismatch = ""

    For j = 1 To daysWidth
        gamaDay = UCase$(Trim$(CStr(wsG.Cells(ROW_DDGAMA, START_DATE_COL_G + j - 1).Value)))
        dayArr(j) = gamaDay

        prevDay = UCase$(Trim$(CStr(wsP.Cells(PREV_DAY_ROW, PREV_FIRST_DAY_COL + j - 1).Value)))
        If Len(prevDay) > 0 And prevDay <> gamaDay Then
            mismatchCount = mismatchCount + 1
            If mismatchCount = 1 Then
                firstMismatch = "j=" & j & " | GAMA=" & gamaDay & " | PREDOGLED=" & prevDay & _
                               " | cell=" & wsP.Cells(PREV_DAY_ROW, PREV_FIRST_DAY_COL + j - 1).Address(0, 0)
            End If
        End If
    Next j

    If mismatchCount > 0 Then
        MsgBox "POZOR: Dnevi niso poravnani med GAMA in PREDOGLED!" & vbCrLf & _
               "Neujemanj: " & mismatchCount & vbCrLf & _
               "Prvo: " & firstMismatch & vbCrLf & vbCrLf & _
               "Prekinjam.", vbCritical
        GoTo CLEANUP_OK
    End If


    ' ============================================================
    ' 8) ASSIGN OFFICE
    ' ============================================================
    modSettings.LogStep "--OFFICE--ASSIGN", "Begin MULTI-UNIT (mode=" & OfficeModelMode & ")"

    Dim addedOfficeTotal As Long
    addedOfficeTotal = 0

    Dim localCntArr() As Double
    Dim localCntArrFinal() As Double
    Dim unitKey As Variant

    Dim unitIdx As Long
    unitIdx = 0
    Const BLOCK_W As Long = 4 ' 3 stolpci + 1 prazen ZA ANALIZO


    For Each unitKey In units.Keys
        unitIdx = unitIdx + 1

        Dim startColReport As Long
        Dim startColGraf As Long

        startColReport = 1 + (unitIdx - 1) * BLOCK_W
        startColGraf = 1 + (unitIdx - 1) * BLOCK_W


        ' --- cntArr per unit ---

        Dim cntArr() As Double
        cntArr = LoadCntArrFromGama(wsG, CStr(unitKey), FIRST_DATE_COL_G, daysWidth)

        ReDim localCntArr(1 To daysWidth)
        ReDim localCntArrFinal(1 To daysWidth)

        For j = 1 To daysWidth
            localCntArr(j) = cntArr(j)
        Next j

        ' --- filter hasPerson -> hasUnit ---
        Dim hasUnit() As Boolean
        ReDim hasUnit(firstRow To lastRow)

        Dim anyInUnit As Boolean: anyInUnit = False
        Dim q As Long
        For q = firstRow To lastRow
            If hasPerson(q) And unitOfRow(q) = CStr(unitKey) Then
                hasUnit(q) = True
                anyInUnit = True
            End If
        Next q

        If Not anyInUnit Then
            Debug.Print "Unit " & unitKey & ": no people mapped"
            GoTo NEXT_UNIT
        End If

        Dim addedOfficeUnit As Long
        addedOfficeUnit = 0

        GlobalOfficeRotacija wsG, firstRow, lastRow, daysWidth, _
                             START_DATE_COL_G, ROW_DDGAMA, ROW_DATES_G, _
                             COL_ID, COL_NAME_G, COL_PCT, _
                             cntArr, shArr, hasUnit, dayArr, dateArr, _
                             ScoreThreshold, OfficeModelMode, addedOfficeUnit, _
                             WeightSurplus, WeightMonthlyPct

        addedOfficeTotal = addedOfficeTotal + addedOfficeUnit

        ' snapshot za graf, e ga e uporablja
        For j = 1 To daysWidth
            localCntArrFinal(j) = cntArr(j)
        Next j

    ' ============================================================
    ' 8b) REPORT + GRAPH
    ' ============================================================
    modSettings.LogStep "--OFFICE-- REPORT", "Analiza + graf"

    Analiza.UstvariPorocilo_Block wsG, firstRow, lastRow, COL_NAME_G, COL_PCT, _
                                  shArr, hasUnit, daysWidth, _
                                  START_DATE_COL_G, ROW_DDGAMA, ROW_DATES_G, _
                                  cntArr, logArr, logCount, _
                                  CStr(unitKey), startColReport, ScoreThreshold

    Analiza.NarediGraf_Analiza_Block localCntArr, localCntArrFinal, daysWidth, _
    wsG, ROW_DATES_G, START_DATE_COL_G, _
    CStr(unitKey), unitIdx

NEXT_UNIT:
    Next unitKey

    Dim addedOffice As Long
    addedOffice = addedOfficeTotal


    ' ============================================================
    ' 10) WRITEBACK + UNDO SNAPSHOT
    ' ============================================================
    modSettings.LogStep "--OFFICE-- WRITEBACK", "Snapshot + write to PREDOGLED"

    Dim targetRange As Range
    Dim nChanged As Long
    Dim firstChangedCell As Range

    Set targetRange = GetPreviewTargetRange(wsP, hasPerson, rowMap, daysWidth, PREV_FIRST_DATE_COL)
    If targetRange Is Nothing Then GoTo CLEANUP_OK

    modUndo.BeginSnapshot targetRange, "OFFICE"

    WriteBackShArrToPreview wsP, firstRow, lastRow, daysWidth, PREV_FIRST_DATE_COL, _
                            shArr, hasPerson, rowMap, nChanged, firstChangedCell

    If Not firstChangedCell Is Nothing Then
        wsP.Activate
        Application.Goto firstChangedCell, True
    End If

    modSettings.LogStep "--OFFICE-- DONE ---", "addedOffice=" & addedOffice & ", changed=" & nChanged

    MsgBox "OFFICE dodeljen." & vbCrLf & _
           "Start: " & Format$(DateValue(gStart), "dd.mm.yyyy") & vbCrLf & _
           "Dodeljenih OFFICE: " & CStr(addedOffice) & vbCrLf & _
           "Spremenjenih celic: " & CStr(nChanged) & vbCrLf & _
           "Threshold: " & CStr(ScoreThreshold) & vbCrLf & vbCrLf & _
           "UNDO: UndoLastAction", vbInformation

CLEANUP_OK:
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalc
    On Error GoTo 0
    Exit Sub

SETTINGS_ERR:
    modSettings.LogStep "--OFFICE-- ERROR", "Read settings failed | " & Err.Number & " - " & Err.Description
    MsgBox "Napaka pri branju nastavitev: " & Err.Description, vbCritical
    GoTo CLEANUP_OK

CLEANUP_ERR:
    errNum = Err.Number
    eDesc = Err.Description

    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalc
    On Error GoTo 0

    modSettings.LogStep "--OFFICE-- ERROR", "ApplyOfficeToPreview | stage=" & modSettings.GetStage() & " | " & errNum & " - " & eDesc

    If errNum <> 0 Then
        MsgBox "Napaka: " & errNum & vbCrLf & eDesc, vbCritical
    End If
End Sub


' =============================================================================
'  B) CORE: GlobalOfficeRotacija (RAM-only)
' =============================================================================

Public Sub GlobalOfficeRotacija( _
    ByVal wsG As Worksheet, _
    ByVal firstRow As Long, ByVal lastRow As Long, _
    ByVal daysWidth As Long, _
    ByVal FIRST_DATE_COL_G As Long, _
    ByVal ROW_DDGAMA As Long, _
    ByVal ROW_DATES_G As Long, _
    ByVal COL_ID As Long, _
    ByVal COL_NAME As Long, _
    ByVal COL_PCT As Long, _
    ByRef cntArr() As Double, _
    ByRef shArr() As Variant, _
    ByRef hasPerson() As Boolean, _
    ByRef dayArr() As String, _
    ByRef dateArr() As Date, _
    ByVal ScoreThreshold As Double, _
    ByVal OfficeModelMode As String, _
    ByRef addedOffice As Long, _
    ByVal WeightSurplus As Double, _
    ByVal WeightMonthlyPct As Double)


    modSettings.LogStep "GlobalOfficeRotacija", "GlobalOfficeRotacija START"

    logCount = 0
    ReDim logArr(1 To 1)
    addedOffice = 0

    Dim officeNeed() As Long
    ReDim officeNeed(firstRow To lastRow)

    Dim totalNeed As Long: totalNeed = 0
    Dim cap As Long
    cap = 0

    modSettings.LogStep "GlobalOfficeRotacija", "ComputeOfficeNeed"

    modOffice_Logic.ComputeOfficeNeed _
        wsG, firstRow, lastRow, daysWidth, COL_PCT, _
        shArr, hasPerson, officeNeed, totalNeed, dateArr


    If totalNeed <= 0 Then
        modSettings.LogStep "GlobalOfficeRotacija", "EXIT (no need)"
        Exit Sub
    End If


    Dim a As Long
    For a = 1 To daysWidth
        If dayArr(a) <> "PR" Then
            If cntArr(a) >= ScoreThreshold Then
                cap = cap + CLng(Fix(cntArr(a) - ScoreThreshold))
            End If
        End If
    Next a

    Dim spreadMode As Boolean
    spreadMode = (cap >= totalNeed)

modSettings.LogStep "GlobalOfficeRotacija", _
    "cap=" & cap & " totalNeed=" & totalNeed & " spreadMode=" & IIf(spreadMode, "DA", "NE")

    modSettings.LogStep "GlobalOfficeRotacija", "Model=" & OfficeModelMode & _
                        " | Threshold=" & ScoreThreshold

    If OfficeModelMode = "GLOBAL" Then

        modSettings.LogStep "GlobalOfficeRotacija", "AssignOffice_FairPerRound"

        modOfficeModels.AssignOffice_FairPerRound _
            wsG, COL_NAME, COL_PCT, _
            firstRow, lastRow, daysWidth, _
            cntArr, shArr, hasPerson, officeNeed, dayArr, _
            ScoreThreshold, totalNeed, addedOffice, dateArr, spreadMode, _
            WeightSurplus, WeightMonthlyPct

    Else

        modSettings.LogStep "GlobalOfficeRotacija", "AssignOffice_GreedySequential"

        modOfficeModels.AssignOffice_GreedySequential _
            wsG, COL_NAME, _
            firstRow, lastRow, daysWidth, _
            cntArr, shArr, hasPerson, officeNeed, dayArr, _
            ScoreThreshold, totalNeed, addedOffice, dateArr, spreadMode

    End If

    modSettings.LogStep "GlobalOfficeRotacija", "END | addedOffice=" & addedOffice

End Sub


' =============================================================================
'  C) PERIODICNI SESTANKI -> avtomatski OFFICE ("O") v PREDOGLED
'     (poenoteno: TryWriteOfficeCell)
'
'  POPRAVKI:
'   - unitKey NI ve roni parameter: za vsak ID ga doloimo iz GAMA (COL_CYCLE/Team tag)
'   - LoadCntArrFromGama zdaj klie za VSE enote, ki nastopajo v idList
'   - sc(cand) je uteeno povpreje score-ov po enotah (CountTblXXXX)
'   - popravljena napaka: "cntArr = cntArr = ..."
' =============================================================================
Public Sub ApplyPeriodicMeetingsOffice()
    Dim errNum As Long, eDesc As String

    Dim wbThis As Workbook: Set wbThis = ThisWorkbook
    Dim wsSet As Worksheet: Set wsSet = wbThis.Worksheets("NASTAVITVE")
    Dim wsP As Worksheet:   Set wsP = wbThis.Worksheets("PREDOGLED")

    modSettings.LogStep "--OFFICE--", "ApplyPeriodicMeetingsOffice START"

    ' --- settings (key-based) ---
    Dim s As modSettings.TMainSettings
    Dim o As modSettings.TOfficeSettings

    On Error GoTo EH_SETTINGS
    s = modSettings.LoadMainSettings(wsSet)
    o = modSettings.LoadOfficeSettings(wsSet)
    On Error GoTo 0

    ' --- PREDOGLED layout iz OfficeSettings ---
    Dim PREV_FIRST_ROW As Long, PREV_ID_COL As Long, PREV_FIRST_DATE_COL As Long
    PREV_FIRST_ROW = o.PrevFirstDataRow
    PREV_ID_COL = o.PrevColID
    PREV_FIRST_DATE_COL = o.PrevFirstDateCol

    ' --- daysWidth iz MainSettings ---
    Dim daysWidth As Long
    daysWidth = s.daysWidth
    If daysWidth <= 0 Then
        MsgBox "DaysWidth mora biti > 0 (NASTAVITVE klju: DaysWidth).", vbCritical
        Exit Sub
    End If

    ' ---------------------------------------------------------------------
    ' Periodicni sestanki (key-based)  vrednosti so v stolpcu B, kljui v C
    ' ---------------------------------------------------------------------
    Dim MeetActive As Boolean
    Dim sName As String, sDow As String, sNth As String, sIds As String, sNote As String
    MeetActive = o.MeetActive
    
    sName = o.MeetName
    sDow = o.MeetDOW
    sNth = o.MeetNthMonth
    sIds = o.MeetIDsSpec
    sNote = o.MeetNote

    If Not MeetActive Then
        MsgBox "Periodini sestanki: AKTIVACIJA je NE.", vbInformation
        Exit Sub
    End If

    If Len(sDow) = 0 Or Len(sIds) = 0 Then
        MsgBox "PeriodiCni sestanki: manjka DAN ali ID-ji." & vbCrLf & _
               "Kljuca: 'Dan v tednu za sestanek' in/ali 'ID-ji zaposlenih'.", vbCritical
        Exit Sub
    End If

    Dim nthList() As Long
    nthList = ParseNthList(sNth)
    If (Not Not nthList) = 0 Then
        MsgBox "Periodicni sestanki: 'Kateri dan v mesecu' ni veljaven (npr. 1,3).", vbCritical
        Exit Sub
    End If

    Dim dowIdx As Long
    dowIdx = ParseSlovenianDOWIndex(sDow)
    If dowIdx = 0 Then
        MsgBox "Periodicni sestanki: DAN v tednu ni veljaven (PO/TO/SR/CE/PE/SO/NE).", vbCritical
        Exit Sub
    End If

    ' --- core nastavitve (GAMA) ---
    Dim pathG As String, sheetG As String
    Dim FIRST_DATE_COL_G As Long
    Dim OfficeModelMode As String

    pathG = s.PathGama
    sheetG = s.GamaSheetName

    ' KLJUNI FIX: ta stolpec je start-date v GAMI (preracunan)
    FIRST_DATE_COL_G = s.GamaStartDateCol

    OfficeModelMode = o.OfficeModelMode
    modSettings.LogStep "--OFFICE-- SETTINGS", "Loaded (daysWidth=" & daysWidth & ", model=" & OfficeModelMode & ")"

    ' --- app state backup ---
    Dim prevScreenUpdating As Boolean, prevEnableEvents As Boolean
    Dim prevCalc As XlCalculation
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    Dim wbG As Workbook, wsG As Worksheet

    On Error GoTo CLEANUP_ERR
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wbG = modSettings.OpenGamaWorkbook(pathG, False, True)
    If wbG Is Nothing Then GoTo CLEANUP_OK
    Set wsG = modSettings.GetWorksheetSafe(wbG, sheetG)
    If wsG Is Nothing Then
        MsgBox "List '" & sheetG & "' ne obstaja v GAMA.", vbCritical
        GoTo CLEANUP_OK
    End If

    ' --- PREDOGLED id->row map ---
    Dim idToPrevRow As Object: Set idToPrevRow = CreateObject("Scripting.Dictionary")
    idToPrevRow.CompareMode = vbTextCompare

    Dim lastPrevRow As Long
    lastPrevRow = wsP.Cells(wsP.Rows.Count, PREV_ID_COL).End(xlUp).Row
    If lastPrevRow < PREV_FIRST_ROW Then
        MsgBox "PREDOGLED je prazen. Najprej naredi PREDOGLED.", vbCritical
        GoTo CLEANUP_OK
    End If

    ' SNAPSHOT PREDEN ZAPISUJEM OFFICE
    Dim rngAll As Range
    Set rngAll = wsP.Range(wsP.Cells(PREV_FIRST_ROW, PREV_FIRST_DATE_COL), _
                           wsP.Cells(lastPrevRow, PREV_FIRST_DATE_COL + daysWidth - 1))
    modUndo.BeginSnapshot rngAll, "MEETINGS"

    Dim r As Long, id As String
    For r = PREV_FIRST_ROW To lastPrevRow
        id = NormID(wsP.Cells(r, PREV_ID_COL))
        If Len(id) > 0 Then
            If Not idToPrevRow.Exists(id) Then idToPrevRow.Add id, r
        End If
    Next r

    ' --- dateArr + dayArr iz PREDOGLEDA ---
    Dim pDateArr() As Date, pDayArr() As String
    ReDim pDateArr(1 To daysWidth)
    ReDim pDayArr(1 To daysWidth)

    Dim j As Long, v As Variant
    For j = 1 To daysWidth
        v = wsP.Cells(1, PREV_FIRST_DATE_COL + j - 1).Value
        If Not IsDate(v) Then
            MsgBox "PREDOGLED nima veljavnih datumov v vrstici 1.", vbCritical
            GoTo CLEANUP_OK
        End If
        pDateArr(j) = DateValue(v)
        pDayArr(j) = UCase$(Trim$(CStr(wsP.Cells(2, PREV_FIRST_DATE_COL + j - 1).Value)))
    Next j

    ' --- preberi PREDOGLED enkrat v RAM ---
    Dim prevArr As Variant
    prevArr = rngAll.Value

    ' --- ID list (CSV ali range spec) ---
    Dim idList() As String
    idList = ReadIdList(wsSet, sIds)
    If (Not Not idList) = 0 Then
        MsgBox "Periodini sestanki: 'ID-ji zaposlenih' ni dal nobenih ID-jev: " & sIds, vbCritical
        GoTo CLEANUP_OK
    End If

    ' -------------------------------------------------------------------------
    ' 1) ID -> enota (unitKey) iz GAMA
    ' -------------------------------------------------------------------------
    Dim idToUnit As Object
    Set idToUnit = BuildIdToUnitKeyFromGama(wsG, s, idList)

    ' -------------------------------------------------------------------------
    ' 2) Koliko udeleencev je iz katere enote
    ' -------------------------------------------------------------------------
    Dim attendeeCountByUnit As Object: Set attendeeCountByUnit = CreateObject("Scripting.Dictionary")
    attendeeCountByUnit.CompareMode = vbTextCompare

    Dim k As Long, uid As String, uk As String
    For k = LBound(idList) To UBound(idList)
        uid = modSettings.NormalizeID(idList(k))
        If Len(uid) = 0 Then GoTo NextK
        If Not idToUnit.Exists(uid) Then GoTo NextK
        uk = CStr(idToUnit(uid))
        If Len(uk) = 0 Then GoTo NextK

        If attendeeCountByUnit.Exists(uk) Then
            attendeeCountByUnit(uk) = CLng(attendeeCountByUnit(uk)) + 1
        Else
            attendeeCountByUnit.Add uk, 1
        End If
NextK:
    Next k

    If attendeeCountByUnit.Count = 0 Then
        MsgBox "Periodini sestanki: noben ID ni dobil enote (UnitFromCycleTag).", vbCritical
        GoTo CLEANUP_OK
    End If

    ' -------------------------------------------------------------------------
    ' 3) Nalozi cntArr za relevantne enote
    ' -------------------------------------------------------------------------
    Dim unitCounts As Object
    Set unitCounts = LoadCntArrByUnit(wsG, attendeeCountByUnit, FIRST_DATE_COL_G, daysWidth)

    ' --- unikaten seznam mesecev ---
    Dim monthDict As Object: Set monthDict = CreateObject("Scripting.Dictionary")
    monthDict.CompareMode = vbTextCompare

    For j = 1 To daysWidth
        Dim keyM As String
        keyM = CStr(Year(pDateArr(j))) & "-" & Format$(Month(pDateArr(j)), "00")
        If Not monthDict.Exists(keyM) Then monthDict.Add keyM, True
    Next j

    Dim totalWrites As Long: totalWrites = 0
    Dim keyMonth As Variant

    For Each keyMonth In monthDict.Keys

        Dim yy As Long, mm As Long
        yy = CLng(Split(CStr(keyMonth), "-")(0))
        mm = CLng(Split(CStr(keyMonth), "-")(1))

        Dim candJs() As Long, wc() As Long, sc() As Double
        Dim candCount As Long: candCount = 0

        Dim t As Long, candJ As Long
        For t = LBound(nthList) To UBound(nthList)

            candJ = FindJ_ForNthWeekdayInMonth(pDateArr, daysWidth, yy, mm, dowIdx, nthList(t))
            If candJ > 0 Then

                If pDayArr(candJ) = "PR" Then GoTo NEXT_T

                Dim dayScore As Double
                If Not TryGetCandidateDayScore(candJ, unitCounts, attendeeCountByUnit, dayScore) Then GoTo NEXT_T

                candCount = candCount + 1
                ReDim Preserve candJs(1 To candCount)
                ReDim Preserve wc(1 To candCount)
                ReDim Preserve sc(1 To candCount)

                candJs(candCount) = candJ
                wc(candCount) = CountConcreteShiftsOnDay_Arr(idList, idToPrevRow, prevArr, PREV_FIRST_ROW, candJ)
                sc(candCount) = dayScore
            End If

NEXT_T:
        Next t

        If candCount = 0 Then GoTo NEXT_MONTH

        Dim bestJ As Long
        bestJ = PickBestJ60(candJs, wc, sc, pDateArr)
        If bestJ = 0 Then GoTo NEXT_MONTH

        ' --- vpii OFFICE za vse ID-je na bestJ ---
        Dim prevRow As Long
        Dim cell As Range
        Dim wrote As Boolean

        For k = LBound(idList) To UBound(idList)

            id = modSettings.NormalizeID(idList(k))
            If Len(id) = 0 Then GoTo NEXT_ID

            If Not idToPrevRow.Exists(id) Then GoTo NEXT_ID
            prevRow = CLng(idToPrevRow(id))

            Set cell = wsP.Cells(prevRow, PREV_FIRST_DATE_COL + bestJ - 1)

            Dim cmt As String
            cmt = BuildMeetingComment(sName, sNote)

            If TryWriteOfficeCell(cell, cmt, wrote) Then
                If wrote Then totalWrites = totalWrites + 1
                ' posodobi RAM copy
                prevArr((prevRow - PREV_FIRST_ROW) + 1, bestJ) = cell.Value
            End If

NEXT_ID:
        Next k

NEXT_MONTH:
    Next keyMonth

    MsgBox "Periodini sestanki: konano." & vbCrLf & _
           "Pravilo: " & sName & " (" & sDow & " / " & sNth & ")" & vbCrLf & _
           "ID-vir: " & sIds & vbCrLf & _
           "Vpisanih OFFICE celic: " & totalWrites & vbCrLf & vbCrLf & _
           "UNDO je pripravljen (makro: UndoLastAction).", vbInformation

CLEANUP_OK:
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalc
    On Error GoTo 0
    Exit Sub

CLEANUP_ERR:
    errNum = Err.Number
    eDesc = Err.Description

    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalc
    On Error GoTo 0

    If errNum <> 0 Then
        MsgBox "Napaka: " & errNum & vbCrLf & eDesc, vbCritical
    End If
    Exit Sub

EH_SETTINGS:
    MsgBox "Napaka pri branju nastavitev (LoadMain/LoadOfficeSettings): " & Err.Description, vbCritical
End Sub

' =============================================================================
' BuildIdToUnitKeyFromGama
' -----------------------------------------------------------------------------
' Namen:
'   - Iz GAMA lista prebere "enoto" za vsak ID udeleenca.
'   - Enoto doloi iz oznake tima/cikla (COL_CYCLE) preko UnitFromCycleTag.
'
' Zakaj:
'   - Periodini sestanek ima ID-je (iz PREDOGLED ali range-a),
'     Count tabela pa je odvisna od enote (CountTblOKZP, CountTblFIS, ...).
'
' Vrne:
'   Dictionary: key = ID (string), value = unitKey (OKZP/BRN/MBX/POW/CEK/FDT/FIS)
'
' Opombe:
'   - e ID v GAMA nima prepoznanega taga (UnitFromCycleTag=""), se ne doda.
'   - Namenoma filtrira samo ID-je, ki so v idList (hitreje + manj smeti).
' =============================================================================
Private Function BuildIdToUnitKeyFromGama( _
    ByVal wsG As Worksheet, _
    ByRef s As modSettings.TMainSettings, _
    ByRef idList() As String) As Object

    Dim need As Object: Set need = CreateObject("Scripting.Dictionary")
    need.CompareMode = vbTextCompare

    Dim k As Long, id As String
    For k = LBound(idList) To UBound(idList)
        id = modSettings.NormalizeID(idList(k))
        If Len(id) > 0 Then
            If Not need.Exists(id) Then need.Add id, True
        End If
    Next k

    Dim out As Object: Set out = CreateObject("Scripting.Dictionary")
    out.CompareMode = vbTextCompare

    Dim r As Long, tag As String, uk As String
    Dim gid As String

    For r = s.firstDataRowG To s.lastDataRowG
        gid = modSettings.NormalizeID(wsG.Cells(r, s.colIdG).Text)
        If Len(gid) = 0 Then GoTo NextR
        If Not need.Exists(gid) Then GoTo NextR

        tag = Trim$(CStr(wsG.Cells(r, s.ColTeam).Value))   ' cycle/team tag
        uk = modOffice_Logic.UnitFromCycleTag(tag)
        If Len(uk) > 0 Then
            If Not out.Exists(gid) Then out.Add gid, uk
        End If

NextR:
    Next r

    Set BuildIdToUnitKeyFromGama = out
End Function

' =============================================================================
' LoadCntArrByUnit
' -----------------------------------------------------------------------------
' Namen:
'   - Enkrat naloi "cntArr" za vsako relevant enoto (CountTblXXXX).
'
' Vhod:
'   - unitsDict: dictionary, kjer so kljui unitKey (OKZP/FIS/...) ki jih rabi
'                (value je lahko karkoli, mi uporabljamo samo Keys).
'
' Vrne:
'   Dictionary: key = unitKey, value = Double() array (1..daysWidth)
'
' Zakaj:
'   - Kandidatni dan ocenjuje glede na udeleence iz ve enot,
'     zato mora imeti score za vsak day posebej za vsako enoto.
'
' Opomba:
'   - LoadCntArrFromGama mora imeti podpis:
'       LoadCntArrFromGama(wsG, unitKey, FIRST_DATE_COL_G, daysWidth)
' =============================================================================
Private Function LoadCntArrByUnit( _
    ByVal wsG As Worksheet, _
    ByVal unitsDict As Object, _
    ByVal FIRST_DATE_COL_G As Long, _
    ByVal daysWidth As Long) As Object

    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    modSettings.LogStep "LoadCntArrByUnit START | sheet=" & wsG.Name & _
                                   " | daysWidth=" & daysWidth & _
                                   " | FIRST_DATE_COL_G=" & FIRST_DATE_COL_G

    If unitsDict Is Nothing Then
        modSettings.LogStep "ERR", "LoadCntArrByUnit: unitsDict IS NOTHING"
        Set LoadCntArrByUnit = d
        Exit Function
    End If

    modSettings.LogStep "DBG", "LoadCntArrByUnit: units count=" & unitsDict.Count

    Dim uk As Variant
    For Each uk In unitsDict.Keys

        modSettings.LogStep "DBG", "LoadCntArrByUnit: processing unit='" & CStr(uk) & "'"
        On Error GoTo ERR_HANDLER

        Dim arr As Variant
        modSettings.LogStep "DBG", "LoadCntArrByUnit: calling LoadCntArrFromGama | unit=" & CStr(uk)
        arr = LoadCntArrFromGama(wsG, CStr(uk), FIRST_DATE_COL_G, daysWidth)
        modSettings.LogStep "DBG", "LoadCntArrByUnit: returned OK | unit=" & CStr(uk)
        d.Add CStr(uk), arr
        modSettings.LogStep "DBG", "LoadCntArrByUnit: stored unit='" & CStr(uk) & "'"
        On Error GoTo 0
        GoTo NEXT_UNIT

ERR_HANDLER:
        modSettings.LogStep "ERR", "LoadCntArrByUnit: ERROR unit=" & CStr(uk) & _
                                  " | Err=" & Err.Number & " | " & Err.Description
        Err.Clear
        On Error GoTo 0

NEXT_UNIT:
    Next uk

    modSettings.LogStep "DBG", "LoadCntArrByUnit END | loaded units=" & d.Count

    Set LoadCntArrByUnit = d
End Function




Private Function TryGetCandidateDayScore( _
    ByVal j As Long, _
    ByVal unitCounts As Object, _
    ByVal attendeeCountByUnit As Object, _
    ByRef outScore As Double) As Boolean

    Dim sumW As Double: sumW = 0
    Dim sumS As Double: sumS = 0

    Dim uk As Variant, w As Long
    For Each uk In attendeeCountByUnit.Keys
        If CDbl(unitCounts(uk)(j)) < 0 Then
            outScore = -1E+99
            TryGetCandidateDayScore = False
            Exit Function
        End If

        w = CLng(attendeeCountByUnit(uk))
        If w > 0 Then
            sumW = sumW + w
            sumS = sumS + CDbl(unitCounts(uk)(j)) * w
        End If
    Next uk

    If sumW <= 0 Then
        outScore = -1E+99
        TryGetCandidateDayScore = False
    Else
        outScore = sumS / sumW
        TryGetCandidateDayScore = True
    End If
End Function

' =============================================================================
'  Helper: hitro tetje konkretnih izmen na dan candJ iz RAM (prevArr)
'  prevArr je rngAll.Value = (1..nRows, 1..daysWidth) za PREDOGLED range
'  PREV_FIRST_ROW je zaetek rngAll v sheetu, da iz prevRow naredimo index v array
' =============================================================================
Private Function CountConcreteShiftsOnDay_Arr( _
    ByRef idList() As String, _
    ByVal idToPrevRow As Object, _
    ByRef prevArr As Variant, _
    ByVal PREV_FIRST_ROW As Long, _
    ByVal candJ As Long) As Long

    Dim k As Long, id As String
    Dim prevRow As Long, idx As Long
    Dim v As Variant
    Dim c As Long

    c = 0

    For k = LBound(idList) To UBound(idList)
        id = Trim$(CStr(idList(k)))
        If Len(id) = 0 Then GoTo NEXT_K
        If Not idToPrevRow.Exists(id) Then GoTo NEXT_K

        prevRow = CLng(idToPrevRow(id))
        idx = (prevRow - PREV_FIRST_ROW) + 1
        If idx < 1 Then GoTo NEXT_K

        v = prevArr(idx, candJ)

        ' teje "konkretne izmene" -> jaz predpostavim: vse kar ni prazno in ni "O"
        ' (e ima drugano definicijo, jo tu poravnaj!)
        If Len(Trim$(CStr(v))) > 0 Then
            If UCase$(Trim$(CStr(v))) <> "O" Then
                c = c + 1
            End If
        End If

NEXT_K:
    Next k

    CountConcreteShiftsOnDay_Arr = c
End Function


' =============================================================================
'  D) WRITEBACK helper (preprost in pregleden)
' =============================================================================
Private Sub WriteBackShArrToPreview( _
    ByVal wsP As Worksheet, _
    ByVal firstRow As Long, ByVal lastRow As Long, _
    ByVal daysWidth As Long, _
    ByVal PREV_FIRST_DATE_COL As Long, _
    ByRef shArr() As Variant, _
    ByRef hasPerson() As Boolean, _
    ByRef rowMap() As Long, _
    ByRef nChanged As Long, _
    ByRef firstChangedCell As Range)

    Dim r As Long, j As Long
    nChanged = 0
    Set firstChangedCell = Nothing

    For r = firstRow To lastRow
        If hasPerson(r) Then
            If rowMap(r) > 0 Then

                For j = 1 To daysWidth
                    Dim c As Range
                    Set c = wsP.Cells(rowMap(r), PREV_FIRST_DATE_COL + j - 1)

                    Dim oldV As String, newV As String
                    oldV = modOffice_Logic.CleanShiftText(c.Value)
                    newV = modOffice_Logic.CleanShiftText(shArr(r, j))

                    If oldV <> newV Then
                        nChanged = nChanged + 1
                        If firstChangedCell Is Nothing Then Set firstChangedCell = c
                        c.Value = newV
                    End If
                Next j

            End If
        End If
    Next r
End Sub


' =============================================================================
'  E) POLICY / RUNTIME HELPERS (settings-driven)
' =============================================================================

' ============================================================
' IsAnyShiftForQuota
' ============================================================
Public Function IsAnyShiftForQuota(ByVal s As String) As Boolean
    Dim k As String
    k = modOffice_Logic.ShiftKey(s)

    If gCountAllShiftsForWorkday Then
        IsAnyShiftForQuota = (Len(Trim$(k)) > 0)
    Else
        IsAnyShiftForQuota = (Len(k) > 0 And Left$(k, 1) = "X")
    End If
End Function

' ============================================================
' IsOverwritableByOffice
' ============================================================
Public Function IsOverwritableByOffice(ByVal s As String) As Boolean
    Dim k As String
    k = modOffice_Logic.ShiftKey(s)

    If gOverwriteAllowed Is Nothing Then
        IsOverwritableByOffice = (k = "X1" Or k = "X2")
    Else
        IsOverwritableByOffice = gOverwriteAllowed.Exists(k)
    End If
End Function

' ============================================================
' BuildAllowedShiftDict
' ============================================================
Public Function BuildAllowedShiftDict(ByVal csv As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim s As String
    s = UCase$(Trim$(csv))
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, " ", "")

    If Len(s) = 0 Then
        Set BuildAllowedShiftDict = d
        Exit Function
    End If

    Dim parts() As String
    parts = Split(s, ",")

    Dim i As Long, p As String
    For i = LBound(parts) To UBound(parts)
        p = UCase$(Trim$(parts(i)))
        If Len(p) > 0 Then
            If Not d.Exists(p) Then d.Add p, True
        End If
    Next i

    Set BuildAllowedShiftDict = d
End Function

' ============================================================
' NormID
' ============================================================
Public Function NormID(ByVal c As Range) As String
    Dim s As String
    s = c.Text
    s = Replace(s, ChrW(160), " ")
    NormID = Trim$(s)
End Function

' ============================================================
' ParseScore
' ============================================================
Public Function ParseScore(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then
        ParseScore = 0
        Exit Function
    End If
    If IsNumeric(v) Then
        ParseScore = CDbl(v)
        Exit Function
    End If

    Dim s As String
    s = Trim$(CStr(v))
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, " ", "")
    s = Replace(s, ",", ".")
    ParseScore = Val(s)
End Function

' ============================================================
' ParsePct
' ============================================================
Public Function ParsePct(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then
        ParsePct = 1
        Exit Function
    End If

    Dim s As String, x As Double
    If IsNumeric(v) Then
        x = CDbl(v)
    Else
        s = Trim$(CStr(v))
        s = Replace(s, ChrW(160), " ")
        s = Replace(s, " ", "")
        s = Replace(s, ",", ".")
        s = Replace(s, "%", "")
        x = Val(s)
    End If

    If x > 1 Then x = x / 100
    If x < 0 Then x = 0
    If x > 1 Then x = 1
    ParsePct = x
End Function

' =============================================================================
'  G) LoadCntArrFromGama
'  - prebere Count tabelo za dano enoto ("CountTbl" & unitKey)
'  - bere 4. vrstico znotraj tabele (Row + 3)
'  - bere po stolpcih tabele (ne po FIRST_DATE_COL_G), da se ne razbije pri zamikih
'  - TRENUTNO JE TA AKTIVNA!
' =============================================================================
Public Function LoadCntArrFromGama( _
    ByVal wsG As Worksheet, _
    ByVal unitKey As String, _
    ByVal FIRST_DATE_COL_G As Long, _
    ByVal daysWidth As Long) As Double()

    modSettings.LogStep "DBG", "LoadCntArrFromGama START | unit=" & unitKey & _
                                   " | daysWidth=" & daysWidth

    Dim out() As Double
    ReDim out(1 To daysWidth)

    unitKey = UCase$(Trim$(unitKey))
    If Len(unitKey) = 0 Then
        modSettings.LogStep "ERR", "LoadCntArrFromGama: prazen unitKey"
        Err.Raise vbObjectError + 1002, , "LoadCntArrFromGama: unitKey je prazen."
    End If

    Dim rngCount As Range
    Dim tblName As String
    tblName = "CountTbl" & unitKey

    modSettings.LogStep "DBG", "LoadCntArrFromGama: looking for range '" & tblName & "'"

    On Error Resume Next
    Set rngCount = wsG.Range(tblName)
    On Error GoTo 0

    If rngCount Is Nothing Then
        modSettings.LogStep "ERR", "LoadCntArrFromGama: range NOT FOUND -> " & tblName
        Err.Raise vbObjectError + 1001, , _
            "Named range '" & tblName & "' ne obstaja v GAMA."
    End If

    modSettings.LogStep "DBG", "LoadCntArrFromGama: range=" & rngCount.Address(0, 0) & _
                                   " | firstRow=" & rngCount.Row & _
                                   " | firstCol=" & rngCount.Column & _
                                   " | rows=" & rngCount.Rows.Count & _
                                   " | cols=" & rngCount.Columns.Count
    If FIRST_DATE_COL_G <= 0 Then
        Err.Raise vbObjectError + 1003, , _
        "LoadCntArrFromGama: FIRST_DATE_COL_G mora biti > 0."
    End If

    
    ' 4. vrstica v tabeli
    Dim countRow As Long
    countRow = rngCount.Row + 3
    
    ' Start column mora slediti START_DATE (Prvi stolpec datuma v GAMA - preraèunan)
    ' da se števec bere od istega dne, kot je ustvarjen DefaultRoster.

    Dim startCol As Long
    startCol = FIRST_DATE_COL_G

    modSettings.LogStep "DBG", "LoadCntArrFromGama: countRow=" & countRow & _
                                   " | startCol=" & startCol

    Dim j As Long, readCol As Long
    Dim cellVal As Variant
    Dim parsedVal As Double

    For j = 1 To daysWidth
        readCol = startCol + (j - 1)
        If j = 1 And readCol <> startCol Then
            Err.Raise vbObjectError + 1004, , _
                "LoadCntArrFromGama self-check failed: j=1 mora mapirati na startDate stolpec."
        End If

        cellVal = wsG.Cells(countRow, readCol).Value

        parsedVal = modOffice.ParseScore(cellVal)
        out(j) = parsedVal

        ' logiraj prvih nekaj dni (da ne ubiješ performance)
        If j <= 10 Then
            modSettings.LogStep "DBG", "LoadCntArrFromGama: j=" & j & _
                                           " | cell=" & wsG.Cells(countRow, readCol).Address(0, 0) & _
                                           " | raw='" & wsG.Cells(countRow, readCol).Text & "'" & _
                                           " | parsed=" & parsedVal
        End If

    Next j

    modSettings.LogStep "DBG", "LoadCntArrFromGama END | unit=" & unitKey & _
                                   " | firstVal=" & out(1)

    LoadCntArrFromGama = out
End Function


' =============================================================================
'  H) Debug / Log
' =============================================================================
Public Sub DebugBlock(ByVal title As String, ByVal msg As String)
    If DEBUG_PRINT Then
        Debug.Print "================ " & title & " ================"
        Debug.Print msg
    End If
    If gDEBUG_UI Then
        MsgBox "=== " & title & " ===" & vbCrLf & msg, vbInformation
    End If
End Sub

Public Sub AppendOfficeLog(ByVal msg As String)
    logCount = logCount + 1
    If logCount = 1 Then
        ReDim logArr(1 To 1)
    Else
        ReDim Preserve logArr(1 To logCount)
    End If
    logArr(logCount) = Format$(Now, "hh:nn:ss") & " | " & msg
End Sub


' =============================================================================
'  I) GetPreviewTargetRange
' =============================================================================
Public Function GetPreviewTargetRange( _
    ByVal wsP As Worksheet, _
    ByRef hasPerson() As Boolean, _
    ByRef rowMap() As Long, _
    ByVal daysWidth As Long, _
    ByVal PREV_FIRST_DATE_COL As Long) As Range

    Dim minRow As Long, maxRow As Long
    Dim r As Long

    minRow = 9999999
    maxRow = 0

    For r = LBound(hasPerson) To UBound(hasPerson)
        If hasPerson(r) Then
            If rowMap(r) > 0 Then
                If rowMap(r) < minRow Then minRow = rowMap(r)
                If rowMap(r) > maxRow Then maxRow = rowMap(r)
            End If
        End If
    Next r

    If maxRow = 0 Then
        Set GetPreviewTargetRange = Nothing
        Exit Function
    End If

    Set GetPreviewTargetRange = wsP.Range( _
        wsP.Cells(minRow, PREV_FIRST_DATE_COL), _
        wsP.Cells(maxRow, PREV_FIRST_DATE_COL + daysWidth - 1) _
    )
End Function

' =============================================================================
'  J) TryWriteOfficeCell (skupni helper za periodic meetings)
' =============================================================================
Private Function TryWriteOfficeCell( _
    ByVal cell As Range, _
    ByVal commentText As String, _
    ByRef wrote As Boolean) As Boolean

    wrote = False

    Dim oldKey As String
    oldKey = modOffice_Logic.ShiftKey(cell.Value) ' UCase + trim (samo za check)

    ' 1) ce je e O -> samo komentar (append)
    If oldKey = "O" Then
        If Len(commentText) > 0 Then AppendCellComment cell, commentText
        TryWriteOfficeCell = True
        Exit Function
    End If

    ' 2) ce NI dovoljena izmena za prepis -> samo komentar
    If Not IsOverwritableByOffice(oldKey) Then
        If Len(commentText) > 0 Then AppendCellComment cell, commentText
        TryWriteOfficeCell = True
        Exit Function
    End If

    ' 3) zapisi O + komentar
    cell.Value = "O"
    wrote = True
    If Len(commentText) > 0 Then AppendCellComment cell, commentText

    TryWriteOfficeCell = True
End Function


' =============================================================================
'  K) Komentarji (threaded + legacy)  append
' =============================================================================
Private Sub AppendCellComment(ByVal cell As Range, ByVal addTxt As String)
    addTxt = Trim$(addTxt)
    If Len(addTxt) = 0 Then Exit Sub

    Dim existing As String
    existing = ""

    ' 1) threaded
    On Error Resume Next
    If Not cell.CommentThreaded Is Nothing Then
        existing = CStr(cell.CommentThreaded.Text)
    End If
    On Error GoTo 0

    ' 2) legacy
    If Len(existing) = 0 Then
        On Error Resume Next
        If Not cell.Comment Is Nothing Then
            existing = CStr(cell.Comment.Text)
        End If
        On Error GoTo 0
    End If

    existing = Trim$(existing)

    ' brez duplikatov
    If Len(existing) > 0 Then
        If InStr(1, existing, addTxt, vbTextCompare) > 0 Then Exit Sub
    End If

    Dim finalTxt As String
    If Len(existing) = 0 Then
        finalTxt = addTxt
    Else
        finalTxt = existing & vbCrLf & addTxt
    End If

    ' pobrii in zapii na novo
    On Error Resume Next
    If Not cell.Comment Is Nothing Then cell.Comment.Delete
    On Error GoTo 0

    On Error Resume Next
    If Not cell.CommentThreaded Is Nothing Then cell.CommentThreaded.Delete
    On Error GoTo 0

    On Error Resume Next
    cell.AddCommentThreaded finalTxt
    If Err.Number <> 0 Then
        Err.Clear
        cell.AddComment finalTxt
    End If
    On Error GoTo 0
End Sub


' =============================================================================
'  L) Helperji za periodic meetings (ostanejo tukaj, ker so UI-specifini)
' =============================================================================

Private Function IsYes(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        IsYes = CBool(v)
        Exit Function
    End If

    Dim s As String
    s = UCase$(Trim$(CStr(v)))
    s = Replace(s, ChrW(160), " ")
    IsYes = (s = "DA" Or s = "YES" Or s = "TRUE" Or s = "1" Or s = "Y")
End Function

Private Function ParseNthList(ByVal v As Variant) As Long()
    Dim s As String: s = Trim$(CStr(v))
    Dim out() As Long
    Dim n As Long: n = 0

    If Len(s) = 0 Then
        ParseNthList = out
        Exit Function
    End If

    s = Replace(s, ";", ",")
    s = Replace(s, " ", ",")
    Do While InStr(s, ",,") > 0
        s = Replace(s, ",,", ",")
    Loop
    If Left$(s, 1) = "," Then s = Mid$(s, 2)
    If Right$(s, 1) = "," Then s = Left$(s, Len(s) - 1)
    If Len(s) = 0 Then
        ParseNthList = out
        Exit Function
    End If

    Dim parts() As String
    parts = Split(s, ",")

    Dim i As Long, x As Long
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For i = LBound(parts) To UBound(parts)
        If IsNumeric(parts(i)) Then
            x = CLng(parts(i))
            If x >= 1 And x <= 5 Then
                If Not dict.Exists(CStr(x)) Then dict.Add CStr(x), x
            End If
        End If
    Next i

    If dict.Count = 0 Then
        ParseNthList = out
        Exit Function
    End If

    ReDim out(1 To dict.Count)
    i = 0
    Dim k As Variant
    For Each k In dict.Keys
        i = i + 1
        out(i) = CLng(dict(k))
    Next k

    ' sort asc
    Dim a As Long, b As Long, tmp As Long
    For a = 1 To UBound(out) - 1
        For b = a + 1 To UBound(out)
            If out(b) < out(a) Then
                tmp = out(a): out(a) = out(b): out(b) = tmp
            End If
        Next b
    Next a

    ParseNthList = out
End Function

Private Function FindJ_ForNthWeekdayInMonth( _
    ByRef dateArr() As Date, _
    ByVal daysWidth As Long, _
    ByVal yy As Long, ByVal mm As Long, _
    ByVal dowIdx As Long, ByVal nth As Long) As Long

    Dim d As Date
    d = NthWeekdayOfMonth(yy, mm, dowIdx, nth)
    If d = 0 Then
        FindJ_ForNthWeekdayInMonth = 0
        Exit Function
    End If

    Dim j As Long
    For j = 1 To daysWidth
        If dateArr(j) = d Then
            FindJ_ForNthWeekdayInMonth = j
            Exit Function
        End If
    Next j

    FindJ_ForNthWeekdayInMonth = 0
End Function

Private Function NthWeekdayOfMonth(ByVal yy As Long, ByVal mm As Long, ByVal dowIdx As Long, ByVal nth As Long) As Date
    Dim firstDay As Date
    firstDay = DateSerial(yy, mm, 1)

    Dim firstDOW As Long
    firstDOW = Weekday(firstDay, vbMonday)

    Dim offset As Long
    offset = (dowIdx - firstDOW + 7) Mod 7

    Dim d As Date
    d = DateAdd("d", offset + (nth - 1) * 7, firstDay)

    If Month(d) <> mm Then
        NthWeekdayOfMonth = 0
    Else
        NthWeekdayOfMonth = d
    End If
End Function

Private Function ParseSlovenianDOWIndex(ByVal s As String) As Long
    Dim t As String
    t = UCase$(Trim$(s))
    t = Replace(t, "", "C")
    t = Replace(t, "", "S")
    t = Replace(t, "", "Z")

    Select Case t
        Case "PO", "PON", "PONEDELJEK": ParseSlovenianDOWIndex = 1
        Case "TO", "TOR", "TOREK":      ParseSlovenianDOWIndex = 2
        Case "SR", "SRE", "SREDA":      ParseSlovenianDOWIndex = 3
        Case "CE", "CET", "CETRTEK":    ParseSlovenianDOWIndex = 4
        Case "PE", "PET", "PETEK":      ParseSlovenianDOWIndex = 5
        Case "SO", "SOB", "SOBOTA":     ParseSlovenianDOWIndex = 6
        Case "NE", "NED", "NEDELJA":    ParseSlovenianDOWIndex = 7
        Case Else:                      ParseSlovenianDOWIndex = 0
    End Select
End Function

Private Function ReadIdList(ByVal wsSet As Worksheet, ByVal spec As String) As String()
    Dim s As String: s = Trim$(spec)
    Dim out() As String
    Dim n As Long: n = 0

    If InStr(1, s, ":", vbTextCompare) > 0 Then
        Dim rng As Range
        On Error Resume Next
        Set rng = wsSet.Range(s)
        On Error GoTo 0
        If rng Is Nothing Then
            ReadIdList = out
            Exit Function
        End If

        Dim c As Range, id As String
        For Each c In rng.Cells
            id = modSettings.NormalizeID(c.Text)
            If Len(id) > 0 Then
                n = n + 1
                ReDim Preserve out(1 To n)
                out(n) = id
            End If
        Next c
    Else
        s = Replace(s, ";", ",")
        s = Replace(s, " ", ",")
        Do While InStr(s, ",,") > 0
            s = Replace(s, ",,", ",")
        Loop
        If Left$(s, 1) = "," Then s = Mid$(s, 2)
        If Right$(s, 1) = "," Then s = Left$(s, Len(s) - 1)

        If Len(s) = 0 Then
            ReadIdList = out
            Exit Function
        End If

        Dim parts() As String
        parts = Split(s, ",")

        Dim i As Long, p As String
        For i = LBound(parts) To UBound(parts)
            p = modSettings.NormalizeID(parts(i))
            If Len(p) > 0 Then
                n = n + 1
                ReDim Preserve out(1 To n)
                out(n) = p
            End If
        Next i
    End If

    ReadIdList = out
End Function

Private Function BuildMeetingComment(ByVal meetingName As String, ByVal noteText As String) As String
    Dim t As String
    t = ""
    If Len(Trim$(meetingName)) > 0 Then t = t & Trim$(meetingName)
    If Len(noteText) > 0 Then
        If Len(t) > 0 Then t = t & vbCrLf
        t = t & noteText
    End If
    BuildMeetingComment = t
End Function

Private Function CountConcreteShiftsOnDay( _
    ByRef idList() As String, _
    ByVal idToPrevRow As Object, _
    ByVal wsP As Worksheet, _
    ByVal PREV_FIRST_DATE_COL As Long, _
    ByVal j As Long) As Long

    Dim k As Long, id As String, pr As Long
    Dim sKey As String
    Dim cnt As Long: cnt = 0

    For k = LBound(idList) To UBound(idList)
        id = modSettings.NormalizeID(idList(k))
        If Len(id) = 0 Then GoTo NEXT_ID
        If Not idToPrevRow.Exists(id) Then GoTo NEXT_ID

        pr = CLng(idToPrevRow(id))
        sKey = modOffice_Logic.ShiftKey(wsP.Cells(pr, PREV_FIRST_DATE_COL + j - 1).Value)

        If (sKey = "X1" Or sKey = "X2" Or sKey = "O") Then
            cnt = cnt + 1
        End If
NEXT_ID:
    Next k

    CountConcreteShiftsOnDay = cnt
End Function

Private Function PickBestJ60( _
    ByRef candJs() As Long, _
    ByRef workCount() As Long, _
    ByRef cntScore() As Double, _
    ByRef dateArr() As Date) As Long

    Dim i As Long
    If (UBound(candJs) < LBound(candJs)) Then
        PickBestJ60 = 0
        Exit Function
    End If

    Dim wcMin As Double, wcMax As Double
    Dim scMin As Double, scMax As Double

    wcMin = 1E+99: wcMax = -1E+99
    scMin = 1E+99: scMax = -1E+99

    For i = LBound(candJs) To UBound(candJs)
        If workCount(i) < wcMin Then wcMin = workCount(i)
        If workCount(i) > wcMax Then wcMax = workCount(i)
        If cntScore(i) < scMin Then scMin = cntScore(i)
        If cntScore(i) > scMax Then scMax = cntScore(i)
    Next i

    Dim wcDen As Double, scDen As Double
    wcDen = wcMax - wcMin
    scDen = scMax - scMin

    Dim bestIdx As Long: bestIdx = 0
    Dim bestU As Double: bestU = -1E+99
    Dim wcNorm As Double, scNorm As Double, u As Double

    For i = LBound(candJs) To UBound(candJs)

        If wcDen = 0 Then wcNorm = 1 Else wcNorm = (workCount(i) - wcMin) / wcDen
        If scDen = 0 Then scNorm = 1 Else scNorm = (cntScore(i) - scMin) / scDen

        u = 0.6 * wcNorm + 0.4 * scNorm

        If bestIdx = 0 Then
            bestIdx = i: bestU = u
        ElseIf u > bestU Then
            bestIdx = i: bestU = u
        ElseIf u = bestU Then
            If dateArr(candJs(i)) < dateArr(candJs(bestIdx)) Then bestIdx = i: bestU = u
        End If
    Next i

    If bestIdx = 0 Then PickBestJ60 = 0 Else PickBestJ60 = candJs(bestIdx)
End Function


' =============================================================================
'  M) RONI UNDO (brez vpraanj)
' =============================================================================
Public Sub UndoLastAction()
    On Error GoTo EH

    If Not modUndo.HasSnapshot Then
        MsgBox "Undo ni na voljo (snapshot ne obstaja).", vbExclamation
        Exit Sub
    End If

    modSettings.LogMsg "INFO", "UndoLastAction -> " & modUndo.SnapshotInfo
    modUndo.Undo
    Exit Sub
EH:
    MsgBox "Undo ni na voljo (ni snapshot-a) ali je prilo do napake." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation
End Sub

' =============================================================================
'  (KONEC modOffice)...
' =============================================================================


