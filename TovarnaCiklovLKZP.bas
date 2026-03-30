Attribute VB_Name = "TovarnaCiklovLKZP"
Option Explicit

' === Opcijske nastavitve  ===================
'Private Const KEY_PLAN_UNITS As String = "Planirane enote (OKZP, FIS, FDT, BRN, MBX, POW, CEK)"
'Private Const KEY_EXCLUDED_TYPES As String = "Izloeni glede na OJT filter"
'Private Const KEY_DAYSWIDTH As String = "DaysWidth"
'Private Const KEY_START_DATE As String = "ZAETNI DATUM"


' ============================================================
'  A) BuildPreviewFromCycles  (CreateDefaultRooster)
' ============================================================
' Namen:
'   - Iz NASTAVITEV prebere, kje je GAMA in kako je strukturirana
'   - Odpre GAMA in prebere seznam zaposlenih + njihove time (cikle)
'   - Na list PREDOGLED zapie cikle za izbrano obdobje (daysWidth dni)
'   - Ne vpisuje OFFICE, samo osnovne cikle

Public Sub CreateDefaultRooster()

    ' --- Ta Excel (kjer je makro) ---
    Dim wbThis As Workbook: Set wbThis = ThisWorkbook

    ' --- delovni listi v tem Excelu ---
    Dim wsC As Worksheet   ' list CIKLI (predloge ciklov)
    Dim wsP As Worksheet   ' list PREDOGLED (kam zapiemo rezultat)
    Dim wsSet As Worksheet ' list NASTAVITVE (od koder beremo parametre)

    ' --- pomoni slovarji:
    ' selectedUnits = katere enote sploh planiramo (ALL ali npr. OKZP, FIS)
    ' unitCfg = posebna pravila po enotah (Overwrite, Allow3NonFL, NoNightFL)
    Dim selectedUnits As Object
    Dim unitCfg As Object

    ' =========================================================
    '   1) SPREMENLJIVKE, KI JIH PREBEREMO IZ NASTAVITEV
    ' =========================================================
    Dim pathG As String, sheetG As String          ' pot do GAMA datoteke + ime lista znotraj GAMA
    Dim s As TMainSettings                         ' deklaracija spremenljivke s za MainSettings
    Dim firstRow As Long, lastRow As Long          ' od katere do katere vrstice so zaposleni v GAMA
    Dim COL_ID As Long, COL_NAME As Long           ' v katerem stolpcu je ID in ime
    Dim COL_TYPE As Long, COL_CYCLE As Long        ' OJT tip + tim/cikel
    Dim COL_PCT As Long                            ' odstotek operative (e ga kdaj rabi)

    Dim GamaStartDateCol As Long                   ' stolpec v GAMA, kjer se zane izbrani start date
    Dim ROW_DATES_G As Long                        ' vrstica v GAMA, kjer so datumi (header)
    Dim GamaFirstDateCol As Long                   ' prvi stolpec, kjer se datumi zanejo (header start)
    Dim ROW_DDGAMA As Long                         ' vrstica s dnevi/oznakami (PO, TO / prazniki)
    Dim startDate As Date

    Dim daysWidth As Long                          ' koliko dni naredimo plan (irina plana)
    Dim COL_LIC As Long                            ' stolpec licence (npr FL)

    Dim PrevFirstDateCol As Long                   ' prvi stolpec datumov v PREDOGLED
    Dim PrevFirstDataRow As Long                   ' prva vrstica zaposlenih v PREDOGLED
    Dim KEY_PLAN_UNITS As String                   ' Seznam enot, ki jih planiramo

    Dim excludedTypes As Object                    ' Zgradimo slovar izkljuenih oseb
    Dim exclCsv As String                          ' Izkljuena enota, pridobljena iz celice

    ' =========================================================
    '   2) GAMA workbook (odpre se kot loen Excel)
    ' =========================================================
    Dim wbG As Workbook, wsG As Worksheet

    ' =========================================================
    '   3) VARNOST - da Excel dela hitreje (brez refresh, brez eventov)
    ' =========================================================
    Dim calcMode As XlCalculation
    Dim prevEvents As Boolean, prevScr As Boolean, prevUpd As Boolean

    LogStep "START", "CreateDefaultRooster"

    On Error GoTo CLEANUP

    ' Excel pospeimo in prepreimo pop-up dialoge med delovanjem
    prevScr = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevUpd = Application.AskToUpdateLinks
    calcMode = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual

    ' NASTAVITVE list (da v Excelu vidi, da makro dela)
    LogStep "INIT", "Nastavitev wsSet"
    Set wsSet = wbThis.Worksheets("NASTAVITVE")

    LogStep "LOAD SETTINGS", "LoadMainSettings"
    s = modSettings.LoadMainSettings(wsSet)


    LogStep "READ OPTIONAL SETTINGS", "PlanUnits / UnitCfg / Excluded"



    ' =========================================================
    '   1) POVEEMO DELOVNE LISTE
    ' =========================================================
    LogStep "BIND SHEETS", "CIKLI + PREDOGLED"
    Set wsC = wbThis.Worksheets("CIKLI")
    Set wsP = wbThis.Worksheets("PREDOGLED")

    ' =========================================================
    '   2) OBVEZNE NASTAVITVE
    '      Tukaj se dejansko bere iz NASTAVITEV (ni samo log!)
    ' =========================================================

    LogStep "SETTINGS MAP", "Prenos main settings -> lokalne spremenljivke"
    pathG = s.PathGama
    sheetG = s.GamaSheetName
    firstRow = s.firstDataRowG
    lastRow = s.lastDataRowG

    COL_ID = s.colIdG
    COL_NAME = s.colNameG
    COL_TYPE = s.ColOJT
    COL_CYCLE = s.ColTeam
    COL_PCT = s.ColOperativePct

    ROW_DATES_G = s.firstDateRowG
    GamaFirstDateCol = s.firstDateColG
    ROW_DDGAMA = s.RowDDGama
    daysWidth = s.daysWidth

    COL_LIC = s.colLicenseG

    PrevFirstDateCol = s.PrevFirstDateCol
    PrevFirstDataRow = s.PrevFirstDataRow

    startDate = s.startDate
    GamaStartDateCol = s.GamaStartDateCol
    KEY_PLAN_UNITS = s.selectedUnitsText
    exclCsv = Trim$(s.excludedTypesCsv)


    ' =========================================================
    '   3) IZKLJUENI TIPI
    '      Zgradimo seznam izkljuenih tipov
    ' =========================================================

    LogStep "BUILD EXCLUDED TYPES", "CsvToIdSet"
    If Len(exclCsv) = 0 Then
        Set excludedTypes = CreateObject("Scripting.Dictionary")
        excludedTypes.CompareMode = vbTextCompare
    Else
        Set excludedTypes = modSettings.CsvToIdSet(exclCsv)
    End If


    ' e je daysWidth napaen, nima smisla nadaljevati
    If daysWidth <= 0 Then
        LogMsg "ERROR", "DaysWidth mora biti > 0."
        MsgBox "DaysWidth mora biti > 0.", vbCritical
        GoTo CLEANUP
    End If

    ' =========================================================
    '      (kaj planiramo + posebna pravila)
    ' =========================================================
    ' PREBEREMO ENOTE (iz s.selectedUnitsText)
    Dim PlanUnitsText As String
    PlanUnitsText = Trim$(s.selectedUnitsText)
    If Len(PlanUnitsText) = 0 Then PlanUnitsText = "ALL"

    Debug.Print "PlanUnitsText(from settings object)='" & PlanUnitsText & "'"

    LogStep "BUILD UNIT DICT", PlanUnitsText
    Set selectedUnits = modSettings.BuildUnitDict(PlanUnitsText)
    LogStep "LOAD UNIT CONFIGS", "NASTAVITVE PO ENOTAH"
    Set unitCfg = modSettings.LoadUnitConfigs(wsSet)

    ' =========================================================
    '   4) ODPREMO GAMA DATOTEKO
    ' =========================================================
    LogStep "OPEN GAMA", pathG

    Set wbG = modSettings.OpenGamaWorkbook(pathG, True, True)
    If wbG Is Nothing Then GoTo CLEANUP

    ' Povezemo pravi list v GAMA
    LogStep "BIND GAMA SHEET", sheetG
    Set wsG = modSettings.GetWorksheetSafe(wbG, sheetG)
    If wsG Is Nothing Then
        MsgBox "List '" & sheetG & "' ne obstaja v GAMA.", vbCritical
        GoTo CLEANUP
    End If

    ' =========================================================
    '   5) IZRAUN, KJE V GAMA SE ZANE START DATE
    '      (GamaStartDateCol je stolpec, ki kae na prvi dan plana)
    ' =========================================================
    startDate = s.startDate
    GamaStartDateCol = s.GamaStartDateCol

    LogStep "ALIGN STARTDATE", "from settings"
    LogMsg "DBG", "startDate=" & Format$(startDate, "dd.mm.yyyy") & " | GamaStartDateCol=" & GamaStartDateCol & "| ROW_DATES_G =" & ROW_DATES_G
    LogMsg "DBG", "GAMA(" & ROW_DATES_G & "," & GamaStartDateCol & _
              ") Value='" & CStr(wsG.Cells(ROW_DATES_G, GamaStartDateCol).Value) & _
              "' | Text='" & wsG.Cells(ROW_DATES_G, GamaStartDateCol).Text & _
              "' | IsDate=" & CStr(IsDate(wsG.Cells(ROW_DATES_G, GamaStartDateCol).Value))


    ' Preverimo, da se ta stolpec res ujema s startDate
    If Not IsDate(wsG.Cells(ROW_DATES_G, GamaStartDateCol).Value) Or DateValue(wsG.Cells(ROW_DATES_G, GamaStartDateCol).Value) <> DateValue(startDate) Then
        MsgBox "Nastavitve/GAMA niso poravnane:" & vbCrLf & _
               "StartDate: " & Format$(startDate, "dd.mm.yyyy") & vbCrLf & _
               "GAMA(" & ROW_DATES_G & "," & GamaStartDateCol & "): " & wsG.Cells(ROW_DATES_G, GamaStartDateCol).Text, vbCritical
        GoTo CLEANUP
    End If

    ' ========================================
    ' 6) Build cikMaps
    ' ========================================
    LogStep "BUILD CIKMAPS", "Read headers + teams from CIKLI col A"

    Dim cikMaps As Object: Set cikMaps = CreateObject("Scripting.Dictionary")
    Dim lastR As Long: lastR = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row

    Dim unit As String
    Dim d As Object
    Set d = Nothing
    unit = ""

    Dim rr As Long, nameA As String
    For rr = 1 To lastR
        nameA = Trim$(CStr(wsC.Cells(rr, 1).Value & ""))
        If nameA = "" Then GoTo NextRR

        If IsUnitHeader(nameA) Then
            unit = UCase$(nameA)
            If Not cikMaps.Exists(unit) Then
                Set d = CreateObject("Scripting.Dictionary")
                d.CompareMode = vbTextCompare
                cikMaps.Add unit, d
            Else
                Set d = cikMaps(unit)
            End If
            GoTo NextRR
        End If

        If unit <> "" Then
            If Not d Is Nothing Then
                If Not d.Exists(nameA) Then d.Add nameA, rr
            End If
        End If
NextRR:
    Next rr

    Dim cikliStartCol As Long
    cikliStartCol = PrevFirstDateCol ' poravnano z datumskim stolpcem (npr. E)

    '========================================
    ' 7) Output matrika
    ' ========================================
    LogStep "PREP OUTPUT", "Allocate outputArr"
    Dim outputArr() As Variant
    Dim outRow As Long
    ReDim outputArr(1 To (lastRow - firstRow + 1), 1 To (3 + daysWidth))
    outRow = 0

    ' ========================================
    ' 8) MOTOR loop
    ' ========================================
    LogStep "ENGINE", "Loop people from GAMA"

    Dim r As Long, j As Long
    Dim tip As String
    Dim vShift As String, lic As String
    Dim cikName As String

    Dim prevRow As Long
    Dim prevCol As Long
    Dim existingVal As String

    For r = firstRow To lastRow

        If Trim$(wsG.Cells(r, COL_NAME).Value & "") = "" Then GoTo NEXT_PERSON

        outRow = outRow + 1

        outputArr(outRow, 1) = wsG.Cells(r, COL_ID).Value
        outputArr(outRow, 2) = wsG.Cells(r, COL_NAME).Value

        tip = UCase$(Trim$(wsG.Cells(r, COL_TYPE).Value & ""))

        prevRow = PrevFirstDataRow + outRow - 1

        If IsExcludedType(tip, excludedTypes) Then
            outputArr(outRow, 3) = ""
            For j = 1 To daysWidth
                outputArr(outRow, 3 + j) = ""
            Next j
            GoTo NEXT_PERSON
        End If

        cikName = Trim$(wsG.Cells(r, COL_CYCLE).Value & "")
        outputArr(outRow, 3) = cikName

        lic = UCase$(Trim$(wsG.Cells(r, COL_LIC).Value & ""))

        Dim unitKey As String
        unitKey = UnitFromTeam(cikName)

        If unitKey <> "" Then
            If Not modSettings.IsUnitSelected(selectedUnits, unitKey) Then
                Dim tmpArr As Variant
                tmpArr = wsP.Cells(prevRow, PrevFirstDateCol).Resize(1, daysWidth).Value
                For j = 1 To daysWidth
                    outputArr(outRow, 3 + j) = tmpArr(1, j)
                Next j
                GoTo NEXT_PERSON
            End If
        End If

        Dim allow3NonFL_U As Boolean
        Dim overwrite_U As Boolean
        Dim noNightFL_U As Object
        Dim keepExistingU As Boolean

        allow3NonFL_U = False
        overwrite_U = True
        Set noNightFL_U = Nothing

        If unitKey <> "" Then
            If unitCfg.Exists(unitKey) Then
                Dim cfgU As Object
                Set cfgU = unitCfg(unitKey)
                allow3NonFL_U = CBool(cfgU("Allow3NonFL"))
                overwrite_U = CBool(cfgU("Overwrite"))
                Set noNightFL_U = modSettings.CsvToIdSet(cfgU("NoNightCsv"))
            End If
        End If

        keepExistingU = (overwrite_U = False)

        Dim prevDaysArr As Variant
        Dim hasPrevDays As Boolean
        hasPrevDays = False

        If keepExistingU Then
            prevDaysArr = wsP.Cells(prevRow, PrevFirstDateCol).Resize(1, daysWidth).Value
            hasPrevDays = True
        End If

        Dim rowCik As Long
        rowCik = 0

        If unitKey <> "" Then
            If cikMaps.Exists(unitKey) Then
                Dim map As Object
                Set map = cikMaps(unitKey)
                If map.Exists(cikName) Then rowCik = CLng(map(cikName))
            End If
        End If

        If rowCik = 0 Then
            LogMsg "WARN", "Team '" & cikName & "' ni najden v CIKLI (unit=" & unitKey & ", GAMA row=" & r & ")."
            For j = 1 To daysWidth
                If keepExistingU And hasPrevDays Then
                    outputArr(outRow, 3 + j) = Trim$(CStr(prevDaysArr(1, j) & ""))
                Else
                    outputArr(outRow, 3 + j) = ""
                End If
            Next j
            GoTo NEXT_PERSON
        End If
        Dim idKey As String
        idKey = modSettings.NormalizeID(wsG.Cells(r, COL_ID).Text)


        Dim cikPatternWidth As Long
        cikPatternWidth = GetCyclePatternWidth(wsC, rowCik, cikliStartCol)
        If cikPatternWidth < 1 Then cikPatternWidth = 1

        For j = 1 To daysWidth

            prevCol = PrevFirstDateCol + j - 1

            If keepExistingU Then
                If hasPrevDays Then
                    existingVal = Trim$(CStr(prevDaysArr(1, j) & ""))
                Else
                    existingVal = vbNullString
                End If

                ' Overwrite=NE: ohrani vse, kar je ze vpisano v PREDOGLED.
                If Len(existingVal) > 0 Then
                    outputArr(outRow, 3 + j) = existingVal
                    GoTo NEXT_DAY_J
                End If
            End If

            Dim cikCol As Long
            cikCol = cikliStartCol + ((j - 1) Mod cikPatternWidth)
            vShift = CStr(wsC.Cells(rowCik, cikCol).Value)

            If Right$(vShift, 1) = "3" Then
                If Not CanWorkNightShift(idKey, lic, allow3NonFL_U, noNightFL_U) Then
                    vShift = Shift3to2(vShift)
                End If
            End If

            outputArr(outRow, 3 + j) = vShift

NEXT_DAY_J:
        Next j

NEXT_PERSON:
    Next r

    ' ========================================
    ' 9) Datumi + dnevi
    ' ----------------------------
    LogStep "WRITE DATES", "Row 1 + Row 2"
    wsP.Cells(1, PrevFirstDateCol).Resize(1, daysWidth).Value = _
        wsG.Cells(ROW_DATES_G, GamaStartDateCol).Resize(1, daysWidth).Value

    wsP.Cells(2, PrevFirstDateCol).Resize(1, daysWidth).Value = _
        wsG.Cells(ROW_DDGAMA, GamaStartDateCol).Resize(1, daysWidth).Value


    ' ========================================
    ' 10) ienje / izpis
    ' ========================================
    LogStep "CLEAR / WRITE OUTPUT", "Clear extra rows + dump outputArr"

    Dim lastColToClear As Long
    lastColToClear = PrevFirstDateCol + daysWidth - 1

    Dim lastUsedRow As Long
    lastUsedRow = wsP.Cells(wsP.Rows.Count, PrevFirstDateCol - 3).End(xlUp).Row

    If lastUsedRow > (PrevFirstDataRow + outRow - 1) Then
        wsP.Range(wsP.Cells(PrevFirstDataRow + outRow, PrevFirstDateCol - 3), _
                  wsP.Cells(lastUsedRow, lastColToClear)).ClearContents
    End If

    wsP.Cells(PrevFirstDataRow, PrevFirstDateCol - 3).Resize(outRow, 3 + daysWidth).Value = outputArr

    LogStep "DONE", "PREDOGLED generiran"
    MsgBox "PREDOGLED generiran (cikli, brez OFFICE)." & vbCrLf & _
           "Overwrite/Allow3 je po ENOTAH (tabela 'NASTAVITVE PO ENOTAH').", vbInformation

CLEANUP:
    ' povrni Excel state
    On Error Resume Next
    Application.Calculation = calcMode
    Application.AskToUpdateLinks = prevUpd
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScr
    Application.DisplayAlerts = True
    On Error GoTo 0

    If Not wsSet Is Nothing Then
        On Error Resume Next
        On Error GoTo 0
    End If

    If Err.Number <> 0 Then
        LogMsg "ERROR", "Stage=" & GetStage() & " | Err=" & Err.Number & " | " & Err.Description
        SetStatus "ERROR @ " & GetStage() & " | " & Err.Number & " | " & Err.Description
        MsgBox "Napaka @ " & GetStage() & vbCrLf & _
               "Err " & Err.Number & vbCrLf & Err.Description, vbCritical
    Else
        ' e se kona brez napake, pusti zadnji step v

        SetStatus "OK @ " & GetStage()
    End If
End Sub

'============================================================================
'======================    HELPERJI ========================================

' ============================================================
'  Izloeni tipi (zaenkrat hard-coded)
' ============================================================
Private Function IsExcludedType(ByVal tip As String, ByVal excludedDict As Object) As Boolean
    tip = UCase$(Trim$(CStr(tip & "")))
    tip = Replace(tip, ChrW(160), " ")

    If Len(tip) = 0 Then
        IsExcludedType = False
        Exit Function
    End If

    If excludedDict Is Nothing Then
        IsExcludedType = False
        Exit Function
    End If

End Function

Private Function GetCyclePatternWidth(ByVal wsC As Worksheet, ByVal rowCik As Long, ByVal startCol As Long) As Long
    Dim lastCol As Long
    lastCol = wsC.Cells(rowCik, wsC.Columns.Count).End(xlToLeft).Column

    If lastCol < startCol Then
        GetCyclePatternWidth = 1
    Else
        GetCyclePatternWidth = (lastCol - startCol + 1)
    End If
End Function


' ============================================================
' ============================================================
' CanWorkNightShift (centralna nona pravila)
' ============================================================
Private Function CanWorkNightShift(ByVal employeeId As String, _
                                   ByVal licenseCode As String, _
                                   ByVal allow3NonFL As Boolean, _
                                   ByVal noNightSet As Object) As Boolean

    employeeId = modSettings.NormalizeID(employeeId)
    licenseCode = UCase$(Trim$(licenseCode))

    If Not noNightSet Is Nothing Then
        If noNightSet.Exists(employeeId) Then
            CanWorkNightShift = False
            Exit Function
        End If
    End If

    If Not allow3NonFL Then
        If licenseCode <> "FL" Then
            CanWorkNightShift = False
            Exit Function
        End If
    End If

    CanWorkNightShift = True
End Function

Private Function Shift3to2(ByVal vShift As String) As String
    vShift = Trim$(CStr(vShift))
    If Len(vShift) >= 2 And Right$(vShift, 1) = "3" Then
        Shift3to2 = Left$(vShift, Len(vShift) - 1) & "2"
    Else
        Shift3to2 = vShift
    End If
End Function


Private Function UnitFromTeam(ByVal team As String) As String
    UnitFromTeam = modSettings.MapUnitFromTeamTag(team)
End Function


Private Function IsUnitHeader(ByVal s As String) As Boolean
    Select Case UCase$(Trim$(s))
        Case "OKZP", "BRN", "MBX", "POW", "CEK", "FDT", "FIS"
            IsUnitHeader = True
        Case Else
            IsUnitHeader = False
    End Select
End Function



