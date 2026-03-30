Attribute VB_Name = "CopyFrmGAMA"
Option Explicit
Private gCoreSet As Object
' ============================================================
'  CopyWeirdShifts_FromGamaToPreview
'  - "VLOOKUP princip":
'     * v PREDOGLEDU prebere ID
'     * najde isto osebo v GAMI
'     * za dni v daysWidth prenese samo "èudne izmene"
'       (vse, kar NI X1/X2/X3; privzeto ignorira prazno)
' ============================================================
Public Sub CopyFrmGAMA()

    Dim wbThis As Workbook: Set wbThis = ThisWorkbook
    Dim wsSet As Worksheet: Set wsSet = wbThis.Worksheets("NASTAVITVE")
    Dim wsP As Worksheet:   Set wsP = wbThis.Worksheets("PREDOGLED")

    ' --- app state backup ---
    Dim prevScreenUpdating As Boolean, prevEnableEvents As Boolean
    Dim prevCalc As XlCalculation
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    On Error GoTo FAIL
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    modSettings.LogStep "CopyWeirdShifts", "START"

    ' ============================================================
    ' 1) PREBERI NASTAVITVE (core + PREDOGLED + Unit)
    ' ============================================================
    Dim s As modSettings.TMainSettings
    s = modSettings.LoadMainSettings(wsSet)
    Dim o As modSettings.TOfficeSettings
    o = modSettings.LoadOfficeSettings(wsSet)
    Dim unitCfg As Object
    Set unitCfg = modSettings.LoadUnitConfigs(wsSet)
    
    modSettings.LogMsg "INFO", "GAMA=" & s.PathGama & " | Sheet=" & s.GamaSheetName
    modSettings.LogMsg "INFO", "daysWidth=" & s.daysWidth & " | PrevFirstRow=" & o.PrevFirstDataRow

    Dim pathG As String, sheetG As String
    Dim firstRow As Long, lastRow As Long
    Dim COL_ID As Long, COL_NAME As Long
    Dim START_DATE_COL_G As Long
    Dim ROW_DATES_G As Long
    Dim ROW_DDGAMA As Long
    Dim daysWidth As Long
    Dim COL_LIC As Long
    Dim PREV_FIRST_ROW As Long
    Dim PREV_FIRST_DATE_COL As Long
    Dim PREV_ID_COL As Long

    
    pathG = s.PathGama
    sheetG = s.GamaSheetName
    firstRow = s.firstDataRowG
    lastRow = s.lastDataRowG
    
    COL_ID = s.colIdG
    COL_NAME = s.colNameG
    ROW_DATES_G = s.firstDateRowG
    ROW_DDGAMA = s.RowDDGama
    START_DATE_COL_G = s.GamaStartDateCol
    daysWidth = s.daysWidth
    COL_LIC = s.colLicenseG ' èe ga rabiš, sicer stran
    
    ' PREDOGLED
    PREV_FIRST_ROW = o.PrevFirstDataRow
    PREV_ID_COL = o.PrevColID
    PREV_FIRST_DATE_COL = o.PrevFirstDateCol
    
    
    ' Slovar vseh CoreShiftsov
    modSettings.LogStep "CopyWeirdShifts", "Build CoreShift set"
    Set gCoreSet = CreateObject("Scripting.Dictionary")
    gCoreSet.CompareMode = vbTextCompare
    
    Dim k As Variant, csv As String
    For Each k In unitCfg.Keys
        csv = Trim$(unitCfg(k)("CoreShiftsCsv") & "")
        AddCsvToSet gCoreSet, csv
    Next k
    
    ' fallback, èe je tabela prazna
    If gCoreSet.Count = 0 Then AddCsvToSet gCoreSet, "X1,X2,X3,B1,B2,B3,M1,M2,P1,P2,C1,C2"
    modSettings.LogMsg "INFO", "CoreShift codes count=" & gCoreSet.Count
        

    ' ============================================================
    ' 2) ODPRE GAMA
    ' ============================================================
    modSettings.LogStep "CopyWeirdShifts", "Open GAMA"
    Dim wbG As Workbook, wsG As Worksheet
    Set wbG = modSettings.OpenGamaWorkbook(pathG, False, True)
    If wbG Is Nothing Then GoTo CLEANUP

    Set wsG = modSettings.GetWorksheetSafe(wbG, sheetG)
    modSettings.LogMsg "INFO", "GAMA workbook opened: " & wbG.Name
    If wsG Is Nothing Then
        MsgBox "List '" & sheetG & "' ne obstaja v GAMA.", vbCritical
        GoTo CLEANUP
    End If

    ' ============================================================
    ' 3) NAREDI MAPO ID -> GAMA ROW (VLOOKUP lookup tabela)
    ' ============================================================
    modSettings.LogStep "CopyWeirdShifts", "Build ID->Row map"
    Dim idToRow As Object: Set idToRow = CreateObject("Scripting.Dictionary")
    idToRow.CompareMode = vbTextCompare

    Dim r As Long, idKey As String
    For r = firstRow To lastRow
        idKey = NormID_Text(wsG.Cells(r, COL_ID))
        If Len(idKey) > 0 Then idToRow(idKey) = r
    Next r
    modSettings.LogMsg "INFO", "Mapped IDs=" & idToRow.Count
    ' ============================================================
    ' 4) PREVERI DATUME (PREDOGLED vs GAMA start) – varnost
    ' ============================================================
    Dim startDate As Date
    Dim gCell As Variant
    
    startDate = DateValue(s.startDate)
    
    modSettings.LogStep "CopyWeirdShifts", "Sanity check dates (GamaStartDateCol)"
    
    gCell = wsG.Cells(ROW_DATES_G, s.GamaStartDateCol).Value
    
    modSettings.LogMsg "DBG", "startDate=" & Format$(startDate, "dd.mm.yyyy") & _
                               " | GAMA(" & ROW_DATES_G & "," & s.GamaStartDateCol & ")=" & _
                               wsG.Cells(ROW_DATES_G, s.GamaStartDateCol).Text
    
    If Not IsDate(gCell) Then
        MsgBox "GAMA start datum ni veljaven v stolpcu GamaStartDateCol.", vbCritical
        GoTo CLEANUP
    End If
    
    If DateValue(gCell) <> startDate Then
        MsgBox "Nastavitve/GAMA niso poravnane:" & vbCrLf & _
               "StartDate: " & Format$(startDate, "dd.mm.yyyy") & vbCrLf & _
               "GAMA(" & ROW_DATES_G & "," & s.GamaStartDateCol & "): " & wsG.Cells(ROW_DATES_G, s.GamaStartDateCol).Text, vbCritical
        GoTo CLEANUP
    End If

    ' ============================================================
    '  5) ITERIRAJ PREDOGLED: ID -> poišèi GAMA row -> kopiraj èudne + prazne èe v ciklusu
    '     Pravila:
    '       - prepise samo "cudne" izmene (kar NI core shift)
    '       - ce je v GAMI prazno: pobrise samo na prost dan cikla
    '       - core shift v GAMI ignoriraj
    ' ============================================================
    Dim prevLastRow As Long
    prevLastRow = wsP.Cells(wsP.Rows.Count, PREV_ID_COL).End(xlUp).Row
    If prevLastRow < PREV_FIRST_ROW Then
        MsgBox "PREDOGLED je prazen (ni IDjev).", vbExclamation
        GoTo CLEANUP
    End If
    Dim pr As Long, gRow As Long, j As Long
    Dim srcShift As String, dstCell As Range
    Dim changed As Long, mapped As Long, missed As Long
    changed = 0: mapped = 0: missed = 0
    
    For pr = PREV_FIRST_ROW To prevLastRow
    
        idKey = NormID_Text(wsP.Cells(pr, PREV_ID_COL))
        If Len(idKey) = 0 Then GoTo NextPr
    
        If Not idToRow.Exists(idKey) Then
            missed = missed + 1
            GoTo NextPr
        End If
    
        gRow = CLng(idToRow(idKey))
        mapped = mapped + 1
        Dim dstShift As String
        
        For j = 1 To daysWidth
    
            Set dstCell = wsP.Cells(pr, PREV_FIRST_DATE_COL + j - 1)
    
            ' raw iz GAMI (od startDate naprej)
            srcShift = GetShiftRaw(wsG.Cells(gRow, s.GamaStartDateCol + j - 1).Value)
            
            ' trenutna vrednost v PREDOGLEDU
            dstShift = GetShiftRaw(dstCell.Value)
            
            ' Mirror sync: èe sta razlièna -> prepiši (tudi prazno)
            If Len(srcShift) = 0 Then
                ' V GAMI je ta dan prost/sprošèen:
                ' praznino vedno zrcalimo v PREDOGLED (tudi èe je trenutno core shift),
                ' da ne ostanejo "visi" X1/X2/X3 po aktivnosti na prost dan.
                If Len(dstShift) > 0 Then
                    dstCell.ClearContents
                    changed = changed + 1
                End If
            ElseIf IsCoreShift(srcShift) Then
                ' NOVO: pobere tudi spremembe znotraj core izmen (npr. X3 -> X1)
                If dstShift <> srcShift Then
                    dstCell.Value = srcShift
                    changed = changed + 1
                End If
            ElseIf Not IsCoreShift(srcShift) Then
                If dstShift <> srcShift Then
                    dstCell.Value = srcShift
                    changed = changed + 1
                End If
            End If
               
        Next j
               
NextPr:
    Next pr



    MsgBox "Prenos èudnih izmen konèan." & vbCrLf & _
           "Najdenih (mapiranih) oseb: " & mapped & vbCrLf & _
           "IDjev brez zadetka v GAMA: " & missed & vbCrLf & _
           "Spremenjenih celic: " & changed, vbInformation

CLEANUP:
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalc
    Exit Sub

FAIL:
    MsgBox "Napaka: " & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume CLEANUP
End Sub

' ============================================================
'  Pomožne funkcije (lokalne, da je makro samostojen)
' ============================================================

Private Function ShouldCopyFreeDayFromGama(ByVal previewShift As String) As Boolean
    Dim k As String
    k = Trim$(previewShift)

    If Len(k) = 0 Then
        ShouldCopyFreeDayFromGama = True
        Exit Function
    End If

    ShouldCopyFreeDayFromGama = Not IsCoreShift(k)
End Function

Private Function ReadPositiveLong_Local(ByVal wsSet As Worksheet, ByVal rowN As Long, ByRef outVal As Long) As Boolean
    Dim v As Variant
    v = wsSet.Range("B" & rowN).Value
    If IsNumeric(v) And CLng(v) > 0 Then
        outVal = CLng(v)
        ReadPositiveLong_Local = True
    Else
        outVal = 0
        ReadPositiveLong_Local = False
    End If
End Function

Private Function NormID_Text(ByVal c As Range) As String
    Dim s As String
    s = c.Text
    s = Replace(s, ChrW(160), " ")
    NormID_Text = Trim$(s)
End Function

Private Function NormShift_Text(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, ChrW(160), " ")
    s = Trim$(s)
    NormShift_Text = UCase$(s)
End Function

Private Function GetShiftRaw(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)

    ' odstranimo “nevidne” smeti
    s = Replace(s, ChrW(160), " ")   ' NBSP
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")

    GetShiftRaw = Trim$(s)
End Function

' Je navadna izmena v GAMI ?
Private Function IsCoreShift(ByVal s As String) As Boolean
    If gCoreSet Is Nothing Then
        ' varnost: èe set ni zgrajen (ne bi se smelo zgodit)
        IsCoreShift = False
        Exit Function
    End If
    
    s = UCase$(Trim$(s))
    If Len(s) = 0 Then
        IsCoreShift = False
    Else
        IsCoreShift = gCoreSet.Exists(s)
    End If
End Function

Private Sub AddCsvToSet(ByVal dict As Object, ByVal csv As String)
    Dim arr, i As Long, t As String
    If Len(Trim$(csv)) = 0 Then Exit Sub
    
    arr = Split(csv, ",")
    For i = LBound(arr) To UBound(arr)
        t = UCase$(Trim$(CStr(arr(i))))
        If Len(t) > 0 Then dict(t) = True
    Next i
End Sub


' the end






