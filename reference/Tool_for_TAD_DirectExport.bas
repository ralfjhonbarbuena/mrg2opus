' ============================================================
' MODULE: Tool_for_TAD (Final — Direct Export)
'
' ALL PREVIOUS FIXES RETAINED (1-6):
'   1. GetData()               - Pipeline error handler (On Error GoTo)
'   2. GetData()               - Comma-formatted rate strings converted to numbers
'   3. GroupedCargoNumbering() - rowMap(i)= avoids key-exists crash
'   4. GenerateRouteNote()     - destRow declared as Long
'   5. GenerateWording()       - fullText/finalText declared at Sub level
'   6. ClearAllSheetsDataPreserveHeaders() - Null-check for empty sheets
'
' CHANGES THIS VERSION:
'   A. GetData()
'      - Removed the three Split* calls. The source workbook never grows
'        any RATES_*, ROUTE_*, or SURCHARGES_* sheets.
'      - Still calls ExportAllToNewWorkbook at the end, which now owns
'        all the splitting work internally.
'
'   B. ExportAllToNewWorkbook()
'      - Creates the new workbook first, then calls the three private
'        Split* helpers, passing the new workbook as their target.
'      - No scope sheets are ever written to the source file.
'
'   C. SplitRatesToScopeSheets(targetWB)   \
'      SplitRouteToScopeSheets(targetWB)    > Now Private, accept a
'      SplitSurchargesToScopeSheets(targetWB)/  Workbook parameter.
'      - Read from the source workbook's master sheets as before.
'      - Write every scope sheet directly into targetWB.
'      - Source workbook is never modified beyond its own master sheets.
'
'   D. ClearAllSheetsDataPreserveHeaders()
'      - RATES_* / ROUTE_* / SURCHARGES_* pre-pass removed entirely.
'        Those sheets never exist in the source workbook any more.
'
'   E. ColLetter(), ExtractScopePrefix(), SheetExists(),
'      CopySheetToWorkbook(), BubbleSortStrings() — all retained as-is.
'
' NEW CASES THIS VERSION:
'   F. MapFAKValuesToAA()
'      - Select Case replaced with If/ElseIf so both full strings
'        ("FAK Alexandria (EGALY20 DEKHEILA PORT)") and bare codes
'        ("EGALY20") are caught by a single InStr check per variant.
'      - "EGALY20" / anything containing "EGALY20" → col AA = "EGALY20"
'      - "EGALY21" / anything containing "EGALY21" → col AA = "EGALY21"
'      - These values flow through ConcatRouteNote into col X (ROUTE NOTE)
'        and then into ROUTE col E exactly like the TRIST cases do.
'
'   G. ExtractNode()
'      - Added ElseIf branches for EGALY20 and EGALY21 after the existing
'        HAYDARPASA/MARPORT block.
'      - No dash-parsing needed: the code is the full node value, so
'        result is set directly to "EGALY20" or "EGALY21".
'      - Populates ROUTE col AE (Node) for these ports the same way
'        TRIST21 / TRIST02 are populated for the Turkish terminals.
'
' SETTINGS CHANGES THIS VERSION:
'   H. EnsureSettingsSheet()  [NEW, Public]
'      - Creates a "SETTINGS" sheet in the source workbook if it does
'        not already exist, pre-populated with defaults and Yes/No
'        dropdown validation. Run this once before first use; GetData
'        will read from it on every subsequent run.
'
'   I. ReadSettings()  [NEW, Private]
'      - Reads the three settings from the SETTINGS sheet into typed
'        ByRef out-parameters. Gracefully falls back to safe defaults
'        (both toggles Off, D7 addon = 700) if the sheet is missing.
'
'   J. AppendDryDangerousToData()  [NEW, Private]
'      - Called from GetData when "Include Dry Dangerous" = Yes.
'      - Appends a copy of every "Dry General" row in DATA to the
'        bottom of DATA, changing col O to "Dry Dangerous".
'        Rows are appended AFTER ConcatRouteNote runs, so col X
'        (ROUTE NOTE) is already populated and gets copied too.
'      - Appended rows then flow through the entire pipeline naturally:
'        bulk copy → RATES, cargo lookup, GroupedCargoNumbering,
'        GenerateRouteNote, GenerateSurcharges, scope split, export.
'      - Cleared by ClearAllSheetsDataPreserveHeaders like any other row.
'
'   K. ApplyD7Rates()  [NEW, Private]
'      - Called from GetData when "Include D7 (OFT 45)" = Yes.
'      - For every RATES data row where col Y (OFT HC) is not empty,
'        sets col AA (OFT 45) = col Y + the configurable add-on value.
'        Existing OFT 45 values are overwritten by the D7 calculation.
'      - Runs BEFORE the USD label loop so col Z ("USD" for OFT 45)
'        is stamped automatically on all newly-populated col AA cells.
' ============================================================


' ============================================================
' MAIN PIPELINE
' ============================================================

Sub GetData()
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim srcCols As Variant, destCols As Variant
    Dim i As Long, rowNum As Long, lastRow As Long
    Dim cargoType As String, matchRow As Variant
    Dim usdCheckCols As Variant, usdSetCols As Variant
    Dim maxRow As Long

    ' Settings read at start of every run
    Dim includeDryDangerous As Boolean
    Dim includeD7 As Boolean
    Dim d7Addon As Double

    ' FIX 1: Pipeline-level error handler
    On Error GoTo PipelineError

    ' CHANGE H/I: Read user settings (safe defaults if SETTINGS sheet missing)
    Call ReadSettings(includeDryDangerous, includeD7, d7Addon)

    Set ws1 = Worksheets("DATA")
    Set ws2 = Worksheets("RATES")
    Set ws3 = Worksheets("CARGO")

    srcCols  = Array("A", "B", "C", "E", "G", "H", "I", "K", "M", "S", "T", "U", "V", "O", "W", "D", "X", "J")
    destCols = Array("AD", "AE", "A", "H", "J", "K", "L", "N", "P", "U", "W", "Y", "AA", "AL", "AB", "D", "AC", "M")

    Call AddTSPortPrefix
    Call AddSLanePrefix
    Call MapFAKValuesToAA
    Call ConcatRouteNote

    ' CHANGE J: Append Dry Dangerous rows to DATA now — after ConcatRouteNote
    ' so col X (ROUTE NOTE) is already populated and gets copied with each row.
    ' Appended rows flow through every subsequent step as normal source data.
    If includeDryDangerous Then Call AppendDryDangerousToData

    ' FIX 2: Convert comma-formatted rate strings to actual numbers
    Dim rateCols As Variant
    rateCols = Array("S", "T", "U", "V")
    Dim rateLastRow As Long
    Dim cellVal As String
    Dim cleanVal As Double
    Dim rc As Long, rr As Long

    For rc = 0 To UBound(rateCols)
        rateLastRow = ws1.Cells(ws1.Rows.Count, rateCols(rc)).End(xlUp).Row
        For rr = 2 To rateLastRow
            With ws1.Cells(rr, rateCols(rc))
                If .Value <> "" Then
                    cellVal = CStr(.Value)
                    If InStr(cellVal, ",") > 0 Then
                        On Error Resume Next
                        cleanVal = CDbl(Replace(cellVal, ",", ""))
                        If Err.Number = 0 Then .Value = cleanVal
                        Err.Clear
                        On Error GoTo PipelineError
                    End If
                End If
            End With
        Next rr
    Next rc

    ' Copy data from DATA to RATES
    For i = 0 To UBound(srcCols)
        lastRow = ws1.Cells(ws1.Rows.Count, srcCols(i)).End(xlUp).Row
        If lastRow >= 2 Then
            ws2.Range(destCols(i) & "4").Resize(lastRow - 1, 1).Value = _
                ws1.Range(srcCols(i) & "2").Resize(lastRow - 1, 1).Value
        End If
    Next i

    ' Cargo lookup
    lastRow = ws1.Cells(ws1.Rows.Count, "O").End(xlUp).Row
    For rowNum = 2 To lastRow
        cargoType = ws1.Cells(rowNum, "O").Value
        If cargoType = "Dry General" Or cargoType = "Reefer" Or _
           cargoType = "Reefer Dry"  Or cargoType = "Dry Dangerous" Then
            matchRow = Application.Match(cargoType, ws3.Columns("A"), 0)
            If Not IsError(matchRow) Then
                ws2.Cells(rowNum + 2, "R").Value = ws3.Cells(matchRow, "B").Value
                ws2.Cells(rowNum + 2, "S").Value = ws3.Cells(matchRow, "C").Value
            Else
                ws2.Cells(rowNum + 2, "R").Value = "Not Found"
                ws2.Cells(rowNum + 2, "S").Value = "Not Found"
            End If
        End If
    Next rowNum

    ' Add "USD" labels
    ' CHANGE K: Apply D7 (OFT 45) values BEFORE the USD loop so that
    ' newly-written col AA cells are automatically stamped with "USD" in col Z.
    If includeD7 Then Call ApplyD7Rates(ws2, d7Addon)

    usdCheckCols = Array("U", "W", "Y", "AA")
    usdSetCols   = Array("T", "V", "X", "Z")
    maxRow = 0
    For i = 0 To UBound(usdCheckCols)
        lastRow = ws2.Cells(ws2.Rows.Count, usdCheckCols(i)).End(xlUp).Row
        If lastRow > maxRow Then maxRow = lastRow
    Next i
    For rowNum = 4 To maxRow
        For i = 0 To UBound(usdCheckCols)
            If ws2.Cells(rowNum, usdCheckCols(i)).Value <> "" Then
                ws2.Cells(rowNum, usdSetCols(i)).Value = "USD"
            End If
        Next i
    Next rowNum

    Call GroupedCargoNumbering
    Call ApplyCountIf
    Call GenerateRouteNote
    Call GenerateSurcharges

    ' CHANGE A: Export directly — no Split* calls here.
    '           ExportAllToNewWorkbook reads the master sheets and writes
    '           scope sheets straight into the new workbook.
    Call ExportAllToNewWorkbook

    On Error GoTo 0
    Exit Sub

PipelineError:
    Dim errMsg As String
    errMsg = "Pipeline failed." & vbCrLf & _
             "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
             "Source: " & Err.Source & vbCrLf & vbCrLf & _
             "The workbook may be in a partially written state. " & _
             "Run ClearAllSheetsDataPreserveHeaders before retrying."
    MsgBox errMsg, vbCritical, "GetData Error"
    On Error GoTo 0
End Sub


' ============================================================
' EXPORT  (creates the output workbook and drives all splitting)
' ============================================================

' -------------------------------------------------------------------
' CHANGE B: ExportAllToNewWorkbook now owns the entire split process.
'
' Sequence:
'   1. Create the new workbook.
'   2. Call each private Split* helper, passing the new workbook.
'      The helpers read from ThisWorkbook's master sheets and write
'      scope sheets directly into the new workbook — the source file
'      is never touched beyond what GetData already wrote.
'   3. Remove blank Sheet* stubs from the new workbook.
'   4. Activate the new workbook and prompt the user to save.
' -------------------------------------------------------------------
Sub ExportAllToNewWorkbook()
    Dim sourceWB  As Workbook
    Dim newWB     As Workbook
    Dim s         As Worksheet
    Dim scopeArr  As Variant   ' filled by the Split* helpers and returned here

    Set sourceWB = ThisWorkbook

    ' Create the destination workbook before any splitting starts
    Application.ScreenUpdating = False
    Set newWB = Workbooks.Add

    ' CHANGE B: pass newWB to each split helper; scopeArr receives the
    ' sorted list of scope names that were actually written, so the
    ' final message can report them accurately.
    Dim rateScopes()      As String
    Dim routeScopes()     As String
    Dim surchargeScopes() As String

    Call SplitRatesToScopeSheets(newWB, rateScopes)
    Call SplitRouteToScopeSheets(newWB, routeScopes)
    Call SplitSurchargesToScopeSheets(newWB, surchargeScopes)

    ' Remove default blank Sheet* stubs added by Workbooks.Add
    Application.DisplayAlerts = False
    For Each s In newWB.Sheets
        If s.Name Like "Sheet*" Then s.Delete
    Next s
    Application.DisplayAlerts = True

    Application.ScreenUpdating = True

    ' Build a deduplicated, sorted union of all scope names for the summary
    Dim allScopes As Object
    Set allScopes = CreateObject("Scripting.Dictionary")
    Dim v As Variant
    For Each v In rateScopes      : allScopes(CStr(v)) = True : Next v
    For Each v In routeScopes     : allScopes(CStr(v)) = True : Next v
    For Each v In surchargeScopes : allScopes(CStr(v)) = True : Next v

    Dim scopeList() As String
    ReDim scopeList(allScopes.Count - 1)
    Dim idx As Long : idx = 0
    For Each v In allScopes.Keys
        scopeList(idx) = CStr(v) : idx = idx + 1
    Next v
    Call BubbleSortStrings(scopeList)

    newWB.Activate
    MsgBox "Export complete." & vbCrLf & _
           newWB.Sheets.Count & " sheets written" & _
           " across " & allScopes.Count & " service scope" & _
           IIf(allScopes.Count = 1, "", "s") & _
           ": " & Join(scopeList, ", ") & "." & vbCrLf & vbCrLf & _
           "Please save the workbook now (Ctrl+S).", _
           vbInformation, "GetData — Done"
End Sub


' ============================================================
' PRIVATE SPLIT HELPERS
' Each sub reads from ThisWorkbook's corresponding master sheet,
' builds per-scope sheets directly inside targetWB, and returns
' a sorted array of the scope names it wrote via the out-param.
' ============================================================

' -------------------------------------------------------------------
' CHANGE C  —  Reads master RATES (3-row header, data from row 4,
' scope in col A). Writes RATES_<scope> sheets into targetWB,
' sorted by CMDT Seq (col B) ascending.
' -------------------------------------------------------------------
Private Sub SplitRatesToScopeSheets(ByRef targetWB As Workbook, _
                                     ByRef scopesOut() As String)
    Dim wsMaster  As Worksheet
    Dim wsScope   As Worksheet
    Dim lastRow   As Long, lastCol As Long
    Dim i As Long, c As Long, dr As Long
    Dim scopeVal  As String, sheetName As String
    Dim scopeKey  As Variant
    Dim scopeDict As Object
    Dim allData   As Variant
    Dim colCount  As Long, scopeCount As Long
    Dim scopeData() As Variant

    Set wsMaster  = ThisWorkbook.Worksheets("RATES")
    Set scopeDict = CreateObject("Scripting.Dictionary")

    lastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row
    lastCol = wsMaster.Cells(1, wsMaster.Columns.Count).End(xlToLeft).Column

    If lastRow < 4 Then
        ReDim scopesOut(0) : scopesOut(0) = "" : Exit Sub
    End If

    ' Collect unique scopes
    For i = 4 To lastRow
        scopeVal = Trim(CStr(wsMaster.Cells(i, "A").Value))
        If scopeVal <> "" And Not scopeDict.Exists(scopeVal) Then
            scopeDict.Add scopeVal, True
        End If
    Next i

    If scopeDict.Count = 0 Then
        ReDim scopesOut(0) : scopesOut(0) = "" : Exit Sub
    End If

    ' Read all data into memory once
    allData  = wsMaster.Range("A4:" & ColLetter(lastCol) & lastRow).Value
    colCount = UBound(allData, 2)

    ' Build sorted scope list for the out-param
    ReDim scopesOut(scopeDict.Count - 1)
    i = 0
    For Each scopeKey In scopeDict.Keys
        scopesOut(i) = CStr(scopeKey) : i = i + 1
    Next scopeKey
    Call BubbleSortStrings(scopesOut)

    ' Write one sheet per scope directly into targetWB
    Dim si As Long
    For si = 0 To UBound(scopesOut)
        scopeKey  = scopesOut(si)
        sheetName = "RATES_" & CStr(scopeKey)

        Set wsScope = targetWB.Sheets.Add(After:=targetWB.Sheets(targetWB.Sheets.Count))
        wsScope.Name = sheetName

        ' Copy 3-row header with formatting from master
        wsMaster.Rows("1:3").Copy
        wsScope.Rows("1").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False

        ' Count and filter rows for this scope
        scopeCount = 0
        For i = 1 To UBound(allData, 1)
            If Trim(CStr(allData(i, 1))) = CStr(scopeKey) Then scopeCount = scopeCount + 1
        Next i
        If scopeCount = 0 Then GoTo NextRatesScope

        ReDim scopeData(1 To scopeCount, 1 To colCount)
        dr = 1
        For i = 1 To UBound(allData, 1)
            If Trim(CStr(allData(i, 1))) = CStr(scopeKey) Then
                For c = 1 To colCount
                    scopeData(dr, c) = allData(i, c)
                Next c
                dr = dr + 1
            End If
        Next i
        wsScope.Range("A4").Resize(scopeCount, colCount).Value = scopeData

        ' Sort by CMDT Seq (col B) ascending
        Dim rSortEnd As Long : rSortEnd = 3 + scopeCount
        With wsScope.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsScope.Range("B4:B" & rSortEnd), Order:=xlAscending
            .SetRange wsScope.Range("A4:" & ColLetter(lastCol) & rSortEnd)
            .Header = xlNo
            .Apply
        End With

NextRatesScope:
    Next si
End Sub

' -------------------------------------------------------------------
' CHANGE C  —  Reads master ROUTE (1-row header, data from row 2,
' scope in col A). Writes ROUTE_<scope> sheets into targetWB,
' sorted by Header Seq (col B) asc then Route Seq (col C) asc.
' -------------------------------------------------------------------
Private Sub SplitRouteToScopeSheets(ByRef targetWB As Workbook, _
                                     ByRef scopesOut() As String)
    Dim wsMaster  As Worksheet
    Dim wsScope   As Worksheet
    Dim lastRow   As Long, lastCol As Long
    Dim i As Long, c As Long, dr As Long
    Dim scopeVal  As String, sheetName As String
    Dim scopeKey  As Variant
    Dim scopeDict As Object
    Dim allData   As Variant
    Dim colCount  As Long, scopeCount As Long
    Dim scopeData() As Variant

    Set wsMaster  = ThisWorkbook.Worksheets("ROUTE")
    Set scopeDict = CreateObject("Scripting.Dictionary")

    lastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row
    lastCol = wsMaster.Cells(1, wsMaster.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        ReDim scopesOut(0) : scopesOut(0) = "" : Exit Sub
    End If

    ' Collect unique scopes
    For i = 2 To lastRow
        scopeVal = Trim(CStr(wsMaster.Cells(i, "A").Value))
        If scopeVal <> "" And Not scopeDict.Exists(scopeVal) Then
            scopeDict.Add scopeVal, True
        End If
    Next i

    If scopeDict.Count = 0 Then
        ReDim scopesOut(0) : scopesOut(0) = "" : Exit Sub
    End If

    ' Read all data into memory once
    allData  = wsMaster.Range("A2:" & ColLetter(lastCol) & lastRow).Value
    colCount = UBound(allData, 2)

    ' Build sorted scope list for the out-param
    ReDim scopesOut(scopeDict.Count - 1)
    i = 0
    For Each scopeKey In scopeDict.Keys
        scopesOut(i) = CStr(scopeKey) : i = i + 1
    Next scopeKey
    Call BubbleSortStrings(scopesOut)

    ' Write one sheet per scope directly into targetWB
    Dim si As Long
    For si = 0 To UBound(scopesOut)
        scopeKey  = scopesOut(si)
        sheetName = "ROUTE_" & CStr(scopeKey)

        Set wsScope = targetWB.Sheets.Add(After:=targetWB.Sheets(targetWB.Sheets.Count))
        wsScope.Name = sheetName

        ' Copy header row with formatting
        wsMaster.Rows("1:1").Copy
        wsScope.Rows("1").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False

        ' Count and filter rows for this scope
        scopeCount = 0
        For i = 1 To UBound(allData, 1)
            If Trim(CStr(allData(i, 1))) = CStr(scopeKey) Then scopeCount = scopeCount + 1
        Next i
        If scopeCount = 0 Then GoTo NextRouteScope

        ReDim scopeData(1 To scopeCount, 1 To colCount)
        dr = 1
        For i = 1 To UBound(allData, 1)
            If Trim(CStr(allData(i, 1))) = CStr(scopeKey) Then
                For c = 1 To colCount
                    scopeData(dr, c) = allData(i, c)
                Next c
                dr = dr + 1
            End If
        Next i
        wsScope.Range("A2").Resize(scopeCount, colCount).Value = scopeData

        ' Sort by Header Seq (col B) asc, then Route Seq (col C) asc
        Dim rtSortEnd As Long : rtSortEnd = 1 + scopeCount
        With wsScope.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsScope.Range("B2:B" & rtSortEnd), Order:=xlAscending
            .SortFields.Add Key:=wsScope.Range("C2:C" & rtSortEnd), Order:=xlAscending
            .SetRange wsScope.Range("A2:" & ColLetter(lastCol) & rtSortEnd)
            .Header = xlNo
            .Apply
        End With

NextRouteScope:
    Next si
End Sub

' -------------------------------------------------------------------
' CHANGE C  —  Reads master SURCHARGES (1-row header, data from row 2).
' Col A contains keys like "AEW1", "AMW14" — the scope is the leading
' letter prefix (trailing digits stripped). Writes SURCHARGES_<scope>
' sheets into targetWB, sorted by scope key (col A) asc then
' Charge Seq (col E) asc.
' Formulas in cols B, C, I are read as their calculated values so the
' output sheets contain plain data with no broken cross-file references.
' -------------------------------------------------------------------
Private Sub SplitSurchargesToScopeSheets(ByRef targetWB As Workbook, _
                                          ByRef scopesOut() As String)
    Dim wsMaster  As Worksheet
    Dim wsScope   As Worksheet
    Dim lastRow   As Long, lastCol As Long
    Dim i As Long, c As Long, dr As Long
    Dim keyVal    As String, scopeVal As String, sheetName As String
    Dim scopeKey  As Variant
    Dim scopeDict As Object
    Dim allData   As Variant
    Dim colCount  As Long, scopeCount As Long
    Dim scopeData() As Variant

    Set wsMaster  = ThisWorkbook.Worksheets("SURCHARGES")
    Set scopeDict = CreateObject("Scripting.Dictionary")

    lastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row
    lastCol = wsMaster.Cells(1, wsMaster.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        ReDim scopesOut(0) : scopesOut(0) = "" : Exit Sub
    End If

    ' Collect unique scope prefixes (strip trailing digits from col A keys)
    For i = 2 To lastRow
        keyVal   = Trim(CStr(wsMaster.Cells(i, "A").Value))
        scopeVal = ExtractScopePrefix(keyVal)
        If scopeVal <> "" And Not scopeDict.Exists(scopeVal) Then
            scopeDict.Add scopeVal, True
        End If
    Next i

    If scopeDict.Count = 0 Then
        ReDim scopesOut(0) : scopesOut(0) = "" : Exit Sub
    End If

    ' Read calculated values (not formulas) into memory once.
    ' This avoids broken CODES! references in the output workbook.
    allData  = wsMaster.Range("A2:" & ColLetter(lastCol) & lastRow).Value
    colCount = UBound(allData, 2)

    ' Build sorted scope list for the out-param
    ReDim scopesOut(scopeDict.Count - 1)
    i = 0
    For Each scopeKey In scopeDict.Keys
        scopesOut(i) = CStr(scopeKey) : i = i + 1
    Next scopeKey
    Call BubbleSortStrings(scopesOut)

    ' Write one sheet per scope directly into targetWB
    Dim si As Long
    For si = 0 To UBound(scopesOut)
        scopeKey  = scopesOut(si)
        sheetName = "SURCHARGES_" & CStr(scopeKey)

        Set wsScope = targetWB.Sheets.Add(After:=targetWB.Sheets(targetWB.Sheets.Count))
        wsScope.Name = sheetName

        ' Copy header row with formatting
        wsMaster.Rows("1:1").Copy
        wsScope.Rows("1").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False

        ' Count and filter rows for this scope
        scopeCount = 0
        For i = 1 To UBound(allData, 1)
            If ExtractScopePrefix(Trim(CStr(allData(i, 1)))) = CStr(scopeKey) Then
                scopeCount = scopeCount + 1
            End If
        Next i
        If scopeCount = 0 Then GoTo NextSurchargesScope

        ReDim scopeData(1 To scopeCount, 1 To colCount)
        dr = 1
        For i = 1 To UBound(allData, 1)
            If ExtractScopePrefix(Trim(CStr(allData(i, 1)))) = CStr(scopeKey) Then
                For c = 1 To colCount
                    scopeData(dr, c) = allData(i, c)
                Next c
                dr = dr + 1
            End If
        Next i
        wsScope.Range("A2").Resize(scopeCount, colCount).Value = scopeData

        ' Sort by scope key (col A) asc, then Charge Seq (col E) asc
        Dim sSortEnd As Long : sSortEnd = 1 + scopeCount
        With wsScope.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsScope.Range("A2:A" & sSortEnd), Order:=xlAscending
            .SortFields.Add Key:=wsScope.Range("E2:E" & sSortEnd), Order:=xlAscending
            .SetRange wsScope.Range("A2:" & ColLetter(lastCol) & sSortEnd)
            .Header = xlNo
            .Apply
        End With

NextSurchargesScope:
    Next si
End Sub


' ============================================================
' HELPERS
' ============================================================

' Converts a column number to Excel letter(s): 1->"A", 38->"AL"
Function ColLetter(colNum As Long) As String
    Dim result As String, n As Long
    n = colNum
    Do While n > 0
        result = Chr(((n - 1) Mod 26) + 65) & result
        n = (n - 1) \ 26
    Loop
    ColLetter = result
End Function

' Returns True if a sheet named sheetName exists in wb.
' Called by EnsureSettingsSheet and ReadSettings.
Private Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

' Strips all trailing digit characters, returning the letter prefix.
' e.g. "AEW1" -> "AEW",  "AMW14" -> "AMW",  "WMW" -> "WMW"
Private Function ExtractScopePrefix(val As String) As String
    Dim i As Long
    i = Len(val)
    Do While i > 0
        If Mid(val, i, 1) >= "0" And Mid(val, i, 1) <= "9" Then
            i = i - 1
        Else
            Exit Do
        End If
    Loop
    ExtractScopePrefix = Left(val, i)
End Function

' Simple alphabetical bubble sort for a 0-based String array
Private Sub BubbleSortStrings(arr() As String)
    Dim i As Long, j As Long, tmp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                tmp = arr(i) : arr(i) = arr(j) : arr(j) = tmp
            End If
        Next j
    Next i
End Sub


' ============================================================
' CLEAR / RESET
' ============================================================

' -------------------------------------------------------------------
' CHANGE D: RATES_* / ROUTE_* / SURCHARGES_* pre-pass removed.
' Those scope sheets are never created in the source workbook now,
' so there is nothing to clean up. Only the master sheets are cleared.
' -------------------------------------------------------------------
Sub ClearAllSheetsDataPreserveHeaders()
    Dim ws       As Worksheet
    Dim lastRow  As Long, lastCol As Long
    Dim findCell As Range

    For Each ws In ThisWorkbook.Worksheets
        With ws
            ' FIX 6: Guard against completely empty sheets
            Set findCell = .Cells.Find("*", , , , xlByRows, xlPrevious)
            If findCell Is Nothing Then GoTo NextSheet

            lastRow = findCell.Row
            Set findCell = .Cells.Find("*", , , , xlByColumns, xlPrevious)
            lastCol = findCell.Column

            Select Case .Name
                Case "DATA", "SURCHARGES", "CODES"
                    If lastRow > 1 And lastCol > 0 Then
                        .Range(.Cells(2, 1), .Cells(lastRow, lastCol)).ClearContents
                    End If
                Case "RATES"
                    If lastRow > 2 And lastCol > 0 Then
                        .Range(.Cells(3, 1), .Cells(lastRow, lastCol)).ClearContents
                    End If
                Case "ROUTE"
                    If lastRow > 1 And lastCol > 0 Then
                        .Range(.Cells(2, 1), .Cells(lastRow, lastCol)).ClearContents
                    End If
                Case Else
                    ' Sheet1, CARGO, DATA etc. left untouched
            End Select
        End With
NextSheet:
    Next ws
End Sub


' ============================================================
' SUPPORTING PIPELINE SUBS (unchanged)
' ============================================================

Sub ApplyCountIf()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long

    Set ws = Worksheets("RATES")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 4 To lastRow
        ws.Cells(i, "AF").Value = WorksheetFunction.CountIfs( _
            ws.Range("A4:A" & i), ws.Cells(i, "A").Value, _
            ws.Range("B4:B" & i), ws.Cells(i, "B").Value)
    Next i
End Sub

' -------------------------------------------------------------------

Sub GroupedCargoNumbering()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long, i As Long, outputRow As Long
    Dim prefix As String, surcharge As String, fullText As String
    Dim sdate As String, edate As String
    Dim groupDict As Object, typeDict As Object
    Dim tempDict As Object, key As Variant, prefixKey As Variant
    Dim sortedKeys() As String
    Dim rowMap As Object

    Set ws1 = Worksheets("DATA")
    Set ws2 = Worksheets("RATES")
    Set groupDict = CreateObject("Scripting.Dictionary")
    Set rowMap    = CreateObject("Scripting.Dictionary")

    lastRow = ws1.Cells(ws1.Rows.Count, "W").End(xlUp).Row

    For i = 2 To lastRow
        prefix    = Trim(ws1.Cells(i, "C").Value)
        surcharge = Trim(ws1.Cells(i, "W").Value)
        sdate     = Trim(ws1.Cells(i, "A").Value)
        edate     = Trim(ws1.Cells(i, "B").Value)

        If prefix <> "" And surcharge <> "" And sdate <> "" And edate <> "" Then
            fullText = prefix & "|" & sdate & "|" & edate & "|" & surcharge

            If Not groupDict.Exists(prefix) Then
                Set tempDict = CreateObject("Scripting.Dictionary")
                groupDict.Add prefix, tempDict
            Else
                Set tempDict = groupDict(prefix)
            End If

            If Not tempDict.Exists(fullText) Then tempDict.Add fullText, 0
            rowMap(i) = fullText   ' FIX 3
        End If
    Next i

    For Each prefixKey In groupDict.Keys
        Set typeDict = groupDict(prefixKey)
        ReDim sortedKeys(typeDict.Count - 1)
        i = 0
        For Each key In typeDict.Keys
            sortedKeys(i) = key : i = i + 1
        Next key
        Call CustomSort(sortedKeys)
        For i = 0 To UBound(sortedKeys)
            typeDict(sortedKeys(i)) = i + 1
        Next i
    Next prefixKey

    outputRow = 4
    For i = 2 To lastRow
        If rowMap.Exists(i) Then
            fullText = rowMap(i)
            prefix   = Split(fullText, "|")(0)
            Set typeDict = groupDict(prefix)
            ws2.Cells(outputRow, "B").Value = typeDict(fullText)
            outputRow = outputRow + 1
        End If
    Next i
End Sub

' -------------------------------------------------------------------

Sub CustomSort(arr() As String)
    Dim i As Long, j As Long, temp As String
    Dim partsI() As String, partsJ() As String

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            partsI = Split(arr(i), "|")
            partsJ = Split(arr(j), "|")
            If partsI(1) > partsJ(1) _
            Or (partsI(1) = partsJ(1) And partsI(2) > partsJ(2)) _
            Or (partsI(1) = partsJ(1) And partsI(2) = partsJ(2) _
                And partsI(3) > partsJ(3)) Then
                temp = arr(i) : arr(i) = arr(j) : arr(j) = temp
            End If
        Next j
    Next i
End Sub

' -------------------------------------------------------------------

Sub AddTSPortPrefix()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long

    Set ws = Worksheets("DATA")
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
    For i = 2 To lastRow
        If Trim(ws.Cells(i, "P").Value) <> "" Then
            ws.Cells(i, "Y").Value = "T/S PORT - " & ws.Cells(i, "P").Value
        Else
            ws.Cells(i, "Y").Value = ""
        End If
    Next i
End Sub

' -------------------------------------------------------------------

Sub AddSLanePrefix()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long

    Set ws = Worksheets("DATA")
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row
    For i = 2 To lastRow
        If Trim(ws.Cells(i, "Q").Value) <> "" Then
            ws.Cells(i, "Z").Value = "Service Lane - " & ws.Cells(i, "Q").Value
        Else
            ws.Cells(i, "Z").Value = ""
        End If
    Next i
End Sub

' -------------------------------------------------------------------

Sub MapFAKValuesToAA()
    ' CHANGE F: Refactored from Select Case to If/ElseIf so that both the
    ' full commodity description strings and bare port codes are matched by
    ' a single InStr check per variant.
    '
    ' Exact-match cases (Turkish terminals — values are fixed):
    '   "HAYDARPASA FAK"  → "TRIST21 - HAYDARPASA"
    '   "MARPORT FAK"     → "TRIST02 - MARPORT"
    '
    ' Substring-match cases (Egyptian terminals — col D may contain the
    ' full description or just the code):
    '   anything containing "EGALY20"  → "EGALY20"
    '     e.g. "FAK Alexandria (EGALY20 DEKHEILA PORT)" or "EGALY20"
    '   anything containing "EGALY21"  → "EGALY21"
    '     e.g. "FAK Alexandria (EGALY21 Old Port)"     or "EGALY21"
    '
    ' The value written to col AA flows into ConcatRouteNote (col X) and
    ' then into ROUTE col E, so it appears in the Route Note automatically.
    ' ExtractNode (below) then picks it up for the Node column (col AE).
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim cellValue As String

    Set ws = Worksheets("DATA")
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    For i = 2 To lastRow
        cellValue = Trim(ws.Cells(i, "D").Value)
        If cellValue = "HAYDARPASA FAK" Then
            ws.Cells(i, "AA").Value = "TRIST21 - HAYDARPASA"
        ElseIf cellValue = "MARPORT FAK" Then
            ws.Cells(i, "AA").Value = "TRIST02 - MARPORT"
        ElseIf InStr(1, cellValue, "EGALY20", vbTextCompare) > 0 Then
            ws.Cells(i, "AA").Value = "EGALY20"
        ElseIf InStr(1, cellValue, "EGALY21", vbTextCompare) > 0 Then
            ws.Cells(i, "AA").Value = "EGALY21"
        Else
            ws.Cells(i, "AA").Value = ""
        End If
    Next i
End Sub

' -------------------------------------------------------------------

Sub ConcatRouteNote()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim yVal As String, zVal As String, aaVal As String, result As String

    Set ws = Worksheets("DATA")
    lastRow = Application.WorksheetFunction.Max( _
        ws.Cells(ws.Rows.Count, "Y").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "Z").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row)

    For i = 2 To lastRow
        yVal   = Trim(ws.Cells(i, "Y").Value)
        zVal   = Trim(ws.Cells(i, "Z").Value)
        aaVal  = Trim(ws.Cells(i, "AA").Value)
        result = ""
        If yVal = "" And zVal = "" Then
            If aaVal <> "" Then result = aaVal
        ElseIf yVal = "" And zVal <> "" Then
            result = zVal
            If aaVal <> "" Then result = result & " " & aaVal
        Else
            result = yVal
            If zVal  <> "" Then result = result & " | " & zVal
            If aaVal <> "" Then result = result & " | " & aaVal
        End If
        ws.Cells(i, "X").Value = result
    Next i
End Sub

' -------------------------------------------------------------------

Sub GenerateRouteNote()
    Dim ws1 As Worksheet, ws3 As Worksheet
    Dim lastRow As Long, r As Long
    Dim destRow As Long   ' FIX 4

    Set ws1 = Worksheets("RATES")
    Set ws3 = Worksheets("ROUTE")

    lastRow = ws1.Cells(ws1.Rows.Count, "AC").End(xlUp).Row
    destRow = 2

    For r = 2 To lastRow
        If Trim(ws1.Range("AC" & r).Value) <> "" Then
            ws3.Range("E" & destRow).Value = ws1.Range("AC" & r).Value
            ws3.Range("A" & destRow).Value = ws1.Range("A"  & r).Value
            ws3.Range("C" & destRow).Value = ws1.Range("AF" & r).Value
            ws3.Range("B" & destRow).Value = ws1.Range("B"  & r).Value
            ws3.Range("H" & destRow).Value = ws1.Range("AD" & r).Value
            ws3.Range("I" & destRow).Value = ws1.Range("AE" & r).Value
            ws3.Range("G" & destRow).Value = "APP"
            ws3.Range("D" & destRow).Value = "1"
            ws3.Range("J" & destRow).Value = "S"
            ws3.Range("F" & destRow).Value = "1"
            destRow = destRow + 1
        End If
    Next r

    Call ExtractTSPort
    Call ExtractNode
    Call ExtractServiceLane
End Sub

' -------------------------------------------------------------------

Sub ExtractTSPort()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceText As String, result As String
    Dim startPos As Long, endPos As Long

    Set ws = ThisWorkbook.Sheets("ROUTE")
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    For i = 2 To lastRow
        On Error GoTo SkipRow
        sourceText = ws.Cells(i, "E").Value
        startPos   = InStr(sourceText, "T/S PORT -")
        If startPos > 0 Then
            startPos = startPos + 10
            endPos   = InStr(startPos, sourceText & "|", "|")
            result   = Trim(Mid(sourceText, startPos, endPos - startPos))
        Else
            result = ""
        End If
        ws.Cells(i, "W").Value = result
SkipRow:
        On Error GoTo 0
    Next i
End Sub

' -------------------------------------------------------------------

Sub ExtractServiceLane()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceText As String, lowerText As String, result As String
    Dim startPos As Long, endPos As Long

    Set ws = ThisWorkbook.Sheets("ROUTE")
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    For i = 2 To lastRow
        On Error GoTo SkipRow
        sourceText = ws.Cells(i, "E").Value
        lowerText  = LCase(sourceText)
        startPos   = InStr(lowerText, "service lane -")
        If startPos > 0 Then
            startPos = startPos + 15
            endPos   = InStr(startPos, lowerText & "|", "|")
            result   = Trim(Mid(sourceText, startPos, endPos - startPos))
        Else
            result = ""
        End If
        ws.Cells(i, "V").Value = result
SkipRow:
        On Error GoTo 0
    Next i
End Sub

' -------------------------------------------------------------------

Sub ExtractNode()
    ' CHANGE G: Added ElseIf branches for EGALY20 and EGALY21.
    '
    ' TRIST ports (HAYDARPASA, MARPORT) embed the node as "TRISTxx - PORT NAME"
    ' so dash-position parsing is required to extract the "TRISTxx" prefix.
    '
    ' EGALY ports are written to col E as just "EGALY20" or "EGALY21" (no dash),
    ' so the node value is the match itself — no parsing needed.
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceText As String, result As String
    Dim dashPos As Long

    Set ws = ThisWorkbook.Sheets("ROUTE")
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    For i = 2 To lastRow
        On Error GoTo SkipRow
        sourceText = ws.Cells(i, "E").Value
        If InStr(1, sourceText, "HAYDARPASA", vbTextCompare) > 0 Or _
           InStr(1, sourceText, "MARPORT",    vbTextCompare) > 0 Then
            ' TRIST terminals: extract the "TRISTxx" code before the dash
            dashPos = InStr(sourceText, " -")
            If dashPos > 0 Then
                If Left(sourceText, 5) = "TRIST" Then
                    result = Trim(Left(sourceText, dashPos - 1))
                Else
                    result = Trim(Mid(sourceText, dashPos + 2))
                End If
            Else
                result = ""
            End If
        ElseIf InStr(1, sourceText, "EGALY20", vbTextCompare) > 0 Then
            ' Alexandria Dekheila Port — code is the full node value
            result = "EGALY20"
        ElseIf InStr(1, sourceText, "EGALY21", vbTextCompare) > 0 Then
            ' Alexandria Old Port — code is the full node value
            result = "EGALY21"
        Else
            result = ""
        End If
        ws.Cells(i, "AE").Value = result
SkipRow:
        On Error GoTo 0
    Next i
End Sub

' -------------------------------------------------------------------

Sub ExpandDataWithHeaderAndCount()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim prefix As String, codes() As String
    Dim outputRow As Long, count As Long

    Set ws1 = Worksheets("CODES")
    Set ws2 = Worksheets("SURCHARGES")

    lastRow   = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    outputRow = 2

    For i = 1 To lastRow
        prefix = ws1.Cells(i, 1).Value
        If prefix <> "" And ws1.Cells(i, 2).Value <> "" Then
            count = 1
            ws2.Cells(outputRow, 4).Value = prefix
            ws2.Cells(outputRow, 5).Value = count
            ws2.Cells(outputRow, 6).Value = "APP"
            outputRow = outputRow + 1 : count = count + 1

            codes = Split(ws2.Cells(i, 2).Value, ",")
            For j = LBound(codes) To UBound(codes)
                ws2.Cells(outputRow, 4).Value = prefix
                ws2.Cells(outputRow, 5).Value = count
                ws2.Cells(outputRow, 6).Value = Trim(codes(j))
                outputRow = outputRow + 1 : count = count + 1
            Next j
        End If
    Next i

    With ws2.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws2.Range("D2:D" & outputRow - 1), Order:=xlAscending
        .SortFields.Add key:=ws2.Range("E2:E" & outputRow - 1), Order:=xlAscending
        .SetRange ws2.Range("D2:F" & outputRow - 1)
        .Header = xlNo
        .Apply
    End With
End Sub

' -------------------------------------------------------------------

Sub CopySurcharges()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim srcCodes As Variant, desCodes As Variant
    Dim lastRow As Long, i As Long

    Set ws1 = Worksheets("RATES")
    Set ws2 = Worksheets("CODES")

    srcCodes = Array("A", "B", "AB", "AD", "AE")
    desCodes = Array("B", "C", "D", "E", "F")

    For i = 0 To UBound(srcCodes)
        lastRow = ws1.Cells(ws1.Rows.Count, srcCodes(i)).End(xlUp).Row
        If lastRow >= 4 Then
            ws2.Range(desCodes(i) & "2").Resize(lastRow - 3, 1).Value = _
            ws1.Range(srcCodes(i) & "4").Resize(lastRow - 3, 1).Value
        End If
    Next i

    lastRow = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRow
        ws2.Cells(i, "A").Value = ws2.Cells(i, "B").Value & ws2.Cells(i, "C").Value
    Next i

    Call RemoveDuplicateRows
End Sub

' -------------------------------------------------------------------

Sub RemoveDuplicateRows()
    Dim ws As Worksheet
    Set ws = Worksheets("CODES")

    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 1 To lastRow
        ws.Cells(i, "G").Value = ws.Cells(i, "A").Value & "|" & _
                                 ws.Cells(i, "B").Value & "|" & _
                                 ws.Cells(i, "C").Value & "|" & _
                                 ws.Cells(i, "D").Value & "|" & _
                                 ws.Cells(i, "E").Value & "|" & _
                                 ws.Cells(i, "F").Value
    Next i

    ws.Range("G1:G" & lastRow).RemoveDuplicates Columns:=1, Header:=xlNo

    Dim newLastRow As Long
    newLastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    If newLastRow < lastRow Then
        ws.Range("A" & newLastRow + 1 & ":G" & lastRow).ClearContents
    End If

    Dim parts() As String
    For i = 1 To newLastRow
        parts = Split(ws.Cells(i, "G").Value, "|")
        If UBound(parts) >= 5 Then
            ws.Cells(i, "A").Value = parts(0)
            ws.Cells(i, "B").Value = parts(1)
            ws.Cells(i, "C").Value = parts(2)
            ws.Cells(i, "D").Value = parts(3)
            ws.Cells(i, "E").Value = parts(4)
            ws.Cells(i, "F").Value = parts(5)
        End If
    Next i

    ws.Range("G1:G" & lastRow).ClearContents
End Sub

' -------------------------------------------------------------------

Sub GenerateSurcharges()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim prefix As String, sdate As String, edate As String
    Dim codes() As String
    Dim outputRow As Long, count As Long

    Call ExpandDataWithHeaderAndCount
    Call CopySurcharges

    Set ws1 = ThisWorkbook.Worksheets("CODES")
    Set ws2 = ThisWorkbook.Worksheets("SURCHARGES")

    lastRow   = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    outputRow = 2

    For i = 1 To lastRow
        prefix = ws1.Cells(i, "A").Value
        sdate  = ws1.Cells(i, "E").Value
        edate  = ws1.Cells(i, "F").Value
        If prefix <> "" And ws1.Cells(i, "D").Value <> "" Then
            count = 1
            ws2.Cells(outputRow, "A").Value   = prefix
            ws2.Cells(outputRow, "E").Value   = count
            ws2.Cells(outputRow, "F").Value   = "APP"
            ws2.Cells(outputRow, "G").Value   = sdate
            ws2.Cells(outputRow, "H").Value   = edate
            ws2.Cells(outputRow, "I").Formula = "=IF(F" & outputRow & "=""APP"",""S"",""I"")"
            ws2.Cells(outputRow, "B").Formula = "=IF(F" & outputRow & "=""APP"",CODES!C" & i & ","""")"
            ws2.Cells(outputRow, "C").Formula = "=IF(F" & outputRow & "=""APP"",1,"""")"
            outputRow = outputRow + 1 : count = count + 1

            codes = Split(ws1.Cells(i, "D").Value, ",")
            For j = LBound(codes) To UBound(codes)
                ws2.Cells(outputRow, "A").Value   = prefix
                ws2.Cells(outputRow, "E").Value   = count
                ws2.Cells(outputRow, "F").Value   = Trim(codes(j))
                ws2.Cells(outputRow, "G").Value   = sdate
                ws2.Cells(outputRow, "H").Value   = edate
                ws2.Cells(outputRow, "I").Formula = "=IF(F" & outputRow & "=""APP"",""S"",""I"")"
                ws2.Cells(outputRow, "B").Formula = "=IF(F" & outputRow & "=""APP"",CODES!C" & i & ","""")"
                ws2.Cells(outputRow, "C").Formula = "=IF(F" & outputRow & "=""APP"",1,"""")"
                outputRow = outputRow + 1 : count = count + 1
            Next j
        End If
    Next i

    With ws2.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws2.Range("A2:A" & outputRow - 1), Order:=xlAscending
        .SortFields.Add key:=ws2.Range("B2:B" & outputRow - 1), Order:=xlAscending
        .SetRange ws2.Range("A2:I" & outputRow - 1)
        .Header = xlNo
        .Apply
    End With

    Call GenerateWording
End Sub

' -------------------------------------------------------------------

Sub GenerateWording()
    Dim wsCodes As Worksheet, wsData As Worksheet
    Dim lastRow As Long, i As Long
    Dim code As String, description As String
    Dim blockText As String
    Dim foundCell As Range
    Dim appRow As Long
    Dim inBlock As Boolean
    Dim validFrom As String, validTo As String
    Dim firstItem As Boolean
    Dim fullText  As String   ' FIX 5
    Dim finalText As String   ' FIX 5

    Set wsCodes = ThisWorkbook.Sheets("SURCHARGES")
    Set wsData  = ThisWorkbook.Sheets("Sheet1")

    lastRow   = wsCodes.Cells(wsCodes.Rows.Count, "F").End(xlUp).Row
    inBlock   = False
    blockText = ""

    For i = 2 To lastRow
        code = Trim(wsCodes.Cells(i, "F").Value)

        If UCase(code) = "APP" Then
            If inBlock And blockText <> "" Then
                validFrom = Format(wsCodes.Cells(appRow, "G").Value, "yyyymmdd")
                validTo   = Format(wsCodes.Cells(appRow, "H").Value, "yyyymmdd")
                fullText  = "Rates are valid from " & validFrom & " to " & validTo & vbCrLf & _
                            "Rates are inclusive of the " & blockText & vbCrLf & _
                            "Rates are subject to all other surcharges including those, if any, " & _
                            "specified in the contract and those published in the Governing Tariff(s) " & _
                            "at the time of shipment."
                With wsCodes.Cells(appRow, "D")
                    .Value     = fullText
                    .WrapText  = True
                    .RowHeight = 30
                    wsCodes.Columns("D").ColumnWidth = 50
                End With
            End If
            blockText = "" : inBlock = True : appRow = i : firstItem = True

        ElseIf inBlock And code <> "" Then
            Set foundCell = wsData.Columns("A").Find(What:=code, LookIn:=xlValues, LookAt:=xlWhole)
            If Not foundCell Is Nothing Then
                description = Trim(foundCell.Offset(0, 1).Value)
                If description <> "" Then
                    If firstItem Then
                        blockText = description & "(" & code & ")"
                        firstItem = False
                    Else
                        blockText = blockText & " and the " & description & "(" & code & ")"
                    End If
                End If
            End If
        End If
    Next i

    ' Flush last block
    If inBlock And blockText <> "" Then
        validFrom = Format(wsCodes.Cells(appRow, "G").Value, "yyyymmdd")
        validTo   = Format(wsCodes.Cells(appRow, "H").Value, "yyyymmdd")
        finalText = "Rates are valid from " & validFrom & " to " & validTo & vbCrLf & _
                    "Rates are inclusive of the " & blockText & vbCrLf & _
                    "Rates are subject to all other surcharges including those, if any, " & _
                    "specified in the contract and those published in the Governing Tariff(s) " & _
                    "at the time of shipment."
        With wsCodes.Cells(appRow, "D")
            .Value     = finalText
            .WrapText  = True
            .RowHeight = 30
            wsCodes.Columns("D").ColumnWidth = 50
        End With
    End If
End Sub


' ============================================================
' SETTINGS  (Changes H, I, J, K)
' ============================================================

' -------------------------------------------------------------------
' CHANGE H — Public setup sub, run once before first use.
' Creates the SETTINGS sheet with labelled input cells and Yes/No
' dropdown validation. Does nothing if the sheet already exists,
' so it is safe to call repeatedly.
'
' SETTINGS sheet layout:
'   Row 1  : Title
'   Row 2  : Column headers
'   Row 4  : Include Dry Dangerous   | Yes / No
'   Row 5  : Include D7 (OFT 45)     | Yes / No
'   Row 6  : D7 Add-on Value         | numeric (e.g. 700)
' -------------------------------------------------------------------
Sub EnsureSettingsSheet()
    If SheetExists(ThisWorkbook, "SETTINGS") Then
        MsgBox "SETTINGS sheet already exists.", vbInformation, "Settings"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = "SETTINGS"

    ' ── Title ─────────────────────────────────────────────────────────
    With ws.Range("A1")
        .Value          = "Tool Settings"
        .Font.Bold      = True
        .Font.Size      = 14
    End With

    ' ── Column headers ────────────────────────────────────────────────
    ws.Range("A2").Value = "Setting"
    ws.Range("B2").Value = "Value"
    ws.Range("C2").Value = "Description"
    ws.Range("A2:C2").Font.Bold = True

    ' ── Row 3: blank separator ────────────────────────────────────────

    ' ── Row 4: Dry Dangerous toggle ───────────────────────────────────
    ws.Range("A4").Value = "Include Dry Dangerous"
    ws.Range("B4").Value = "No"
    ws.Range("C4").Value = "Duplicate every Dry General row as Dry Dangerous " & _
                           "(same rates, same route notes, cargo type changed)"

    ' ── Row 5: D7 toggle ──────────────────────────────────────────────
    ws.Range("A5").Value = "Include D7 (OFT 45)"
    ws.Range("B5").Value = "No"
    ws.Range("C5").Value = "Populate OFT 45 column (col AA in RATES) = OFT HC + Add-on Value"

    ' ── Row 6: D7 add-on value ────────────────────────────────────────
    ws.Range("A6").Value = "D7 Add-on Value"
    ws.Range("B6").Value = 700
    ws.Range("C6").Value = "Amount added to each OFT HC rate to produce the OFT 45 rate " & _
                           "(only used when Include D7 = Yes)"

    ' ── Yes/No dropdown validation on B4 and B5 ───────────────────────
    Dim r As Range
    For Each r In ws.Range("B4:B5")
        With r.Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:="Yes,No"
            .IgnoreBlank    = True
            .ShowInput      = True
            .ShowError      = True
            .ErrorTitle     = "Invalid value"
            .ErrorMessage   = "Please select Yes or No from the dropdown."
        End With
    Next r

    ' ── Numeric validation on B6 ──────────────────────────────────────
    With ws.Range("B6").Validation
        .Delete
        .Add Type:=xlValidateDecimal, _
             Operator:=xlGreaterEqual, _
             Formula1:="0"
        .IgnoreBlank  = True
        .ShowError    = True
        .ErrorTitle   = "Invalid value"
        .ErrorMessage = "D7 Add-on Value must be a number >= 0."
    End With

    ' ── Formatting ────────────────────────────────────────────────────
    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 65
    ws.Range("B4:B6").HorizontalAlignment = xlCenter

    MsgBox "SETTINGS sheet created with default values." & vbCrLf & _
           "Review and adjust before running GetData.", _
           vbInformation, "Settings Ready"
End Sub

' -------------------------------------------------------------------
' CHANGE I — Private; called at the top of GetData every run.
' Reads the three user settings from the SETTINGS sheet into typed
' ByRef out-parameters. If the SETTINGS sheet is missing, all
' parameters retain their safe defaults (both toggles False, addon 700).
' -------------------------------------------------------------------
Private Sub ReadSettings(ByRef includeDryDangerous As Boolean, _
                          ByRef includeD7 As Boolean, _
                          ByRef d7Addon As Double)
    ' Safe defaults — used when SETTINGS sheet is absent
    includeDryDangerous = False
    includeD7           = False
    d7Addon             = 700

    If Not SheetExists(ThisWorkbook, "SETTINGS") Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SETTINGS")

    If UCase(Trim(CStr(ws.Range("B4").Value))) = "YES" Then includeDryDangerous = True
    If UCase(Trim(CStr(ws.Range("B5").Value))) = "YES" Then includeD7 = True

    Dim raw As Variant
    raw = ws.Range("B6").Value
    If IsNumeric(raw) And CDbl(raw) >= 0 Then d7Addon = CDbl(raw)
End Sub

' -------------------------------------------------------------------
' CHANGE J — Private; called from GetData when Include Dry Dangerous = Yes.
'
' Appends a copy of every "Dry General" row in DATA to the bottom of
' DATA, changing col O (Cargo/Container Type) to "Dry Dangerous".
'
' Why append to DATA rather than insert rows directly into RATES:
'   - GroupedCargoNumbering reads DATA to assign CMDT Seq numbers.
'     Rows appended here are seen by GroupedCargoNumbering and get
'     their own correct sequence numbers automatically.
'   - GenerateRouteNote, GenerateSurcharges and the export split all
'     read from RATES which is populated from DATA — so the new rows
'     propagate through the entire pipeline with zero extra handling.
'
' Why run after ConcatRouteNote:
'   - By this point cols Y, Z, AA (T/S port, service lane, FAK node)
'     and col X (ROUTE NOTE concat) are already written for all Dry
'     General rows. The copy captures them so Dry Dangerous rows carry
'     identical route notes without any re-processing.
'
' Cleanup: ClearAllSheetsDataPreserveHeaders clears DATA from row 2
' onwards, which removes these appended rows automatically.
' -------------------------------------------------------------------
Private Sub AppendDryDangerousToData()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, destRow As Long, srcRow As Long
    Dim count As Long
    Dim dryGenRows() As Long

    Set ws = Worksheets("DATA")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' ── Pass 1: collect Dry General row numbers BEFORE appending ──────
    ' Collecting first avoids reading our own newly-appended rows.
    count = 0
    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, "O").Value)) = "Dry General" Then
            ReDim Preserve dryGenRows(count)
            dryGenRows(count) = i
            count = count + 1
        End If
    Next i

    If count = 0 Then Exit Sub

    ' ── Pass 2: append copies with "Dry Dangerous" in col O ───────────
    destRow = lastRow + 1
    For i = 0 To count - 1
        srcRow = dryGenRows(i)
        ' Copy full row then immediately override col O
        ws.Range(ws.Cells(srcRow, 1), ws.Cells(srcRow, lastCol)).Copy _
            Destination:=ws.Cells(destRow, 1)
        ws.Cells(destRow, "O").Value = "Dry Dangerous"
        destRow = destRow + 1
    Next i

    Application.CutCopyMode = False
End Sub

' -------------------------------------------------------------------
' CHANGE K — Private; called from GetData when Include D7 = Yes.
'
' Populates RATES col AA (OFT 45) for every data row where col Y
' (OFT HC) contains a numeric value:
'   col AA  =  col Y  +  d7Addon
'
' Existing OFT 45 values (copied from DATA col V via the bulk mapping)
' are overwritten — D7 mode always computes OFT 45 from OFT HC.
'
' This sub runs BEFORE the USD label loop in GetData, so col Z
' ("USD" currency label for OFT 45) is automatically stamped on every
' row that receives a D7 value without any extra handling.
'
' Parameters:
'   ws       — the RATES worksheet (passed in to avoid re-resolving it)
'   d7Addon  — the numeric add-on read from SETTINGS row 6 (e.g. 700)
' -------------------------------------------------------------------
Private Sub ApplyD7Rates(ByRef ws As Worksheet, ByVal d7Addon As Double)
    ' D7 (OFT 45) exclusion rules — a row is SKIPPED if ANY of these are true:
    '   1. Service Scope (col A) is not AEW or AMW
    '   2. Cargo/Container Type (col AL) contains "Reefer" (covers "Reefer",
    '      "Reefer Dry" — case-insensitive)
    '   3. POD location (col H) contains "JP" (Japan ports — case-insensitive)
    '
    ' Column references in RATES (data starts row 4):
    '   col A  = Service Scope      (from DATA col C)
    '   col AL = Cargo Type         (from DATA col O)
    '   col H  = POD / Location     (from DATA col E) — JP check applied here
    '   col Y  = OFT HC (source value for D7 calculation)
    '   col AA = OFT 45 (target — written only when all rules pass)
    Dim lastRow As Long, i As Long
    Dim hcVal    As Variant
    Dim scope    As String
    Dim cargo    As String
    Dim location As String

    lastRow = ws.Cells(ws.Rows.Count, "Y").End(xlUp).Row

    For i = 4 To lastRow
        ' Rule 1: Scope must be AEW or AMW
        scope = UCase(Trim(CStr(ws.Cells(i, "A").Value)))
        If scope <> "AEW" And scope <> "AMW" Then GoTo SkipRow

        ' Rule 2: Skip all Reefer cargo types
        cargo = LCase(Trim(CStr(ws.Cells(i, "AL").Value)))
        If InStr(cargo, "reefer") > 0 Then GoTo SkipRow

        ' Rule 3: Skip JP (Japan) locations
        location = UCase(Trim(CStr(ws.Cells(i, "H").Value)))
        If InStr(location, "JP") > 0 Then GoTo SkipRow

        ' All rules passed — apply D7 value
        hcVal = ws.Cells(i, "Y").Value
        If IsNumeric(hcVal) And CStr(hcVal) <> "" Then
            ws.Cells(i, "AA").Value = CDbl(hcVal) + d7Addon
        End If

SkipRow:
    Next i
End Sub
