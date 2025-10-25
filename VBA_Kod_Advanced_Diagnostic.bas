' ============================================
' ADVANCED DIAGNOSTIC - Detaljniji alat
' ============================================
' Pokazuje TAČNO šta je problem na svakom sheet-u
' ============================================

Option Explicit

' Glavna diagnostic funkcija - prikazuje detaljnu analizu
Sub DiagnostikujDetaljno()
    Dim ws As Worksheet
    Dim izvestaj As String
    Dim wsReport As Worksheet
    Dim row As Long

    ' Kreiraj novi sheet za izveštaj
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("DIAGNOSTIC_REPORT").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsReport = ThisWorkbook.Worksheets.Add
    wsReport.Name = "DIAGNOSTIC_REPORT"

    ' Header
    row = 1
    wsReport.Cells(row, 1).Value = "DETALJNI DIAGNOSTIC IZVEŠTAJ"
    wsReport.Cells(row, 1).Font.Bold = True
    wsReport.Cells(row, 1).Font.Size = 14
    row = row + 2

    wsReport.Cells(row, 1).Value = "Sheet"
    wsReport.Cells(row, 2).Value = "Status"
    wsReport.Cells(row, 3).Value = "Redovi"
    wsReport.Cells(row, 4).Value = "Numeričke"
    wsReport.Cells(row, 5).Value = "Tekst brojevi"
    wsReport.Cells(row, 6).Value = "Formule"
    wsReport.Cells(row, 7).Value = "Kolone sa cenama"
    wsReport.Cells(row, 8).Value = "Problem"
    wsReport.Range(wsReport.Cells(row, 1), wsReport.Cells(row, 8)).Font.Bold = True
    row = row + 1

    ' Analiziraj svaki sheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "DIAGNOSTIC_REPORT" Then
            Dim analiza As SheetAnalysis
            analiza = AnalizirajSheet(ws)

            wsReport.Cells(row, 1).Value = ws.Name
            wsReport.Cells(row, 2).Value = analiza.status
            wsReport.Cells(row, 3).Value = analiza.brojRedova
            wsReport.Cells(row, 4).Value = analiza.brojNumerickihCelija
            wsReport.Cells(row, 5).Value = analiza.brojTekstualnihBrojeva
            wsReport.Cells(row, 6).Value = analiza.brojFormula
            wsReport.Cells(row, 7).Value = analiza.koloneSaCenama
            wsReport.Cells(row, 8).Value = analiza.problem

            ' Oboji red u zavisnosti od statusa
            If analiza.status = "PROBLEM" Then
                wsReport.Range(wsReport.Cells(row, 1), wsReport.Cells(row, 8)).Interior.Color = RGB(255, 200, 200)
            ElseIf analiza.status = "OK" Then
                wsReport.Range(wsReport.Cells(row, 1), wsReport.Cells(row, 8)).Interior.Color = RGB(200, 255, 200)
            Else
                wsReport.Range(wsReport.Cells(row, 1), wsReport.Cells(row, 8)).Interior.Color = RGB(255, 255, 200)
            End If

            row = row + 1
        End If
    Next ws

    ' Formatiraj kolone
    wsReport.Columns("A:H").AutoFit

    ' Dodaj detaljne informacije ispod
    row = row + 2
    wsReport.Cells(row, 1).Value = "DETALJNI PRIMERI PODATAKA PO SHEET-OVIMA:"
    wsReport.Cells(row, 1).Font.Bold = True
    row = row + 2

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "DIAGNOSTIC_REPORT" Then
            wsReport.Cells(row, 1).Value = "=== " & ws.Name & " ==="
            wsReport.Cells(row, 1).Font.Bold = True
            row = row + 1

            Dim primeri As String
            primeri = PrikaziPrimereIzSheeta(ws)

            ' Split na redove i prikaži
            Dim lines() As String
            lines = Split(primeri, vbCrLf)

            Dim i As Long
            For i = LBound(lines) To UBound(lines)
                wsReport.Cells(row, 1).Value = lines(i)
                row = row + 1
            Next i

            row = row + 1
        End If
    Next ws

    wsReport.Activate
    wsReport.Range("A1").Select

    MsgBox "Diagnostic završen!" & vbCrLf & vbCrLf & _
           "Pogledaj sheet 'DIAGNOSTIC_REPORT' za detaljne informacije." & vbCrLf & vbCrLf & _
           "Sheet-ovi označeni CRVENOM imaju problem!", vbInformation
End Sub

' Struktura za analizu sheet-a
Type SheetAnalysis
    status As String
    brojRedova As Long
    brojNumerickihCelija As Long
    brojTekstualnihBrojeva As Long
    brojFormula As Long
    koloneSaCenama As String
    problem As String
End Type

' Analiziraj jedan sheet
Function AnalizirajSheet(ws As Worksheet) As SheetAnalysis
    Dim result As SheetAnalysis
    Dim cell As Range
    Dim i As Long, j As Long

    result.status = "OK"
    result.brojRedova = 0
    result.brojNumerickihCelija = 0
    result.brojTekstualnihBrojeva = 0
    result.brojFormula = 0
    result.koloneSaCenama = ""
    result.problem = "Nema"

    ' Proveri zaštitu
    If ws.ProtectContents Then
        result.status = "ZAŠTIĆEN"
        result.problem = "Sheet je zaštićen!"
        AnalizirajSheet = result
        Exit Function
    End If

    ' Nađi broj redova
    On Error Resume Next
    result.brojRedova = ws.UsedRange.Rows.Count
    On Error GoTo 0

    If result.brojRedova <= 1 Then
        result.status = "PRAZAN"
        result.problem = "Nema podataka (samo header ili prazan)"
        AnalizirajSheet = result
        Exit Function
    End If

    ' Detektuj kolone sa cenama
    Dim maxCol As Long
    maxCol = ws.UsedRange.Columns.Count

    For i = 1 To maxCol
        Dim header As String
        On Error Resume Next
        header = LCase(Trim(ws.Cells(1, i).Value))
        On Error GoTo 0

        If InStr(header, "cena") > 0 Or InStr(header, "cijena") > 0 Or _
           InStr(header, "price") > 0 Or InStr(header, "cost") > 0 Or _
           InStr(header, "iznos") > 0 Or InStr(header, "vrednost") > 0 Then
            If result.koloneSaCenama <> "" Then result.koloneSaCenama = result.koloneSaCenama & ", "
            result.koloneSaCenama = result.koloneSaCenama & GetColumnLetterHelper(i)
        End If
    Next i

    ' Analiziraj podatke (samo prvih 100 redova za brzinu)
    Dim maxRowToCheck As Long
    maxRowToCheck = Application.WorksheetFunction.Min(100, result.brojRedova)

    For i = 2 To maxRowToCheck
        For j = 1 To maxCol
            On Error Resume Next
            Set cell = ws.Cells(i, j)
            On Error GoTo 0

            If Not cell Is Nothing Then
                If cell.HasFormula Then
                    result.brojFormula = result.brojFormula + 1
                ElseIf IsNumeric(cell.Value) And cell.Value <> "" And Not IsEmpty(cell.Value) Then
                    result.brojNumerickihCelija = result.brojNumerickihCelija + 1
                ElseIf VarType(cell.Value) = vbString And cell.Value <> "" Then
                    ' Proveri da li je broj u tekstu
                    Dim testVal As String
                    testVal = Trim(cell.Value)
                    testVal = Replace(testVal, " ", "")
                    testVal = Replace(testVal, ",", ".")

                    If IsNumeric(testVal) Then
                        result.brojTekstualnihBrojeva = result.brojTekstualnihBrojeva + 1
                    End If
                End If
            End If
        Next j
    Next i

    ' Odredi status i problem
    If result.brojNumerickihCelija = 0 And result.brojTekstualnihBrojeva = 0 Then
        result.status = "PROBLEM"
        result.problem = "Nema numeričkih vrednosti!"
    ElseIf result.brojTekstualnihBrojeva > result.brojNumerickihCelija Then
        result.status = "UPOZORENJE"
        result.problem = "Većina brojeva je u tekstualnom formatu!"
    ElseIf result.brojTekstualnihBrojeva > 0 Then
        result.status = "UPOZORENJE"
        result.problem = "Neki brojevi su u tekstualnom formatu"
    ElseIf result.brojFormula > 0 And result.brojNumerickihCelija = 0 Then
        result.status = "PROBLEM"
        result.problem = "Sve vrednosti su formule!"
    End If

    AnalizirajSheet = result
End Function

' Prikaži primere podataka iz sheet-a
Function PrikaziPrimereIzSheeta(ws As Worksheet) As String
    Dim result As String
    Dim i As Long, j As Long
    Dim maxRow As Long, maxCol As Long

    result = ""

    On Error Resume Next
    maxRow = Application.WorksheetFunction.Min(5, ws.UsedRange.Rows.Count)
    maxCol = Application.WorksheetFunction.Min(10, ws.UsedRange.Columns.Count)
    On Error GoTo 0

    If maxRow = 0 Then
        result = "  (Sheet je prazan)"
        PrikaziPrimereIzSheeta = result
        Exit Function
    End If

    ' Prikaži header
    result = "  HEADER (Red 1):" & vbCrLf
    For j = 1 To maxCol
        Dim headerVal As Variant
        On Error Resume Next
        headerVal = ws.Cells(1, j).Value
        On Error GoTo 0

        If headerVal <> "" Then
            result = result & "    " & GetColumnLetterHelper(j) & ": " & headerVal & vbCrLf
        End If
    Next j

    ' Prikaži primere podataka
    result = result & vbCrLf & "  PRIMERI PODATAKA (Red 2-3):" & vbCrLf
    For i = 2 To Application.WorksheetFunction.Min(3, maxRow)
        result = result & "  Red " & i & ":" & vbCrLf
        For j = 1 To maxCol
            Dim cellVal As Variant
            On Error Resume Next
            cellVal = ws.Cells(i, j).Value
            On Error GoTo 0

            If cellVal <> "" Then
                Dim cellInfo As String
                cellInfo = "    " & GetColumnLetterHelper(j) & ": "

                ' Dodaj vrednost
                cellInfo = cellInfo & cellVal

                ' Dodaj tip
                cellInfo = cellInfo & " [" & TypeName(cellVal) & "]"

                ' Proveri specijalne slučajeve
                If ws.Cells(i, j).HasFormula Then
                    cellInfo = cellInfo & " (FORMULA)"
                ElseIf VarType(cellVal) = vbString And IsNumeric(Replace(Replace(Trim(cellVal), " ", ""), ",", ".")) Then
                    cellInfo = cellInfo & " ⚠️ BROJ KAO TEKST!"
                End If

                result = result & cellInfo & vbCrLf
            End If
        Next j
        result = result & vbCrLf
    Next i

    PrikaziPrimereIzSheeta = result
End Function

' Helper funkcija za dobijanje slova kolone
Function GetColumnLetterHelper(colNum As Long) As String
    Dim result As String
    Dim num As Long

    num = colNum
    Do While num > 0
        Dim remainder As Long
        remainder = (num - 1) Mod 26
        result = Chr(65 + remainder) & result
        num = (num - remainder) \ 26
    Loop

    GetColumnLetterHelper = result
End Function

' Brza provera jednog sheet-a
Sub ProveraJednogSheeta()
    Dim sheetName As String
    Dim ws As Worksheet

    sheetName = InputBox("Unesi tačno ime sheet-a koji ne radi:", "Provera Sheet-a")

    If sheetName = "" Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' ne postoji!", vbCritical
        Exit Sub
    End If

    Dim analiza As SheetAnalysis
    analiza = AnalizirajSheet(ws)

    Dim msg As String
    msg = "ANALIZA SHEET-A: " & sheetName & vbCrLf & vbCrLf
    msg = msg & "Status: " & analiza.status & vbCrLf
    msg = msg & "Problem: " & analiza.problem & vbCrLf & vbCrLf
    msg = msg & "Broj redova: " & analiza.brojRedova & vbCrLf
    msg = msg & "Numeričke vrednosti: " & analiza.brojNumerickihCelija & vbCrLf
    msg = msg & "Brojevi kao tekst: " & analiza.brojTekstualnihBrojeva & vbCrLf
    msg = msg & "Formule: " & analiza.brojFormula & vbCrLf
    msg = msg & "Kolone sa cenama: " & analiza.koloneSaCenama & vbCrLf & vbCrLf
    msg = msg & "=================" & vbCrLf & vbCrLf
    msg = msg & PrikaziPrimereIzSheeta(ws)

    ' Kreiraj text file sa analizom
    Dim wsTemp As Worksheet
    Set wsTemp = ThisWorkbook.Worksheets.Add
    wsTemp.Name = "ANALIZA_" & Left(sheetName, 20)
    wsTemp.Range("A1").Value = msg
    wsTemp.Columns("A:A").ColumnWidth = 100
    wsTemp.Range("A1").WrapText = False

    MsgBox "Analiza kreirana u novom sheet-u!", vbInformation
End Sub
