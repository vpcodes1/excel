' ============================================
' DIAGNOSTIC ALAT - Provera Excel fajla
' ============================================
' Ova skripta analizira sheet-ove i prikazuje
' detalje koji mogu pomoći u debugging-u
' ============================================

Option Explicit

Sub DiagnostikujFajl()
    Dim ws As Worksheet
    Dim izvestaj As String
    Dim brojNumerickihCelija As Long
    Dim brojTekstualnihCelija As Long
    Dim brojFormula As Long
    Dim i As Long, j As Long
    Dim cell As Range

    izvestaj = "========================================" & vbCrLf
    izvestaj = izvestaj & "DIAGNOSTIC IZVEŠTAJ" & vbCrLf
    izvestaj = izvestaj & "========================================" & vbCrLf & vbCrLf

    For Each ws In ThisWorkbook.Worksheets
        izvestaj = izvestaj & "📄 SHEET: " & ws.Name & vbCrLf
        izvestaj = izvestaj & "----------------------------------------" & vbCrLf

        ' Proveri da li je zaštićen
        If ws.ProtectContents Then
            izvestaj = izvestaj & "⚠️  ZAŠTIĆEN: Da" & vbCrLf
        Else
            izvestaj = izvestaj & "✅ ZAŠTIĆEN: Ne" & vbCrLf
        End If

        ' Dimenzije
        izvestaj = izvestaj & "📊 Redovi: " & ws.UsedRange.Rows.Count & vbCrLf
        izvestaj = izvestaj & "📊 Kolone: " & ws.UsedRange.Columns.Count & vbCrLf

        ' Analiziraj prvi red (header)
        izvestaj = izvestaj & "📋 HEADER (Red 1):" & vbCrLf
        For i = 1 To WorksheetFunction.Min(10, ws.UsedRange.Columns.Count)
            Dim headerVal As String
            headerVal = ws.Cells(1, i).Value
            If headerVal <> "" Then
                izvestaj = izvestaj & "   " & GetColumnLetter(i) & ": " & headerVal & vbCrLf
            End If
        Next i

        ' Brojanje tipova podataka
        brojNumerickihCelija = 0
        brojTekstualnihCelija = 0
        brojFormula = 0

        For i = 2 To WorksheetFunction.Min(100, ws.UsedRange.Rows.Count)
            For j = 1 To ws.UsedRange.Columns.Count
                Set cell = ws.Cells(i, j)

                If cell.HasFormula Then
                    brojFormula = brojFormula + 1
                ElseIf IsNumeric(cell.Value) And cell.Value <> "" Then
                    brojNumerickihCelija = brojNumerickihCelija + 1
                ElseIf cell.Value <> "" Then
                    ' Proveri da li je "broj kao tekst"
                    If IsNumeric(Replace(Replace(cell.Value, ".", ""), ",", ".")) Then
                        brojTekstualnihCelija = brojTekstualnihCelija + 1
                    End If
                End If
            Next j
        Next i

        izvestaj = izvestaj & "🔢 Numeričke vrednosti: " & brojNumerickihCelija & vbCrLf
        izvestaj = izvestaj & "📝 Tekstualne brojeve: " & brojTekstualnihCelija & vbCrLf
        izvestaj = izvestaj & "📐 Formule: " & brojFormula & vbCrLf

        ' Proveri da li ima kolone sa cenama
        Dim koloneSaCenama As String
        koloneSaCenama = ""
        For i = 1 To ws.UsedRange.Columns.Count
            Dim header As String
            header = LCase(Trim(ws.Cells(1, i).Value))

            If InStr(header, "cena") > 0 Or InStr(header, "cijena") > 0 Or _
               InStr(header, "price") > 0 Or InStr(header, "cost") > 0 Or _
               InStr(header, "iznos") > 0 Or InStr(header, "vrednost") > 0 Then
                If koloneSaCenama <> "" Then koloneSaCenama = koloneSaCenama & ", "
                koloneSaCenama = koloneSaCenama & GetColumnLetter(i)
            End If
        Next i

        If koloneSaCenama <> "" Then
            izvestaj = izvestaj & "💰 Detektovane kolone sa cenama: " & koloneSaCenama & vbCrLf
        Else
            izvestaj = izvestaj & "⚠️  Nisu detektovane kolone sa cenama (ažuriraće SVE numeričke)" & vbCrLf
        End If

        ' Primeri podataka iz prvog reda
        izvestaj = izvestaj & "📋 PRIMER PODATAKA (Red 2):" & vbCrLf
        For i = 1 To WorksheetFunction.Min(5, ws.UsedRange.Columns.Count)
            Dim val As Variant
            val = ws.Cells(2, i).Value
            If val <> "" Then
                izvestaj = izvestaj & "   " & GetColumnLetter(i) & ": " & val
                izvestaj = izvestaj & " [" & TypeName(val) & "]"
                If ws.Cells(2, i).HasFormula Then
                    izvestaj = izvestaj & " (FORMULA)"
                End If
                izvestaj = izvestaj & vbCrLf
            End If
        Next i

        izvestaj = izvestaj & vbCrLf & vbCrLf
    Next ws

    ' Prikaži izveštaj
    Dim frmReport As Object
    On Error Resume Next

    ' Pokušaj da prikaže u message box-u (limitirano na 1024 karaktera)
    If Len(izvestaj) < 1024 Then
        MsgBox izvestaj, vbInformation, "Diagnostic Izveštaj"
    Else
        ' Ako je predugačak, sačuvaj u novi sheet
        Dim wsReport As Worksheet
        On Error Resume Next
        Set wsReport = ThisWorkbook.Worksheets("Diagnostic_Report")
        If wsReport Is Nothing Then
            Set wsReport = ThisWorkbook.Worksheets.Add
            wsReport.Name = "Diagnostic_Report"
        Else
            wsReport.Cells.Clear
        End If
        On Error GoTo 0

        wsReport.Range("A1").Value = izvestaj
        wsReport.Range("A1").WrapText = False
        wsReport.Columns("A:A").ColumnWidth = 100
        wsReport.Activate

        MsgBox "Izveštaj je sačuvan u sheet-u 'Diagnostic_Report'", vbInformation, "Diagnostic završen"
    End If
End Sub

' Pomoćna funkcija za dobijanje slova kolone
Function GetColumnLetter(colNum As Long) As String
    Dim result As String
    Dim num As Long

    num = colNum
    Do While num > 0
        Dim remainder As Long
        remainder = (num - 1) Mod 26
        result = Chr(65 + remainder) & result
        num = (num - remainder) \ 26
    Loop

    GetColumnLetter = result
End Function

' Dodatna funkcija - Testiranje na jednom sheet-u
Sub TestirajJedanSheet()
    Dim sheetName As String
    Dim ws As Worksheet
    Dim koeficijent As Double
    Dim odgovor As String

    ' Pitaj korisnika koji sheet da testira
    sheetName = InputBox("Unesi ime sheet-a koji želiš da testiraš:", "Test Sheet-a")

    If sheetName = "" Then Exit Sub

    ' Pronađi sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' ne postoji!", vbCritical
        Exit Sub
    End If

    ' Pitaj za koeficijent
    odgovor = InputBox("Unesi koeficijent za test:", "Koeficijent", "1.1")

    If odgovor = "" Then Exit Sub

    On Error Resume Next
    koeficijent = CDbl(Replace(odgovor, ",", "."))
    On Error GoTo 0

    If koeficijent <= 0 Then
        MsgBox "Neispravan koeficijent!", vbCritical
        Exit Sub
    End If

    ' Testiraj
    MsgBox "Testiram sheet '" & sheetName & "' sa koeficijentom " & koeficijent & vbCrLf & vbCrLf & _
           "Proverite Debug panel za detalje.", vbInformation

    Dim brojAzuriranih As Long
    brojAzuriranih = TestAzurirajCeneUSheetU(ws, koeficijent)

    MsgBox "Test završen!" & vbCrLf & vbCrLf & _
           "Ažurirano ćelija: " & brojAzuriranih, vbInformation
End Sub

Function TestAzurirajCeneUSheetU(ws As Worksheet, koeficijent As Double) As Long
    Dim cell As Range
    Dim brojAzuriranih As Long
    Dim i As Long
    Dim izvestaj As String

    brojAzuriranih = 0
    izvestaj = "TEST REZULTATI ZA: " & ws.Name & vbCrLf & vbCrLf

    ' Prolazi kroz sve ćelije
    For Each cell In ws.UsedRange
        If cell.Row > 1 Then ' Preskačemo header
            ' Debug info
            Dim tipPodatka As String
            tipPodatka = TypeName(cell.Value)

            If cell.HasFormula Then
                izvestaj = izvestaj & cell.Address & ": FORMULA - " & cell.Formula & vbCrLf
            ElseIf IsNumeric(cell.Value) And cell.Value <> "" Then
                Dim starVal As Double
                starVal = cell.Value
                ' cell.Value = cell.Value * koeficijent ' Zakomentarisano jer je test
                brojAzuriranih = brojAzuriranih + 1
                izvestaj = izvestaj & cell.Address & ": " & starVal & " → " & (starVal * koeficijent) & " [OK]" & vbCrLf
            ElseIf cell.Value <> "" Then
                ' Možda je broj kao tekst?
                Dim testVal As String
                testVal = Replace(Replace(cell.Value, ".", ""), ",", ".")
                If IsNumeric(testVal) Then
                    izvestaj = izvestaj & cell.Address & ": '" & cell.Value & "' [BROJ KAO TEKST - PROBLEM!]" & vbCrLf
                End If
            End If

            If brojAzuriranih >= 20 Then Exit For ' Ograniči na prvih 20
        End If
    Next cell

    Debug.Print izvestaj

    TestAzurirajCeneUSheetU = brojAzuriranih
End Function
