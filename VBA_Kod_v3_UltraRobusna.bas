' ============================================
' AŽURIRANJE CENA - VBA MACRO v3 (ULTRA ROBUSNA)
' ============================================
' Autor: Claude
' Verzija: 3.0
' Opis: Najrobustnija verzija - pokušava SVE načine
'       da detektuje i konvertuje brojeve
' ============================================

Option Explicit

' ============================================
' STRUKTURE PODATAKA (moraju biti na vrhu!)
' ============================================

Type UpdateResult
    brojAzuriranih As Long
    brojKonvertovanih As Long
    brojFormula As Long
    brojPreskocenih As Long
    zastita As Boolean
    detaljniLog As String
End Type

Type CellUpdateResult
    azurirano As Long
    konvertovano As Long
    formula As Long
    preskoceno As Long
End Type

' ============================================
' FUNKCIJE
' ============================================

' Glavna funkcija
Sub AzurirajCene()
    Dim koeficijent As Double
    Dim odgovor As String
    Dim ws As Worksheet
    Dim ukupno As Long
    Dim korisnickiOdabir As String
    Dim odabraniSheetovi As String
    Dim azurirajSve As Boolean
    Dim celanLista As String
    Dim izvestaj As String
    Dim brojProblema As Long
    Dim detaljniLog As String

    ' Prikupi listu svih sheet-ova
    celanLista = ""
    For Each ws In ThisWorkbook.Worksheets
        celanLista = celanLista & ws.Name & vbCrLf
    Next ws

    ' Pitaj korisnika
    korisnickiOdabir = MsgBox("Dostupni sheet-ovi:" & vbCrLf & vbCrLf & celanLista & vbCrLf & _
                               "Da li želiš da ažuriraš SVE sheet-ove?" & vbCrLf & vbCrLf & _
                               "Klikni DA za sve, NE za odabrane sheet-ove", _
                               vbYesNoCancel + vbQuestion, "Izbor sheet-ova")

    If korisnickiOdabir = vbCancel Then Exit Sub

    azurirajSve = (korisnickiOdabir = vbYes)

    If Not azurirajSve Then
        odabraniSheetovi = InputBox("Unesi imena sheet-ova odvojene zarezom:", "Odabir sheet-ova")
        If odabraniSheetovi = "" Then Exit Sub
    End If

    ' Pitaj za koeficijent
    odgovor = InputBox("Unesi koeficijent za ažuriranje cena:" & vbCrLf & vbCrLf & _
                       "Primeri:" & vbCrLf & _
                       "  1.1  = povećanje za 10%" & vbCrLf & _
                       "  1.25 = povećanje za 25%" & vbCrLf & _
                       "  0.9  = smanjenje za 10%" & vbCrLf & _
                       "  0.8  = smanjenje za 20%" & vbCrLf & _
                       "  2.0  = udvostručenje", _
                       "Ažuriranje cena", "1.1")

    If odgovor = "" Then Exit Sub

    On Error Resume Next
    koeficijent = CDbl(Replace(odgovor, ",", "."))
    On Error GoTo 0

    If koeficijent <= 0 Then
        MsgBox "Neispravan koeficijent! Mora biti pozitivan broj.", vbCritical, "Greška"
        Exit Sub
    End If

    ' Potvrda
    Dim porukaPotvrde As String
    If azurirajSve Then
        porukaPotvrde = "Ažuriraću cene u SVIM sheet-ovima sa koeficijentom " & koeficijent & "."
    Else
        porukaPotvrde = "Ažuriraću cene u sheet-ovima: " & odabraniSheetovi & vbCrLf & _
                       "Koeficijent: " & koeficijent
    End If

    Dim potvrda As VbMsgBoxResult
    potvrda = MsgBox(porukaPotvrde & vbCrLf & vbCrLf & "Da li želiš da nastaviš?", _
                     vbYesNo + vbExclamation, "Potvrda")

    If potvrda <> vbYes Then Exit Sub

    ' Isključi osvežavanje
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ukupno = 0
    brojProblema = 0
    izvestaj = "IZVEŠTAJ O AŽURIRANJU (v3 Ultra):" & vbCrLf & vbCrLf
    detaljniLog = ""

    ' Prolaziš kroz sheet-ove
    For Each ws In ThisWorkbook.Worksheets
        Dim treba As Boolean
        treba = False

        If azurirajSve Then
            treba = True
        Else
            If InStr(1, odabraniSheetovi, ws.Name, vbTextCompare) > 0 Then
                treba = True
            End If
        End If

        If treba Then
            Dim result As UpdateResult
            result = AzurirajCeneUSheetUV3(ws, koeficijent)

            ukupno = ukupno + result.brojAzuriranih

            izvestaj = izvestaj & "✓ " & ws.Name & ": " & result.brojAzuriranih & " ćelija"

            If result.brojKonvertovanih > 0 Then
                izvestaj = izvestaj & " (+" & result.brojKonvertovanih & " konvertovano)"
            End If

            If result.brojFormula > 0 Then
                izvestaj = izvestaj & " [" & result.brojFormula & " formula preskočeno]"
            End If

            If result.brojPreskocenih > 0 Then
                izvestaj = izvestaj & " [" & result.brojPreskocenih & " preskočeno]"
            End If

            If result.zastita Then
                izvestaj = izvestaj & " [ZAŠTIĆEN]"
                brojProblema = brojProblema + 1
            End If

            If result.brojAzuriranih = 0 And Not result.zastita Then
                izvestaj = izvestaj & " ⚠️ NEMA PROMENA!"
                brojProblema = brojProblema + 1
            End If

            izvestaj = izvestaj & vbCrLf

            ' Dodaj u detaljni log
            If result.detaljniLog <> "" Then
                detaljniLog = detaljniLog & vbCrLf & "=== " & ws.Name & " ===" & vbCrLf
                detaljniLog = detaljniLog & result.detaljniLog & vbCrLf
            End If
        End If
    Next ws

    ' Uključi nazad osvežavanje
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Prikaži rezultat
    izvestaj = izvestaj & vbCrLf & "===================" & vbCrLf
    izvestaj = izvestaj & "UKUPNO: " & ukupno & " ćelija ažurirano" & vbCrLf

    If brojProblema > 0 Then
        izvestaj = izvestaj & vbCrLf & "⚠️ UPOZORENJE: " & brojProblema & " sheet(ova) sa problemima!" & vbCrLf
        izvestaj = izvestaj & vbCrLf & "Klikni OK da vidiš detaljni log..."
    End If

    MsgBox izvestaj, vbInformation, "Ažuriranje završeno"

    ' Ako ima problema, prikaži detaljni log
    If brojProblema > 0 And detaljniLog <> "" Then
        Dim odgovor2 As VbMsgBoxResult
        odgovor2 = MsgBox("Da li želiš da vidiš DETALJNI LOG problematičnih sheet-ova?", vbYesNo + vbQuestion)

        If odgovor2 = vbYes Then
            Dim wsLog As Worksheet
            On Error Resume Next
            Application.DisplayAlerts = False
            ThisWorkbook.Worksheets("DETALJNI_LOG").Delete
            Application.DisplayAlerts = True
            On Error GoTo 0

            Set wsLog = ThisWorkbook.Worksheets.Add
            wsLog.Name = "DETALJNI_LOG"
            wsLog.Range("A1").Value = detaljniLog
            wsLog.Columns("A:A").ColumnWidth = 120
            wsLog.Activate

            MsgBox "Detaljni log kreiran u sheet-u 'DETALJNI_LOG'", vbInformation
        End If
    End If
End Sub

' Poboljšana funkcija v3
Function AzurirajCeneUSheetUV3(ws As Worksheet, koeficijent As Double) As UpdateResult
    Dim cell As Range
    Dim result As UpdateResult
    Dim i As Long, j As Long
    Dim maxRow As Long, maxCol As Long

    result.brojAzuriranih = 0
    result.brojKonvertovanih = 0
    result.brojFormula = 0
    result.brojPreskocenih = 0
    result.zastita = False
    result.detaljniLog = ""

    ' Proveri zaštitu
    If ws.ProtectContents Then
        result.zastita = True
        result.detaljniLog = "Sheet je zaštićen - ne mogu ažurirati!" & vbCrLf
        AzurirajCeneUSheetUV3 = result
        Exit Function
    End If

    ' Nađi granice
    On Error Resume Next
    maxRow = ws.UsedRange.Rows.Count
    maxCol = ws.UsedRange.Columns.Count
    On Error GoTo 0

    If maxRow <= 1 Then
        result.detaljniLog = "Sheet je prazan ili ima samo header!" & vbCrLf
        AzurirajCeneUSheetUV3 = result
        Exit Function
    End If

    ' Detektuj kolone sa cenama
    Dim koloneSaCenama As Collection
    Set koloneSaCenama = New Collection

    For i = 1 To maxCol
        Dim header As String
        On Error Resume Next
        header = LCase(Trim(ws.Cells(1, i).Value))
        On Error GoTo 0

        If InStr(header, "cena") > 0 Or InStr(header, "cijena") > 0 Or _
           InStr(header, "price") > 0 Or InStr(header, "cost") > 0 Or _
           InStr(header, "iznos") > 0 Or InStr(header, "vrednost") > 0 Or _
           InStr(header, "vrijednost") > 0 Then
            koloneSaCenama.Add i
        End If
    Next i

    ' Ako nema kolona sa cenama, ažuriraj SVE numeričke
    Dim obradjujeSveKolone As Boolean
    obradjujeSveKolone = (koloneSaCenama.Count = 0)

    If obradjujeSveKolone Then
        result.detaljniLog = result.detaljniLog & "Nisam pronašao kolone sa 'cena' u headeru - ažuriram sve numeričke!" & vbCrLf
    Else
        result.detaljniLog = result.detaljniLog & "Detektovane kolone sa cenama: "
        Dim col As Variant
        For Each col In koloneSaCenama
            result.detaljniLog = result.detaljniLog & GetColLetter(col) & " "
        Next col
        result.detaljniLog = result.detaljniLog & vbCrLf
    End If

    ' Obrađuj svaku ćeliju
    For i = 2 To maxRow ' Preskačemo header
        For j = 1 To maxCol
            ' Ako imamo specifične kolone, preskači one koje nisu u listi
            If Not obradjujeSveKolone Then
                Dim jeUListi As Boolean
                jeUListi = False
                For Each col In koloneSaCenama
                    If col = j Then
                        jeUListi = True
                        Exit For
                    End If
                Next col

                If Not jeUListi Then GoTo NextCell
            End If

            On Error Resume Next
            Set cell = ws.Cells(i, j)
            On Error GoTo 0

            If Not cell Is Nothing Then
                Dim cellResult As CellUpdateResult
                cellResult = AzurirajCelijuV3(cell, koeficijent)

                result.brojAzuriranih = result.brojAzuriranih + cellResult.azurirano
                result.brojKonvertovanih = result.brojKonvertovanih + cellResult.konvertovano
                result.brojFormula = result.brojFormula + cellResult.formula
                result.brojPreskocenih = result.brojPreskocenih + cellResult.preskoceno
            End If

NextCell:
        Next j
    Next i

    ' Ako nema promena, dodaj u log
    If result.brojAzuriranih = 0 Then
        result.detaljniLog = result.detaljniLog & "⚠️ NEMA AŽURIRANIH ĆELIJA!" & vbCrLf
        result.detaljniLog = result.detaljniLog & "  - Numeričke vrednosti: 0" & vbCrLf
        result.detaljniLog = result.detaljniLog & "  - Preskočeno (formule/tekst): " & result.brojPreskocenih & vbCrLf
        result.detaljniLog = result.detaljniLog & vbCrLf & "PROVERI:" & vbCrLf
        result.detaljniLog = result.detaljniLog & "  1. Da li su brojevi sačuvani kao TEKST?" & vbCrLf
        result.detaljniLog = result.detaljniLog & "  2. Da li su sve vrednosti FORMULE?" & vbCrLf
        result.detaljniLog = result.detaljniLog & "  3. Da li sheet uopšte ima numeričke vrednosti?" & vbCrLf
        result.detaljniLog = result.detaljniLog & vbCrLf & "Koristi 'DiagnostikujDetaljno' za više informacija!" & vbCrLf
    End If

    AzurirajCeneUSheetUV3 = result
End Function

' Funkcija za ažuriranje pojedinačne ćelije - Ultra robusna verzija
Function AzurirajCelijuV3(cell As Range, koeficijent As Double) As CellUpdateResult
    Dim result As CellUpdateResult
    Dim numVal As Double

    result.azurirano = 0
    result.konvertovano = 0
    result.formula = 0
    result.preskoceno = 0

    On Error GoTo ErrorHandler

    ' Preskači prazne
    If IsEmpty(cell.Value) Or cell.Value = "" Then
        GoTo Cleanup
    End If

    ' Preskači formule
    If cell.HasFormula Then
        result.formula = 1
        GoTo Cleanup
    End If

    ' Pokušaj 1: Direktan broj
    If IsNumeric(cell.Value) And VarType(cell.Value) <> vbString Then
        numVal = CDbl(cell.Value)
        cell.Value = numVal * koeficijent
        result.azurirano = 1
        GoTo Cleanup
    End If

    ' Pokušaj 2: String koji je broj
    If VarType(cell.Value) = vbString Then
        Dim cleanVal As String
        cleanVal = Trim(CStr(cell.Value))

        ' Ukloni sve razmake
        cleanVal = Replace(cleanVal, " ", "")
        cleanVal = Replace(cleanVal, Chr(160), "") ' Non-breaking space

        ' Zameni zarez sa tačkom
        cleanVal = Replace(cleanVal, ",", ".")

        ' Ukloni valutne simbole ako postoje
        cleanVal = Replace(cleanVal, "$", "")
        cleanVal = Replace(cleanVal, "€", "")
        cleanVal = Replace(cleanVal, "din", "")
        cleanVal = Replace(cleanVal, "RSD", "")
        cleanVal = Replace(cleanVal, "EUR", "")

        ' Pokušaj konverziju
        If IsNumeric(cleanVal) Then
            numVal = CDbl(cleanVal)
            cell.Value = numVal * koeficijent
            result.azurirano = 1
            result.konvertovano = 1
            GoTo Cleanup
        End If
    End If

    ' Pokušaj 3: Val funkcija (za posebne formate)
    On Error Resume Next
    numVal = Val(CStr(cell.Value))
    If numVal <> 0 Then
        cell.Value = numVal * koeficijent
        result.azurirano = 1
        result.konvertovano = 1
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler

    ' Ako ništa nije uspelo, preskači
    result.preskoceno = 1

Cleanup:
    AzurirajCelijuV3 = result
    Exit Function

ErrorHandler:
    result.azurirano = 0
    result.konvertovano = 0
    result.formula = 0
    result.preskoceno = 1
    Resume Cleanup
End Function

' Helper funkcija
Function GetColLetter(colNum As Long) As String
    Dim result As String
    Dim num As Long

    num = colNum
    Do While num > 0
        Dim remainder As Long
        remainder = (num - 1) Mod 26
        result = Chr(65 + remainder) & result
        num = (num - remainder) \ 26
    Loop

    GetColLetter = result
End Function

' Backup funkcija
Sub NapraviBackup()
    Dim originalPath As String
    Dim backupPath As String
    Dim fso As Object

    If ThisWorkbook.Path = "" Then
        MsgBox "Molim te prvo sačuvaj fajl!", vbExclamation
        Exit Sub
    End If

    originalPath = ThisWorkbook.FullName
    backupPath = Replace(originalPath, ".xlsm", "_backup_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsm")
    backupPath = Replace(backupPath, ".xlsx", "_backup_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx")

    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    ThisWorkbook.Save
    fso.CopyFile originalPath, backupPath

    If Err.Number = 0 Then
        MsgBox "Backup kreiran uspešno:" & vbCrLf & backupPath, vbInformation
    Else
        MsgBox "Greška pri kreiranju backup-a: " & Err.Description, vbCritical
    End If
    On Error GoTo 0
End Sub
