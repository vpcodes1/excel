' ============================================
' AŽURIRANJE CENA - VBA MACRO v2 (POBOLJŠANA)
' ============================================
' Autor: Claude
' Verzija: 2.0
' Opis: Automatski ažurira cene u svim sheet-ovima
'       sa boljom podrškom za različite formate
' ============================================

Option Explicit

' Glavna funkcija - pokreće se kada korisnik klikne dugme
Sub AzurirajCene()
    Dim koeficijent As Double
    Dim odgovor As String
    Dim ws As Worksheet
    Dim brojAzuriranih As Long
    Dim ukupno As Long
    Dim korisnickiOdabir As String
    Dim odabraniSheetovi As String
    Dim azurirajSve As Boolean
    Dim celanLista As String
    Dim izvestaj As String
    Dim brojProblema As Long

    ' Prikupi listu svih sheet-ova
    celanLista = ""
    For Each ws In ThisWorkbook.Worksheets
        celanLista = celanLista & ws.Name & vbCrLf
    Next ws

    ' Pitaj korisnika da li želi sve sheet-ove ili odabrane
    korisnickiOdabir = MsgBox("Dostupni sheet-ovi:" & vbCrLf & vbCrLf & celanLista & vbCrLf & _
                               "Da li želiš da ažuriraš SVE sheet-ove?" & vbCrLf & vbCrLf & _
                               "Klikni DA za sve, NE za odabrane sheet-ove", _
                               vbYesNoCancel + vbQuestion, "Izbor sheet-ova")

    If korisnickiOdabir = vbCancel Then
        Exit Sub
    End If

    azurirajSve = (korisnickiOdabir = vbYes)

    If Not azurirajSve Then
        odabraniSheetovi = InputBox("Unesi imena sheet-ova koje želiš da ažuriraš, odvojene zarezom:" & vbCrLf & _
                                     "Primer: Januar 2024, Februar 2024, Mart 2024", _
                                     "Odabir sheet-ova")
        If odabraniSheetovi = "" Then
            Exit Sub
        End If
    End If

    ' Pitaj korisnika za koeficijent
    odgovor = InputBox("Unesi koeficijent za ažuriranje cena:" & vbCrLf & vbCrLf & _
                       "Primeri:" & vbCrLf & _
                       "  1.1  = povećanje za 10%" & vbCrLf & _
                       "  1.25 = povećanje za 25%" & vbCrLf & _
                       "  0.9  = smanjenje za 10%" & vbCrLf & _
                       "  0.8  = smanjenje za 20%" & vbCrLf & _
                       "  2.0  = udvostručenje", _
                       "Ažuriranje cena", "1.1")

    ' Proveri da li je korisnik otkazao
    If odgovor = "" Then
        Exit Sub
    End If

    ' Konvertuj u broj
    On Error Resume Next
    koeficijent = CDbl(Replace(odgovor, ",", "."))
    On Error GoTo 0

    ' Validacija
    If koeficijent <= 0 Then
        MsgBox "Neispravan koeficijent! Mora biti pozitivan broj.", vbCritical, "Greška"
        Exit Sub
    End If

    ' Potvrda pre izvršenja
    Dim porukaPotvrde As String
    If azurirajSve Then
        porukaPotvrde = "Ažuriraću cene u SVIM sheet-ovima sa koeficijentom " & koeficijent & "." & vbCrLf & vbCrLf & _
                       "Da li želiš da nastaviš?"
    Else
        porukaPotvrde = "Ažuriraću cene u sledećim sheet-ovima: " & odabraniSheetovi & vbCrLf & _
                       "Koeficijent: " & koeficijent & vbCrLf & vbCrLf & _
                       "Da li želiš da nastaviš?"
    End If

    Dim potvrda As VbMsgBoxResult
    potvrda = MsgBox(porukaPotvrde, vbYesNo + vbExclamation, "Potvrda")

    If potvrda <> vbYes Then
        Exit Sub
    End If

    ' Isključi osvežavanje ekrana za brži rad
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ukupno = 0
    brojProblema = 0
    izvestaj = "IZVEŠTAJ O AŽURIRANJU:" & vbCrLf & vbCrLf

    ' Prolaziš kroz sheet-ove
    For Each ws In ThisWorkbook.Worksheets
        ' Proveri da li treba ažurirati ovaj sheet
        Dim treba As Boolean
        treba = False

        If azurirajSve Then
            treba = True
        Else
            ' Proveri da li je ovaj sheet u listi
            If InStr(1, odabraniSheetovi, ws.Name, vbTextCompare) > 0 Then
                treba = True
            End If
        End If

        If treba Then
            Dim result As UpdateResult
            result = AzurirajCeneUSheetU(ws, koeficijent)

            ukupno = ukupno + result.brojAzuriranih

            ' Dodaj u izveštaj
            izvestaj = izvestaj & "✓ " & ws.Name & ": " & result.brojAzuriranih & " ćelija"

            If result.brojKonvertovanih > 0 Then
                izvestaj = izvestaj & " (+" & result.brojKonvertovanih & " konvertovano iz teksta)"
            End If

            If result.brojFormula > 0 Then
                izvestaj = izvestaj & " [" & result.brojFormula & " formula preskočeno]"
            End If

            If result.zastita Then
                izvestaj = izvestaj & " [ZAŠTIĆEN - PRESKOČEN]"
                brojProblema = brojProblema + 1
            End If

            izvestaj = izvestaj & vbCrLf
        End If
    Next ws

    ' Uključi nazad osvežavanje
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Prikaži rezultat
    izvestaj = izvestaj & vbCrLf & "===================" & vbCrLf
    izvestaj = izvestaj & "UKUPNO: " & ukupno & " ćelija ažurirano"

    If brojProblema > 0 Then
        izvestaj = izvestaj & vbCrLf & "UPOZORENJE: " & brojProblema & " sheet(ova) sa problemima"
    End If

    MsgBox izvestaj, vbInformation, "Ažuriranje završeno"
End Sub

' Struktura za rezultate ažuriranja
Type UpdateResult
    brojAzuriranih As Long
    brojKonvertovanih As Long
    brojFormula As Long
    zastita As Boolean
End Type

' Poboljšana funkcija koja ažurira cene u jednom sheet-u
Function AzurirajCeneUSheetU(ws As Worksheet, koeficijent As Double) As UpdateResult
    Dim cell As Range
    Dim kolone As Collection
    Dim col As Variant
    Dim result As UpdateResult
    Dim i As Long
    Dim poslednjiredRed As Long
    Dim poslednjaKolona As Long

    result.brojAzuriranih = 0
    result.brojKonvertovanih = 0
    result.brojFormula = 0
    result.zastita = False

    ' Proveri da li je sheet zaštićen
    If ws.ProtectContents Then
        result.zastita = True
        AzurirajCeneUSheetU = result
        Exit Function
    End If

    ' Nađi poslednji red i kolonu
    On Error Resume Next
    poslednjiredRed = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    poslednjaKolona = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    On Error GoTo 0

    If poslednjiredRed < 2 Then
        ' Nema podataka
        AzurirajCeneUSheetU = result
        Exit Function
    End If

    ' Detektuj kolone sa cenama
    Set kolone = New Collection

    For i = 1 To poslednjaKolona
        Dim header As String
        header = ""
        On Error Resume Next
        header = LCase(Trim(ws.Cells(1, i).Value))
        On Error GoTo 0

        ' Proveri da li header sadrži reči koje ukazuju na cenu
        If InStr(header, "cena") > 0 Or _
           InStr(header, "cijena") > 0 Or _
           InStr(header, "price") > 0 Or _
           InStr(header, "cost") > 0 Or _
           InStr(header, "iznos") > 0 Or _
           InStr(header, "vrednost") > 0 Or _
           InStr(header, "vrijednost") > 0 Then
            kolone.Add i
        End If
    Next i

    ' Ako nije pronađena nijedna kolona sa cenom, ažuriraj sve numeričke ćelije
    If kolone.Count = 0 Then
        For Each cell In ws.UsedRange
            If cell.Row > 1 Then ' Preskačemo header
                Dim updateResult As CellUpdateResult
                updateResult = AzurirajCeliju(cell, koeficijent)

                result.brojAzuriranih = result.brojAzuriranih + updateResult.azurirano
                result.brojKonvertovanih = result.brojKonvertovanih + updateResult.konvertovano
                result.brojFormula = result.brojFormula + updateResult.formula
            End If
        Next cell
    Else
        ' Ažuriraj samo kolone sa cenama
        For Each col In kolone
            For i = 2 To poslednjiredRed ' Počinjemo od reda 2 (preskačemo header)
                On Error Resume Next
                Set cell = ws.Cells(i, col)
                On Error GoTo 0

                If Not cell Is Nothing Then
                    Dim updateResult2 As CellUpdateResult
                    updateResult2 = AzurirajCeliju(cell, koeficijent)

                    result.brojAzuriranih = result.brojAzuriranih + updateResult2.azurirano
                    result.brojKonvertovanih = result.brojKonvertovanih + updateResult2.konvertovano
                    result.brojFormula = result.brojFormula + updateResult2.formula
                End If
            Next i
        Next col
    End If

    AzurirajCeneUSheetU = result
End Function

' Struktura za rezultat ažuriranja pojedinačne ćelije
Type CellUpdateResult
    azurirano As Long
    konvertovano As Long
    formula As Long
End Type

' Funkcija za ažuriranje pojedinačne ćelije
Function AzurirajCeliju(cell As Range, koeficijent As Double) As CellUpdateResult
    Dim result As CellUpdateResult
    Dim numVal As Double

    result.azurirano = 0
    result.konvertovano = 0
    result.formula = 0

    On Error GoTo ErrorHandler

    ' Proveri da li je formula
    If cell.HasFormula Then
        result.formula = 1
        GoTo Cleanup
    End If

    ' Proveri da li je broj
    If IsNumeric(cell.Value) And cell.Value <> "" And Not IsEmpty(cell.Value) Then
        numVal = CDbl(cell.Value)
        cell.Value = numVal * koeficijent
        result.azurirano = 1
        GoTo Cleanup
    End If

    ' Pokušaj da konvertuješ iz teksta
    If VarType(cell.Value) = vbString And cell.Value <> "" Then
        Dim cleanedVal As String
        cleanedVal = Trim(cell.Value)

        ' Pokušaj različite formate
        cleanedVal = Replace(cleanedVal, " ", "") ' Ukloni razmake
        cleanedVal = Replace(cleanedVal, ",", ".") ' Zameni zarez sa tačkom

        If IsNumeric(cleanedVal) Then
            numVal = CDbl(cleanedVal)
            cell.Value = numVal * koeficijent
            result.azurirano = 1
            result.konvertovano = 1
        End If
    End If

Cleanup:
    AzurirajCeliju = result
    Exit Function

ErrorHandler:
    ' Ako je greška, samo preskoči
    result.azurirano = 0
    result.konvertovano = 0
    result.formula = 0
    AzurirajCeliju = result
End Function

' Dodatna funkcija - pravi backup trenutnog fajla
Sub NapraviBackup()
    Dim originalPath As String
    Dim backupPath As String
    Dim fso As Object

    ' Proveri da li je fajl sačuvan
    If ThisWorkbook.Path = "" Then
        MsgBox "Molim te prvo sačuvaj fajl!", vbExclamation
        Exit Sub
    End If

    originalPath = ThisWorkbook.FullName
    backupPath = Replace(originalPath, ".xlsm", "_backup_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsm")
    backupPath = Replace(backupPath, ".xlsx", "_backup_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx")

    ' Kopiraj fajl
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
