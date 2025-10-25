' ============================================
' AŽURIRANJE CENA - VBA MACRO
' ============================================
' Autor: Claude
' Opis: Automatski ažurira cene u svim sheet-ovima
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
            brojAzuriranih = AzurirajCeneUSheetU(ws, koeficijent)
            ukupno = ukupno + brojAzuriranih
        End If
    Next ws

    ' Uključi nazad osvežavanje
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Prikaži rezultat
    MsgBox "Uspešno ažurirano!" & vbCrLf & vbCrLf & _
           "Ukupno ažuriranih ćelija: " & ukupno, _
           vbInformation, "Završeno"
End Sub

' Pomoćna funkcija koja ažurira cene u jednom sheet-u
Function AzurirajCeneUSheetU(ws As Worksheet, koeficijent As Double) As Long
    Dim cell As Range
    Dim kolone As Collection
    Dim col As Variant
    Dim brojAzuriranih As Long
    Dim i As Long
    Dim poslednjiredRed As Long
    Dim poslednjaKolona As Long

    brojAzuriranih = 0

    ' Nađi poslednji red i kolonu
    poslednjiredRed = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    poslednjaKolona = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If poslednjiredRed < 2 Then
        ' Nema podataka
        AzurirajCeneUSheetU = 0
        Exit Function
    End If

    ' Detektuj kolone sa cenama
    Set kolone = New Collection

    For i = 1 To poslednjaKolona
        Dim header As String
        header = LCase(Trim(ws.Cells(1, i).Value))

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
                If IsNumeric(cell.Value) And cell.Value <> "" Then
                    cell.Value = cell.Value * koeficijent
                    brojAzuriranih = brojAzuriranih + 1
                End If
            End If
        Next cell
    Else
        ' Ažuriraj samo kolone sa cenama
        For Each col In kolone
            For i = 2 To poslednjiredRed ' Počinjemo od reda 2 (preskačemo header)
                Set cell = ws.Cells(i, col)
                If IsNumeric(cell.Value) And cell.Value <> "" Then
                    cell.Value = cell.Value * koeficijent
                    brojAzuriranih = brojAzuriranih + 1
                End If
            Next i
        Next col
    End If

    AzurirajCeneUSheetU = brojAzuriranih
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
