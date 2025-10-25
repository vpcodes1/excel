# 📝 Kako dodati VBA Macro u Excel

Ovo su detaljne instrukcije kako da dodaš automatizaciju za ažuriranje cena direktno u tvoj Excel fajl.

## 🎯 Prednosti VBA rešenja

✅ **Bez instalacije** - Radi u svakom Excel-u
✅ **Jednostavno** - Samo klikneš dugme
✅ **Za ceo tim** - Svi mogu da koriste bez tehničkog znanja
✅ **Bezbedno** - Ne menja originalni fajl dok ne potvrdis

---

## 📋 Korak po korak

### 1️⃣ Otvori Excel fajl

Otvori Excel fajl u kome želiš da ažuriraš cene.

### 2️⃣ Omogući Developer tab (ako nije već omogućen)

1. Klikni na **File** → **Options** (Opcije)
2. U levom meniju izaberi **Customize Ribbon** (Prilagodi Ribbon)
3. Na desnoj strani **štikliraj** opciju **Developer** (Programer)
4. Klikni **OK**

### 3️⃣ Otvori VBA Editor

1. Klikni na **Developer** tab u ribbon-u
2. Klikni na **Visual Basic** (ili pritisni **Alt + F11**)

### 4️⃣ Dodaj novi modul

1. U VBA Editor-u, desni klik na **VBAProject (tvoj_fajl.xlsx)**
2. Izaberi **Insert** → **Module**
3. Otvoriće se novi prazan modul

### 5️⃣ Kopiraj VBA kod

1. Otvori fajl **VBA_Kod.bas** iz ovog projekta
2. **Kopiraj sav kod** (Ctrl+A, Ctrl+C)
3. **Nalepi kod** u novi modul u VBA Editor-u (Ctrl+V)

### 6️⃣ Dodaj dugme u Excel (opciono, ali preporučeno)

**Opcija A: Quick Access Toolbar** (najbrže)

1. Zatvori VBA Editor (Alt + Q ili X)
2. Desni klik na **Quick Access Toolbar** (mala traka iznad)
3. Izaberi **Customize Quick Access Toolbar**
4. Iz dropdown-a izaberi **Macros**
5. Odaberi **AzurirajCene**
6. Klikni **Add**
7. Klikni **OK**

**Opcija B: Dugme na sheet-u** (lepše izgleda)

1. U Excel-u, idi na **Developer** tab
2. Klikni **Insert** → **Button (Form Control)**
3. Nacrtaj dugme gde god želiš
4. U dijalogu izaberi **AzurirajCene**
5. Klikni **OK**
6. Desni klik na dugme → **Edit Text** i promeni tekst u "Ažuriraj Cene"

### 7️⃣ Sačuvaj fajl kao Macro-Enabled

⚠️ **VAŽNO**: Moraš sačuvati fajl sa podrškom za makroe!

1. **File** → **Save As** (Sačuvaj kao)
2. U **Save as type** izaberi: **Excel Macro-Enabled Workbook (*.xlsm)**
3. Klikni **Save**

---

## 🚀 Kako koristiti

### Osnovna upotreba:

1. **Klikni** na dugme koje si napravio (ili pokreni macro iz Developer → Macros)
2. **Odaberi** da li želiš SVE sheet-ove ili samo neke
3. **Unesi koeficijent** (npr. 1.1 za +10%, 0.9 za -10%)
4. **Potvrdi** i gotovo!

### Primeri koeficijenata:

| Šta želiš | Unesi |
|-----------|-------|
| Povećaj za 10% | **1.1** |
| Povećaj za 25% | **1.25** |
| Smanji za 10% | **0.9** |
| Smanji za 20% | **0.8** |
| Udvostručenje | **2.0** |
| Prepolovljenje | **0.5** |

---

## 🛡️ Sigurnost

### Backup funkcija

Makro uključuje i funkciju za backup! Pre nego što ažuriraš cene možeš da:

1. Otvori **Developer** → **Macros**
2. Odaberi **NapraviBackup**
3. Klikni **Run**
4. Backup fajl će biti kreiran sa datumom i vremenom

### Automatski backup pre ažuriranja

Možeš dodati automatski backup tako što ćeš dodati ovaj red na početku `AzurirajCene` funkcije:

```vba
Call NapraviBackup
```

---

## ❓ Česta pitanja

**P: Da li mogu da podelim ovaj fajl sa kolegama?**
O: Da! Samo pošalji .xlsm fajl i makro će raditi kod njih.

**P: Excel mi govori da su makroi onemogućeni?**
O: Klikni na **Enable Content** (Omogući sadržaj) u žutoj traci na vrhu.

**P: Šta ako ne želim da makro detektuje kolone automatski?**
O: Možeš modifikovati kod da uvek ažurira specifične kolone (npr. samo C i D).

**P: Mogu li da vratim izmene?**
O: Koristi **Ctrl+Z** odmah nakon izvršenja, ili koristi backup funkciju pre ažuriranja.

**P: Da li mogu da koristim makro na Mac-u?**
O: Da, VBA radi i na Mac verziji Excel-a.

**P: Kako da uklonim makro?**
O: U VBA Editor-u (Alt+F11), desni klik na modul → Remove.

---

## 💡 Saveti

1. **Prvo testiraj** na kopiji fajla
2. **Pravi backup** pre važnih ažuriranja
3. **Proveri rezultate** nakon ažuriranja
4. **Podeli sa timom** - Šalji .xlsm fajl direktno

---

## 🔧 Prilagođavanje

Ako želiš da prilagodiš kako makro radi, možeš modifikovati kod:

### Ažuriraj uvek samo određene kolone (npr. C i D)

Nađi funkciju `AzurirajCeneUSheetU` i izmeni deo gde se detektuju kolone.

### Dodaj keyboard shortcut

1. U Excel-u: **Developer** → **Macros**
2. Odaberi **AzurirajCene**
3. Klikni **Options**
4. Unesi shortcut (npr. **Ctrl+Shift+P**)

---

## 📞 Podrška

Ako imaš problema ili pitanja, konsultuj se sa IT timom ili osobom koja je napravila ovaj makro.
