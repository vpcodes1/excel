# 📊 Excel Price Updater - Automatsko ažuriranje cena

Alat za automatsko ažuriranje cena u Excel fajlovima sa više sheet-ova. **Umesto da ručno menjaš cene na svakom sheet-u, ovaj alat to radi automatski za tebe!**

---

## 🎯 Tri načina korišćenja

Izaberi opciju koja ti najbolje odgovara:

| Opcija | Za koga | Prednosti | Instalacija |
|--------|---------|-----------|-------------|
| **🥇 Excel Macro (VBA)** | **Za sve, preporučeno za timove** | ✅ Bez instalacije<br>✅ Klikneš dugme i gotovo<br>✅ Radi u svakom Excel-u | Kopiraj kod u Excel |
| **🥈 GUI Aplikacija** | Za one koji ne vole komandnu liniju | ✅ Grafički interfejs<br>✅ Drag & drop | Potreban Python |
| **🥉 Komandna linija** | Za programere i IT | ✅ Automatizacija<br>✅ Fleksibilno | Potreban Python |

---

## 🥇 OPCIJA 1: Excel Macro (VBA) - **PREPORUČENO**

### ✨ Najprostije rešenje za ceo tim!

**Zašto ova opcija?**
- ❌ **Nema instalacije** - Radi u svakom Excel-u
- ❌ **Nije potreban Python** - Ni tehničko znanje
- ✅ **Samo klikneš dugme** i to je to!
- ✅ **Ceo tim može da koristi** bez problema

### 📝 Kako dodati macro u tvoj Excel fajl

👉 **[Detaljne instrukcije ovde: INSTRUKCIJE_VBA.md](INSTRUKCIJE_VBA.md)**

**Ukratko:**

1. Otvori tvoj Excel fajl
2. Pritisni **Alt + F11** (otvara VBA Editor)
3. **Insert** → **Module**
4. Kopiraj kod iz `VBA_Kod.bas` i nalepi
5. Dodaj dugme u Excel-u
6. Sačuvaj kao `.xlsm` fajl
7. **Gotovo!** Sada samo klikneš dugme

### 🚀 Kako koristiti macro

1. **Klikni dugme** "Ažuriraj Cene" u Excel-u
2. **Odaberi sheet-ove** (ili ostavi sve)
3. **Unesi koeficijent** (npr. 1.1 za +10%)
4. **Potvrdi** i gotovo!

**Jednostavno!** Pogledaj [INSTRUKCIJE_VBA.md](INSTRUKCIJE_VBA.md) za detaljnije korake sa slikama.

---

## 🥈 OPCIJA 2: GUI Aplikacija

### 🖥️ Grafički interfejs za one koji ne vole terminal

**Instalacija:**

```bash
pip install openpyxl
```

**Pokretanje:**

```bash
python gui_app.py
```

**Kako koristiti:**

1. Otvori aplikaciju
2. Klikni "Odaberi" ili prevuci Excel fajl
3. Odaberi sheet-ove (opciono)
4. Unesi koeficijent
5. Klikni "Pregled" da vidiš šta će se promeniti
6. Klikni "Ažuriraj Cene"

![GUI Screenshot](https://via.placeholder.com/600x400?text=GUI+Preview)

---

## 🥉 OPCIJA 3: Komandna linija (Python)

### ⌨️ Za programere i automatizaciju

**Instalacija:**

```bash
pip install -r requirements.txt
```

### Osnovno korišćenje

```bash
# Povećaj sve cene za 10%
python update_prices.py cenovnik.xlsx 1.1

# Smanji sve cene za 20%
python update_prices.py cenovnik.xlsx 0.8

# Prvo vidi šta bi se promenilo (dry run)
python update_prices.py cenovnik.xlsx 1.1 --dry-run
```

### Napredne opcije

```bash
# Samo određene kolone
python update_prices.py cenovnik.xlsx 1.1 -c C D

# Samo određeni sheet-ovi
python update_prices.py cenovnik.xlsx 1.1 -s "Januar 2024" "Februar 2024"

# Sačuvaj u novi fajl
python update_prices.py cenovnik.xlsx 1.1 -o cenovnik_novi.xlsx

# Kombinovano
python update_prices.py cenovnik.xlsx 1.15 -c C D E -s "Januar" "Februar" --dry-run
```

### Sve opcije

```
python update_prices.py [-h] [-c COLUMNS [COLUMNS ...]]
                        [-s SHEETS [SHEETS ...]]
                        [-o OUTPUT]
                        [--dry-run]
                        file coefficient

Pozicioni argumenti:
  file                  Putanja do Excel fajla
  coefficient           Koeficijent (1.1 = +10%, 0.9 = -10%)

Opcioni argumenti:
  -h, --help            Help poruka
  -c, --columns         Kolone za ažuriranje (npr. C D E)
  -s, --sheets          Sheet-ovi za ažuriranje
  -o, --output          Čuvaj u novi fajl
  --dry-run             Pregled bez promena
```

---

## 📊 Primeri koeficijenata

| Šta želiš | Koeficijent | Razlog |
|-----------|-------------|--------|
| Povećaj za 10% | **1.1** | 100 → 110 |
| Povećaj za 25% | **1.25** | 100 → 125 |
| Povećaj za 50% | **1.5** | 100 → 150 |
| Smanji za 10% | **0.9** | 100 → 90 |
| Smanji za 20% | **0.8** | 100 → 80 |
| Smanji za 50% | **0.5** | 100 → 50 |
| Udvostručenje | **2.0** | 100 → 200 |
| Utrostručenje | **3.0** | 100 → 300 |

---

## 🎯 Kako radi

Alat automatski:

1. **Detektuje kolone sa cenama** - Traži kolone koje u hederu sadrže: "cena", "cijena", "price", "cost", "iznos", "vrednost", itd.
2. **Prolazi kroz sve sheet-ove** - Obrađuje sve sheet-ove odjednom (ili one koje odabereš)
3. **Ažurira numeričke vrednosti** - Množi sve cene sa koeficijentom
4. **Čuva izmene** - Sačuva ažurirani fajl

---

## 💡 Saveti za sve opcije

1. **📁 Prvo napravi backup** - Uvek imaj kopiju originalnog fajla!
2. **👁 Koristi pregled** - U VBA i GUI verzijama, prvo vidi šta će se promeniti
3. **🧪 Testiraj na malom primeru** - Prvo testiraj na jednom sheet-u
4. **✅ Proveri rezultate** - Nakon ažuriranja, uvek proveri da li su cene tačne

---

## 📦 Fajlovi u projektu

| Fajl | Opis |
|------|------|
| **VBA fajlovi:** | |
| `VBA_Kod_v2_Poboljsan.bas` | ⭐ VBA v2 - Preporučeno! Rešava probleme sa formatom |
| `VBA_Kod.bas` | VBA originalna verzija |
| `VBA_Kod_Diagnostic.bas` | Diagnostic alat za pronalaženje problema |
| `INSTRUKCIJE_VBA.md` | Detaljne instrukcije za VBA setup |
| `TROUBLESHOOTING.md` | Rešavanje problema - Obavezno pročitati ako nešto ne radi! |
| **Python fajlovi:** | |
| `gui_app.py` | GUI aplikacija sa grafičkim interfejsom |
| `update_prices.py` | Python CLI skripta |
| `requirements.txt` | Python biblioteke |
| **Test fajlovi:** | |
| `cenovnik_primer.xlsx` | Test Excel fajl sa primerima |
| `create_sample.py` | Pomoćna skripta za kreiranje test fajlova |

---

## ❓ Česta pitanja

### Za VBA verziju

**P: Da li mogu da podelim .xlsm fajl sa kolegama?**
O: Da! Samo pošalji fajl i makro će raditi kod njih.

**P: Excel mi govori da su makroi onemogućeni?**
O: Klikni **Enable Content** u žutoj traci na vrhu.

**P: Radi li na Mac-u?**
O: Da! VBA radi i na Mac verziji Excel-a.

**P: Na nekim sheet-ovima radi, a na nekim ne?**
O: Najčešće su **brojevi sačuvani kao tekst**. Koristi `VBA_Kod_v2_Poboljsan.bas` koja automatski rešava ovaj problem! Pogledaj [TROUBLESHOOTING.md](TROUBLESHOOTING.md) za detaljno rešavanje.

### Za sve verzije

**P: Šta ako nemam header sa imenima kolona?**
O: Alat će ažurirati sve numeričke vrednosti koje pronađe.

**P: Da li mogu da vratim izmene?**
O: Koristi **Ctrl+Z** odmah nakon izmene, ili koristi backup funkciju/opciju.

**P: Šta ako imam cene u dinarima i eurima?**
O: Možeš odabrati samo određene kolone za ažuriranje.

**P: Da li radi sa .xls fajlovima?**
O: Ne, samo .xlsx i .xlsm. Možeš konvertovati .xls u .xlsx kroz Excel.

**P: Kako izračunam koeficijent?**
O:
- Za povećanje: novi_procenat / 100 (npr. 110% = 1.10)
- Za smanjenje: novi_procenat / 100 (npr. 80% = 0.80)
- Ili: 1 + promena (npr. +15% = 1.15, -15% = 0.85)

---

## 🎓 Primjer test fajla

Projekat uključuje `cenovnik_primer.xlsx` sa 4 sheet-a za testiranje:

- **Januar 2024** - Proizvodi sa osnovnim cenama
- **Februar 2024** - Više proizvoda
- **Mart 2024** - Maloprodajne i veleprodajne cene
- **Specijalna ponuda** - Redovne i akcijske cene

Možeš kreirati novi test fajl sa:

```bash
python create_sample.py
```

---

## 🏆 Preporuke

| Situacija | Najbolja opcija |
|-----------|-----------------|
| **Tim koji nije tehnički** | 🥇 VBA Macro |
| **Česta upotreba u firmi** | 🥇 VBA Macro |
| **Ne voliš komandnu liniju** | 🥈 GUI App |
| **Automatizacija / scripting** | 🥉 CLI |
| **Potrebna fleksibilnost** | 🥉 CLI |

---

## 📝 Licenca

Slobodno koristi i modifikuj!

---

## 🤝 Podrška

Za dodatna pitanja ili prilagođavanja, konsultuj osobu koja je napravila ovaj alat ili tvoj IT tim.

**Srećno sa ažuriranjem cena! 🎉**
