# Excel Price Updater - Automatsko ažuriranje cena

Alat za automatsko ažuriranje cena u Excel fajlovima sa više sheet-ova. Umesto da ručno menjaš cene na svakom sheet-u, ova skripta to radi automatski za tebe!

## 🚀 Brza instalacija

```bash
# Instaliraj potrebne biblioteke
pip install -r requirements.txt
```

## 📖 Kako koristiti

### Osnovno korišćenje

```bash
# Povećaj sve cene za 10% (množi sa 1.1)
python update_prices.py cenovnik.xlsx 1.1

# Smanji sve cene za 20% (množi sa 0.8)
python update_prices.py cenovnik.xlsx 0.8

# Povećaj cene za 15%
python update_prices.py cenovnik.xlsx 1.15
```

### Napredne opcije

#### 1. Ažuriranje samo određenih kolona

Ako želiš da ažuriraš samo određene kolone (npr. samo kolone C i D):

```bash
python update_prices.py cenovnik.xlsx 1.1 -c C D
```

#### 2. Ažuriranje samo određenih sheet-ova

Ako imaš mnogo sheet-ova, ali želiš da ažuriraš samo neke:

```bash
python update_prices.py cenovnik.xlsx 1.1 -s "Januar 2024" "Februar 2024"
```

#### 3. Čuvanje u novi fajl

Ako ne želiš da prepisuješ originalni fajl:

```bash
python update_prices.py cenovnik.xlsx 1.1 -o cenovnik_novi.xlsx
```

#### 4. Dry run - proveri šta bi se promenilo

Pre nego što stvarno promeniš fajl, možeš da vidiš šta bi se desilo:

```bash
python update_prices.py cenovnik.xlsx 1.1 --dry-run
```

### Kombinovanje opcija

Možeš kombinovati više opcija:

```bash
python update_prices.py cenovnik.xlsx 1.15 -c C D E -s "Januar" "Februar" --dry-run
```

## 🎯 Kako radi

Skripta automatski:

1. **Detektuje kolone sa cenama** - Traži kolone koje u hederu sadrže reči kao što su: "cena", "cijena", "price", "cost", "iznos", itd.
2. **Prolazi kroz sve sheet-ove** - Automatski obrađuje sve sheet-ove u fajlu (ili one koje navedete)
3. **Ažurira numeričke vrednosti** - Množi sve numeričke vrednosti sa zadatim koeficijentom
4. **Čuva izmene** - Sačuva ažurirani fajl

## 📊 Primeri koeficijenata

| Šta želiš | Koeficijent | Primer komande |
|-----------|-------------|----------------|
| Povećaj za 10% | 1.1 | `python update_prices.py file.xlsx 1.1` |
| Povećaj za 25% | 1.25 | `python update_prices.py file.xlsx 1.25` |
| Smanji za 10% | 0.9 | `python update_prices.py file.xlsx 0.9` |
| Smanji za 20% | 0.8 | `python update_prices.py file.xlsx 0.8` |
| Udvostruči cene | 2.0 | `python update_prices.py file.xlsx 2.0` |
| Prepolovi cene | 0.5 | `python update_prices.py file.xlsx 0.5` |

## 💡 Saveti

1. **Prvo koristi --dry-run** - Uvek prvo proveri šta bi se promenilo
2. **Napravi backup** - Uvek imaj kopiju originalnog fajla pre ažuriranja
3. **Ili čuvaj u novi fajl** - Koristi `-o` opciju da sačuvaš u novi fajl
4. **Proveri sheet-ove** - Ako neki sheet-ovi nemaju cene, možeš ih preskočiti sa `-s` opcijom

## ⚙️ Sve opcije

```
python update_prices.py [-h] [-c COLUMNS [COLUMNS ...]]
                        [-s SHEETS [SHEETS ...]]
                        [-o OUTPUT]
                        [--dry-run]
                        file coefficient

Pozicioni argumenti:
  file                  Putanja do Excel fajla
  coefficient           Koeficijent za množenje (npr. 1.1 za +10%, 0.9 za -10%)

Opcioni argumenti:
  -h, --help            Prikaži help poruku
  -c, --columns         Kolone za ažuriranje (npr. C D E)
  -s, --sheets          Imena sheet-ova za ažuriranje
  -o, --output          Putanja za čuvanje novog fajla
  --dry-run             Prikaži šta bi se promenilo bez promena
```

## ❓ Česta pitanja

**P: Šta ako nemam header sa imenima kolona?**
O: Skripta će ažurirati sve numeričke vrednosti koje pronađe.

**P: Da li mogu da vratim izmene?**
O: Ako nisi koristio `-o` opciju, prepisao si originalni fajl. Zato preporučujem da prvo koristiš `--dry-run` ili `-o` da sačuvaš u novi fajl.

**P: Šta ako imam i cene u dinarima i u eurima?**
O: Možeš koristiti `-c` opciju da navedete samo određene kolone koje treba ažurirati.

**P: Da li radi sa .xls fajlovima?**
O: Ne, trenutno podržava samo .xlsx format. Možeš konvertovati .xls u .xlsx kroz Excel.

## 📝 Licenca

Slobodno koristi i modifikuj!
