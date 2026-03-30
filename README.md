# TovarnaCiklov

VBA orodje za pripravo dolgoročnega in operativnega predogleda planov v Excelu, z integracijo podatkov iz GAME in dodatno avtomatiko za dodeljevanje `OFFICE`, periodičnih `OFFICE` dogodkov, drugih izmen ter osnovno analizo uspešnosti dodeljevanja.

## Namen projekta

Ta repozitorij vsebuje izvožene VBA module iz Excelove datoteke **TovarnaCiklov**. Orodje je namenjeno predvsem planiranju v okolju SZP, kjer je treba:

- iz GAME prebrati zaposlene, njihove time/cikle in posebnosti,
- zgraditi list **PREDOGLED** za izbrano obdobje,
- v plan vključiti obstoječe izmene iz GAME,
- dodeliti `OFFICE` glede na odstotek operative in dnevne potrebe,
- dodeliti periodične `OFFICE` dogodke,
- dopolniti druge izmene po timskih predlogah,
- pripraviti osnovno poročilo in graf analize,
- omogočiti ročni `undo` zadnje avtomatske akcije.

Projekt je zgrajen modularno: en del skrbi za nastavitve, en del za generiranje predogleda, en del za `OFFICE` logiko, en del za analizo in en del za `undo` mehanizem.

---

## Kaj makri trenutno pokrivajo

### 1. Ustvarjanje splošnih ciklusov za delavce
Glavni makro za to je:

- `TovarnaCiklovLKZP.CreateDefaultRooster`

Naloga makra:
- prebere ključne nastavitve iz lista **NASTAVITVE**,
- odpre datoteko **GAMA**,
- prebere zaposlene, njihove ID-je, time/cikle in osnovne podatke,
- na list **PREDOGLED** zapiše osnovne cikluse za izbrano obdobje,
- ne dodeljuje `OFFICE`, temveč pripravi osnovno matriko dela.

To je izhodiščni korak celotnega procesa.

### 2. Kopiranje ciklusov in izmen iz GAME
Za sinhronizacijo s podatki iz GAME se uporablja:

- `CopyFrmGAMA.CopyFrmGAMA`

Makro po ID-ju poveže osebo v **PREDOGLEDU** z isto osebo v **GAMI** in prenese izmene iz GAME na ustrezne dni v predogledu. V praksi deluje kot “mirror sync” med virom in predogledom.

Uporabno je predvsem po tem, ko je osnovni ciklus že zgrajen, saj nato v predogled pridejo še posebne ali že vnaprej določene izmene.

### 3. Dodeljevanje office glede na odstotek officev
Osrednji makro za ta del je:

- `modOffice.ApplyOfficeToPreview`

Pod njim delujejo še pomožni moduli:
- `modOffice_Logic`
- `modOfficeModels`
- `modSettings`
- `Analiza`
- `modUndo`

Makro:
- prebere nastavitve,
- v RAM naloži stanje iz **PREDOGLED** in **GAMA**,
- izračuna potrebo po `OFFICE` glede na odstotek operative,
- izbere model dodeljevanja (`GLOBAL` ali `SEQUENTIAL`),
- zapiše rezultat nazaj v **PREDOGLED**,
- pripravi `undo snapshot`,
- izdela poročilo in graf analize.

Podprta sta vsaj dva pristopa dodeljevanja:
- `AssignOffice_FairPerRound` – bolj enakomerna razporeditev po osebah,
- `AssignOffice_GreedySequential` – bolj sekvenčen/grabežljiv model.

Dodeljevanje je vodeno prek nastavitev in uporablja pravila, kot so:
- dovoljene izmene za prepis,
- prag števcev,
- obravnava delovnih dni,
- razporeditev po enotah,
- blokirani dnevi (`SO`, `NE`, `PR`),
- možnost širjenja `OFFICE` čez obdobje, kadar je kapacitete dovolj.

### 4. Dodeljevanje periodičnih office dogodkov
Za periodične dogodke obstaja ločen makro:

- `modOffice.ApplyPeriodicMeetingsOffice`

Makro prebere nastavitve za periodične sestanke iz lista **NASTAVITVE** in poskuša zapisati `OFFICE` oziroma opombe na predvidene termine za izbrane zaposlene.

Tipični parametri:
- aktivacija makra,
- ime sestanka,
- dan v tednu,
- kateri pojav v mesecu,
- seznam ID-jev zaposlenih,
- opomba.

Ta del je uporaben za ponavljajoče se dogodke, npr. sestanke vodij ali druge redne obveznosti, ki jih želimo sistematično umestiti v plan.

### 5. Dodeljevanje drugih izmen
Za dopolnjevanje drugih izmen obstaja:

- `DodajDrugeIzmene.DodajDrugeIzmene`

Makro deluje po logiki timske predloge:
- poišče vrstice, kjer že obstaja referenčni vzorec cikla (`X1`, `X2`, `X3`, `O`),
- iz teh vrstic zgradi predlogo po timu,
- v drugih vrsticah z istim timom zapolni prazna mesta z izbrano šifro,
- zna tudi resetirati vnose, če je v kodi znak `-`.

Ta korak je uporaben za hitre serijske vnose dodatnih tipov izmen po že obstoječem vzorcu.

### 6. Analiza dodeljevanja officev
Poročilni del je implementiran v modulu:

- `Analiza.bas`

Glavni proceduri sta:
- `UstvariPorocilo_Block`
- `NarediGraf_Analiza_Block`

Analiza pripravi:
- tabelo najboljših dni (`BEST DAYS`),
- log dodelitev,
- pregled po osebah,
- odstotek uspešnosti,
- skupne pokazatelje uspešnosti,
- grafični prikaz rezultatov.

Rezultat je list **POROČILO**, ki je namenjen hitremu preverjanju, kako uspešno je bila opravljena dodelitev `OFFICE`.

### 7. Undo zadnje akcije
Undo je ločen v modulu:

- `modUndo.bas`
- sprožitev prek `modOffice.UndoLastAction`

Trenutna logika uporablja **single-level snapshot**:
- shrani se obseg celic,
- shranijo se `Value2` in `FormulaR1C1`,
- ne shranjujejo se formati, komentarji, barve in podobno.

To pomeni, da undo ni splošen Excel undo, ampak lasten VBA mehanizem za zadnji avtomatski zapis.

> Trenutno stanje: po opisu uporabnika `undo` v praksi še ni zanesljiv oziroma “ne dela”, zato je ta del smiselno obravnavati kot **known issue**.

---

## Arhitektura modulov

### `TovarnaCiklovLKZP.bas`
Vstopna točka za gradnjo osnovnega predogleda iz ciklov.

### `modSettings.bas`
Centralno branje nastavitev iz lista **NASTAVITVE**.

Pomembni tipi:
- `TMainSettings`
- `TOfficeSettings`
- `TUnitConfig`

Ta modul skrbi za:
- branje ključev iz nastavitev,
- pretvorbe tipov (`Text`, `Long`, `Date`, `Bool`),
- logiranje faz izvajanja,
- nalaganje nastavitev po enotah.

### `modOffice.bas`
Glavni orkestrator za `OFFICE` in periodične sestanke.

Skrbi za:
- odpiranje GAME,
- preverjanje poravnave datumov,
- branje predogleda v RAM,
- pisanje rezultatov nazaj,
- pripravo `undo snapshot-a`,
- zagon analize.

### `modOffice_Logic.bas`
Poslovna logika in normalizacija šifer.

Skrbi za:
- čiščenje in kanonizacijo šifer izmen,
- določitev enote iz oznake tima,
- izračun potreb po `OFFICE`.

### `modOfficeModels.bas`
Algoritmi za dodeljevanje `OFFICE` v RAM poljih.

### `CopyFrmGAMA.bas`
Sinhronizacija izmen iz GAME v **PREDOGLED**.

### `DodajDrugeIzmene.bas`
Množično polnjenje dodatnih izmen po timskih predlogah.

### `Analiza.bas`
Poročilo in grafi o uspešnosti dodeljevanja.

### `modUndo.bas`
Lastni mehanizem za ročni povratek zadnje akcije.

---

## Predviden potek uporabe

Tipičen vrstni red dela je tak:

1. Urediš nastavitve na listu **NASTAVITVE**.
2. Zaženeš `CreateDefaultRooster`.
3. Po potrebi sinhroniziraš izmene z `CopyFrmGAMA`.
4. Zaženeš `ApplyOfficeToPreview`.
5. Po potrebi zaženeš `ApplyPeriodicMeetingsOffice`.
6. Po potrebi dopolniš še druge izmene z `DodajDrugeIzmene`.
7. Pregledaš list **POROČILO**.
8. Če rezultat ni ustrezen, poskusiš `UndoLastAction`.

---

## Odvisnosti in predpostavke

Projekt predpostavlja, da obstajajo oziroma so pravilno poimenovani vsaj naslednji listi:

- `NASTAVITVE`
- `PREDOGLED`
- `CIKLI`
- `POROČILO` (ga lahko makro tudi ustvari)

Poleg tega je potrebna dostopna datoteka **GAMA**, katere pot in struktura se bereta iz **NASTAVITEV**.

Makri so občutljivi na:
- pravilne indekse stolpcev in vrstic,
- pravilno nastavljen začetni datum,
- poravnavo datumov med `NASTAVITVE`, `GAMA` in `PREDOGLED`,
- veljavne ID-je zaposlenih,
- dosledne oznake timov/ciklov.

---

## Ključne nastavitve

Pomembni parametri na listu **NASTAVITVE** vključujejo:

- pot do datoteke GAMA,
- ime lista v GAMA,
- prva/zadnja vrstica zaposlenih,
- stolpec ID,
- stolpec imena,
- stolpec OJT,
- stolpec tima,
- odstotek operative,
- stolpec licence,
- začetni in končni datum,
- prvi stolpec datumov v GAMA,
- prvi stolpec datumov v PREDOGLED,
- prva vrstica zaposlenih v PREDOGLED,
- izbor enot za planiranje,
- izključene tipe,
- prag za števce,
- katere izmene se sme prepisati z `OFFICE`,
- model dodeljevanja `OFFICE` (`GLOBAL` / `SEQUENTIAL`),
- nastavitve za periodične sestanke.

Če te nastavitve niso usklajene, bodo makri praviloma vrgli `MsgBox` napako in se ustavili.

---

## Znane omejitve / known issues

### 1. Undo ni popolnoma zanesljiv
`modUndo` je enonivojski in ne vrača formatiranja. Po trenutnem stanju je to najbolj očitna točka za izboljšavo.

### 2. Močna odvisnost od strukture workbooka
Makri so tesno vezani na konkretna imena listov, stolpce, vrstice in organizacijo nastavitev.

### 3. Veliko poslovne logike je implicitne
Del pravil je skritih v šifrah izmen, timskih oznakah, strukturi GAME in nastavitvah po enotah. Brez poznavanja procesa je kodo težje vzdrževati.

### 4. Omejena testabilnost
Ker gre za VBA v Excel okolju, ni enostavnega avtomatskega testiranja. Priporočljivo je testiranje na kopiji workbooka.

---

## Predlog strukture repozitorija

Praktična struktura za Git repozitorij:

```text
TovarnaCiklov/
├─ README.md
├─ src/
│  ├─ TovarnaCiklovLKZP.bas
│  ├─ modSettings.bas
│  ├─ modOffice.bas
│  ├─ modOffice_Logic.bas
│  ├─ modOfficeModels.bas
│  ├─ CopyFrmGAMA.bas
│  ├─ DodajDrugeIzmene.bas
│  ├─ Analiza.bas
│  └─ modUndo.bas
├─ docs/
│  ├─ PLANIRANJE-SZP-NAVODILA.md
│  ├─ setup.md
│  ├─ workflow.md
│  └─ known-issues.md
└─ samples/
   └─ primer-nastavitev.md
```

Če v repozitorij dodajaš tudi `.xlsm` datoteko, je priporočljivo, da ostane ločeno od izvornih `.bas` modulov. Za verzioniranje kode je bistveno bolj pregledno, da Git sledi izvoženim modulom, ne samo binarnemu workbooku.

---

## Priporočila za nadaljnji razvoj

1. **Stabilizacija `undo` mehanizma**
   - preveriti, ali se snapshot vedno zajame nad pravim obsegom,
   - dodati večnivojski undo ali vsaj bolj jasen status snapshot-a,
   - po potrebi shranjevati tudi komentarje oziroma formate.

2. **Ločitev poslovne logike od UI / workbook plasti**
   - še več logike premakniti iz `modOffice` v čiste helperje,
   - zmanjšati število mest, kjer se neposredno bere/piše po worksheetih.

3. **Dokumentiranje nastavitev**
   - narediti ločen dokument z vsemi ključi iz `NASTAVITEV`,
   - zraven dopisati primer vrednosti in pomen vsake nastavitve.

4. **Uvedba testnih scenarijev**
   - vsaj ročni testni checklist za tipične primere:
     - gradnja predogleda,
     - copy iz GAME,
     - office rotacija,
     - periodični office,
     - analiza,
     - undo.

5. **Poenotenje poimenovanj**
   - `CreateDefaultRooster` verjetno pomeni `Roster`, ne `Rooster`, zato je smiselno razmisliti o preimenovanju zaradi jasnosti.

---

## Začetek dela

Če želiš projekt zagnati na novem workbooku:

1. Uvozi vse `.bas` module v VBA projekt.
2. Pripravi liste `NASTAVITVE`, `PREDOGLED` in `CIKLI`.
3. Na listu `NASTAVITVE` izpolni vse obvezne ključe.
4. Preveri, da je pot do GAME pravilna.
5. Najprej zaženi `CreateDefaultRooster`.
6. Nadaljuj po običajnem workflowu.

---

## Povzetek

TovarnaCiklov je specializirano VBA orodje za planerski workflow, ki povezuje **GAMO**, **PREDOGLED**, pravila po enotah in več avtomatskih korakov za pripravo plana. Največja vrednost projekta je v tem, da:

- pohitri gradnjo plana,
- standardizira dodeljevanje `OFFICE`,
- omogoča sinhronizacijo z obstoječimi podatki,
- doda osnovno analitiko,
- zmanjša ročno delo pri ponavljajočih se korakih.

Hkrati pa projekt zaradi tesne vezave na Excel okolje in konkretno strukturo podatkov zahteva dobro dokumentirane nastavitve, previdno testiranje in nadaljnjo refaktorizacijo za lažje vzdrževanje.
