# Handleiding voor de leraar — Les 06: Mijn NovaDepot-map

**Vak:** Toegepaste Informatica
**Doelgroep:** 3ORLO, 4ORLOa en 4ORLOb — 2de graad Organisatie en logistiek, arbeidsmarktgerichte finaliteit
**Lesduur:** 1 × 50 minuten: 10 minuten instructie + 40 minuten keuzewerktijd
**Context:** NovaDepot, fictief logistiek bedrijf en groothandel — module 2 van het jaarplan
(`Leerplannen/Lessenplan_TOINFO_ORLO_arbeidsmarkt_2026-2027.md`)
**Toestel:** **Chromebook, in de studie** (op jouw vraag, 08-10-2026) — Google Workspace in Chrome en de app *Bestanden*
**Lesdag:** uur 2 van week 6: 3ORLO wo 7 oktober (2de uur) · 4ORLOb wo 7 oktober (3de uur) · 4ORLOa do 8 oktober (7de uur)
**Kernleerplandoelen:** `BV2_04.03` (digitale inhouden beheren) en `BK2_02.06` / `BK2_02.06.02` (logische mappenstructuur) — toepassen
**Deadline:** vrijdag 9 oktober 2026, 20.00 uur · 20 punten

De leerling maakt in Google Drive een map *NovaDepot* met drie mappen, haalt vier rommelige bestanden uit
een zip-bestand, geeft ze een naam volgens de afspraak **Wat - Van wie** en zet ze in de juiste map. Op
vraag van Jonas **heel eenvoudig**: zelfs bij de doorstroomklas (les 2 van 3MWWE) was het benoemen van de
bestanden het moeilijkste. Daarom geen datum, geen versie en geen lage streepjes in de naam.

> [!NOTE]
> **Chromebook-versie (08-10-2026).** De leerlingen maken deze taak in de studie, met alleen een
> Chromebook en zonder klassikale uitleg. Daarom:
> - **Stap 3** gaat via de app **Bestanden**: 2 vingers op het zip-bestand › *Alles uitpakken*, dan
>   de vier bestanden kopiëren en plakken in *Google Drive › Mijn Drive › NovaDepot*.
> - **Rechtsklikken** heet overal *klikken met 2 vingers* (of *Alt + klik*).
> - **Stap 4 en 5** hebben een **filmpje** van 12 à 13 seconden: het voorbeeld uit rij 1, hernoemd en
>   verplaatst in een demomap, opgenomen met Claude in Chrome in jouw profiel *Jonas*.
> - Twee dingen die pas bij het opnemen bleken: **Drive selecteert bij *Naam wijzigen* de hele naam,
>   ook `.docx`** (wie meteen typt, verliest de extensie: de leerling typt ze er dus zelf achter), en
>   het venster **Verplaatsen opent op *Voorgesteld*** met andere mappen (de leerling klikt eerst naast
>   *Huidige locatie* op *NovaDepot*).

---

## 1. Inhoud van het pakket

```text
W06 - Les 06 - ORLO - Mijn NovaDepot-map/
├── index.html                   # de lespagina: route · één stap · checklist
├── presentatie.html             # 7 dia's voor de lesstart en "Ik doe"
├── css/style.css, css/slides.css
├── js/script.js, js/slides.js
├── assets/
│   ├── novadepot-logo.svg/.png, novadepot-icon.svg, dalton-gent-logo.png
│   ├── fonts/                   # Atkinson Hyperlegible + Montserrat (OFL, zelf gehost)
│   ├── video/                   # zip-downloaden.mp4 + startbeeld (uit les 2 van 3MWb en 3MWWE)
│   └── screenshots/LEESMIJ.md   # vier optionele schermafbeeldingen
├── werkdocument/
│   ├── WP1_Mijn-NovaDepot-map.docx   # het werkdocument dat de leerling INLEVERT
│   ├── NovaDepot-bestanden.zip       # DIT hang je aan de opdracht
│   ├── rommel/                       # de vier losse bestanden die in de zip zitten
│   └── maak_werkdocumenten.py        # maakt alles opnieuw (python-docx, openpyxl, Pillow)
├── lesvoorbereiding.md          # volgens §9.2 van de AI-lesplanner v2.1
├── dalton-lesfiche.html         # Dalton-lesfiche in de kleurcode: openen, Kopieer, plakken in je planner
├── classroom.json               # de opdracht voor _tools/zet_opdracht_klaar.py
├── lesdoelen.json               # leerplandoelen voor je jaaroverzicht
└── README.md                    # deze handleiding
```

### 1b. Hoe de lespagina werkt

Dezelfde opbouw als de vorige ORLO-lessen (AI-lesplanner v2.1): links de **route** (zes stappen in twee
groepen: *Je map in Drive* · *Ordenen en inleveren*), midden **één stap** met vijf vaste blokken, rechts
de **checklist** met 17 concrete taken. Op een smal venster staat alles onder elkaar en zie je alleen de
taken van de huidige stap; in stap 6 staat de hele lijst open.

- Kleine stappen: 3 tot 5 handelingen per stap.
- De **naamkaart** in stap 4 toont de afspraak in kleur: *Wat is het?* (Leveringsbon, Factuur, Prijslijst
  of Foto) en *Van wie?* (de naam van het bedrijf, die bovenaan in elk bestand staat).
- Het **filmpje** bij *Hulp nodig?* in stap 3 toont hoe je de zip downloadt uit de opdracht.
- De **theoriekaart** heeft een kaartje *Zo werkt elke les*: de leerlingen moeten nog wennen aan het
  systeem (Classroom › lespagina › werkdocument › Inleveren).
- Alleen Chromebook (de taak gebeurt in de studie). `localStorage` bewaart alleen de vinkjes en de laatste stap
  (voorvoegsel `novadepot_map_v1_`), met een wisknop.

---

## 2. Klaarzetten

### Stap 1 — Publiceren via GitHub Pages ✅ *gebeurd op 04-10-2026*

Repository [`jonasdaltongent/Mappen-NovaDepot`](https://github.com/jonasdaltongent/Mappen-NovaDepot),
Pages op branch `main`, map `/ (root)`. De lespagina staat op
[jonasdaltongent.github.io/Mappen-NovaDepot](https://jonasdaltongent.github.io/Mappen-NovaDepot/); dat
adres staat ook op dia 6, in `classroom.json` en in `lesdoelen.json`.

### Stap 2 — De opdracht in Classroom, met de koppeling ✅ *gepubliceerd op 04-10-2026*

Gepubliceerd in [3ORLO](https://classroom.google.com/c/MTYyNjY1NjU1OTEx/a/MjU0MzkzNjAwMDJa/details),
[4ORLOa](https://classroom.google.com/c/MjUzMTgzNDM2NjZa/a/ODg4OTg5NjY4NDE0/details) en
[4ORLOb](https://classroom.google.com/c/MjUzMTc2OTkwNzNa/a/MjU0MzkzNzc5MzRa/details). Nagekeken vóór
het publiceren: het werkdocument is een Google-document (kopie per leerling), de zip is alleen te
bekijken, deadline en punten kloppen.

`classroom.json` zet de opdracht klaar in 3ORLO, 4ORLOa en 4ORLOb (zie `_tools/CLASSROOM-KOPPELING.md`):

| | Opdracht: **Werkplek 1 — Mijn NovaDepot-map** |
|---|---|
| **Onderwerp** | Digitale competenties (bestaat al, zoals bij les 01) |
| **Bijlage 1** | de link naar de lespagina |
| **Bijlage 2** | `WP1_Mijn-NovaDepot-map` (Google-document) — **Een kopie maken voor elke leerling** |
| **Bijlage 3** | `NovaDepot-bestanden.zip` — **Leerlingen kunnen bestand bekijken** |
| **Punten** | 20 |
| **Deadline** | vrijdag 9 oktober 2026, 20.00 uur |

```bash
python3 "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/_tools/zet_opdracht_klaar.py" "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/2026-2027/W06 - Les 06 - ORLO - Mijn NovaDepot-map" --maak
```

Daarna `--publiceer` als je wil dat de opdracht meteen gepubliceerd wordt, of zelf **Toewijzen** in
Classroom. Publiceer pas in Classroom als de lespagina online staat (stap 1).

Instructietekst (vult het script in):

```text
1. Open de lespagina (link) en je werkdocument WP1_Mijn-NovaDepot-map.
2. Het zip-bestand heb je nodig in stap 3.
3. Klik op Inleveren, ten laatste vrijdag 9 oktober 2026 om 20.00 uur.
```

### Stap 3 — Met de hand, als de koppeling niet werkt

1. Zet eenmalig in Drive **Uploads converteren naar de indeling van een Editor van Google Documenten** aan
   ([Drive-help](https://support.google.com/drive/answer/2424368?hl=nl)).
2. Upload `werkdocument/WP1_Mijn-NovaDepot-map.docx` **in Drive zelf** (**Nieuw** › **Bestand uploaden**).
   Staat er geen `.docx` meer achter de naam? Dan is het een Google-document.
3. Voeg het in de opdracht toe met **Bijvoegen** › **Drive** en kies **Een kopie maken voor elke
   leerling** ([Classroom-help](https://support.google.com/edu/classroom/answer/6020265?hl=nl)).
4. Voeg `werkdocument/NovaDepot-bestanden.zip` toe met **Bijvoegen** › **Uploaden** en kies **Leerlingen
   kunnen bestand bekijken**. Een zip-bestand wordt niet omgezet; dat hoeft ook niet.
5. Voeg de lespagina toe met **Link**.

> [!IMPORTANT]
> Alleen het werkdocument krijgt *Een kopie maken voor elke leerling*, en dat kan alleen **vóór** je de
> opdracht post. Het zip-bestand niet: één gedeelde kopie volstaat, de leerlingen downloaden het toch.

### Afvinklijst vóór de les

- [ ] De lespagina staat online en het adres op dia 6 klopt.
- [ ] De opdracht staat in de drie klassen, met drie bijlagen en de juiste instelling per bijlage.
- [ ] Een testleerling krijgt een eigen kopie van het werkdocument, met de eigen naam in de titel.
- [ ] Met een **leerlingaccount** getest: de link naar het zip-bestand in stap 3 werkt, en de terugweg via de
      opdracht (Openen met › Openen in nieuw tabblad) ook.
- [ ] De map NovaDepot staat bij jou al klaar voor de demo (mappen maken doe je niet voor).
- [ ] De Dalton-lesfiche staat in je planner (open `dalton-lesfiche.html`, klik op **Kopieer de fiche**, plak).
- [ ] `presentatie.html` opent op de beamer; `N` toont je notities, `F` is volledig scherm.

### 2b. Nagelezen klikpaden

Alle knopnamen komen uit de tabel van §7 van de lesplanner (nagelezen op 23 en 27-09-2026) en uit les 2
van 3MWb en 3MWWE:

| Handeling | Klikpad / naam | Bron |
|---|---|---|
| Map maken | **Mijn Drive** › **Nieuw** › **Nieuwe map**, naam typen, **Maken** | jouw scherm (04-10-2026); de [Drive-help](https://support.google.com/drive/answer/2375091?hl=nl) schrijft *Map* |
| Zip downloaden | de link in stap 3 opent het zip-bestand uit de opdracht in een nieuw tabblad › de pijl linksboven (**Downloaden**). Terugweg (bij *Hulp nodig?*): in de opdracht klik op het bestand › **Openen met** › **Openen in nieuw tabblad** › de pijl linksboven | Jonas' schermopname (27-09-2026), lesplanner §7; de link opent dezelfde Drive-pagina (`/file/d/…/view`) |
| Rechtsklikken (Chromebook) | druk of tik met **2 vingers** op de touchpad, of **Alt** + klik | [Chromebook 1047367](https://support.google.com/chromebook/answer/1047367?hl=nl) |
| App Bestanden openen | **Launcher** (hoek van het scherm) › **Bestanden**, of **Shift + Alt + m**; links **Downloads** | [Chromebook 1700055](https://support.google.com/chromebook/answer/1700055?hl=nl) · [sneltoetsen](https://support.google.com/chromebook/answer/183101?hl=nl) |
| Uitpakken (Chromebook) | in **Bestanden** 2 vingers op het zip-bestand › **Alles uitpakken**; een dubbelklik opent het zip-bestand links als een map (**Uitwerpen**) | [Chromebook 1700055](https://support.google.com/chromebook/answer/1700055?hl=nl) |
| Naar Drive (Chromebook) | nieuwe map openen › **Ctrl + A** › **Ctrl + C** › links **Google Drive** › **Mijn Drive** › NovaDepot › **Ctrl + V**; terugweg: Drive › **Nieuw** › **Bestand uploaden** | [Chromebook 1700055](https://support.google.com/chromebook/answer/1700055?hl=nl) (linkerkolom *Google Drive › Mijn Drive*) |
| Naam wijzigen | 2 vingers › **Naam wijzigen**; de **hele naam** staat geselecteerd, ook `.docx`: typ de extensie mee › **OK**. Anders: bestand aanklikken › **Meer acties** › **Naam wijzigen** | [Drive 2424384](https://support.google.com/drive/answer/2424384?hl=nl) · proef in Drive, 08-10-2026 |
| Verplaatsen | 2 vingers › **Ordenen** › **Verplaatsen** › naast *Huidige locatie* op **NovaDepot** › map › **Verplaatsen** (het venster opent op *Voorgesteld*) | [Drive 2375091](https://support.google.com/drive/answer/2375091?hl=nl) · proef in Drive, 08-10-2026 |
| Kleur en ster (extra) | 2 vingers › **Ordenen** › **Kleur van map** · **Toevoegen aan Met ster** | [Drive 2375091](https://support.google.com/drive/answer/2375091?hl=nl) |
| Inleveren | **Inleveren** · **Inleveren ongedaan maken** | [Classroom 6020285](https://support.google.com/edu/classroom/answer/6020285?hl=nl) |

---

## 3. Het verloop van de les

| Fase | Tijd | Wat |
|---|---|---|
| **Instructie** | **10'** | Dia 1–2 lesstart (3'): *waar zit de factuur?* Dia 3–5 (5'): lesdoel, voor en na, en de demo: de zip uit de opdracht naar Drive (downloaden, uitpakken, uploaden uit de **nieuwe map**). Dia 6 *Zo werk je verder* (2') |
| **Keuzewerktijd** | **40'** | Dia 6 blijft staan; de leerlingen werken stap 1 tot 6 af, inleveren inbegrepen. Dia 7 in de laatste minuut. |

**Stap 4 doe je niet klassikaal voor** ("dan letten ze niet op"): bestand 1 staat als voorbeeld in rij 1
van het werkdocument, en je helpt individueel.

Keuzewerktijd = 50 minuten − instructietijd. De minuten per stap staan op dia 6 en in
`dalton-lesfiche.html`. Zeg aan het einde mondeling dat wie niet klaar is, toch inlevert.

**Eerste rondgang, kijk alleen naar twee dingen.** Ze blokkeren alles wat erna komt:

1. Staat de map NovaDepot in **Mijn Drive** en niet in de map *Classroom*?
2. Zijn de vier bestanden van **Downloads** naar **Drive** geraakt?

**Tweede rondgang, rond stap 4:** laat een nieuwe naam luidop lezen. *"Weet ik nu wat erin zit, en van
wie?"* Staat `.docx` er nog achter?

---

## 4. Verbetersleutel

De leerling levert alleen het werkdocument in. De mappen zie je niet: kijk ernaar tijdens de rondgang. In
les 08 (*Delen met de juiste rechten*) delen de leerlingen hun map met jou.

### Deel 2 — de vier bestanden

| Oude naam | Nieuwe naam | Map |
|---|---|---|
| `Document zonder titel.docx` | `Leveringsbon - Bakkerij Korstjes.docx` | Leveringen (voorbeeld, al ingevuld) |
| `factuur def nieuw (2).docx` | `Factuur - Papierhandel Vellekens.docx` | Facturen |
| `lijst.xlsx` | `Prijslijst - Koffiebranderij De Bonenbaas.xlsx` | Prijslijsten |
| `IMG_20261013_131502.png` | `Foto - Speelgoed Tolletje.png` | Leveringen |

Kleine verschillen zijn **goed**: hoofdletters, een spatie meer of minder, *Foto levering* in plaats van
*Foto*. Wat moet kloppen: eerst wat het is, dan van wie, en de extensie staat er nog.

### Deel 3 — de vragen

| Vraag | Waar het om gaat |
|---|---|
| 1 | Na het uitpakken: **op deze computer, in Downloads** (lokaal). Na het uploaden: **in Google Drive** (cloud). |
| 2 | De nieuwe naam zegt wat erin zit en van wie: je vindt het terug zonder het te openen, ook een collega. |

### Essentiële fouten — geef hier altijd feedback op

- De map NovaDepot staat in de map *Classroom* in plaats van in *Mijn Drive*.
- De bestanden staan nog alleen in Downloads.
- Een naam zegt nog altijd niet wat erin zit of van wie (*factuur 2*, *lijst nieuw*).

---

## 5. Schermafbeeldingen (optioneel)

De schermafbeelding van stap 2 staat er al; stap 4 en 5 hebben een filmpje. In stap 3 zijn twee plaatsen
voor een schermafbeelding van een **Chromebook** (de app *Bestanden*): die kan ik vanaf je Mac niet
maken. Open `index.html?leraar` om te zien waar ze komen; de lijst staat in
`assets/screenshots/LEESMIJ.md`. Zolang ze ontbreken, zie je twee 404-meldingen in de console.

---

## 6. Het materiaal opnieuw maken

```bash
python3 "werkdocument/maak_werkdocumenten.py"
```

Maakt het werkdocument, de vier rommelbestanden en de zip opnieuw. Vereist `python-docx`, `openpyxl` en
`Pillow`. Lees eerst de waarschuwing bovenaan het script: de slechte bestandsnamen zijn lesmateriaal.

> [!IMPORTANT]
> **De link in stap 3** wijst naar het zip-bestand dat de koppeling op 04-10-2026 in Drive zette en aan de
> drie opdrachten hing (één bestand, gedeeld met 3ORLO, 4ORLOa en 4ORLOb). Maak je de opdracht opnieuw aan,
> of vervang je het zip-bestand in Classroom, dan krijgt het een nieuw adres. Pas de link in `index.html`
> dan aan (zoek op `drive.google.com/file`). Leerlingen moeten aangemeld zijn met hun schoolaccount.

---

## 7. Leerplandoelen in je jaaroverzicht

De repository krijgt dezelfde `pre-push` hook als de andere lessen: bij elke push roept hij
`_tools/update_leerdoelen.py` aan met `lesdoelen.json`. De regels zijn bij de push van 04-10-2026 toegevoegd. Het blad
**ORLO** heeft sinds die dag ook een rij voor `BV2_04.03`, dus dat doel telt mee.

---

## 8. Wat nog moet blijken in de klas

1. **Haalbaarheid.** Downloaden, uitpakken en uploaden is het riskante stuk (10 minuten). Loopt het
   vast, laat de leerlingen dan in stap 4 maar twee bestanden hernoemen.
2. **Het venster *Naam wijzigen*.** Nagegaan op 08-10-2026: Drive selecteert de hele naam, ook `.docx`.
   De lespagina en het filmpje zeggen daarom: typ de extensie er zelf achter.
6. **Op een Chromebook in de studie** (nog niet nagekeken op een schoolchromebook): heet de nieuwe map na
   *Alles uitpakken* echt `NovaDepot-bestanden`, en staat *Google Drive* links in de app *Bestanden*?
   Zo niet, dan staat bij *Hulp nodig?* de terugweg via *Nieuw › Bestand uploaden*.
3. **Het filmpje** komt uit les 2 van 3MWb en 3MWWE en toont een ander zip-bestand. Het klikpad is
   hetzelfde. Wil je een eigen filmpje, zet het dan in `assets/video/` met dezelfde naam.
4. **4ORLOa** heeft deze les op donderdag (7de uur): tot de deadline op vrijdag blijft er weinig tijd.
5. **Knopnamen van jouw scherm.** *Nieuwe map* en *Bestand uploaden* komen uit jouw screenshot van
   04-10-2026; de Drive-help schrijft *Map* en *Bestanden uploaden*. De les van 3MWb en 3MWWE gebruikt
   nog de namen uit de help.
