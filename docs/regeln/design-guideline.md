> **Hinweis (CDE-Änderungsübernahme):** Diese Datei ist eine unveränderte Kopie der Design-Guideline der bess-pipeline.
> Gilt als Stilvorbild. Unser Dashboard (`frontend/`) übernimmt Aufbau, Variablen-Prinzip, Dark Mode und Apple-Stil,
> nutzt aber ein neutrales Blau (`#0071E3`) statt RR-Blau und **kein** Rolls-Royce-Logo (eigenes Seitenprojekt).
> Abweichungen bewusst: Port 5100, kein Logo, andere Tabs. Nico Lehr.

# BESS-Pipeline Dashboard — Design Guideline

Diese Datei ist die verbindliche Design-Referenz für alles, was an diesem Dashboard neu gebaut oder überarbeitet wird. Sie wurde aus dem bestehenden Frontend (`index.html`, `style.css`, `script.js`) extrahiert — **beschreibend, nicht erfunden**. Ziel: jede neue Komponente fügt sich nahtlos neben das Bestehende, als käme sie aus einem Guss.

**Leitprinzip: Apple-Design-Sprache, RR-Branding.** Minimalistisch, ruhig, hochwertig. Viel Weißraum, klare Typo-Hierarchie, dezente Schatten statt harte Kanten, weiche Animationen. Rolls-Royce-Blau ist die einzige Marken-Akzentfarbe — nichts schreit, alles fügt sich ein.

---

## 1. Farbsystem

Alle Farben werden **ausschließlich über CSS-Variablen** referenziert, nie als Hex-Wert direkt im Markup oder in Komponenten-CSS hartcodiert (Ausnahme: die `:root`-Definition selbst und die washed-out Status-Hintergründe, siehe 1.3).

### 1.1 Markenfarbe (RR-Blau)

```css
--rr-blue:       #00498F;   /* Primäre Akzentfarbe: CTAs, aktive Zustände, Links, Fokus */
--rr-blue-dark:  #00335E;   /* Gradient-Partner für Primary-Button, Hover-Vertiefung */
--rr-blue-light: #EAF1F8;   /* Washed Background für aktive Flächen (z. B. Stage "active", Privacy-Note) */
```

Reserviert für: Primary-Button-Gradient, aktive Pipeline-Stage, Fokus-Ringe (`box-shadow: 0 0 0 3px var(--rr-blue-light)`), Icon-Akzente, Badges/Counter, Links (`.doc-link`). Sparsam einsetzen — es ist die einzige "laute" Farbe im System.

### 1.2 Ampel-/Status-Farben (die zentrale fachliche Semantik)

Anders als im alten Sport-Dashboard gibt es hier **keine dekorativen Kategorie-Farben**. Das Ampel-System *ist* die Farbsemantik des Projekts:

```css
--green: #2E7D46;   /* Grün: Wert verifiziert / machbar / Confidence hoch (≥0.85) */
--amber: #B8790C;   /* Gelb: unsicher / Review nötig / Confidence mittel (0.60–0.84) */
--red:   #B3261E;   /* Rot: nicht machbar / Confidence niedrig (<0.6) / Konflikt / Fehler */
```

**Diese drei Farben werden ausschließlich für ihre fachliche Bedeutung verwendet** (Ampel-Dots, Confidence-Badges, Pipeline-Stage-Status, Tabellen-Konfliktmarkierung, Toggle-„aktiv"-Zustand). Nie dekorativ oder für sportfremde/beliebige Kategorien zweckentfremden — das war der Fehler im alten Guide, den es hier zu vermeiden gilt.

**Zwei parallele Farbwasch-Muster existieren aktuell nebeneinander** (beide gültig, aber bei neuen Komponenten das passende wählen):
- **Fixe Hex-Washes** für große Flächen: `#E7F3EA` (Grün-Wash), `#FDF1E3` (Amber-Wash), `#FBEAE9` (Rot-Wash) — genutzt bei `.status-green/-amber/-red` und `.stage.complete/.error`.
- **Dynamische `rgba()`-Washes** (10 %) direkt aus der Variable abgeleitet, z. B. `rgba(46, 125, 70, 0.1)` für `.confidence-badge.high`, `rgba(179, 38, 30, 0.05)` für `tr.conflict` — genutzt bei kleinteiligeren/inline Elementen.

### 1.3 Neutrale Skala (Grau + Ink)

```css
--white:   #FFFFFF;   /* Seiten-/Karten-Hintergrund */
--grey-50: #FAFAFA;   /* Subtile Flächen-Abstufung (Gradients, Footer) */
--grey-100:#F5F5F7;   /* Standard-Sekundärfläche (Cards, Dropzone-Hintergrund, Pills) */
--grey-150:#E8E8ED;   /* Borders, Trennlinien, Hover-Flächen */
--grey-300:#D2D2D7;   /* Deaktivierte Elemente, Input-Borders, stärkere Trennlinien */
--grey-500:#86868B;   /* Sekundärtext, Meta-Infos, Labels */
--ink:     #1D1D1F;   /* Primärtext */
```

### 1.4 Dark Mode (`[data-theme="dark"]` sowie `@media (prefers-color-scheme: dark)` als Fallback)

```css
--white:   #1C1C1E;
--grey-50: #2C2C2E;
--grey-100:#232325;
--grey-150:#323234;
--grey-300:#48484A;
--grey-500:#98989D;
--ink:     #F5F5F7;
```

**Wichtiger aktueller Zustand — bewusst beim Facelift prüfen:** `--rr-blue`, `--green`, `--amber`, `--red` sowie die Hex-Washes (`#E7F3EA` etc.) bekommen aktuell **keine eigenen Dark-Mode-Werte** — sie bleiben in beiden Themes identisch. Das kann so bleiben (Markenfarbe/Ampelfarben sollen erkennbar konstant sein), aber die *Kontrastwirkung* auf dunklem Hintergrund sollte beim Überarbeiten geprüft werden, insbesondere die Hex-Washes (`#E7F3EA` etc.), die auf `--white: #1C1C1E` vermutlich zu hell/unpassend wirken. Falls angepasst: dark-mode-spezifische `--*-wash`-Varianten einführen statt der Basisfarbe selbst.

Jede neue Komponente muss in **beiden** Modi geprüft werden. Nie eine Farbe nur für ein Theme definieren, ohne das Gegenstück zu testen.

### 1.5 Regeln

- `--rr-blue` ist reserviert für Marken-/Interaktionsakzente, nicht für Status. Ampelfarben sind reserviert für Status, nicht für Dekoration. Diese Trennung nie vermischen.
- Rot/Grün/Gelb nur für echte Ampel-/Confidence-/Fehlerzustände, nicht dekorativ (z. B. nicht als beliebige Chart-Füllfarbe ohne Bezug zur Machbarkeitsaussage).
- Keine neuen Grundfarben ohne Rücksprache einführen — das bestehende Set (RR-Blau + Ampel + Neutral) deckt praktisch alle Fälle ab.

---

## 2. Typografie

### 2.1 Font-Stack

```css
font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
```

Systemfont, kein Web-Font-Import — bewusst leicht, schnell, nativ aussehend. Nicht ändern.

### 2.2 Schriftgrößen & -gewichte (bestehende Skala, wiederverwenden statt neue erfinden)

| Verwendung | Größe | Weight | Letter-Spacing |
|---|---|---|---|
| Hero-Titel (`.intro h1`) | 42px | 700 | -1px |
| Settings-Panel-Titel (`h2`) | 20px | default (700 via Browser-H2) | — |
| Dropzone-Titel | 18px | 600 | -0.3px |
| Hero-Beschreibung (`.intro p`) | 16px | normal | — |
| Progress-Header (`h3`) | 16px | default | — |
| Brand-Name (Header-Logo-Text) | 16px | 600 | -0.3px |
| Datei-/Feld-Titel (`.file-details h3`) | 15px | 600 | — |
| Button-Text (`.btn`) | 15px | 600 | -0.2px |
| Stat-Wert (KPI, `.stat-value`) | 32px | 700 | -1px |
| Footer-Titel | 15px | 700 | -0.3px |
| Stage-Label | 13px | 600 | — |
| Field-Name (Tabelle) | Body-Größe | 600 | — |
| Status-Pill / Progress-Percent | 13–14px | 500–600 | — |
| Label / Kicker (uppercase, z. B. `.field-label`, `.settings-group h3`) | 12–13px | 700 | 0.05em, uppercase |
| Badges (`.status-badge`, `.confidence-badge`) | 11–12px | 600 | 0.05em (nur `.status-badge`), uppercase (nur `.status-badge`) |
| Meta-/Sekundärtext (`.brand-sub`, `.file-details p`, `.footer-text`) | 12–13px | 500 (brand-sub) / normal | — |

**Hinweis:** Anders als im alten Guide gibt es aktuell **kein** `tabular-nums` für Kennzahlen (z. B. `.stat-value`, Confidence-Werte). Da hier viele numerische Werte (Kapazität, kW, Prozentsätze) nebeneinander stehen, lohnt sich `font-variant-numeric: tabular-nums` als Ergänzung beim Facelift, damit Zahlen nicht "wackeln" — bislang nicht umgesetzt, aber konsistent mit dem Apple-Prinzip.

Gewichte liegen bewusst oft zwischen Standardstufen (500, 600, 620 wirken in der Praxis austauschbar mit 600 — im Code wird durchgängig auf **600** als "betont" und **700** als "stark betont" gesetzt, kein 620/640 wie im alten Guide). Bei neuen Komponenten bei 400/500/600/700 bleiben.

---

## 3. Layout & Abstände

- Max-Breite Content-Container (`.app-main`): `900px`, zentriert, Padding `48px 24px` (Desktop) / `32px 16px` (Mobile).
- Header-Padding: `20px 32px` (Desktop) / `16px 16px` (Mobile).
- Sektions-Abstand (`margin-bottom` zwischen Hero/Filelist/Pipeline/Action-Row/Progress/Results/Stats): durchgängig `48px`, Hero etwas größer (`60px`).
- Karten-/Listen-Grid-Abstand (`gap`): `12–24px` je nach Dichte (File-List `12px`, Stats-Grid `16px`, Action-Row `16px`).
- Karten-Innenabstand: Stat-Card/Stage/File-Item `16–24px`, Dropzone `48px 24px` (bewusst großzügig als visueller Einstiegspunkt), Settings-Panel-Body `24px`.
- Settings-Panel: fixe Breite `max-width: 420px`, slide-in von rechts (`transform: translateX(100%)` → `0`).
- Alles responsiv: Grids fallen auf Mobile auf 1 Spalte (Stats-Grid, Stage-List, Action-Row), siehe Abschnitt 8.
---

## 4. Radien (Border-Radius) — Systematik

| Element | Radius |
|---|---|
| Dropzone | 16px |
| Karten/Sektionen (Stat-Card, Stage, File-Item, Progress-Section, Table-Wrapper, Settings-Header-Inputs) | 12px |
| Buttons (`.btn`), Info-Icon | 12px / 8px (Info-Icon kleiner) |
| Kleine Icons/Badges (File-Icon, Confidence-Badge, Toast) | 6–8px |
| Pills / Buttons / Badges (Status-Pill, Count-Badge, Status-Badge, Toggle-Track) | `999px` bzw. `100px` (voll rund) |
| Kreise (Ampel-Dots, Stage-Dots, Icon-Buttons, Toggle-Thumb) | `50%` |

Nie einen beliebigen Radius erfinden — sich an diese Stufen halten. Grundmuster: **große Flächen 12–16px, Pills/Badges voll rund, funktionale Icons/Kreise rund**.

---

## 5. Schatten & Tiefe

Fünf definierte Schattenstufen (mehr als im alten Guide, hier feiner gestaffelt) — kein sechstes Set erfinden:

```css
--shadow-xs   /* Ruhezustand von Pills/Header: kaum sichtbar */
--shadow-sm   /* Standard-Karten-Ruhezustand */
--shadow-md   /* Hover-Zustand der meisten Karten/Buttons */
--shadow-lg   /* Hover von Primary-Button, Stat-Card */
--shadow-xl   /* Settings-Panel (schwebt über allem) */
```

Toast und Info-Popup nutzen eigene, stärkere Ad-hoc-Schatten (`0 20px 40px rgba(0,0,0,0.3)` bzw. `0 12px 32px rgba(0,0,0,0.15)`), analog zur alten Regel „Modals/Overlays dürfen einen stärkeren eigenen Schatten haben, da sie über allem schweben".

---

## 6. Interaktion & Bewegung

- Standard-Transition für die meisten Elemente: `all 0.3s cubic-bezier(0.4, 0, 0.2, 1)` (Apple-typisches sanftes Ease-Out).
- Einfache Link-/Farbübergänge (`.doc-link`, Info-Popup-Close): `0.2s ease` / `0.2s`.
- Hover auf Karten/Buttons: leichtes Anheben (`translateY(-1px)` bis `-4px`, gestaffelt nach Element-Gewicht: Brand/Pill `-1px`, Buttons/File-Item/Stage `-2px`, Stat-Card `-4px`) + nächsthöhere Schattenstufe.
- Klick/Active-Zustand: `scale(0.92–0.97)` für Icon-Buttons; Primary/Secondary-Buttons nutzen stattdessen einen Ripple-Effekt (`::before` mit expandierendem `rgba(255,255,255,0.3)`-Kreis).
- Eintritts-Animationen (`slideInUp`/`slideInDown`, 0.6s, leicht gestaffelt mit `0.1–0.3s`-Delay) für Hero, Dropzone, File-Section, Pipeline-Tracker, Action-Row, Progress, Results, Stats — jede Hauptsektion "kommt herein", wenn sie erscheint. Bei neuen Sektionen dieses Muster fortführen.
- Dezente Loop-Animationen als Markenmoment: `pulse` (Status-Dot, 2s), `pulse-border` (aktive Pipeline-Stage, 2s) — sparsam, nur für „live"/„aktiv"-Zustände, nie für statische Elemente.
- `prefers-reduced-motion` wird aktuell nicht berücksichtigt — beim Facelift als Verbesserung einplanen.

---

## 7. Komponenten-Bausteine (wiederverwenden, nicht neu erfinden)

Vor dem Bau einer neuen Komponente immer prüfen, ob eines dieser Muster passt:

- **`.stat-card`** — KPI-Kachel mit großem `--rr-blue`-Wert und Label darunter.
- **`.btn-primary`** (RR-Blau-Gradient) / **`.btn-secondary`** (neutral, Border) — primäre vs. sekundäre Aktion.
- **`.status-badge` / `.confidence-badge`** — für Ampel- bzw. Confidence-Zustände, Farbwash-Hintergrund statt voller Farbfläche (siehe 1.2).
- **`.ampel` / `.dot` / `.stage-dot`** — runde Statuspunkte in Grün/Amber/Rot, immer über `--green/--amber/--red`, nie eigene Rottöne.
- **`.info-icon` + `.info-popup`** — einziger Mechanismus für erklärenden Zusatztext (analog zum alten „ⓘ"-Prinzip), kein Fließtext direkt in der Karte.
- **`.stage` (Pipeline-Tracker)** — Zustände `active` (RR-Blau-Wash + Puls-Border), `complete` (Grün-Wash), `error` (Rot-Wash) — dieses Dreier-Zustandsmuster (aktiv/erledigt/Fehler) ist auch für neue Prozessanzeigen das Referenzmuster.
- **`.toast`** — dunkles, schwebendes Feedback-Element unten links, nicht für Dauerzustände, nur für kurze Bestätigungen/Fehler.
- **`.settings-panel`** — Slide-in-Overlay von rechts für Konfiguration, mit `.overlay`-Backdrop-Blur dahinter.

---

## 8. Responsive Design

- Breakpoint: `768px`. Darunter: Stats-Grid und Stage-List auf 1 Spalte, Action-Row und Buttons auf volle Breite, Settings-Panel auf `100%` Breite, Tabellenschrift auf `12px` verkleinert.
- Touch-Targets (Buttons, Icon-Buttons) mindestens `32–40px`.
- Header bleibt sticky (`position: sticky; top: 0`) auch auf Mobile, mit reduziertem Padding.
- Keinen Platz verschwenden. Nutze für die darstellung der Inhalte von den verschiedenen Seiten immer die volle Bildschirmbreite
---

## 9. Ton & Sprache

- Deutschsprachiges internes RR-Tool, kein Sprach-Toggle nötig (anders als beim alten zweisprachigen Landingpage-Konzept).
- UI-Text kurz, sachlich, fachlich präzise (z. B. „Confidence", „Ampel-Verteilung") statt Marketing-Sprache — passt zum bestehenden Ton in `index.html`.

---

## 10. Bekannte technische Unsauberkeiten im aktuellen Code (Hinweis fürs Facelift)

Nicht Teil der Design-Regeln selbst, aber beim Überarbeiten sinnvoll mitzunehmen:

- `.status-badge`, `.status-green`, `.status-amber`, `.status-red` sind aktuell **doppelt** in `style.css` definiert (identischer Inhalt). Beim Facelift konsolidieren.
- Zwei parallele Farbwasch-Muster (fixe Hex-Washes vs. dynamische `rgba()`-Washes, siehe 1.2) könnten auf ein einheitliches Muster (z. B. durchgängig `rgba()` aus der jeweiligen Variable) vereinheitlicht werden — aktuell kein Fehler, aber Inkonsistenz.
- Fehlende Dark-Mode-Werte für RR-Blau/Ampelfarben/Washes (siehe 1.4) — prüfen, ob Kontrast im Dark Mode ausreicht.
- Kein `tabular-nums` für Zahlenwerte (siehe 2.2).
- `prefers-reduced-motion` nicht berücksichtigt (siehe 6.).

---

## 11. Goldene Regel für die KI

> Vor dem Bau von etwas Neuem: Passt eine bestehende Variable / Komponente / Radius-/Schatten-Stufe? Wenn ja → wiederverwenden. Wenn nein → neue Variable nach demselben Namensschema (`--rr-*`, `--grey-*`, `--shadow-*` etc.) ergänzen, nicht hartcodieren.

Diese Guideline ist bewusst **beschreibend** (aus dem echten `index.html`/`style.css` extrahiert), nicht erfunden — jede neue Komponente sollte sich nahtlos ins bestehende Dashboard einfügen, als käme sie aus einem Guss.
