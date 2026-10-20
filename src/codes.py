"""Pruefcodes in einfachem Deutsch (Auftrag 003, Schema 4 `Aenderung.grund_einfach`).

EINE Tabelle (`SAETZE`): Pruefcode -> kurzer Satz fuer Menschen ohne Fachwissen. Zusammengefuehrt aus den frueheren Tabellen
des Planers (`_GRUND`, `_CODE_SATZ`, `_A1_SATZ`) und den Hinweis-Codes des PDF-Lesers. Die Saetze enthalten nie Dokumentinhalte
(Regel 16). Abgeleitete Codes: "a1_<art>" (die Excel weicht vom alten PDF-Text ab) und "sperre:<hinweis>" (eine unsichere Stelle
im Absatz sperrt die Automatik).

`grund_einfach(ae)` bildet daraus EINEN Satz je Aenderung, der wichtigste Grund zuerst (`RANG`).
"""
from __future__ import annotations

from typing import Any

from .models import NICHT, PRUEFEN, UEBERNOMMEN

SAETZE: dict[str, str] = {
    # ---- Zuordnung (Reissverschluss, Planer)
    "mehrdeutig": "Der alte Absatz steht mehrmals in der Excel; die Reihenfolge entscheidet nicht, welche Stelle gemeint ist.",
    "widerspruch": "Der alte Absatz steht in der Excel an einer Stelle, die nicht zur Reihenfolge im PDF passt.",
    "keiner": "Der alte Absatz wurde in der Excel nicht genau gefunden.",
    "excel_weicht_ab": "Der alte Absatz aus dem PDF steht so nicht in der Excel; die Stelle ist nur ähnlich.",
    "nicht_gefunden": "Der Absatz wurde in der Excel nicht gefunden.",
    "zu_unsicher": "Die passende Stelle in der Excel ist zu unsicher.",
    "schon_eingetragen": "Die Änderung steht schon so in der Excel.",
    "vielleicht_doppelt": "Ein gleicher kurzer Absatz steht schon in der Excel; vielleicht wäre er doppelt.",
    "kurz": "Sehr kurzer Text, die Stelle ist nicht sicher.",
    "section_unpassend": "Die Section passt nicht zur gefundenen Stelle.",
    "section_unsicher": "Die Section-Nummer steht als Zahl in der Excel und ist nicht eindeutig (2.1 oder 2.10?).",
    "versteckt": "Die Stelle in der Excel ist ausgeblendet.",
    "kopfspalte": "Die Stelle liegt in der Nummern- oder Titelspalte und nicht im Fließtext.",
    "layout_unsicher": "Der Aufbau des Blatts wurde nur unsicher erkannt.",
    "label_spalte": "Die Nummer oder der Buchstabe steht auch in den Zeilen darunter; dort bleibt er unverändert. Bitte prüfen.",
    # ---- Stelle im Absatz (Abgleich)
    "naht_unklar": "Die genaue Stelle im Excel-Absatz konnte nicht sicher bestimmt werden.",
    "excel_umbruch": "In der Excel steht an der Stelle ein Zeilenumbruch.",
    "mehrere_zellen": "Die Änderung betrifft mehrere Excel-Zellen.",
    "leer": "Die Änderung hat keinen Inhalt.",
    "teil_absatz": "In der Excel steht im selben Absatz mehr Text, der im PDF nicht gestrichen ist.",
    "gruppe": "Gehört zu anderen Änderungen, die nur gemeinsam eingetragen werden.",
    # ---- Neue Absaetze und Zeilen
    "anker_unsicher": "Der Absatz davor konnte in der Excel nicht sicher zugeordnet werden; die Einfügestelle ist nur ungefähr "
                      "bestimmbar.",
    "anker_teil": "Der PDF-Text deckt nur ein Stück des längeren Excel-Absatzes ab (Seitenumbruch?); der neue Absatz könnte "
                  "mitten im Satz landen.",
    "anker_luecke": "Zwischen der Ankerstelle und dem nächsten zugeordneten Absatz steht in der Excel Text, der keinem PDF-Absatz "
                    "zugeordnet ist.",
    "kein_anker": "Für den neuen Absatz gibt es keine Stelle in der Excel (kein Nachbar-Absatz gefunden).",
    "blatt_ohne_anker": "Die Einfügestelle läge in einem Blatt, in dem kein Absatz sicher zugeordnet ist; dort kommt keine neue Zeile hin.",
    "keine_inhaltsspalte": "Die Spalte mit dem Text wurde in der Excel nicht erkannt.",
    "zeilen_risiko": "Beim Einfügen der neuen Zeile könnte in der Excel etwas verrutschen. Bitte prüfen.",
    "ausserhalb_cde": "Gehört nicht zur CDE (Deckblatt, Vorspann oder Anhang des PDF).",
    "spec_unklar": "Die Spec des PDFs wurde in der CDE nicht eindeutig gefunden. Bitte prüfen.",
    # ---- Neuer Text
    "steuerzeichen": "Der neue Text enthält Steuerzeichen und kann nicht in die Excel geschrieben werden.",
    "neutext_verdaechtig": "Der neue Text könnte aus dem PDF falsch gelesen worden sein.",
    "zahlformat": "Die Zelle ist eine Zahl, der neue Text nicht; das Zahlenformat ginge verloren.",
    # ---- Eintragen und Gegenprobe (Applier)
    "abgewaehlt": "Ein Mensch hat die Änderung abgewählt.",
    "ohne_absatz": "Zur Änderung fehlt der PDF-Absatz für die Nachkontrolle; sie wird nicht automatisch eingetragen.",
    "gegenprobe_text": "Nachkontrolle: Der Text in der Excel passt danach nicht zum PDF.",
    "gegenprobe_aussen": "Nachkontrolle: Außerhalb des Absatzes hat sich etwas verändert.",
    "gegenprobe_umbruch": "Nachkontrolle: Es ist ein Zeilenumbruch dazugekommen.",
    "gegenprobe_zelle": "Nachkontrolle: Die Zelle wurde nicht gefunden.",
    "gegenprobe_gruppe": "Nachkontrolle: Eine Änderung derselben Gruppe passt nicht zum PDF.",
    "gegenprobe_neuer_absatz": "Nachkontrolle: Der neue Absatz steht nicht als eigener Absatz in der Zelle.",
    "gegenprobe_verlust": "Nachkontrolle: Ein Zeichen der Excel wäre verloren gegangen.",
    "gegenprobe_alt": "Nachkontrolle: Der alte Text (durchgestrichen) passt nicht zum PDF.",
    "gegenprobe_format": "Nachkontrolle: Die Markierung (rot / blau durchgestrichen) passt nicht zum PDF.",
    # ---- Hinweise des PDF-Lesers (als "sperre:<code>" oder nur Hinweis)
    "unbekannte_farbe": "Im Absatz gibt es eine Farbe, die nicht in der Legende steht.",
    "teilweise_markiert": "Ein Wort ist nur zum Teil markiert.",
    "gemischt_blau_strich": "Ein Wort ist blau und durchgestrichen (widersprüchlich).",
    "gemischt_widerspruch": "Ein Wort ist zugleich als neu und als gestrichen markiert.",
    "linie_unklar": "Eine Linie im Text ist nicht eindeutig (Zierlinie oder Markierung?).",
    "linienfarbe_unklar": "Die Farbe einer Linie im Text konnte nicht gelesen werden.",
    "kommentar": "Auf dem Text liegt ein PDF-Kommentar mit unklarer Bedeutung.",
    "tabelle": "Der Text steht in einer Tabelle; Zellen können vermischt sein.",
    "mehrspaltig": "Der Text steht in mehreren Spalten nebeneinander; Sätze können vermischt sein.",
    "design_flaeche": "Eine Farbfläche hinter dem Text sieht nach Gestaltung aus.",
    "design_blau": "Die blaue Überschrift ist Gestaltung, keine Änderung.",
    "farbe_uneinig": "Zwei Leseverfahren sehen die Farbe an dieser Stelle verschieden.",
    "tinte_unerklaert": "Im Seitenbild gibt es Farbe, die keiner Markierung zugeordnet werden konnte.",
    "fortsetzung_unklar": "Der Absatz geht vielleicht über einen Seitenumbruch weiter; das ist nicht sicher.",
    "absatzgrenze_unklar": "Die Grenze zwischen zwei Absätzen ist im PDF nicht sicher erkennbar.",
    "trennung_unklar": "Eine Silbentrennung am Zeilenende konnte nicht sicher aufgelöst werden.",
    "rand_unklar": "Die Markierung endet nicht sauber an einer Wortgrenze.",
    "hochzahl_markiert": "Eine kleine hochgestellte Marke (z. B. Fußnote) ist markiert.",
    "ligatur": "Das PDF enthält verbundene Buchstaben (Ligaturen); der Text kann anders geschrieben sein.",
    "nicht_lesbar": "Einzelne Zeichen im PDF konnten nicht gelesen werden.",
    "link": "Ein Link im PDF: seine blaue Farbe zählt nicht als Änderung.",
    "unterstrichen": "Eine schwarze Unterstreichung hat keine Bedeutung und zählt nicht.",
    "farbe_ohne_linie": "Ein Wort ist farbig, aber nicht unter- oder durchgestrichen; es zählt als unverändert.",
    "hinterlegt": "Ein Wort ist farbig hinterlegt; das zählt nicht als Änderung.",
    "verschmolzen": "Zwei Wörter wurden im PDF zusammen gelesen.",
    # ---- Auftrag 004: kein Vorschlag statt eines unsicheren Vorschlags
    "kein_vorschlag": "Für diese Stelle gibt es keinen sicheren Vorschlag; bitte von Hand eintragen.",
}

# Teil-Saetze fuer "a1_<art>" (worin die Excel vom alten PDF-Text abweicht); frueher planner._A1_SATZ
A1_ART: dict[str, str] = {
    "leerraum": "Leerzeichen", "satzzeichen": "ein Satzzeichen", "gross_klein": "Groß-/Kleinschreibung",
    "ziffer": "eine Ziffer", "buchstabe": "ein Buchstabe", "pdf_laenger": "im PDF steht mehr Text",
    "excel_laenger": "in der Excel steht mehr Text", "abweichung": "der Text",
}

# Wichtigster Grund zuerst: Gruende, die das Eintragen ganz verhindern, vor Gruenden fuer "pruefen", vor reinen Hinweisen
RANG: tuple[str, ...] = (
    "ausserhalb_cde", "spec_unklar", "steuerzeichen", "schon_eingetragen", "abgewaehlt", "nicht_gefunden", "zu_unsicher", "kein_anker",
    "blatt_ohne_anker", "keine_inhaltsspalte", "gegenprobe_", "ohne_absatz", "sperre:", "excel_weicht_ab", "a1_", "widerspruch",
    "mehrdeutig", "keiner", "label_spalte", "teil_absatz", "mehrere_zellen", "excel_umbruch", "naht_unklar", "neutext_verdaechtig",
    "zahlformat", "zeilen_risiko", "anker_", "layout_unsicher", "section_", "kopfspalte", "versteckt", "kurz",
    "vielleicht_doppelt", "gruppe",
)


def satz(code: str) -> str:
    """Satz zu einem Pruefcode (leer, wenn es keinen gibt)."""
    if code in SAETZE:
        return SAETZE[code]
    if code.startswith("a1_"):
        art = A1_ART.get(code[3:], "")
        return f"Der alte Text steht in der Excel etwas anders ({art})." if art else ""
    if code.startswith("sperre:"):
        innen = satz(code[len("sperre:"):])
        return f"Unsichere Stelle im Absatz: {innen}" if innen else ""
    if code.startswith("kommentar_"):
        return SAETZE["kommentar"]
    return ""


def _rang(code: str) -> int:
    for i, r in enumerate(RANG):
        if code == r or (r.endswith(("_", ":")) and code.startswith(r)):
            return i
    return len(RANG)


# Kurzfassung fuer die Dashboard-Karten (hoechstens KURZ_MAX_WOERTER Woerter, Welle 3): wo der Satz oben laenger ist
KURZ: dict[str, str] = {
    "mehrdeutig": "Der alte Absatz steht mehrmals in der Excel; unklar, welche Stelle gemeint ist.",
    "widerspruch": "Der alte Absatz steht in der Excel an einer unerwarteten Stelle.",
    "excel_weicht_ab": "Die Excel weicht hier vom alten PDF-Text ab. Vorschlag bitte prüfen.",
    "label_spalte": "Nummer oder Buchstabe steht auch in den Zeilen darunter. Bitte prüfen.",
    "anker_unsicher": "Die Stelle für den neuen Absatz ist nur ungefähr bekannt.",
    "anker_teil": "Der neue Absatz könnte mitten in einem längeren Excel-Absatz landen.",
    "anker_luecke": "Vor dem neuen Absatz steht in der Excel Text ohne Gegenstück im PDF.",
    "blatt_ohne_anker": "Die Stelle läge in einem Nebenblatt; dort kommt keine neue Zeile hin.",
    "sperre": "Eine Stelle im Absatz ist im PDF unklar markiert. Bitte im PDF prüfen.",
}
KURZ_MAX_WOERTER = 15


def kurz(code: str) -> str:
    """Satz zu einem Pruefcode fuer die Dashboard-Karte (kurz, ohne Kennungen)."""
    if code in KURZ:
        return KURZ[code]
    if code.startswith("sperre:"):
        return KURZ["sperre"] if satz(code) else ""
    return satz(code)


def _kurz_genug(text: str) -> str:
    """Erster Satzteil eines frueheren Grunds ohne Klammerzusatz -- nur, wenn er kurz genug ist."""
    t = text.split("; ")[0].split(" (")[0].strip()
    t = t if t.endswith((".", "!", "?")) else t + "."
    return t if 1 < len(t.split()) <= KURZ_MAX_WOERTER else ""


def grund_einfach(ae: Any) -> str:
    """EIN kurzer Satz zum Status einer Aenderung (wichtigster Grund zuerst), fuer die Dashboard-Karten."""
    gruende = sorted((c for c in ae.pruefcodes if kurz(c)), key=_rang)
    if gruende:
        haupt = kurz(gruende[0])
        if ae.status == UEBERNOMMEN:
            return "Wird eingetragen. " + haupt if len(haupt.split()) <= KURZ_MAX_WOERTER - 2 else haupt
        return haupt
    if ae.status == UEBERNOMMEN:
        return "Wird eingetragen: der alte Text steht genau so in der Excel."
    erster = _kurz_genug(ae.grund or "")
    if ae.status == PRUEFEN:
        return erster or "Vorschlag: bitte prüfen und freigeben."
    if ae.status == NICHT:
        return erster or "Wird nicht eingetragen; bitte die Stelle von Hand prüfen."
    return erster
