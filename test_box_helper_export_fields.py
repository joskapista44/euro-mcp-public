#!/usr/bin/env python3
"""box_helper.extract_export_fields / _fully_unescape.

MIERT LETEZIK EZ A TESZT: a `text` valtozo (a mentett word/document.xml tag-mentesitett valtozata)
minden bekezdest ELVALASZTO JEL NELKUL fuz ossze -- az egyetlen megbizhato hatar a NEGY ISMERT
marker-prefix sajat maga. ES a marker-tartalom a forrasatol fuggoen KULONBOZO melysegben van
XML/HTML-entitas-escapelve (elo Document Serverrel MERVE, 2026-08-17): a `documentStats`/
`GetCustomProperties` JSON.stringify-kimenete EGY reteget kap (csak az AddText() sajat
XML-escapeleset), a `ToMarkdown()`/`ToHtml()` eredmenye MAR ONMAGABAN escapelve erkezik, tehat
KETTO-HAROM reteget kap -- egy fix passz-szam ROSSZ volna az egyikre. A `_fully_unescape` ezert
egy hatarolt fixpont-ciklus, nem egy hardkodolt ismetlesszam.
"""
import importlib.util
import os
import sys

HELPER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "box-helper.py")
spec = importlib.util.spec_from_file_location("box_helper", HELPER)
box_helper = importlib.util.module_from_spec(spec)
spec.loader.exec_module(box_helper)

hibak = []
osszes = 0


def check(cimke, felteteles, reszlet=""):
    global osszes
    osszes += 1
    print(f"  {'ok  ' if felteteles else 'BUKAS'}  {cimke}" + (f"   {reszlet}" if reszlet and not felteteles else ""))
    if not felteteles:
        hibak.append(cimke)


print("[1] nincs marker a szovegben -> mindharom mezo None")
fields = box_helper.extract_export_fields("EURO-MCP bemeneti dokumentumsima szoveg, marker nelkul")
check("exportResult None", fields["exportResult"] is None, str(fields))
check("stats None", fields["stats"] is None, str(fields))
check("customProperties None", fields["customProperties"] is None, str(fields))

print("\n[2] egyetlen __STATS_JSON__ marker, EGY retegnyi escapeles (JSON.stringify -> AddText)")
# MERT elo alak (2026-08-17): egy `"` az XML-ben `&quot;`-kent all, EGY escape-reteg.
raw = 'elozo szoveg__STATS_JSON__:{&quot;PageCount&quot;:1,&quot;WordsCount&quot;:5}'
fields = box_helper.extract_export_fields(raw)
check("stats parszolt JSON, nem string", fields["stats"] == {"PageCount": 1, "WordsCount": 5}, str(fields["stats"]))
check("exportResult None marad (nincs export-marker)", fields["exportResult"] is None)

print("\n[3] __CUSTOM_PROPERTIES_JSON__ marker a STATS UTAN, hatarolva a kovetkezo elott")
raw = (
    'szia__STATS_JSON__:{&quot;a&quot;:1}'
    '__CUSTOM_PROPERTIES_JSON__:{&quot;b&quot;:2}'
)
fields = box_helper.extract_export_fields(raw)
check("stats CSAK a sajat tartalmat kapja, nem folyik at a customPropertiesbe",
      fields["stats"] == {"a": 1}, str(fields["stats"]))
check("customProperties a maga tartalmat kapja", fields["customProperties"] == {"b": 2}, str(fields["customProperties"]))

print("\n[3b] __STATS_JSON__ marker, STRING-ERTEK amp ES escapelt idezojellel (korabbi lelet, 2026-08-17)")
# JSON.stringify({cim: 'R&D "terv"', szo: 12}) JS-kimenete: {"cim":"R&D \"terv\"","szo":12}
# -- a beagyazott idezojelet a JSON MAGA \"-kent escapeli, EZUTAN escapeli az AddText() a
# TELJES stringet XML-re (minden szo szerinti " -> &quot;, minden & -> &amp;, a backslash
# erintetlen). A raw ertek tehat, PONTOSAN ahogy az XML-ben all:
raw = '__STATS_JSON__:{&quot;cim&quot;:&quot;R&amp;D \\&quot;terv\\&quot;&quot;,&quot;szo&quot;:12}'
fields = box_helper.extract_export_fields(raw)
check("stats.cim HELYESEN 'R&D \"terv\"' (nem marad entitas, es a JSON szerkezet nem serul)",
      fields["stats"] == {"cim": 'R&D "terv"', "szo": 12}, str(fields["stats"]))

print("\n[3c] ugyanaz, a MASIK realis rezsimben: csak &/</> escapelve, az idezojelek NYERSEK")
# Egy korabbi opcionalis kiegeszites (2026-08-17): a ket VALODI escapeles-rezsim kozul ez a masik --
# a strukturalis ES a tartalom-belso idezojel EGYFORMAN nyers marad, csak a `&` lesz `&amp;`.
raw = '__STATS_JSON__:{"cim":"R&amp;D \\"terv\\"","szo":12}'
fields = box_helper.extract_export_fields(raw)
check("stats.cim UGYANAZ az ertek, ha az idezojelek eleve nyersen erkeznek",
      fields["stats"] == {"cim": 'R&D "terv"', "szo": 12}, str(fields["stats"]))

print("\n[4] __EXPORT_MARKDOWN__, KETTO retegnyi escapeles (ToMarkdown() mar escapel, AddText() meg egyszer)")
# MERT elo alak (2026-08-17, elo Document Server): egy eredeti '\"' -> raw XML-ben &amp;quot;
raw = 'elozo__EXPORT_MARKDOWN__:Hello &amp;quot;quoted&amp;quot; world &amp;amp; more'
fields = box_helper.extract_export_fields(raw)
check("exportResult.format markdown", fields["exportResult"]["format"] == "markdown")
check("exportResult.content TELJESEN feloldva (nem all meg entitas)",
      fields["exportResult"]["content"] == 'Hello "quoted" world & more', str(fields["exportResult"]))

print("\n[5] __EXPORT_HTML__, VEGYES melyseg egy stringen belul (a strukturalis tag 2, a tartalom-kar 3 reteg)")
# MERT elo alak (2026-08-17): a <p> tag ket, a szoveg-& harom escape-reteget kap -- lasd a
# fuggveny sajat kommentjet arrol, MIERT ter el a ketto.
raw = 'x__EXPORT_HTML__:&amp;lt;p&amp;gt;HTML export &amp;amp;amp; &amp;amp;lt;tags&amp;amp;gt; test&amp;lt;/p&amp;gt;'
fields = box_helper.extract_export_fields(raw)
check("exportResult.format html", fields["exportResult"]["format"] == "html")
check("exportResult.content: a strukturalis tag ES a tartalom-kar EGYARANT teljesen feloldva",
      fields["exportResult"]["content"] == "<p>HTML export & <tags> test</p>", str(fields["exportResult"]))

print("\n[6] hibas/nem-JSON payload a STATS markeren -> nevesitett parseError, NEM dob kivetelt")
raw = "__STATS_JSON__:ez nem json"
fields = box_helper.extract_export_fields(raw)
check("stats.parseError True", isinstance(fields["stats"], dict) and fields["stats"].get("parseError") is True, str(fields["stats"]))
check("stats.raw megorzi az eredeti (feloldott) szoveget", fields["stats"].get("raw") == "ez nem json", str(fields["stats"]))

print("\n[7] _fully_unescape: mar-tiszta szovegen NO-OP (nulla escape-reteg is biztonsagos)")
check("mar tiszta szoveg valtozatlan", box_helper._fully_unescape("sima szoveg, semmi entitas") == "sima szoveg, semmi entitas")

print("\n[8] _fully_unescape: hatarolt (nem vegtelen ciklus egy patologikus bemeneten)")
# Egy string, ami SOHA nem stabilizalodik (minden passz utan MEG mindig van benne "&amp;" -- ide
# szandekosan egy olyan bemenetet adunk, ami tobbszori unescape utan is valtozna, ha nem lenne
# hatar). A hatar maga a MERCE: 5 iteracion belul terjunk vissza, ne fagyjon le.
import time
kezdet = time.monotonic()
eredmeny = box_helper._fully_unescape("&amp;" * 50)
eltelt = time.monotonic() - kezdet
check("hatarolt ido alatt visszater (nem vegtelen ciklus)", eltelt < 1.0, f"{eltelt:.3f}s")

print(f"\nellenorzesek: {osszes - len(hibak)} ok, {len(hibak)} bukas")
sys.exit(1 if hibak else 0)
