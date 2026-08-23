#!/usr/bin/env python3
"""A kimenet TIPUS-FELISMERESE (box-helper.detect_kind) -- mindket iranyban.

MIERT LETEZIK EZ A TESZT: a valasz-ut korabban HARDKODOLTAN
`word/document.xml`-t olvasott a legyartott fajlbol, es egy xlsx/pptx eseten korai `return`-t
adott -- a kimenet `outcome="nem-docx"`, `written=null` lett, MIKOZBEN a fajl mar ott volt a
lemezen. Az a csendes null volt az, ami miatt ket fajl eszrevetlenul mas uton keszult el.
A felismeres tehat nem kenyelmi funkcio: ez az, ami a valasz-utat beszelove teszi.

*** A NEGATIV AG ITT NEM DISZITES: *** ha a felismeres MINDENRE mondana valamit, akkor a
pozitiv talalatai sem allitananak semmit. Ezert a szemet-bemenet es a "zip, de egyik sem"
eset ugyanolyan sullyal szerepel, mint a harom valodi formatum.
"""
import importlib.util
import io
import os
import sys
import zipfile

HELPER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "box-helper.py")
spec = importlib.util.spec_from_file_location("box_helper", HELPER)
box_helper = importlib.util.module_from_spec(spec)
spec.loader.exec_module(box_helper)

hibak = []
osszes = 0


def check(cimke, felteteles, reszlet=""):
    # A darabszamot a `check` SZAMOLJA, nem en irom bele a vegen: egy kezzel karbantartott
    # osszeg akkor csuszik el, amikor egy esetet hozzaadnak -- es akkor a teszt sajat
    # jelentese lesz az elso hamis allitas benne.
    global osszes
    osszes += 1
    print(f"  {'ok  ' if felteteles else 'BUKAS'}  {cimke}" + (f"   {reszlet}" if reszlet and not felteteles else ""))
    if not felteteles:
        hibak.append(cimke)


def csomag(*reszek):
    """Egy minimalis ZIP a megadott resz-nevekkel. A tartalom nem szamit: a felismeres a
    RESZ-NEVEKBOL dolgozik, mert azok azonositjak az OOXML-formatumot."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for r in reszek:
            z.writestr(r, "<a/>")
    return buf.getvalue()


print("[1] a harom valodi formatum")
for varT, resz in (("docx", "word/document.xml"), ("xlsx", "xl/workbook.xml"), ("pptx", "ppt/presentation.xml")):
    kind, fo = box_helper.detect_kind(csomag("[Content_Types].xml", resz))
    check(f"{resz} -> {varT}", kind == varT and fo == resz, f"kapott: {kind}/{fo}")

print("\n[2] NEG. KONTROLL: amit NEM szabad felismernie")
kind, fo = box_helper.detect_kind(b"ez egyaltalan nem zip")
check("nem-zip bemenet -> ismeretlen", kind == "ismeretlen" and fo is None, f"kapott: {kind}")
kind, fo = box_helper.detect_kind(csomag("valami/mas.xml"))
check("ervenyes zip, de egyik formatum sem -> ismeretlen", kind == "ismeretlen", f"kapott: {kind}")
kind, fo = box_helper.detect_kind(b"")
check("ures bemenet -> ismeretlen", kind == "ismeretlen", f"kapott: {kind}")

print("\n[3] a sorrend szamit, ha tobb fo resz is jelen van")
# Egy csomag, ami MINDHAROM fo reszt hordozza, nem valos eset -- de ha eloall, a felismeresnek
# DETERMINISZTIKUSAN kell valaszolnia, nem a zip belso sorrendjetol fuggoen.
elso = box_helper.detect_kind(csomag("word/document.xml", "xl/workbook.xml", "ppt/presentation.xml"))[0]
masodik = box_helper.detect_kind(csomag("ppt/presentation.xml", "xl/workbook.xml", "word/document.xml"))[0]
check("a valasz fuggetlen a zip belso sorrendjetol", elso == masodik, f"{elso} vs {masodik}")

print(f"\nellenorzesek: {osszes - len(hibak)} ok, {len(hibak)} bukas")
sys.exit(1 if hibak else 0)
