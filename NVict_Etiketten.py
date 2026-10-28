# -*- coding: utf-8 -*-
"""
NVict Etiketten Maker - Professional Edition
Ontwikkeld door NVict Service

Website: www.nvict.nl
Versie: 4.3.0
"""

import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from PIL import Image, ImageTk
import urllib.request
import urllib.error
import json
import re
import threading
import webbrowser
import tempfile
import subprocess
import platform
from datetime import datetime

# Applicatie versie
APP_VERSION = "4.3.0"
UPDATE_CHECK_URL = "https://www.nvict.nl/software/updates/etiketten_version.json"
UPDATE_CHECK_URL_FALLBACK = "https://www.nvict.nl/software/updates/etiketten_version.php"

try:
    import winreg
except ImportError:
    winreg = None

def get_resource_path(relative_path):
    """Geef het absolute pad naar resource bestanden (werkt met PyInstaller)"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class Theme:
    """Bevat de kleurenschema's voor lichte en donkere thema's."""
    
    def __init__(self):
        self.is_dark_mode = self.detect_windows_theme()
        self.update_colors()
    
    def detect_windows_theme(self):
        """Detecteer Windows donker/licht thema"""
        try:
            if winreg:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                                  r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
                    value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                    return value == 0
        except:
            pass
        return True
    
    def update_colors(self):
        """Update kleuren op basis van thema"""
        if self.is_dark_mode:
            self.bg_primary = "#202020"
            self.bg_secondary = "#2d2d2d"
            self.bg_tertiary = "#1a1a1a"
            self.text_primary = "#ffffff"
            self.text_secondary = "#b0b0b0"
            self.accent = "#0078d4"
            self.accent_hover = "#1084d8"
            self.success_color = "#107c10"
            self.border = "#3c3c3c"
            self.card_bg = "#2d2d2d"
            # Invulvelden: duidelijk lichter dan de kaart, met zichtbare rand
            self.input_bg = "#3a3a3a"
            self.input_border = "#6a6a6a"
        else:
            self.bg_primary = "#f3f3f3"
            self.bg_secondary = "#ffffff"
            self.bg_tertiary = "#e8e8e8"
            self.text_primary = "#202020"
            self.text_secondary = "#5a5a5a"
            self.accent = "#0078d4"
            self.accent_hover = "#1084d8"
            self.success_color = "#107c10"
            self.border = "#d1d1d1"
            self.card_bg = "#ffffff"
            # Invulvelden: wit met een grijze rand, anders vallen ze weg op de witte kaart
            self.input_bg = "#ffffff"
            self.input_border = "#8a8a8a"

    def toggle(self):
        """Wissel tussen licht en donker thema"""
        self.is_dark_mode = not self.is_dark_mode
        self.update_colors()

# ---------------------------------------------------------------------------
# Etiketformaten
#
# Elk formaat wordt beschreven in millimeters, zoals op de verpakking van de
# etiketten staat: afmeting van één etiket, aantal naast/onder elkaar, marges
# en tussenruimte. maak_formaat_specs() rekent dat om naar de cm-waarden die
# het Word-document nodig heeft. Eigen formaten van de gebruiker staan in een
# JSON-bestand in %APPDATA% en worden bij het starten erbij geladen.
# ---------------------------------------------------------------------------

PAGINA_FORMATEN = {
    'A4': (210.0, 297.0),
    'A5': (148.0, 210.0),
    'Letter': (215.9, 279.4),
}

# Vaste formaten (mm). 'links'/'boven' leeg = gecentreerd op het vel.
# De 2-koloms 99,1 mm-vellen (Herma/Avery) hebben een tussenruimte van 2,5 mm
# en een linkermarge van 4,65 mm; de tweede kolom begint dus op 106,25 mm.
STANDAARD_FORMATEN = {
    'Herma 10825': dict(breedte=99.1, hoogte=33.8, per_rij=2, per_kolom=8,
                        boven=12.7, links=4.65, tussen_h=2.5, tussen_v=0),
    'Herma 4625': dict(breedte=66.0, hoogte=33.8, per_rij=3, per_kolom=8,
                       boven=12.7, links=2.9, tussen_h=0, tussen_v=0),
    'Herma 4360': dict(breedte=70.0, hoogte=36.0, per_rij=3, per_kolom=8,
                       boven=4.5, links=0, tussen_h=0, tussen_v=0),
    'Herma 4267': dict(breedte=99.1, hoogte=42.3, per_rij=2, per_kolom=6,
                       boven=21.6, links=4.65, tussen_h=2.5, tussen_v=0),
    'Herma 4473 / Avery L7166': dict(breedte=99.1, hoogte=93.1, per_rij=2, per_kolom=3,
                                     boven=8.85, links=4.65, tussen_h=2.5, tussen_v=0),
    'Herma 4425': dict(breedte=105.0, hoogte=148.5, per_rij=2, per_kolom=2,
                       boven=0, links=0, tussen_h=0, tussen_v=0),
    'Avery L7163': dict(breedte=99.1, hoogte=38.1, per_rij=2, per_kolom=7,
                        boven=15.15, links=4.65, tussen_h=2.5, tussen_v=0),
}


def _mm(waarde, standaard=None):
    """Tekst/getal uit de instellingen naar float (mm); leeg → standaard."""
    if waarde is None:
        return standaard
    if isinstance(waarde, str):
        waarde = waarde.strip().replace(',', '.')
        if waarde == '':
            return standaard
    return float(waarde)


def formaat_beschrijving(inst):
    """'99,1 × 33,8 mm (16 etiketten)' uit de mm-instellingen."""
    b = _mm(inst.get('breedte'), 0.0)
    h = _mm(inst.get('hoogte'), 0.0)
    aantal = int(inst.get('per_rij', 0)) * int(inst.get('per_kolom', 0))
    tekst = f"{b:.1f} × {h:.1f} mm ({aantal} etiketten)"
    return tekst.replace('.', ',')


def maak_formaat_specs(inst, eigen=False):
    """Reken mm-instellingen om naar de specs (cm) voor het Word-document.

    Linker-/bovenmarge leeg → etiketten gecentreerd op het vel. De rechter- en
    ondermarge volgen uit de rest (nooit kleiner dan 0).
    """
    pagina = inst.get('pagina', 'A4')
    pw, ph = PAGINA_FORMATEN.get(pagina, PAGINA_FORMATEN['A4'])
    b = _mm(inst['breedte'])
    h = _mm(inst['hoogte'])
    nr = int(inst['per_rij'])
    nk = int(inst['per_kolom'])
    th = _mm(inst.get('tussen_h'), 0.0)
    tv = _mm(inst.get('tussen_v'), 0.0)
    totaal_b = nr * b + (nr - 1) * th
    totaal_h = nk * h + (nk - 1) * tv
    links = _mm(inst.get('links'), None)
    boven = _mm(inst.get('boven'), None)
    if links is None:
        links = (pw - totaal_b) / 2
    if boven is None:
        boven = (ph - totaal_h) / 2
    rechts = max(0.0, pw - links - totaal_b)
    onder = max(0.0, ph - boven - totaal_h)
    return {
        'beschrijving': formaat_beschrijving(inst),
        'pagina': pagina,
        'page_width': pw / 10,
        'page_height': ph / 10,
        'top_margin': boven / 10,
        'bottom_margin': onder / 10,
        'left_margin': links / 10,
        'right_margin': rechts / 10,
        'etiketten_per_rij': nr,
        'etiketten_per_kolom': nk,
        'etiket_breedte': b / 10,
        'etiket_hoogte': h / 10,
        'h_gap': th / 10,
        'v_gap': tv / 10,
        'eigen': eigen,
        'instellingen': dict(inst),
    }


def controleer_formaat(inst):
    """Geef een lijst met problemen (leeg = in orde)."""
    fouten = []
    try:
        b = _mm(inst.get('breedte'))
        h = _mm(inst.get('hoogte'))
        nr = int(inst.get('per_rij'))
        nk = int(inst.get('per_kolom'))
        th = _mm(inst.get('tussen_h'), 0.0)
        tv = _mm(inst.get('tussen_v'), 0.0)
        links = _mm(inst.get('links'), None)
        boven = _mm(inst.get('boven'), None)
    except (TypeError, ValueError):
        return ["Vul bij alle afmetingen een getal in (bijvoorbeeld 99,1)."]
    if b is None or h is None or b <= 0 or h <= 0:
        fouten.append("Breedte en hoogte van het etiket moeten groter dan 0 zijn.")
    if nr < 1 or nk < 1:
        fouten.append("Aantal etiketten naast en onder elkaar moet minimaal 1 zijn.")
    if th < 0 or tv < 0 or (links is not None and links < 0) or (boven is not None and boven < 0):
        fouten.append("Marges en tussenruimte kunnen niet negatief zijn.")
    if fouten:
        return fouten
    pw, ph = PAGINA_FORMATEN.get(inst.get('pagina', 'A4'), PAGINA_FORMATEN['A4'])
    totaal_b = nr * b + (nr - 1) * th + (links or 0)
    totaal_h = nk * h + (nk - 1) * tv + (boven or 0)
    if totaal_b > pw + 0.05:
        fouten.append(f"De etiketten zijn samen {totaal_b:.1f} mm breed, het vel is {pw:.0f} mm.".replace('.', ','))
    if totaal_h > ph + 0.05:
        fouten.append(f"De etiketten zijn samen {totaal_h:.1f} mm hoog, het vel is {ph:.0f} mm.".replace('.', ','))
    return fouten


def instellingen_pad():
    """Bestand met eigen formaten + laatst gekozen formaat (in %APPDATA%)."""
    basis = os.environ.get('APPDATA') or os.path.expanduser('~')
    map_ = os.path.join(basis, 'NVict Service', 'Etiketten Maker')
    try:
        os.makedirs(map_, exist_ok=True)
    except OSError:
        pass
    return os.path.join(map_, 'etiketten_instellingen.json')


def laad_instellingen():
    """{'eigen_formaten': {naam: mm-instellingen}, 'laatst_gekozen': naam}"""
    try:
        with open(instellingen_pad(), 'r', encoding='utf-8') as f:
            data = json.load(f)
    except (OSError, ValueError):
        data = {}
    if not isinstance(data, dict):
        data = {}
    eigen = data.get('eigen_formaten')
    if not isinstance(eigen, dict):
        eigen = {}
    goed = {}
    for naam, inst in eigen.items():
        if isinstance(inst, dict) and not controleer_formaat(inst):
            goed[naam] = inst
    data['eigen_formaten'] = goed
    return data


def bewaar_instellingen(data):
    try:
        with open(instellingen_pad(), 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except OSError:
        return False


def bestandsnaam_code(formaat_naam):
    """Formaatnaam geschikt maken voor in een bestandsnaam."""
    code = re.sub(r'[^0-9A-Za-z_\-]+', '_', formaat_naam).strip('_')
    return code or 'etiketten'


def stel_pagina_in(doc, specs):
    """Papierformaat en marges van het document op het etiketformaat zetten."""
    section = doc.sections[0]
    section.page_height = Cm(specs['page_height'])
    section.page_width = Cm(specs['page_width'])
    section.top_margin = Cm(specs['top_margin'])
    section.left_margin = Cm(specs['left_margin'])
    section.right_margin = Cm(specs['right_margin'])
    # De ondermarge iets kleiner dan berekend: Word schuift anders door
    # afronding de onderste rij etiketten naar een volgende pagina als de
    # rijen het vel precies vullen. De positie van de etiketten verandert
    # hier niet door (die volgt uit de bovenmarge en de rijhoogtes).
    section.bottom_margin = Cm(max(0.0, specs['bottom_margin'] - 0.3))


def _verborgen_alinea(doc):
    """Lege, verborgen alinea van 1 pt tussen twee vellen (neemt geen ruimte in).

    Word eist een alinea tussen twee tabellen (anders worden ze één tabel) en
    ná de laatste tabel; een gewone lege alinea past niet meer op een vol vel
    en levert dan een lege pagina op.
    """
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(1)
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(r'<w:rPr %s><w:vanish/><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>' % nsdecls('w')))
    return p


def maak_etiket_vel(doc, specs, nieuwe_pagina):
    """Voeg één vel (tabel) toe aan het document.

    Tussenruimte tussen de etiketten wordt gemaakt met lege kolommen/rijen in
    de tabel. Geeft de etiketcellen terug in leesvolgorde (links→rechts,
    boven→onder).
    """
    nr = specs['etiketten_per_rij']
    nk = specs['etiketten_per_kolom']
    h_gap = specs.get('h_gap', 0) or 0
    v_gap = specs.get('v_gap', 0) or 0

    kolommen = []   # (breedte cm, is etiket)
    for k in range(nr):
        if k > 0 and h_gap > 0:
            kolommen.append((h_gap, False))
        kolommen.append((specs['etiket_breedte'], True))
    rijen = []      # (hoogte cm, is etiket)
    for r in range(nk):
        if r > 0 and v_gap > 0:
            rijen.append((v_gap, False))
        rijen.append((specs['etiket_hoogte'], True))

    table = doc.add_table(rows=len(rijen), cols=len(kolommen))
    table.allow_autofit = False

    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = parse_xml(r'<w:tblPr %s/>' % nsdecls('w'))
        tbl.insert(0, tblPr)
    tblPr.append(parse_xml(r'<w:tblCellSpacing %s w:w="0" w:type="dxa"/>' % nsdecls('w')))
    tblPr.append(parse_xml(r'<w:tblLayout %s w:type="fixed"/>' % nsdecls('w')))
    tblPr.append(parse_xml(
        r'<w:tblBorders %s>'
        r'<w:top w:val="none"/><w:left w:val="none"/>'
        r'<w:bottom w:val="none"/><w:right w:val="none"/>'
        r'<w:insideH w:val="none"/><w:insideV w:val="none"/>'
        r'</w:tblBorders>' % nsdecls('w')
    ))
    # Tabelrand precies op de paginamarge (Word schuift een tabel anders
    # 1,9 mm naar links, ter grootte van de standaard celmarge).
    tblPr.append(parse_xml(r'<w:tblInd %s w:w="0" w:type="dxa"/>' % nsdecls('w')))
    tblPr.append(parse_xml(
        r'<w:tblCellMar %s>'
        r'<w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
        r'<w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>'
        r'</w:tblCellMar>' % nsdecls('w')
    ))

    if nieuwe_pagina:
        table.rows[0].cells[0].paragraphs[0].paragraph_format.page_break_before = True

    for idx, (breedte, _) in enumerate(kolommen):
        table.columns[idx].width = Cm(breedte)
        for cell in table.columns[idx].cells:
            cell.width = Cm(breedte)

    cellen = []
    for r, (hoogte, rij_is_etiket) in enumerate(rijen):
        row = table.rows[r]
        if r == len(rijen) - 1:
            # Onderste rij een fractie lager: Word zet een rij die het vel tot
            # op de twip precies vult anders op de volgende pagina.
            hoogte = max(0.1, hoogte - 0.05)
        row.height = Cm(hoogte)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        for k, (_, kol_is_etiket) in enumerate(kolommen):
            cell = row.cells[k]
            tc = cell._element
            tcPr = tc.tcPr
            if tcPr is None:
                tcPr = parse_xml(r'<w:tcPr %s/>' % nsdecls('w'))
                tc.insert(0, tcPr)
            # Geen celmarges (tcMar): Word maakt een rij met vaste hoogte dan
            # ~0,5 mm hoger dan ingesteld, waardoor de onderste rij van het vel
            # valt. De witruimte in het etiket komt van de alinea-inspringing
            # in vul_etiket_cel().
            if rij_is_etiket and kol_is_etiket:
                cellen.append(cell)
            else:
                # Tussenruimte: een zo klein mogelijke lege alinea
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.add_run("").font.size = Pt(1)
    _verborgen_alinea(doc)
    return cellen


def vul_etiket_cel(cell, regels, font_grootte, alignment=WD_ALIGN_PARAGRAPH.LEFT):
    """Zet de adresregels in een etiketcel."""
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    p.alignment = alignment
    for regel in regels:
        run = p.add_run(regel + "\n")
        run.font.size = Pt(font_grootte)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.line_spacing = 1.0
    # Witruimte links/rechts in het etiket (i.p.v. celmarges, zie maak_etiket_vel)
    p.paragraph_format.left_indent = Cm(0.2)
    p.paragraph_format.right_indent = Cm(0.1)
    cell.vertical_alignment = 1


def etiket_regels_uit_rij(row_data, regels_config):
    """Bouw de tekstregels voor één adres uit een Excel-rij."""
    regels = []
    for regel_velden in regels_config:
        delen = []
        for veld in regel_velden:
            if veld in row_data:
                waarde = str(row_data[veld]).strip()
                if waarde and waarde != 'nan':
                    delen.append(waarde)
        if delen:
            regels.append(" ".join(delen))
    return regels


def preview_afmeting(specs, max_breedte, max_hoogte, min_hoogte=60):
    """Pixelafmeting van een voorbeeldetiket met de verhouding van het formaat."""
    try:
        verhouding = float(specs['etiket_hoogte']) / float(specs['etiket_breedte'])
    except (KeyError, TypeError, ValueError, ZeroDivisionError):
        verhouding = 1 / 3
    breedte = max_breedte
    hoogte = int(round(breedte * verhouding))
    if hoogte > max_hoogte:
        hoogte = max_hoogte
        breedte = int(round(hoogte / verhouding))
    if hoogte < min_hoogte:
        hoogte = min_hoogte
    return breedte, hoogte


def maak_invoerveld(parent, theme, textvariable, width=None, **extra):
    """Tekstveld dat in beide thema's duidelijk zichtbaar is (rand + eigen achtergrond)."""
    opties = dict(
        textvariable=textvariable,
        font=("Segoe UI", 10),
        bg=theme.input_bg,
        fg=theme.text_primary,
        relief=tk.SOLID,
        borderwidth=1,
        highlightthickness=1,
        highlightbackground=theme.input_border,
        highlightcolor=theme.accent,
        insertbackground=theme.text_primary,
    )
    if width is not None:
        opties['width'] = width
    opties.update(extra)
    return tk.Entry(parent, **opties)


def maak_spinbox(parent, theme, textvariable, from_, to, width=5, command=None, font_size=12):
    """Spinbox in dezelfde stijl als maak_invoerveld."""
    return tk.Spinbox(parent,
                      textvariable=textvariable,
                      from_=from_, to=to,
                      font=("Segoe UI", font_size),
                      bg=theme.input_bg,
                      fg=theme.text_primary,
                      buttonbackground=theme.input_bg,
                      relief=tk.SOLID,
                      borderwidth=1,
                      highlightthickness=1,
                      highlightbackground=theme.input_border,
                      highlightcolor=theme.accent,
                      insertbackground=theme.text_primary,
                      width=width,
                      command=command)


def detecteer_adres_velden(kolommen):
    """Detecteer automatisch welke velden bij elkaar horen"""
    kolommen_lower = [k.lower() for k in kolommen]
    regels = []
    
    # Regel 1: Naam velden (aanhef, voorletters, achternaam)
    naam_velden = []
    naam_keywords = ['aanspr', 'aanhef', 'dhr', 'mevr', 'titel']
    voorletter_keywords = ['voorletter', 'initiaal', 'initial']
    achternaam_keywords = ['achternaam', 'naam', 'surname', 'lastname']
    
    for i, kol_lower in enumerate(kolommen_lower):
        if any(kw in kol_lower for kw in naam_keywords):
            naam_velden.append(kolommen[i])
        elif any(kw in kol_lower for kw in voorletter_keywords):
            naam_velden.append(kolommen[i])
        elif any(kw in kol_lower for kw in achternaam_keywords):
            naam_velden.append(kolommen[i])
    
    if naam_velden:
        regels.append(naam_velden)
    
    # Regel 2: Extra (alleen als het bestaat)
    extra_velden = []
    extra_keywords = ['extra', 'tussenvoegsel', 'voorvoegsel', 'toevoeging']
    
    for i, kol_lower in enumerate(kolommen_lower):
        if any(kw in kol_lower for kw in extra_keywords):
            extra_velden.append(kolommen[i])
    
    if extra_velden:
        regels.append(extra_velden)
    
    # Regel 3: Adres velden (straat, huisnummer)
    adres_velden = []
    straat_keywords = ['adres', 'straat', 'street', 'weg', 'laan', 'plein']
    huisnr_keywords = ['huisnummer', 'huisnr', 'nummer', 'nr', 'number']
    
    for i, kol_lower in enumerate(kolommen_lower):
        if any(kw in kol_lower for kw in straat_keywords):
            adres_velden.append(kolommen[i])
        elif any(kw in kol_lower for kw in huisnr_keywords):
            adres_velden.append(kolommen[i])
    
    if adres_velden:
        regels.append(adres_velden)
    
    # Regel 4: Postcode en plaats
    postcode_velden = []
    postcode_keywords = ['postcode', 'zip', 'postal']
    plaats_keywords = ['woonplaats', 'plaats', 'stad', 'city', 'town']
    
    for i, kol_lower in enumerate(kolommen_lower):
        if any(kw in kol_lower for kw in postcode_keywords):
            postcode_velden.append(kolommen[i])
        elif any(kw in kol_lower for kw in plaats_keywords):
            postcode_velden.append(kolommen[i])
    
    if postcode_velden:
        regels.append(postcode_velden)
    
    # Regel 5: Land (indien aanwezig)
    land_velden = []
    land_keywords = ['land', 'country', 'nation']
    
    for i, kol_lower in enumerate(kolommen_lower):
        if any(kw in kol_lower for kw in land_keywords):
            land_velden.append(kolommen[i])
    
    if land_velden:
        regels.append(land_velden)
    
    return regels

class EtiketAanpassingWindow:
    """Window voor beperkte aanpassingen"""
    def __init__(self, parent, theme, df, excel_bestand, auto_regels, formaat_specs=None):
        self.parent = parent
        self.theme = theme
        self.df = df
        self.excel_bestand = excel_bestand
        self.auto_regels = auto_regels
        self.formaat_specs = formaat_specs or {}
        self.result = None
        
        self.settings = {
            'regels': auto_regels,
            'font_grootte': 10,
            'uitlijning': 'links'
        }
        
        self.window = tk.Toplevel(parent)
        self.window.title("Etiket Preview")
        self.window.geometry("1000x800")
        self.window.minsize(900, 700)
        self.window.transient(parent)
        self.window.grab_set()
        self.window.configure(bg=theme.bg_primary)
        
        # Centreer op scherm
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (1000 // 2)
        y = (self.window.winfo_screenheight() // 2) - (800 // 2)
        self.window.geometry(f"1000x800+{x}+{y}")
        
        # Icoon instellen
        try:
            icon_path = get_resource_path("favicon.ico")
            if os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except:
            pass
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup de gebruikersinterface"""
        # Header
        header_frame = tk.Frame(self.window, bg=self.theme.bg_secondary, height=70)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame,
                              text="✓  Preview & Aanpassen",
                              font=("Segoe UI", 18, "bold"),
                              fg=self.theme.text_primary,
                              bg=self.theme.bg_secondary)
        title_label.pack(pady=20, padx=30, anchor=tk.W)
        
        # Main content
        main_container = tk.Frame(self.window, bg=self.theme.bg_primary)
        main_container.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Links: Preview
        left_frame = tk.Frame(main_container, bg=self.theme.bg_primary)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))
        
        # Uitleg
        uitleg_card = tk.Frame(left_frame, bg=self.theme.card_bg)
        uitleg_card.pack(fill=tk.X, pady=(0, 15))
        
        uitleg_content = tk.Frame(uitleg_card, bg=self.theme.card_bg)
        uitleg_content.pack(padx=25, pady=20)
        
        tk.Label(uitleg_content,
                text="📋 Automatisch gedetecteerd",
                font=("Segoe UI", 14, "bold"),
                fg=self.theme.text_primary,
                bg=self.theme.card_bg).pack(anchor=tk.W, pady=(0, 10))
        
        # Toon welke velden op welke regel komen
        for i, regel_velden in enumerate(self.auto_regels):
            regel_text = f"Regel {i+1}: {', '.join(regel_velden)}"
            tk.Label(uitleg_content,
                    text=regel_text,
                    font=("Segoe UI", 10),
                    fg=self.theme.text_secondary,
                    bg=self.theme.card_bg,
                    anchor=tk.W).pack(fill=tk.X, pady=2)
        
        # Preview card
        preview_card = tk.Frame(left_frame, bg=self.theme.card_bg)
        preview_card.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        preview_header = tk.Frame(preview_card, bg=self.theme.card_bg)
        preview_header.pack(fill=tk.X, padx=25, pady=(20, 15))
        
        tk.Label(preview_header,
                text="👁️ Voorbeeld",
                font=("Segoe UI", 14, "bold"),
                fg=self.theme.text_primary,
                bg=self.theme.card_bg).pack(anchor=tk.W)
        
        beschrijving = self.formaat_specs.get('beschrijving', '')
        tk.Label(preview_header,
                text="Eerste adres uit uw Excel" + (f"  •  {beschrijving}" if beschrijving else ""),
                font=("Segoe UI", 10),
                fg=self.theme.text_secondary,
                bg=self.theme.card_bg).pack(anchor=tk.W)

        # Container voor preview met vaste verhoudingen
        preview_container = tk.Frame(preview_card, bg=self.theme.card_bg)
        preview_container.pack(fill=tk.X, padx=25, pady=(10, 25))

        # Preview etiket met de verhouding van het gekozen formaat
        pv_breedte, pv_hoogte = preview_afmeting(self.formaat_specs, 360, 320, 100)
        preview_etiket = tk.Frame(preview_container, bg="#ffffff",
                                 relief=tk.SOLID, borderwidth=2,
                                 width=pv_breedte, height=pv_hoogte)
        preview_etiket.pack()
        preview_etiket.pack_propagate(False)
        
        self.preview_label = tk.Label(preview_etiket,
                                      text="",
                                      font=("Segoe UI", 9),
                                      fg="#000000",
                                      bg="#ffffff",
                                      justify=tk.LEFT,
                                      anchor=tk.NW,
                                      padx=20,
                                      pady=20)
        self.preview_label.pack(fill=tk.BOTH, expand=True)
        
        # Rechts: Opties
        right_frame = tk.Frame(main_container, bg=self.theme.bg_primary, width=280)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y)
        right_frame.pack_propagate(False)
        
        opties_card = tk.Frame(right_frame, bg=self.theme.card_bg)
        opties_card.pack(fill=tk.BOTH, expand=True)
        
        opties_content = tk.Frame(opties_card, bg=self.theme.card_bg)
        opties_content.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)
        
        tk.Label(opties_content,
                text="⚙️ Opties",
                font=("Segoe UI", 14, "bold"),
                fg=self.theme.text_primary,
                bg=self.theme.card_bg).pack(anchor=tk.W, pady=(0, 15))
        
        # Font grootte
        tk.Label(opties_content,
                text="Lettergrootte:",
                font=("Segoe UI", 11),
                fg=self.theme.text_primary,
                bg=self.theme.card_bg).pack(anchor=tk.W, pady=(5, 5))
        
        self.font_var = tk.IntVar(value=10)
        font_scale = tk.Scale(opties_content,
                             from_=8,
                             to=14,
                             orient=tk.HORIZONTAL,
                             variable=self.font_var,
                             command=lambda x: self.update_preview(),
                             bg=self.theme.card_bg,
                             fg=self.theme.text_primary,
                             troughcolor=self.theme.bg_secondary,
                             highlightthickness=0)
        font_scale.pack(fill=tk.X, pady=(0, 15))
        
        # Uitlijning
        tk.Label(opties_content,
                text="Uitlijning:",
                font=("Segoe UI", 11),
                fg=self.theme.text_primary,
                bg=self.theme.card_bg).pack(anchor=tk.W, pady=(5, 5))
        
        self.align_var = tk.StringVar(value='links')
        for text, val in [("Links", "links"), ("Midden", "midden"), ("Rechts", "rechts")]:
            rb = tk.Radiobutton(opties_content,
                               text=text,
                               variable=self.align_var,
                               value=val,
                               font=("Segoe UI", 10),
                               fg=self.theme.text_primary,
                               bg=self.theme.card_bg,
                               selectcolor=self.theme.bg_secondary,
                               activebackground=self.theme.card_bg,
                               command=self.update_preview)
            rb.pack(anchor=tk.W, pady=2)
        
        # Footer met knoppen
        footer_frame = tk.Frame(self.window, bg=self.theme.bg_secondary, height=85)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        btn_container = tk.Frame(footer_frame, bg=self.theme.bg_secondary)
        btn_container.pack(expand=True)
        
        annuleer_btn = tk.Button(btn_container,
                                text="Annuleren",
                                command=self.annuleren,
                                font=("Segoe UI", 12),
                                bg=self.theme.card_bg,
                                fg=self.theme.text_primary,
                                relief=tk.FLAT,
                                cursor="hand2",
                                padx=35,
                                pady=14)
        annuleer_btn.pack(side=tk.LEFT, padx=8)
        
        maak_btn = tk.Button(btn_container,
                           text="✓ Etiketten maken",
                           command=self.bevestigen,
                           font=("Segoe UI", 12, "bold"),
                           bg=self.theme.accent,
                           fg="#ffffff",
                           activebackground=self.theme.accent_hover,
                           activeforeground="#ffffff",
                           relief=tk.FLAT,
                           cursor="hand2",
                           padx=35,
                           pady=14)
        maak_btn.pack(side=tk.LEFT, padx=8)
        
        # Initial preview
        self.update_preview()
    
    def update_preview(self):
        """Update het preview etiket"""
        if len(self.df) == 0:
            self.preview_label.config(text="Geen data beschikbaar")
            return
        
        eerste_rij = self.df.iloc[0]
        preview_text = ""
        
        # Bouw regels
        for regel_velden in self.auto_regels:
            regel_tekst = []
            for veld in regel_velden:
                if veld in eerste_rij:
                    waarde = str(eerste_rij[veld]).strip()
                    if waarde and waarde != 'nan':
                        regel_tekst.append(waarde)
            
            if regel_tekst:
                preview_text += " ".join(regel_tekst) + "\n"
        
        # Gebruik de actuele font grootte van de slider
        font_grootte = self.font_var.get()
        self.preview_label.config(text=preview_text, font=("Segoe UI", font_grootte))
        
        # Update uitlijning
        align_map = {"links": tk.W, "midden": tk.CENTER, "rechts": tk.E}
        justify_map = {"links": tk.LEFT, "midden": tk.CENTER, "rechts": tk.RIGHT}
        self.preview_label.config(anchor=align_map[self.align_var.get()],
                                 justify=justify_map[self.align_var.get()])
    
    def annuleren(self):
        """Annuleer zonder te maken"""
        self.result = None
        self.window.destroy()
    
    def bevestigen(self):
        """Bevestig en sla instellingen op"""
        self.settings['font_grootte'] = self.font_var.get()
        self.settings['uitlijning'] = self.align_var.get()
        
        self.result = self.settings
        self.window.destroy()
    
    def wait_for_result(self):
        """Wacht tot window wordt gesloten en return result"""
        self.window.wait_window()
        return self.result

class EénAdresWindow:
    """Dialog voor het invoeren van één adres op alle etiketten"""
    def __init__(self, parent, theme, formaat_naam, formaat_specs):
        self.parent = parent
        self.theme = theme
        self.formaat_naam = formaat_naam
        self.formaat_specs = formaat_specs
        self.result = None

        self.etiketten_per_vel = formaat_specs['etiketten_per_rij'] * formaat_specs['etiketten_per_kolom']

        self.window = tk.Toplevel(parent)
        self.window.title("Één adres op alle etiketten")
        self.window.geometry("820x680")
        self.window.minsize(700, 600)
        self.window.transient(parent)
        self.window.grab_set()
        self.window.configure(bg=theme.bg_primary)

        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (820 // 2)
        y = (self.window.winfo_screenheight() // 2) - (680 // 2)
        self.window.geometry(f"820x680+{x}+{y}")

        try:
            icon_path = get_resource_path("favicon.ico")
            if os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except:
            pass

        self.setup_ui()

    def setup_ui(self):
        # Header
        header_frame = tk.Frame(self.window, bg=self.theme.bg_secondary, height=70)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        tk.Label(header_frame,
                 text="📝  Één adres op alle etiketten",
                 font=("Segoe UI", 18, "bold"),
                 fg=self.theme.text_primary,
                 bg=self.theme.bg_secondary).pack(pady=20, padx=30, anchor=tk.W)

        # Main content
        main_container = tk.Frame(self.window, bg=self.theme.bg_primary)
        main_container.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)

        # Links: invoervelden
        left_frame = tk.Frame(main_container, bg=self.theme.bg_primary)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))

        # Adresregels card
        adres_card = tk.Frame(left_frame, bg=self.theme.card_bg)
        adres_card.pack(fill=tk.X, pady=(0, 15))

        adres_content = tk.Frame(adres_card, bg=self.theme.card_bg)
        adres_content.pack(padx=20, pady=20, fill=tk.X)

        tk.Label(adres_content,
                 text="Adresregels:",
                 font=("Segoe UI", 12, "bold"),
                 fg=self.theme.text_primary,
                 bg=self.theme.card_bg).pack(anchor=tk.W, pady=(0, 10))

        self.regel_vars = []
        regel_labels = [
            "Naam:",
            "Adres:",
            "Postcode + Woonplaats:",
            "Regel 4 (optioneel):",
            "Regel 5 (optioneel):"
        ]

        for label_text in regel_labels:
            row_frame = tk.Frame(adres_content, bg=self.theme.card_bg)
            row_frame.pack(fill=tk.X, pady=4)

            tk.Label(row_frame,
                     text=label_text,
                     font=("Segoe UI", 10),
                     fg=self.theme.text_secondary,
                     bg=self.theme.card_bg,
                     width=22,
                     anchor=tk.W).pack(side=tk.LEFT)

            var = tk.StringVar()
            entry = maak_invoerveld(row_frame, self.theme, var)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5, padx=(0, 5))
            var.trace('w', lambda *args: self.update_preview())
            self.regel_vars.append(var)

        # Aantal vellen card
        vellen_card = tk.Frame(left_frame, bg=self.theme.card_bg)
        vellen_card.pack(fill=tk.X, pady=(0, 15))

        vellen_content = tk.Frame(vellen_card, bg=self.theme.card_bg)
        vellen_content.pack(padx=20, pady=20, fill=tk.X)

        tk.Label(vellen_content,
                 text="Aantal vellen:",
                 font=("Segoe UI", 12, "bold"),
                 fg=self.theme.text_primary,
                 bg=self.theme.card_bg).pack(anchor=tk.W, pady=(0, 8))

        vellen_row = tk.Frame(vellen_content, bg=self.theme.card_bg)
        vellen_row.pack(fill=tk.X)

        self.vellen_var = tk.IntVar(value=1)
        vellen_spin = maak_spinbox(vellen_row, self.theme, self.vellen_var,
                                   from_=1, to=999, width=5,
                                   command=self.update_vellen_info)
        vellen_spin.pack(side=tk.LEFT, ipady=5)

        self.vellen_info = tk.Label(vellen_row,
                                    text=f"× {self.etiketten_per_vel} per vel = {self.etiketten_per_vel} etiketten totaal",
                                    font=("Segoe UI", 10),
                                    fg=self.theme.text_secondary,
                                    bg=self.theme.card_bg)
        self.vellen_info.pack(side=tk.LEFT, padx=12)

        # Rechts: preview + lettergrootte
        right_frame = tk.Frame(main_container, bg=self.theme.bg_primary, width=260)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y)
        right_frame.pack_propagate(False)

        preview_card = tk.Frame(right_frame, bg=self.theme.card_bg)
        preview_card.pack(fill=tk.BOTH, expand=True)

        preview_content = tk.Frame(preview_card, bg=self.theme.card_bg)
        preview_content.pack(padx=15, pady=15, fill=tk.BOTH, expand=True)

        tk.Label(preview_content,
                 text="👁️ Voorbeeld",
                 font=("Segoe UI", 12, "bold"),
                 fg=self.theme.text_primary,
                 bg=self.theme.card_bg).pack(anchor=tk.W, pady=(0, 2))

        tk.Label(preview_content,
                 text=self.formaat_specs.get('beschrijving', ''),
                 font=("Segoe UI", 9),
                 fg=self.theme.text_secondary,
                 bg=self.theme.card_bg).pack(anchor=tk.W, pady=(0, 8))

        pv_breedte, pv_hoogte = preview_afmeting(self.formaat_specs, 220, 220, 70)
        preview_etiket = tk.Frame(preview_content, bg="#ffffff",
                                  relief=tk.SOLID, borderwidth=2,
                                  width=pv_breedte, height=pv_hoogte)
        preview_etiket.pack()
        preview_etiket.pack_propagate(False)

        self.preview_label = tk.Label(preview_etiket,
                                      text="(vul adresregels in)",
                                      font=("Segoe UI", 9),
                                      fg="#000000",
                                      bg="#ffffff",
                                      justify=tk.LEFT,
                                      anchor=tk.NW,
                                      padx=10,
                                      pady=8)
        self.preview_label.pack(fill=tk.BOTH, expand=True)

        tk.Label(preview_content,
                 text="Lettergrootte:",
                 font=("Segoe UI", 10),
                 fg=self.theme.text_primary,
                 bg=self.theme.card_bg).pack(anchor=tk.W, pady=(15, 3))

        self.font_var = tk.IntVar(value=10)
        font_scale = tk.Scale(preview_content,
                              from_=8, to=14,
                              orient=tk.HORIZONTAL,
                              variable=self.font_var,
                              command=lambda x: self.update_preview(),
                              bg=self.theme.card_bg,
                              fg=self.theme.text_primary,
                              troughcolor=self.theme.bg_secondary,
                              highlightthickness=0)
        font_scale.pack(fill=tk.X, pady=(0, 10))

        # Footer
        footer_frame = tk.Frame(self.window, bg=self.theme.bg_secondary, height=85)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)

        btn_container = tk.Frame(footer_frame, bg=self.theme.bg_secondary)
        btn_container.pack(expand=True)

        tk.Button(btn_container,
                  text="Annuleren",
                  command=self.annuleren,
                  font=("Segoe UI", 12),
                  bg=self.theme.card_bg,
                  fg=self.theme.text_primary,
                  relief=tk.FLAT,
                  cursor="hand2",
                  padx=35, pady=14).pack(side=tk.LEFT, padx=8)

        tk.Button(btn_container,
                  text="✓ Etiketten maken",
                  command=self.bevestigen,
                  font=("Segoe UI", 12, "bold"),
                  bg=self.theme.accent,
                  fg="#ffffff",
                  activebackground=self.theme.accent_hover,
                  activeforeground="#ffffff",
                  relief=tk.FLAT,
                  cursor="hand2",
                  padx=35, pady=14).pack(side=tk.LEFT, padx=8)

    def update_preview(self):
        lines = [v.get().strip() for v in self.regel_vars if v.get().strip()]
        preview_text = "\n".join(lines) if lines else "(vul adresregels in)"
        self.preview_label.config(text=preview_text, font=("Segoe UI", self.font_var.get()))
        self.update_vellen_info()

    def update_vellen_info(self):
        try:
            n = self.vellen_var.get()
            totaal = n * self.etiketten_per_vel
            self.vellen_info.config(
                text=f"× {self.etiketten_per_vel} per vel = {totaal} etiketten totaal"
            )
        except:
            pass

    def annuleren(self):
        self.result = None
        self.window.destroy()

    def bevestigen(self):
        regels = [v.get().strip() for v in self.regel_vars if v.get().strip()]
        if not regels:
            messagebox.showwarning("Geen adres", "Vul minimaal één adresregel in.", parent=self.window)
            return
        self.result = {
            'regels': regels,
            'aantal_vellen': self.vellen_var.get(),
            'font_grootte': self.font_var.get()
        }
        self.window.destroy()

    def wait_for_result(self):
        self.window.wait_window()
        return self.result


class EigenFormaatWindow:
    """Dialoog om een eigen etiketformaat in te voeren of te bewerken (alles in mm)."""

    VELDEN = [
        # (sleutel, label, toelichting)
        ('breedte', "Breedte etiket (mm):", "bijv. 99,1"),
        ('hoogte', "Hoogte etiket (mm):", "bijv. 93,1"),
        ('per_rij', "Aantal naast elkaar:", "kolommen op het vel"),
        ('per_kolom', "Aantal onder elkaar:", "rijen op het vel"),
        ('links', "Linkermarge (mm):", "leeg = centreren"),
        ('boven', "Bovenmarge (mm):", "leeg = centreren"),
        ('tussen_h', "Tussenruimte breedte (mm):", "tussen de kolommen"),
        ('tussen_v', "Tussenruimte hoogte (mm):", "tussen de rijen"),
    ]

    def __init__(self, parent, theme, bestaande_namen, naam=None, inst=None):
        self.parent = parent
        self.theme = theme
        self.bestaande_namen = set(bestaande_namen)
        self.oude_naam = naam
        self.result = None

        inst = dict(inst or {})
        self.naam_var = tk.StringVar(value=naam or "")
        self.pagina_var = tk.StringVar(value=inst.get('pagina', 'A4'))
        self.vars = {}
        for sleutel, _, _ in self.VELDEN:
            waarde = inst.get(sleutel)
            if waarde is None:
                tekst = ""
            elif sleutel in ('per_rij', 'per_kolom'):
                tekst = str(int(waarde))
            else:
                tekst = f"{float(waarde):g}".replace('.', ',')
            self.vars[sleutel] = tk.StringVar(value=tekst)
        if not inst:
            self.vars['per_rij'].set("2")
            self.vars['per_kolom'].set("3")
            self.vars['tussen_h'].set("0")
            self.vars['tussen_v'].set("0")

        self.window = tk.Toplevel(parent)
        self.window.title("Eigen etiketformaat" if naam is None else "Etiketformaat bewerken")
        self.window.geometry("860x640")
        self.window.minsize(800, 600)
        self.window.transient(parent)
        self.window.grab_set()
        self.window.configure(bg=theme.bg_primary)

        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (860 // 2)
        y = (self.window.winfo_screenheight() // 2) - (640 // 2)
        self.window.geometry(f"860x640+{x}+{y}")

        try:
            icon_path = get_resource_path("favicon.ico")
            if os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except:
            pass

        self.setup_ui()
        for var in list(self.vars.values()) + [self.pagina_var]:
            var.trace_add('write', lambda *args: self.update_preview())
        self.update_preview()

    def setup_ui(self):
        header_frame = tk.Frame(self.window, bg=self.theme.bg_secondary, height=70)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        titel = "➕  Eigen etiketformaat" if self.oude_naam is None else "✏️  Etiketformaat bewerken"
        tk.Label(header_frame,
                 text=titel,
                 font=("Segoe UI", 18, "bold"),
                 fg=self.theme.text_primary,
                 bg=self.theme.bg_secondary).pack(pady=20, padx=30, anchor=tk.W)

        main_container = tk.Frame(self.window, bg=self.theme.bg_primary)
        main_container.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)

        # Links: invoervelden
        left_card = tk.Frame(main_container, bg=self.theme.card_bg)
        left_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))

        form = tk.Frame(left_card, bg=self.theme.card_bg)
        form.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)

        tk.Label(form,
                 text="Neem de maten over van de verpakking van de etiketten.",
                 font=("Segoe UI", 10),
                 fg=self.theme.text_secondary,
                 bg=self.theme.card_bg,
                 wraplength=420,
                 justify=tk.LEFT).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 12))

        rij = 1
        tk.Label(form, text="Naam:", font=("Segoe UI", 10),
                 fg=self.theme.text_primary, bg=self.theme.card_bg,
                 anchor=tk.W).grid(row=rij, column=0, sticky=tk.W, pady=4)
        self.naam_entry = maak_invoerveld(form, self.theme, self.naam_var, width=24)
        self.naam_entry.grid(row=rij, column=1, sticky=tk.EW, pady=4, ipady=4)
        tk.Label(form, text="bijv. Avery L7166", font=("Segoe UI", 9),
                 fg=self.theme.text_secondary, bg=self.theme.card_bg,
                 anchor=tk.W).grid(row=rij, column=2, sticky=tk.W, padx=(10, 0))

        rij += 1
        tk.Label(form, text="Papier:", font=("Segoe UI", 10),
                 fg=self.theme.text_primary, bg=self.theme.card_bg,
                 anchor=tk.W).grid(row=rij, column=0, sticky=tk.W, pady=4)
        pagina_combo = ttk.Combobox(form,
                                    textvariable=self.pagina_var,
                                    values=list(PAGINA_FORMATEN.keys()),
                                    state='readonly',
                                    style='Custom.TCombobox',
                                    font=("Segoe UI", 10),
                                    width=10)
        pagina_combo.grid(row=rij, column=1, sticky=tk.W, pady=4)
        tk.Label(form, text="A4 = 210 × 297 mm", font=("Segoe UI", 9),
                 fg=self.theme.text_secondary, bg=self.theme.card_bg,
                 anchor=tk.W).grid(row=rij, column=2, sticky=tk.W, padx=(10, 0))

        for sleutel, label_tekst, toelichting in self.VELDEN:
            rij += 1
            tk.Label(form, text=label_tekst, font=("Segoe UI", 10),
                     fg=self.theme.text_primary, bg=self.theme.card_bg,
                     anchor=tk.W).grid(row=rij, column=0, sticky=tk.W, pady=4)
            entry = maak_invoerveld(form, self.theme, self.vars[sleutel], width=10)
            entry.grid(row=rij, column=1, sticky=tk.W, pady=4, ipady=4)
            tk.Label(form, text=toelichting, font=("Segoe UI", 9),
                     fg=self.theme.text_secondary, bg=self.theme.card_bg,
                     anchor=tk.W).grid(row=rij, column=2, sticky=tk.W, padx=(10, 0))

        form.columnconfigure(1, weight=1)
        form.columnconfigure(0, minsize=190)

        # Rechts: voorbeeld van het vel
        right_card = tk.Frame(main_container, bg=self.theme.card_bg, width=300)
        right_card.pack(side=tk.RIGHT, fill=tk.Y)
        right_card.pack_propagate(False)

        preview_content = tk.Frame(right_card, bg=self.theme.card_bg)
        preview_content.pack(padx=15, pady=15, fill=tk.BOTH, expand=True)

        tk.Label(preview_content,
                 text="👁️ Indeling van het vel",
                 font=("Segoe UI", 12, "bold"),
                 fg=self.theme.text_primary,
                 bg=self.theme.card_bg).pack(anchor=tk.W, pady=(0, 8))

        self.canvas = tk.Canvas(preview_content, width=260, height=340,
                                bg=self.theme.card_bg, highlightthickness=0)
        self.canvas.pack()

        self.info_label = tk.Label(preview_content,
                                   text="",
                                   font=("Segoe UI", 9),
                                   fg=self.theme.text_secondary,
                                   bg=self.theme.card_bg,
                                   justify=tk.LEFT,
                                   anchor=tk.NW,
                                   wraplength=260)
        self.info_label.pack(fill=tk.X, pady=(10, 0))

        # Footer
        footer_frame = tk.Frame(self.window, bg=self.theme.bg_secondary, height=85)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)

        btn_container = tk.Frame(footer_frame, bg=self.theme.bg_secondary)
        btn_container.pack(expand=True)

        tk.Button(btn_container,
                  text="Annuleren",
                  command=self.annuleren,
                  font=("Segoe UI", 12),
                  bg=self.theme.card_bg,
                  fg=self.theme.text_primary,
                  relief=tk.FLAT,
                  cursor="hand2",
                  padx=35, pady=14).pack(side=tk.LEFT, padx=8)

        tk.Button(btn_container,
                  text="✓ Opslaan",
                  command=self.bevestigen,
                  font=("Segoe UI", 12, "bold"),
                  bg=self.theme.accent,
                  fg="#ffffff",
                  activebackground=self.theme.accent_hover,
                  activeforeground="#ffffff",
                  relief=tk.FLAT,
                  cursor="hand2",
                  padx=35, pady=14).pack(side=tk.LEFT, padx=8)

        self.naam_entry.focus_set()

    def lees_instellingen(self):
        """Huidige invoer als mm-instellingen (tekst → getal). ValueError bij onzin."""
        inst = {'pagina': self.pagina_var.get() or 'A4'}
        for sleutel, _, _ in self.VELDEN:
            tekst = self.vars[sleutel].get().strip().replace(',', '.')
            if sleutel in ('per_rij', 'per_kolom'):
                inst[sleutel] = int(tekst) if tekst else 0
            elif sleutel in ('links', 'boven'):
                inst[sleutel] = float(tekst) if tekst else None
            else:
                inst[sleutel] = float(tekst) if tekst else 0.0
        return inst

    def update_preview(self):
        self.canvas.delete("all")
        try:
            inst = self.lees_instellingen()
            fouten = controleer_formaat(inst)
        except ValueError:
            self.info_label.config(text="Vul bij alle afmetingen een getal in (bijvoorbeeld 99,1).")
            return
        if any("groter dan 0" in f or "minimaal 1" in f for f in fouten):
            self.info_label.config(text="\n".join(fouten))
            return

        specs = maak_formaat_specs(inst, eigen=True)
        pw_mm = specs['page_width'] * 10
        ph_mm = specs['page_height'] * 10
        cw = int(self.canvas['width'])
        ch = int(self.canvas['height'])
        schaal = min((cw - 10) / pw_mm, (ch - 10) / ph_mm)
        x0 = (cw - pw_mm * schaal) / 2
        y0 = (ch - ph_mm * schaal) / 2
        self.canvas.create_rectangle(x0, y0, x0 + pw_mm * schaal, y0 + ph_mm * schaal,
                                     fill="#ffffff", outline="#9a9a9a")

        b = specs['etiket_breedte'] * 10
        h = specs['etiket_hoogte'] * 10
        th = specs['h_gap'] * 10
        tv = specs['v_gap'] * 10
        links = specs['left_margin'] * 10
        boven = specs['top_margin'] * 10
        kleur = "#d13438" if fouten else self.theme.accent
        for r in range(specs['etiketten_per_kolom']):
            for k in range(specs['etiketten_per_rij']):
                lx = x0 + (links + k * (b + th)) * schaal
                ly = y0 + (boven + r * (h + tv)) * schaal
                self.canvas.create_rectangle(lx, ly, lx + b * schaal, ly + h * schaal,
                                             fill="#dbe9f7" if not fouten else "#f7dbdb",
                                             outline=kleur)

        aantal = specs['etiketten_per_rij'] * specs['etiketten_per_kolom']
        regels = [
            f"{specs['beschrijving']}",
            f"Marges: links {links:.2f} mm, rechts {specs['right_margin'] * 10:.2f} mm,".replace('.', ','),
            f"boven {boven:.2f} mm, onder {specs['bottom_margin'] * 10:.2f} mm".replace('.', ','),
        ]
        if fouten:
            regels.append("")
            regels.extend("⚠ " + f for f in fouten)
        self.info_label.config(text="\n".join(regels), fg="#d13438" if fouten else self.theme.text_secondary)

    def annuleren(self):
        self.result = None
        self.window.destroy()

    def bevestigen(self):
        naam = self.naam_var.get().strip()
        if not naam:
            messagebox.showwarning("Naam ontbreekt", "Geef het formaat een naam.", parent=self.window)
            return
        if naam != self.oude_naam and naam in self.bestaande_namen:
            messagebox.showwarning("Naam bestaat al",
                                   f"Er bestaat al een formaat met de naam '{naam}'.",
                                   parent=self.window)
            return
        try:
            inst = self.lees_instellingen()
        except ValueError:
            messagebox.showwarning("Ongeldige invoer",
                                   "Vul bij alle afmetingen een getal in (bijvoorbeeld 99,1).",
                                   parent=self.window)
            return
        fouten = controleer_formaat(inst)
        if fouten:
            messagebox.showwarning("Formaat klopt niet", "\n".join(fouten), parent=self.window)
            return
        self.result = (naam, inst)
        self.window.destroy()

    def wait_for_result(self):
        self.window.wait_window()
        return self.result


class NVictEtikettenMaker:
    def __init__(self):
        self.root = tk.Tk()
        self.theme = Theme()
        # Vaste formaten + eigen formaten uit het instellingenbestand
        self.instellingen = laad_instellingen()
        self.formaten = {}
        self.herlaad_formaten()
        start = self.instellingen.get('laatst_gekozen')
        if start not in self.formaten:
            start = next(iter(self.formaten))
        self.geselecteerd_formaat = tk.StringVar(value=start)
        self.subtitle_label = None
        self.update_available = False
        self.latest_version = None
        self.download_url = None
        self.update_label = None
        self.stored_version = None
        self.stored_release_notes = None
        
        self.root.title("NVict Etiketten Maker")
        self.root.geometry("1000x800")
        self.root.minsize(900, 700)
        self.root.state('zoomed')
        
        # Stel icoon in
        try:
            icon_path = get_resource_path("favicon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            pass
        
        self.setup_ui()
        self.root.configure(bg=self.theme.bg_primary)
        
    def setup_ui(self):
        """Setup de gebruikersinterface"""
        # Header
        header_frame = tk.Frame(self.root, bg=self.theme.bg_secondary, height=90)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        # Logo in header
        logo_container = tk.Frame(header_frame, bg=self.theme.bg_secondary)
        logo_container.pack(side=tk.LEFT, padx=25)
        
        try:
            logo_path = get_resource_path("Logo.png")
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                logo_img.thumbnail((65, 65), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_img)
                logo_label = tk.Label(logo_container, image=self.logo_photo, 
                                     bg=self.theme.bg_secondary)
                logo_label.pack()
        except:
            pass
        
        # Titel in header
        title_frame = tk.Frame(header_frame, bg=self.theme.bg_secondary)
        title_frame.pack(side=tk.LEFT, padx=15, expand=True, fill=tk.BOTH)
        
        title_label = tk.Label(title_frame, 
                              text="NVict Etiketten Maker",
                              font=("Segoe UI", 22, "bold"),
                              fg=self.theme.text_primary,
                              bg=self.theme.bg_secondary)
        title_label.pack(anchor=tk.W, pady=(20, 0))
        
        self.subtitle_label = tk.Label(title_frame,
                                 text=self.huidige_specs()['beschrijving'],
                                 font=("Segoe UI", 11),
                                 fg=self.theme.text_secondary,
                                 bg=self.theme.bg_secondary)
        self.subtitle_label.pack(anchor=tk.W)
        
        # Thema toggle knop
        theme_btn = tk.Button(header_frame,
                             text="🌙" if not self.theme.is_dark_mode else "☀️",
                             command=self.toggle_theme,
                             font=("Segoe UI", 16),
                             bg=self.theme.bg_secondary,
                             fg=self.theme.text_primary,
                             relief=tk.FLAT,
                             cursor="hand2",
                             padx=12)
        theme_btn.pack(side=tk.RIGHT, padx=25)
        
        # Main content
        main_frame = tk.Frame(self.root, bg=self.theme.bg_primary)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=50, pady=40)
        
        # Etiket selectie card
        select_card = tk.Frame(main_frame, bg=self.theme.card_bg)
        select_card.pack(fill=tk.X, pady=(0, 25))
        
        select_content = tk.Frame(select_card, bg=self.theme.card_bg)
        select_content.pack(padx=35, pady=25)
        
        tk.Label(select_content,
                text="🏷️  Selecteer Etiket Formaat",
                font=("Segoe UI", 14, "bold"),
                fg=self.theme.text_primary,
                bg=self.theme.card_bg).pack(anchor=tk.W, pady=(0, 12))
        
        # Dropdown voor etiket keuze
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Custom.TCombobox',
                       fieldbackground=self.theme.bg_secondary,
                       background=self.theme.bg_secondary,
                       foreground=self.theme.text_primary,
                       arrowcolor=self.theme.text_primary,
                       bordercolor=self.theme.border,
                       lightcolor=self.theme.bg_secondary,
                       darkcolor=self.theme.bg_secondary,
                       borderwidth=1,
                       relief='flat')
        style.map('Custom.TCombobox',
                 fieldbackground=[('readonly', self.theme.bg_secondary)],
                 selectbackground=[('readonly', self.theme.bg_secondary)],
                 selectforeground=[('readonly', self.theme.text_primary)])
        
        combo_frame = tk.Frame(select_content, bg=self.theme.card_bg)
        combo_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.formaat_combo = ttk.Combobox(combo_frame,
                                         textvariable=self.geselecteerd_formaat,
                                         values=list(self.formaten.keys()),
                                         state='readonly',
                                         style='Custom.TCombobox',
                                         font=("Segoe UI", 11),
                                         width=40)
        self.formaat_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.formaat_combo.bind('<<ComboboxSelected>>', self.update_formaat_info)

        # Knoppen voor eigen formaten
        def kleine_knop(tekst, command):
            return tk.Button(combo_frame,
                             text=tekst,
                             command=command,
                             font=("Segoe UI", 10),
                             bg=self.theme.bg_secondary,
                             fg=self.theme.text_primary,
                             activebackground=self.theme.bg_tertiary,
                             activeforeground=self.theme.text_primary,
                             disabledforeground=self.theme.border,
                             relief=tk.FLAT,
                             cursor="hand2",
                             padx=12,
                             pady=6)

        self.btn_formaat_nieuw = kleine_knop("➕ Eigen formaat…", self.nieuw_formaat)
        self.btn_formaat_nieuw.pack(side=tk.LEFT, padx=(0, 6))
        self.btn_formaat_bewerk = kleine_knop("✏️ Bewerken", self.bewerk_formaat)
        self.btn_formaat_bewerk.pack(side=tk.LEFT, padx=(0, 6))
        self.btn_formaat_verwijder = kleine_knop("🗑️ Verwijderen", self.verwijder_formaat)
        self.btn_formaat_verwijder.pack(side=tk.LEFT)

        # Info over geselecteerd formaat
        self.format_info_label = tk.Label(select_content,
                                         text="",
                                         font=("Segoe UI", 10),
                                         fg=self.theme.text_secondary,
                                         bg=self.theme.card_bg)
        self.format_info_label.pack(anchor=tk.W, pady=(5, 0))
        self.update_formaat_info()
        
        # Info card
        info_card = tk.Frame(main_frame, bg=self.theme.card_bg)
        info_card.pack(fill=tk.X, pady=(0, 25))
        
        info_content = tk.Frame(info_card, bg=self.theme.card_bg)
        info_content.pack(padx=35, pady=30)
        
        info_icon = tk.Label(info_content,
                            text="⚡",
                            font=("Segoe UI", 40),
                            bg=self.theme.card_bg)
        info_icon.pack()
        
        info_title = tk.Label(info_content,
                             text="Excel naar Etiketten",
                             font=("Segoe UI", 18, "bold"),
                             fg=self.theme.text_primary,
                             bg=self.theme.card_bg)
        info_title.pack(pady=(12, 8))
        
        info_text = tk.Label(info_content,
                            text="Converteer uw Excel adressenlijst automatisch naar etiketten.",
                            font=("Segoe UI", 11),
                            fg=self.theme.text_secondary,
                            bg=self.theme.card_bg,
                            justify=tk.CENTER)
        info_text.pack()
        
        # Features card
        features_card = tk.Frame(main_frame, bg=self.theme.card_bg)
        features_card.pack(fill=tk.X, pady=(0, 25))
        
        features_content = tk.Frame(features_card, bg=self.theme.card_bg)
        features_content.pack(padx=35, pady=25)
        
        tk.Label(features_content,
                text="✨ Automatisch gedetecteerd:",
                font=("Segoe UI", 14, "bold"),
                fg=self.theme.text_primary,
                bg=self.theme.card_bg).pack(anchor=tk.W, pady=(0, 15))
        
        features = [
            ("📝 Regel 1", "Aanhef, voorletters en achternaam"),
            ("📌 Regel 2", "Extra (indien aanwezig)"),
            ("🏠 Regel 3", "Adres en huisnummer"),
            ("📮 Regel 4", "Postcode en woonplaats")
        ]
        
        for icon_text, desc in features:
            feature_frame = tk.Frame(features_content, bg=self.theme.card_bg)
            feature_frame.pack(fill=tk.X, pady=5)
            
            tk.Label(feature_frame,
                    text=icon_text,
                    font=("Segoe UI", 11, "bold"),
                    fg=self.theme.accent,
                    bg=self.theme.card_bg).pack(side=tk.LEFT, padx=(0, 10))
            
            tk.Label(feature_frame,
                    text=desc,
                    font=("Segoe UI", 11),
                    fg=self.theme.text_secondary,
                    bg=self.theme.card_bg).pack(side=tk.LEFT)
        
        # Action buttons
        btn_frame = tk.Frame(main_frame, bg=self.theme.bg_primary)
        btn_frame.pack(pady=25)

        self.start_btn = tk.Button(btn_frame,
                                   text="📂  Selecteer Excel bestand",
                                   command=self.maak_etiketten,
                                   font=("Segoe UI", 13, "bold"),
                                   bg=self.theme.accent,
                                   fg="#ffffff",
                                   activebackground=self.theme.accent_hover,
                                   activeforeground="#ffffff",
                                   relief=tk.FLAT,
                                   cursor="hand2",
                                   padx=45,
                                   pady=18)
        self.start_btn.pack()

        self.start_btn.bind("<Enter>", lambda e: self.start_btn.config(bg=self.theme.accent_hover))
        self.start_btn.bind("<Leave>", lambda e: self.start_btn.config(bg=self.theme.accent))

        self.één_adres_btn = tk.Button(btn_frame,
                                       text="📝  Één adres op alle etiketten",
                                       command=self.maak_etiketten_één_adres,
                                       font=("Segoe UI", 11),
                                       bg=self.theme.card_bg,
                                       fg=self.theme.text_primary,
                                       activebackground=self.theme.bg_tertiary,
                                       activeforeground=self.theme.text_primary,
                                       relief=tk.FLAT,
                                       cursor="hand2",
                                       padx=45,
                                       pady=12)
        self.één_adres_btn.pack(pady=(10, 0))

        self.één_adres_btn.bind("<Enter>", lambda e: self.één_adres_btn.config(bg=self.theme.bg_tertiary))
        self.één_adres_btn.bind("<Leave>", lambda e: self.één_adres_btn.config(bg=self.theme.card_bg))
        
        # Footer
        footer_frame = tk.Frame(self.root, bg=self.theme.bg_secondary, height=50)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        footer_content = tk.Frame(footer_frame, bg=self.theme.bg_secondary)
        footer_content.pack(expand=True)
        
        # Versie info links
        version_label = tk.Label(footer_content,
                                text=f"v{APP_VERSION}",
                                font=("Segoe UI", 9),
                                fg=self.theme.text_secondary,
                                bg=self.theme.bg_secondary)
        version_label.pack(side=tk.LEFT, padx=(0, 15))
        
        # Update notificatie (verborgen tot update beschikbaar is)
        self.update_label = tk.Label(footer_content,
                                     text="",
                                     font=("Segoe UI", 10, "bold", "underline"),
                                     fg=self.theme.accent,
                                     bg=self.theme.bg_secondary,
                                     cursor="hand2")
        self.update_label.pack(side=tk.LEFT, padx=(0, 15))
        
        # Main footer text
        footer_label = tk.Label(footer_content,
                               text="NVict Service  •  www.nvict.nl",
                               font=("Segoe UI", 10),
                               fg=self.theme.text_secondary,
                               bg=self.theme.bg_secondary)
        footer_label.pack(side=tk.LEFT)
        
        # Start update check in achtergrond (na 2 seconden)
        self.root.after(2000, self.check_for_updates_on_startup)
    
    def check_for_updates_on_startup(self):
        """Check voor updates bij opstarten"""
        threading.Thread(target=self.check_for_updates, daemon=True).start()
    
    def check_for_updates(self):
        """Check voor updates (draait in achtergrond thread)"""
        urls_to_try = [UPDATE_CHECK_URL, UPDATE_CHECK_URL_FALLBACK]
        
        for url_index, url in enumerate(urls_to_try):
            try:
                req = urllib.request.Request(
                    url,
                    headers={'User-Agent': 'NVict-Etiketten-Maker'}
                )
                
                with urllib.request.urlopen(req, timeout=5) as response:
                    data = json.loads(response.read().decode())
                    
                    latest_version = data.get('version', '')
                    download_url = data.get('download_url', '')
                    release_notes = data.get('release_notes', '')
                    
                    # Vergelijk versies
                    if self.is_newer_version(latest_version, APP_VERSION):
                        self.latest_version = latest_version
                        self.download_url = download_url
                        self.update_available = True
                        
                        # Update UI in main thread
                        self.root.after(0, self.show_update_notification, latest_version, release_notes)
                    
                    return
            
            except urllib.error.HTTPError as e:
                if e.code == 403 and url_index < len(urls_to_try) - 1:
                    continue
            except urllib.error.URLError:
                if url_index < len(urls_to_try) - 1:
                    continue
            except:
                pass
    
    def is_newer_version(self, latest, current):
        """Vergelijk versie nummers"""
        try:
            latest_parts = [int(x) for x in latest.split('.')]
            current_parts = [int(x) for x in current.split('.')]
            
            while len(latest_parts) < 3:
                latest_parts.append(0)
            while len(current_parts) < 3:
                current_parts.append(0)
            
            return latest_parts > current_parts
        except:
            return False
    
    def show_update_notification(self, version, release_notes):
        """Toon update notificatie in de UI"""
        self.stored_version = version
        self.stored_release_notes = release_notes
        
        if self.update_label:
            update_text = f"🔔 Update beschikbaar: v{version}"
            self.update_label.config(text=update_text)
            
            self.update_label.unbind("<Button-1>")
            
            def on_click(event):
                self.show_update_dialog(self.stored_version, self.stored_release_notes)
            
            self.update_label.bind("<Button-1>", on_click)
        
        # Automatisch dialog tonen
        self.show_update_dialog(version, release_notes)
    
    def show_update_dialog(self, version, release_notes):
        """Toon update dialog met details"""
        try:
            update_window = tk.Toplevel(self.root)
            update_window.title("Update Beschikbaar")
            update_window.geometry("600x550")
            update_window.minsize(550, 500)
            update_window.transient(self.root)
            update_window.grab_set()
            update_window.configure(bg=self.theme.bg_secondary)
            
            # Icoon instellen
            try:
                icon_path = get_resource_path("favicon.ico")
                if os.path.exists(icon_path):
                    update_window.iconbitmap(icon_path)
            except:
                pass
            
            # Centreer
            update_window.update_idletasks()
            x = (update_window.winfo_screenwidth() // 2) - (600 // 2)
            y = (update_window.winfo_screenheight() // 2) - (550 // 2)
            update_window.geometry(f"600x550+{x}+{y}")
        except:
            return
        
        content_frame = tk.Frame(update_window, bg=self.theme.bg_secondary)
        content_frame.pack(expand=True, fill=tk.BOTH, padx=30, pady=30)
        
        # Icon
        icon_label = tk.Label(content_frame,
                             text="🔔",
                             font=("Segoe UI", 48),
                             bg=self.theme.bg_secondary)
        icon_label.pack(pady=(0, 15))
        
        # Titel
        title_label = tk.Label(content_frame,
                              text="Update Beschikbaar!",
                              font=("Segoe UI", 16, "bold"),
                              fg=self.theme.text_primary,
                              bg=self.theme.bg_secondary)
        title_label.pack(pady=(0, 10))
        
        # Versie info
        version_text = f"Huidige versie: v{APP_VERSION}\nNieuwe versie: v{version}"
        version_label = tk.Label(content_frame,
                                text=version_text,
                                font=("Segoe UI", 11),
                                fg=self.theme.text_secondary,
                                bg=self.theme.bg_secondary,
                                justify=tk.CENTER)
        version_label.pack(pady=(0, 15))
        
        # Release notes
        if release_notes:
            notes_frame = tk.Frame(content_frame, bg=self.theme.card_bg)
            notes_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
            
            notes_label = tk.Label(notes_frame,
                                  text="Wat is er nieuw:",
                                  font=("Segoe UI", 10, "bold"),
                                  fg=self.theme.text_primary,
                                  bg=self.theme.card_bg)
            notes_label.pack(anchor=tk.W, padx=15, pady=(10, 5))
            
            notes_text = tk.Text(notes_frame,
                               font=("Segoe UI", 10),
                               fg=self.theme.text_secondary,
                               bg=self.theme.card_bg,
                               height=6,
                               wrap=tk.WORD,
                               relief=tk.FLAT)
            notes_text.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))
            notes_text.insert("1.0", release_notes)
            notes_text.config(state=tk.DISABLED)
        
        # Knoppen
        btn_frame = tk.Frame(content_frame, bg=self.theme.bg_secondary)
        btn_frame.pack(pady=(10, 0))
        
        def download_and_install():
            if self.download_url:
                update_window.destroy()
                self.download_and_install_update(self.download_url, version)
        
        def download_only():
            if self.download_url:
                webbrowser.open(self.download_url)
                update_window.destroy()
        
        download_install_btn = tk.Button(btn_frame,
                                        text="Download & Installeer",
                                        command=download_and_install,
                                        font=("Segoe UI", 11, "bold"),
                                        bg=self.theme.success_color,
                                        fg="#ffffff",
                                        relief=tk.FLAT,
                                        cursor="hand2",
                                        padx=20,
                                        pady=10)
        download_install_btn.pack(side=tk.LEFT, padx=3)
        
        download_btn = tk.Button(btn_frame,
                                text="Alleen Download",
                                command=download_only,
                                font=("Segoe UI", 11),
                                bg=self.theme.accent,
                                fg="#ffffff",
                                relief=tk.FLAT,
                                cursor="hand2",
                                padx=20,
                                pady=10)
        download_btn.pack(side=tk.LEFT, padx=3)
        
        later_btn = tk.Button(btn_frame,
                             text="Later",
                             command=update_window.destroy,
                             font=("Segoe UI", 11),
                             bg=self.theme.card_bg,
                             fg=self.theme.text_primary,
                             relief=tk.FLAT,
                             cursor="hand2",
                             padx=20,
                             pady=10)
        later_btn.pack(side=tk.LEFT, padx=3)
    
    def download_and_install_update(self, download_url, version):
        """Download update en start installatie automatisch"""
        try:
            # Toon voortgang dialoog
            progress_dialog = tk.Toplevel(self.root)
            progress_dialog.title("Update Downloaden")
            progress_dialog.geometry("400x150")
            progress_dialog.configure(bg=self.theme.bg_primary)
            progress_dialog.transient(self.root)
            progress_dialog.resizable(False, False)
            
            try:
                icon_path = get_resource_path('favicon.ico')
                if os.path.exists(icon_path):
                    progress_dialog.iconbitmap(icon_path)
            except:
                pass
            
            tk.Label(progress_dialog, text="Update downloaden...", 
                    font=("Segoe UI", 12, "bold"),
                    bg=self.theme.bg_primary, 
                    fg=self.theme.text_primary).pack(pady=20)
            
            status_label = tk.Label(progress_dialog, text="Bezig met downloaden...",
                                   font=("Segoe UI", 9),
                                   bg=self.theme.bg_primary,
                                   fg=self.theme.text_secondary)
            status_label.pack(pady=10)
            
            progress_dialog.update()
            
            # Download in achtergrond thread
            def download_thread():
                try:
                    temp_dir = tempfile.gettempdir()
                    filename = f"NVict_Etiketten_v{version}_Setup.exe"
                    filepath = os.path.join(temp_dir, filename)
                    
                    urllib.request.urlretrieve(download_url, filepath)
                    
                    self.root.after(0, lambda: self._finish_download(progress_dialog, filepath))
                    
                except Exception as e:
                    self.root.after(0, lambda: self._download_error(progress_dialog, str(e)))
            
            thread = threading.Thread(target=download_thread, daemon=True)
            thread.start()
            
        except Exception as e:
            messagebox.showerror("Download Fout", 
                f"Kan update niet downloaden:\n{str(e)}\n\nProbeer handmatig te downloaden via de website.")
    
    def _finish_download(self, progress_dialog, filepath):
        """Voltooi download en start installer"""
        try:
            progress_dialog.destroy()
            
            if os.path.exists(filepath):
                if messagebox.askyesno("Update Download Voltooid",
                    f"Update succesvol gedownload!\n\n"
                    f"Wilt u de installer nu starten?\n\n"
                    f"Let op: Sluit eerst NVict Etiketten Maker af voordat u de installatie voltooit."):
                    
                    if platform.system() == "Windows":
                        os.startfile(filepath)
                    elif platform.system() == "Darwin":
                        subprocess.run(["open", filepath])
                    else:
                        subprocess.run(["xdg-open", filepath])
                    
                    self.root.after(1000, self.root.destroy)
            else:
                messagebox.showerror("Fout", "Download bestand niet gevonden")
                
        except Exception as e:
            messagebox.showerror("Fout", f"Kan installer niet starten:\n{str(e)}")
    
    def _download_error(self, progress_dialog, error_msg):
        """Toon download fout"""
        progress_dialog.destroy()
        messagebox.showerror("Download Fout",
            f"Kan update niet downloaden:\n{error_msg}\n\n"
            f"Probeer handmatig te downloaden via de website.")
    
    def toggle_theme(self):
        """Wissel tussen licht en donker thema"""
        self.theme.toggle()
        for widget in self.root.winfo_children():
            widget.destroy()
        self.setup_ui()
    
    def herlaad_formaten(self):
        """Vul self.formaten: eerst de vaste formaten, daarna de eigen formaten."""
        formaten = {}
        for naam, inst in STANDAARD_FORMATEN.items():
            formaten[naam] = maak_formaat_specs(inst, eigen=False)
        for naam, inst in self.instellingen.get('eigen_formaten', {}).items():
            if naam in formaten:
                continue
            try:
                formaten[naam] = maak_formaat_specs(inst, eigen=True)
            except (KeyError, TypeError, ValueError):
                continue
        self.formaten = formaten

    def huidige_specs(self):
        naam = self.geselecteerd_formaat.get()
        if naam not in self.formaten:
            naam = next(iter(self.formaten))
            self.geselecteerd_formaat.set(naam)
        return self.formaten[naam]

    def update_formaat_info(self, event=None):
        """Update de formaat info labels wanneer er een nieuw formaat wordt gekozen"""
        formaat = self.geselecteerd_formaat.get()
        specs = self.huidige_specs()
        eigen = bool(specs.get('eigen'))

        tekst = specs['beschrijving']
        if eigen:
            tekst += "  •  eigen formaat"
        if self.subtitle_label:
            self.subtitle_label.config(text=tekst)

        if hasattr(self, 'format_info_label'):
            inst = specs.get('instellingen', {})
            detail = ""
            try:
                nr = specs['etiketten_per_rij']
                nk = specs['etiketten_per_kolom']
                detail = (f"  —  {nr} naast elkaar × {nk} onder elkaar op {specs.get('pagina', 'A4')}, "
                          f"marge links {specs['left_margin'] * 10:.2f} mm, boven {specs['top_margin'] * 10:.2f} mm"
                          ).replace('.', ',')
                if specs.get('h_gap') or specs.get('v_gap'):
                    detail += (f", tussenruimte {specs.get('h_gap', 0) * 10:.1f} × "
                               f"{specs.get('v_gap', 0) * 10:.1f} mm").replace('.', ',')
            except (KeyError, TypeError):
                pass
            self.format_info_label.config(text=f"✓ {tekst}{detail}")

        for knop in (getattr(self, 'btn_formaat_bewerk', None), getattr(self, 'btn_formaat_verwijder', None)):
            if knop is not None:
                knop.config(state=tk.NORMAL if eigen else tk.DISABLED,
                            cursor="hand2" if eigen else "arrow")

        if formaat in self.formaten and self.instellingen.get('laatst_gekozen') != formaat:
            self.instellingen['laatst_gekozen'] = formaat
            bewaar_instellingen(self.instellingen)

    def ververs_formaat_keuze(self, selecteer=None):
        """Dropdown opnieuw vullen na toevoegen/bewerken/verwijderen."""
        self.herlaad_formaten()
        if hasattr(self, 'formaat_combo'):
            self.formaat_combo.config(values=list(self.formaten.keys()))
        if selecteer in self.formaten:
            self.geselecteerd_formaat.set(selecteer)
        elif self.geselecteerd_formaat.get() not in self.formaten:
            self.geselecteerd_formaat.set(next(iter(self.formaten)))
        self.update_formaat_info()

    def nieuw_formaat(self):
        dialog = EigenFormaatWindow(self.root, self.theme, self.formaten.keys())
        result = dialog.wait_for_result()
        if result is None:
            return
        naam, inst = result
        self.instellingen.setdefault('eigen_formaten', {})[naam] = inst
        if not bewaar_instellingen(self.instellingen):
            messagebox.showwarning("Niet opgeslagen",
                                   "Het formaat kon niet worden bewaard in\n" + instellingen_pad() +
                                   "\n\nHet blijft wel beschikbaar tot u het programma afsluit.")
        self.ververs_formaat_keuze(naam)

    def bewerk_formaat(self):
        naam = self.geselecteerd_formaat.get()
        specs = self.formaten.get(naam)
        if not specs or not specs.get('eigen'):
            messagebox.showinfo("Vast formaat",
                                "Dit is een vast formaat en kan niet bewerkt worden.\n"
                                "Maak met '➕ Eigen formaat…' een eigen variant.")
            return
        dialog = EigenFormaatWindow(self.root, self.theme, self.formaten.keys(),
                                    naam=naam, inst=specs.get('instellingen', {}))
        result = dialog.wait_for_result()
        if result is None:
            return
        nieuwe_naam, inst = result
        eigen = self.instellingen.setdefault('eigen_formaten', {})
        eigen.pop(naam, None)
        eigen[nieuwe_naam] = inst
        bewaar_instellingen(self.instellingen)
        self.ververs_formaat_keuze(nieuwe_naam)

    def verwijder_formaat(self):
        naam = self.geselecteerd_formaat.get()
        specs = self.formaten.get(naam)
        if not specs or not specs.get('eigen'):
            messagebox.showinfo("Vast formaat", "Vaste formaten kunnen niet verwijderd worden.")
            return
        if not messagebox.askyesno("Formaat verwijderen",
                                   f"Eigen formaat '{naam}' verwijderen?"):
            return
        self.instellingen.setdefault('eigen_formaten', {}).pop(naam, None)
        bewaar_instellingen(self.instellingen)
        self.ververs_formaat_keuze(next(iter(STANDAARD_FORMATEN)))
    
    def get_voorbeelden_map(self):
        """Bepaal de locatie van de voorbeelden map"""
        try:
            if os.name == 'nt':
                import winreg
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                   r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders") as key:
                    docs_folder = winreg.QueryValueEx(key, "Personal")[0]
                    voorbeelden_map = os.path.join(docs_folder, "Etiketten Maker")
            else:
                docs_folder = os.path.expanduser("~/Documents")
                voorbeelden_map = os.path.join(docs_folder, "Etiketten Maker")
            
            if not os.path.exists(voorbeelden_map):
                os.makedirs(voorbeelden_map, exist_ok=True)
            
            return voorbeelden_map
        except:
            return os.path.expanduser("~")
    
    def maak_etiketten(self):
        """Hoofdfunctie om etiketten te maken"""
        start_dir = self.get_voorbeelden_map()
        
        excel_bestand = filedialog.askopenfilename(
            title="Selecteer Excel bestand met adressen",
            initialdir=start_dir,
            filetypes=[("Excel bestanden", "*.xlsx *.xls"), ("Alle bestanden", "*.*")]
        )
        
        if not excel_bestand:
            return

        _ok = None
        try:
            df = pd.read_excel(excel_bestand)
            
            if len(df) == 0:
                messagebox.showerror("Fout", "Het Excel bestand bevat geen gegevens.")
                return
            
            auto_regels = detecteer_adres_velden(df.columns.tolist())
            
            if not auto_regels:
                messagebox.showerror("Fout", 
                    "Kon geen adresvelden detecteren in het Excel bestand.\n\n" +
                    "Zorg ervoor dat de kolommen herkenbare namen hebben zoals:\n" +
                    "Naam, Adres, Postcode, Woonplaats, etc.")
                return
            
            formaat_naam = self.geselecteerd_formaat.get()
            specs = self.huidige_specs()

            preview_window = EtiketAanpassingWindow(self.root, self.theme, df, excel_bestand,
                                                    auto_regels, specs)
            settings = preview_window.wait_for_result()
            
            if settings is None:
                return
            
            # Progress window
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Etiketten maken...")
            progress_window.geometry("450x180")
            progress_window.transient(self.root)
            progress_window.grab_set()
            progress_window.configure(bg=self.theme.bg_secondary)
            
            try:
                icon_path = get_resource_path("favicon.ico")
                if os.path.exists(icon_path):
                    progress_window.iconbitmap(icon_path)
            except:
                pass
            
            progress_window.update_idletasks()
            x = (progress_window.winfo_screenwidth() // 2) - (450 // 2)
            y = (progress_window.winfo_screenheight() // 2) - (180 // 2)
            progress_window.geometry(f"450x180+{x}+{y}")
            
            progress_label = tk.Label(progress_window,
                                     text="⏳ Etiketten worden aangemaakt...",
                                     font=("Segoe UI", 13),
                                     fg=self.theme.text_primary,
                                     bg=self.theme.bg_secondary)
            progress_label.pack(expand=True)
            progress_window.update()
            
            # Maak Word document
            doc = Document()
            stel_pagina_in(doc, specs)

            etiketten_per_vel = specs['etiketten_per_rij'] * specs['etiketten_per_kolom']
            aantal_vellen = (len(df) + etiketten_per_vel - 1) // etiketten_per_vel

            regels_config = settings['regels']
            font_grootte = settings['font_grootte']
            uitlijning = settings['uitlijning']

            align_map = {
                'links': WD_ALIGN_PARAGRAPH.LEFT,
                'midden': WD_ALIGN_PARAGRAPH.CENTER,
                'rechts': WD_ALIGN_PARAGRAPH.RIGHT
            }
            word_alignment = align_map.get(uitlijning, WD_ALIGN_PARAGRAPH.LEFT)

            # Maak etiketten: per vel één tabel, cellen in leesvolgorde
            etiket_index = 0
            vel_nummer = 0

            while etiket_index < len(df):
                cellen = maak_etiket_vel(doc, specs, nieuwe_pagina=vel_nummer > 0)
                vel_nummer += 1
                for cell in cellen:
                    if etiket_index >= len(df):
                        break
                    regels = etiket_regels_uit_rij(df.iloc[etiket_index], regels_config)
                    vul_etiket_cel(cell, regels, font_grootte, word_alignment)
                    etiket_index += 1

            # Opslaan
            formaat_code = bestandsnaam_code(formaat_naam)
            output_bestand = excel_bestand.replace('.xlsx', f'_etiketten_{formaat_code}.docx').replace('.xls', f'_etiketten_{formaat_code}.docx')
            doc.save(output_bestand)
            progress_window.destroy()
            _ok = (output_bestand, len(df), aantal_vellen)

        except Exception as e:
            if 'progress_window' in locals():
                try:
                    progress_window.destroy()
                except:
                    pass
            _ok = None
            messagebox.showerror("Fout", f"Er is een fout opgetreden:\n\n{str(e)}")

        if _ok:
            self.toon_success_dialog(*_ok)
    
    def maak_etiketten_één_adres(self):
        """Maak etiketten met één adres herhaald op alle posities"""
        formaat_naam = self.geselecteerd_formaat.get()
        specs = self.huidige_specs()

        dialog = EénAdresWindow(self.root, self.theme, formaat_naam, specs)
        result = dialog.wait_for_result()

        if result is None:
            return

        regels = result['regels']
        aantal_vellen = result['aantal_vellen']
        font_grootte = result['font_grootte']

        etiketten_per_rij = specs['etiketten_per_rij']
        etiketten_per_kolom = specs['etiketten_per_kolom']
        etiketten_per_vel = etiketten_per_rij * etiketten_per_kolom
        totaal_etiketten = aantal_vellen * etiketten_per_vel

        # Vraag opslaglocatie
        start_dir = self.get_voorbeelden_map()
        formaat_code = bestandsnaam_code(formaat_naam)

        output_bestand = filedialog.asksaveasfilename(
            title="Sla etiketten op als...",
            initialdir=start_dir,
            initialfile=f"etiketten_{formaat_code}.docx",
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx"), ("Alle bestanden", "*.*")]
        )

        if not output_bestand:
            return

        # Progress window
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Etiketten maken...")
        progress_window.geometry("450x180")
        progress_window.transient(self.root)
        progress_window.grab_set()
        progress_window.configure(bg=self.theme.bg_secondary)

        try:
            icon_path = get_resource_path("favicon.ico")
            if os.path.exists(icon_path):
                progress_window.iconbitmap(icon_path)
        except:
            pass

        progress_window.update_idletasks()
        x = (progress_window.winfo_screenwidth() // 2) - (450 // 2)
        y = (progress_window.winfo_screenheight() // 2) - (180 // 2)
        progress_window.geometry(f"450x180+{x}+{y}")

        tk.Label(progress_window,
                 text="⏳ Etiketten worden aangemaakt...",
                 font=("Segoe UI", 13),
                 fg=self.theme.text_primary,
                 bg=self.theme.bg_secondary).pack(expand=True)
        progress_window.update()

        _ok = None
        try:
            doc = Document()
            stel_pagina_in(doc, specs)

            etiket_index = 0
            vel_nummer = 0

            while etiket_index < totaal_etiketten:
                cellen = maak_etiket_vel(doc, specs, nieuwe_pagina=vel_nummer > 0)
                vel_nummer += 1
                for cell in cellen:
                    if etiket_index >= totaal_etiketten:
                        break
                    vul_etiket_cel(cell, regels, font_grootte, WD_ALIGN_PARAGRAPH.LEFT)
                    etiket_index += 1

            doc.save(output_bestand)
            progress_window.destroy()
            _ok = (output_bestand, totaal_etiketten, aantal_vellen)

        except Exception as e:
            try:
                progress_window.destroy()
            except:
                pass
            _ok = None
            messagebox.showerror("Fout", f"Er is een fout opgetreden:\n\n{str(e)}")

        if _ok:
            self.toon_success_dialog(*_ok)

    def toon_success_dialog(self, output_bestand, aantal_adressen, aantal_vellen):
        """Toon success dialog"""
        success_window = tk.Toplevel(self.root)
        success_window.title("✅ Gereed")
        success_window.geometry("550x350")
        success_window.transient(self.root)
        success_window.grab_set()
        success_window.configure(bg=self.theme.bg_secondary)
        
        try:
            icon_path = get_resource_path("favicon.ico")
            if os.path.exists(icon_path):
                success_window.iconbitmap(icon_path)
        except:
            pass
        
        success_window.update_idletasks()
        x = (success_window.winfo_screenwidth() // 2) - (550 // 2)
        y = (success_window.winfo_screenheight() // 2) - (350 // 2)
        success_window.geometry(f"550x350+{x}+{y}")
        
        content_frame = tk.Frame(success_window, bg=self.theme.bg_secondary)
        content_frame.pack(expand=True, padx=45, pady=35)
        
        success_icon = tk.Label(content_frame,
                               text="✅",
                               font=("Segoe UI", 52),
                               bg=self.theme.bg_secondary)
        success_icon.pack(pady=(0, 18))
        
        success_title = tk.Label(content_frame,
                                text="Etiketten zijn aangemaakt!",
                                font=("Segoe UI", 17, "bold"),
                                fg=self.theme.text_primary,
                                bg=self.theme.bg_secondary)
        success_title.pack(pady=(0, 12))
        
        file_label = tk.Label(content_frame,
                             text=os.path.basename(output_bestand),
                             font=("Segoe UI", 11, "underline"),
                             fg=self.theme.accent,
                             bg=self.theme.bg_secondary,
                             cursor="hand2")
        file_label.pack(pady=8)
        file_label.bind("<Button-1>", lambda e: os.startfile(output_bestand))
        
        stats_frame = tk.Frame(content_frame, bg=self.theme.card_bg)
        stats_frame.pack(pady=18, fill=tk.X)
        
        stats_text = f"📊 {aantal_adressen} adressen  •  📄 {aantal_vellen} {'vel' if aantal_vellen == 1 else 'vellen'}"
        stats_label = tk.Label(stats_frame,
                              text=stats_text,
                              font=("Segoe UI", 11),
                              fg=self.theme.text_secondary,
                              bg=self.theme.card_bg,
                              padx=25,
                              pady=12)
        stats_label.pack()
        
        btn_frame = tk.Frame(content_frame, bg=self.theme.bg_secondary)
        btn_frame.pack(pady=(18, 0))
        
        open_btn = tk.Button(btn_frame,
                            text="📂 Open bestand",
                            command=lambda: [os.startfile(output_bestand), success_window.destroy()],
                            font=("Segoe UI", 11, "bold"),
                            bg=self.theme.accent,
                            fg="#ffffff",
                            relief=tk.FLAT,
                            cursor="hand2",
                            padx=25,
                            pady=10)
        open_btn.pack(side=tk.LEFT, padx=6)
        
        close_btn = tk.Button(btn_frame,
                             text="Sluiten",
                             command=success_window.destroy,
                             font=("Segoe UI", 11),
                             bg=self.theme.card_bg,
                             fg=self.theme.text_primary,
                             relief=tk.FLAT,
                             cursor="hand2",
                             padx=25,
                             pady=10)
        close_btn.pack(side=tk.LEFT, padx=6)
    
    def run(self):
        """Start de applicatie"""
        self.root.mainloop()

if __name__ == "__main__":
    app = NVictEtikettenMaker()
    app.run()
