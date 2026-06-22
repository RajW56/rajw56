# -*- coding: utf-8 -*-
"""
Clean area data + identify tower/wing for Advay June projects, FILLED IN PLACE.

Ground truth = Marathi 'Property Discription' (most consistent), English as fallback.
Filled columns:
  Flat Area              = flat RERA carpet area, sq.m
  Balcony area           = balcony area, sq.m
  Additional Area        = other attached carpet area (adjacent/utility/EBVT/other), sq.m
  Actual Areas           = (Flat + Balcony + Additional) * 10.764, sq.ft  (the true total RERA carpet)
  Tower Name             = tower / wing identified from description
  Area and Price Remarks = extraction notes / data-quality flags
"""
import re
import openpyxl

SRC = "/root/.claude/uploads/e76392c9-fd1b-51e2-8503-d2aabea764e9/19e351e4-Advay_June_All_projects.xlsx"
OUT = "/home/user/rajw56/advay-june-area-cleaning/Advay_June_All_projects.xlsx"
SQM2SQFT = 10.764

MLET = {'ए': 'A', 'बी': 'B', 'सी': 'C', 'डी': 'D', 'ई': 'E', 'एफ': 'F',
        'जी': 'G', 'एच': 'H', 'आय': 'I', 'जे': 'J', 'के': 'K', 'एल': 'L'}

RIVALI = {'मूनराईज': 'Moonrise', 'मुनराईस': 'Moonrise', 'स्टारगेझ': 'Stargaze',
          'सनबर्स्ट': 'Sunburst', 'सनबस्ट': 'Sunburst', 'स्कायलीप': 'Skyleap'}

# number + चौ + unit ; unit first char म => sq.m, फ => sq.ft (handles मी/मीटर/फुट/फूट/फ़ुट)
NUMUNIT = r'(\d+(?:\.\d+)?)\s*चौ\.?\s*(म|फ)[ऀ-ॿ]*'


def _unit(c):
    return 'sqm' if c == 'म' else 'sqft'

def _to_sqm(val, unit):
    return round(val, 2) if unit == 'sqm' else round(val / SQM2SQFT, 2)

def _first_after(text, anchor, window=45):
    """First number+unit within `window` chars after anchor -> (val, unit) or None."""
    m = re.search(anchor, text)
    if not m:
        return None
    n = re.search(NUMUNIT, text[m.end():m.end() + window])
    return (float(n.group(1)), _unit(n.group(2))) if n else None

def _carpet(text):
    """First '<num> चौ.<unit> [रेरा] कार[्]पेट' -> flat carpet (val, unit).
    Tolerates '.', spaces and '(रेरा)' between the unit and कारपेट."""
    m = re.search(NUMUNIT + r'[\s.()]*(?:रेरा[\s.()]*)?कार[्]?पेट', text)
    return (float(m.group(1)), _unit(m.group(2))) if m else None

def _total_doc(text):
    """Deed's stated total RERA carpet: 'एकूण क्षेत्र(फळ) <num> चौ.. [रेरा] कारपेट'.
    Returns total in sq.m, or None. Requires a carpet keyword nearby so the
    parking 'एकूण क्षेत्रफळ 111.35 चौ.फुट' is NOT picked up."""
    m = re.search(r'(?:एकूण|एकुण)\s*क्षेत्र(?:फळ)?\s*:?-?\s*' + NUMUNIT, text)
    if m and 'कार' in text[m.end():m.end() + 30]:
        return _to_sqm(float(m.group(1)), _unit(m.group(2)))
    return None


def wing_marathi(text):
    for pat in (r'([ऀ-ॿ]+)\s*विंग', r'विंग\s*[-:]?\s*([ऀ-ॿ]+)', r'विंग\s*[-:]?\s*([0-9]+)'):
        for m in re.finditer(pat, text):
            tok = m.group(1)
            if tok.isdigit():
                return tok
            if tok in MLET:
                return MLET[tok]
    return None

def wing_english(text):
    m = re.search(r'\bWing\s*[:\-]?\s*([A-Z0-9]+)\b', text) or re.search(r'\b([A-Z0-9])\s*Wing\b', text)
    return m.group(1) if m else None


def extract(proj, mar, eng):
    t = mar if isinstance(mar, str) else ''
    e = eng if isinstance(eng, str) else ''
    # fix OCR artefact "5. 79" / "105. 26" -> "5.79" / "105.26"
    t = re.sub(r'(\d)\.\s+(\d)', r'\1.\2', t)
    flat = balc = add = None
    tower = None
    remark = []

    # ---------------- AREAS ----------------
    if proj == 'Codename Live 360 (Skyline Icon)':
        bu = _first_after(t, r'क्षेत्रफळ', 25)
        if bu and 'बिल्ट' in t:
            remark.append('description gives built-up only (%.2f sqm); carpet not stated' % _to_sqm(*bu))
        else:
            remark.append('no flat area stated in description')
        tower = wing_marathi(t) or wing_english(e)

    elif proj == 'Narang Vivenda':
        c = _first_after(t, r'चटई\s*क्षेत्रफळ', 30)          # carpet (sqft)
        if c:
            flat = _to_sqm(*c)
        o = _first_after(t, r'इतर\s*क्षेत्रफळ', 30)          # other area
        if o:
            add = _to_sqm(*o)
        tower = wing_marathi(t) or wing_english(e)

    elif proj == 'Borivali Project (Lodha Altus)':
        # redevelopment: new flat carpet stated as "क्षेत्रफळ कारपेट 493 चौ. फूट"
        m = re.search(r'क्षेत्रफळ\s*कारपेट\s*' + NUMUNIT, t)
        if m:
            flat = _to_sqm(float(m.group(1)), _unit(m.group(2)))
            remark.append('new-flat carpet (redevelopment); old area ignored')
        else:
            remark.append('no carpet area stated in description')

    else:
        # for redevelopment/supplementary deeds, read the NEW flat: slice at the
        # 'नवीन' that is immediately followed by its क्षेत्र (skip the old-flat area)
        csrc = t
        if re.search(r'जुन', t):
            mnew = re.search(r'नव(?:ीन|िन)[^.]{0,30}?क्षेत्र', t)
            if mnew:
                csrc = t[mnew.start():]
        c = _carpet(csrc)
        if c:
            flat = _to_sqm(*c)
        # balcony (बाल्कनी / बालकनी)
        b = _first_after(t, r'बाल्?कनी', 45)
        if b and b[0] > 0:
            balc = _to_sqm(*b)
        # additional attached carpet (लगत/उपयोगिता/EBVT/utility/purchased area)
        parts = []
        for anc in (r'लगत', r'उपयोगिता', r'ई\.?\s*बी\.?\s*व्ही\.?\s*टी',
                    r'युट[िी]लिटी', r'विकत\s*घेतलेले'):
            a = _first_after(t, anc, 35)
            if a:
                parts.append(_to_sqm(*a))
        if parts:
            add = round(sum(parts), 2)
        if flat is None and 'बिल्ट' not in t and 'बांधीव' not in t:
            remark.append('no carpet area stated in description')
        elif flat is None:
            remark.append('description gives built-up only; carpet not stated')

    # ---------------- TOWER ----------------
    if proj == 'Sky City':
        me = re.search(r'Tower\s*([A-Z])', e)
        if me:
            tower = me.group(1)
        else:
            mm = re.search(r'टॉवर\s*([ऀ-ॿ]+)', t)
            tower = MLET.get(mm.group(1)) if mm else None

    elif proj == 'Rivali Park':
        for k, v in RIVALI.items():
            if k in t:
                tower = v
                break

    elif proj == 'Borivali Project (Lodha Altus)':
        mm = re.search(r'नविन\s*सदनिका\s*नं\.?\s*([ऀ-ॿ]+)', t)   # new (re-dev) flat wing
        if mm and mm.group(1) in MLET:
            tower = MLET[mm.group(1)]
        else:
            tower = wing_marathi(t) or wing_english(e)

    elif proj == 'The Grand Residences':
        mm = re.search(r'सेल\s*बिल्डिंग\s*(?:नं|क्र)\.?\s*(\d+)', t)
        tower = ('Bldg ' + mm.group(1)) if mm else None

    elif proj == 'Verve Elina':
        mm = re.search(r'सदनिका\s*नं:?\s*\d+\s*([ऀ-ॿ]+)', t)
        tower = MLET.get(mm.group(1)) if mm else None

    elif proj in ('Sumit KMR Param', 'Sumit Garden Grove', 'Girivar Avenue'):
        tower = None   # single building / no wing in description

    elif tower is None:
        tower = wing_marathi(t) or wing_english(e)

    # ---------------- ACTUAL AREA (sq.ft) ----------------
    # = max(deed's stated total RERA carpet, flat+balcony+additional).
    # Handles extras that sit INSIDE the deed total (Godrej लगत/उपयोगिता) as well
    # as a balcony stated OUTSIDE it (Sumit KMR), and the deed's own rounding.
    actual = None
    tot = _total_doc(t)
    comp = None
    if flat is not None:
        comp = flat + (balc or 0) + (add or 0)
    cands = [v for v in (tot, comp) if v is not None]
    if cands:
        actual = round(max(cands) * SQM2SQFT, 2)
        if flat is None:                      # derive a flat figure from the total
            flat = round(max(cands) - (balc or 0) - (add or 0), 2)

    return dict(flat=flat, balc=balc, add=add, actual=actual, tower=tower,
                remark='; '.join(remark))


def main():
    wb = openpyxl.load_workbook(SRC)
    ws = wb['Index II']
    headers = {c.value: c.column for c in ws[1]}   # header -> column index (1-based)

    col_proj = headers['Project Name- As per Our Database']
    col_mar = headers['Property Discription']
    col_eng = headers['Property Discription English']
    col_tower = headers['Tower Name']
    col_actual = headers['Actual Areas']
    col_flat = headers['Flat Area']
    col_balc = headers['Balcony area']
    col_add = headers['Additional Area']
    col_rem = headers['Area and Price Remarks']

    # annotate units in headers for clarity
    ws.cell(row=1, column=col_flat).value = 'Flat Area (Sqm)'
    ws.cell(row=1, column=col_balc).value = 'Balcony area (Sqm)'
    ws.cell(row=1, column=col_add).value = 'Additional Area (Sqm)'
    ws.cell(row=1, column=col_actual).value = 'Actual Areas (Sqft)'

    n_flat = n_act = n_tower = 0
    for r in range(2, ws.max_row + 1):
        proj = ws.cell(row=r, column=col_proj).value
        if proj is None:
            continue
        res = extract(proj, ws.cell(row=r, column=col_mar).value,
                      ws.cell(row=r, column=col_eng).value)
        ws.cell(row=r, column=col_tower).value = res['tower']
        ws.cell(row=r, column=col_flat).value = res['flat']
        ws.cell(row=r, column=col_balc).value = res['balc']
        ws.cell(row=r, column=col_add).value = res['add']
        ws.cell(row=r, column=col_actual).value = res['actual']
        if res['remark']:
            ws.cell(row=r, column=col_rem).value = res['remark']
        n_flat += res['flat'] is not None
        n_act += res['actual'] is not None
        n_tower += res['tower'] is not None

    wb.save(OUT)
    print("Saved:", OUT)
    print("Rows:", ws.max_row - 1)
    print("Flat Area filled:", n_flat)
    print("Actual Areas filled:", n_act)
    print("Tower filled:", n_tower)


if __name__ == '__main__':
    main()
