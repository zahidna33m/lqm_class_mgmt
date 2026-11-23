#!/usr/bin/env python3
"""
Certificate generator — LQM Al-Falah 26

Two steps, with ONE manual print in between (PowerPoint's "Export as PDF" mangles the ChopinScript
font; Print → Save as PDF renders correctly):

  1. build   CSV + template  →  <output_dir>/<csv name>.pptx  (+ <csv name>_manifest.csv)
  2. (you)   open <csv name>.pptx in PowerPoint, Print → PDF → save as
             <output_dir>/<csv name>.pdf
  3. split   <csv name>.pdf  →  <output_dir>/pdf/{StudentID}_{Name}.pdf, one per student

Settings (template, CSV, output folder) live in make_certificates.txt next to this script. Each can be
overridden on the command line (--template, --csv, --output-dir). There are no built-in defaults:
if neither the config nor a flag gives a value, the script stops and says which one is missing.

One-time setup (needs Python 3; requirements-certificates.txt lists python-pptx and pypdf):
  python3 -m venv .venv-certs
  .venv-certs/bin/pip install -r requirements-certificates.txt

Usage (run from the folder that holds this script; use the venv's python):
  .venv-certs/bin/python make_certificates.py build                      # uses make_certificates.txt
  .venv-certs/bin/python make_certificates.py build --csv other.csv      # override one setting
  .venv-certs/bin/python make_certificates.py split --csv other.csv      # same --csv as the build

Options
  build  --csv FILE          student CSV            (overrides "csv" in the config file)
         --template FILE     template .pptx         (overrides "template")
         --output-dir DIR    output folder          (overrides "output_dir")
         --limit N           only the first N students, for a quick test
         --urdu-to-arabic    replace Urdu letters (ٹ ھ ں …) with the standard Arabic letters so they
                             join in the template font. CHANGES the spelling on the certificate.
  split  --csv FILE          the CSV used for the build; it decides the file names (config "csv" if omitted)
         --output-dir DIR    output folder          (overrides "output_dir")
         --pdf FILE          the printed PDF (default: <output_dir>/<csv name>.pdf)
         --overwrite         replace per-student PDFs that already exist

Config file make_certificates.txt (same folder as this script; plain text, open in TextEdit):
  template = <the .pptx template>        csv = <the student CSV>        output_dir = <output folder>
  Format "name = value", no quotes, "#" starts a comment. Relative paths are relative to the config
  file's folder; paths given on the command line are relative to where you run the command.

Checklist for a new batch (e.g. Book 2)
  1. In the Progress sheet, set column Z = "Y" for each student who gets a certificate, then
     LQM Progress → Export Certificate CSV, and save the CSV next to this script.
  2. Put the template .pptx there too. If it changed, check the markers and the Brother/Sister layouts.
  3. Edit make_certificates.txt: set "template" and "csv" (or pass --csv / --template instead).
  4. build. Read the output: the NOTE lines (names shrunk to fit, very long names), the
     "Found non-Arabic letter(s)" lines, and the SUMMARY counts. Fix the sheet and re-export if needed.
  5. Open <csv name>.pptx in PowerPoint → File → Print → PDF → save as <output_dir>/<csv name>.pdf.
     Do NOT use Export as PDF: it garbles the ChopinScript font.
  6. split. It refuses to run if the PDF's page count differs from the manifest. Results are in
     <output_dir>/pdf/{StudentID}_{Name}.pdf.

Output order: all male certificates first, then all female (the CSV order is kept within each group).
The manifest and the split follow the same order. Names are trimmed of leading/trailing spaces and
invisible characters. A name wider than its box is shrunk (down to 65%) and reported as a NOTE.
The Arabic width is only estimated (the template's Arabic font may not be installed), so check the
flagged certificates on the printed PDF.

Template: the .pptx has one slide per gender. Which slide is male / female is detected from the
slide's layout text ("Brother" / "Sister"). The name boxes hold the markers `name_en` and `name_ar`
(no brackets). Each student gets a copy of their gender's slide with the markers replaced.

CSV columns (header row required): Student ID, Gender, Student Name, Arabic Name.
Read as UTF-8 (a byte-order mark is fine).
"""

import argparse
import copy
import csv
import math
import re
import sys
import unicodedata
from pathlib import Path

from PIL import ImageFont
from pptx import Presentation
from pptx.util import Pt

CONFIG_FILE = Path(__file__).with_suffix('.txt')   # make_certificates.txt, next to this script


# ─── Settings (config file + command-line overrides) ───────────────────────────

def load_config():
    """
    Reads make_certificates.txt: one "name = value" per line, "#" for comments, no quotes needed.
    Quotes around a value (straight or the curly ones TextEdit may insert) are tolerated and removed.
    Paths in it are resolved against the config file's folder.
    """
    if not CONFIG_FILE.exists():
        return {}
    allowed = {'template', 'csv', 'output_dir'}
    config = {}
    for n, line in enumerate(CONFIG_FILE.read_text(encoding='utf-8-sig').splitlines(), start=1):
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        if '=' not in line:
            sys.exit(f'{CONFIG_FILE.name} line {n}: expected "name = value", got: {line}')
        key, value = (x.strip() for x in line.split('=', 1))
        value = value.strip('"\'\u201c\u201d\u2018\u2019').strip()
        if key not in allowed:
            sys.exit(f'{CONFIG_FILE.name} line {n}: unknown setting "{key}". Allowed: {", ".join(sorted(allowed))}.')
        if value:
            config[key] = CONFIG_FILE.parent / value
    return config


def setting(args, name, config):
    """Command-line value wins over the config file; stops with a clear message if neither has it."""
    flag = getattr(args, name, None)
    if flag:
        return Path(flag)
    if name in config:
        return config[name]
    sys.exit(f'No {name.replace("_", " ")} given. Set "{name}" in {CONFIG_FILE.name} '
             f'or pass --{name.replace("_", "-")}.')


class Paths:
    """Everything the script reads/writes inside the output folder."""
    def __init__(self, out_dir, csv_path):
        stem = Path(csv_path).stem                       # outputs are named after the input CSV
        self.out_dir       = out_dir
        self.combined_pptx = out_dir / f'{stem}.pptx'
        self.combined_pdf  = out_dir / f'{stem}.pdf'
        self.manifest      = out_dir / f'{stem}_manifest.csv'
        self.pdf_dir       = out_dir / 'pdf'


TOKEN_EN = 'name_en'
TOKEN_AR = 'name_ar'
REQUIRED = ['Student ID', 'Gender', 'Student Name', 'Arabic Name']

# Text fit. A name wider than its box (minus the box's left/right padding) is shrunk, down to MIN_SCALE.
#  - English: MEASURED with the real Lucida Calligraphy font file that ships with PowerPoint.
#  - Arabic: Andalus is not installed on this Mac, so this is only an estimate (avg glyph width / font size).
DEFAULT_SIZE_PT = {TOKEN_EN: 18.0, TOKEN_AR: 28.0}   # used when the run has no explicit size
AVG_CHAR_EM     = {TOKEN_EN: 0.70, TOKEN_AR: 0.42}   # estimate, used when the font file is not found
MIN_SCALE       = 0.65
DEFAULT_INSET_EMU = 91440                            # PowerPoint's default left/right text padding (0.1 in)
ENGLISH_FONT_FILE = Path('/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts/LucidaCalligraphyItalic.ttf')

# Urdu/Persian letters → the standard Arabic letter the template's Arabic font has. Optional (--urdu-to-arabic):
# without it, letters outside the standard block fall back to another font and may not join to their neighbours.
URDU_TO_ARABIC = {
    '\u0679': '\u062A',  # ٹ → ت
    '\u0688': '\u062F',  # ڈ → د
    '\u0691': '\u0631',  # ڑ → ر
    '\u06BA': '\u0646',  # ں → ن
    '\u06BE': '\u0647',  # ھ → ه
    '\u06C1': '\u0647',  # ہ → ه
    '\u06C3': '\u0629',  # ۃ → ة
    '\u06CC': '\u064A',  # ی → ي
    '\u06D2': '\u064A',  # ے → ي
    '\u06A9': '\u0643',  # ک → ك
}

# Letters of the standard Arabic block (plus diacritics) the template's Arabic font is expected to have.
ARABIC_OK = set(chr(c) for c in range(0x0621, 0x064B)) | set(chr(c) for c in range(0x064B, 0x0653)) | {' ', 'ـ'}


# ─── CSV ───────────────────────────────────────────────────────────────────────

# Invisible characters that str.split() does not treat as spaces (zero-width, bidi marks, BOM).
INVISIBLE = dict.fromkeys(map(ord, '\u200b\u200c\u200d\u200e\u200f\u202a\u202b\u202c\u202d\u202e\u2066\u2067\u2068\u2069\ufeff\u2060'))


def clean_name(text):
    """Removes invisible characters, trims leading/trailing spaces (incl. non-breaking), collapses inner runs."""
    return ' '.join((text or '').translate(INVISIBLE).split())


def read_students(csv_path):
    """Returns (students, problems). Problems are fatal; the caller stops if there are any."""
    problems = []
    with open(csv_path, encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        missing = [c for c in REQUIRED if c not in (reader.fieldnames or [])]
        if missing:
            sys.exit(f'CSV is missing column(s): {missing}. Found: {reader.fieldnames}')
        students, seen = [], {}
        for n, row in enumerate(reader, start=2):
            sid    = (row['Student ID'] or '').strip()
            gender = (row['Gender'] or '').strip().lower()
            en     = clean_name(row['Student Name'])
            ar     = clean_name(row['Arabic Name'])
            where  = f'CSV line {n} (ID {sid or "?"})'
            if not sid:                      problems.append(f'{where}: Student ID is blank.')
            if sid in seen:                  problems.append(f'{where}: duplicate Student ID (also line {seen[sid]}).')
            if gender not in ('male', 'female'):
                problems.append(f'{where}: Gender "{row["Gender"]}" is not Male or Female.')
            if not en:                       problems.append(f'{where}: Student Name is blank.')
            if not ar:                       problems.append(f'{where}: Arabic Name is blank.')
            seen[sid] = n
            students.append({'id': sid, 'gender': gender, 'en': en, 'ar': ar, 'line': n})
    return students, problems


def arabic_warnings(students):
    """Names containing letters outside the standard Arabic block (e.g. Urdu ٹ ھ). Returns the affected students."""
    return [s for s in students if any(c not in ARABIC_OK for c in s['ar'])]


# ─── Template handling ─────────────────────────────────────────────────────────

def layout_text(layout):
    return ' '.join(sh.text_frame.text for sh in layout.shapes if sh.has_text_frame)


def find_gender_slides(prs):
    """Maps 'male' / 'female' to the template slide whose layout says Brother / Sister."""
    found = {}
    for slide in prs.slides:
        text = layout_text(slide.slide_layout)
        if 'Brother' in text:  found.setdefault('male', slide)
        if 'Sister' in text:   found.setdefault('female', slide)
    if set(found) != {'male', 'female'}:
        sys.exit('Could not find both a Brother (male) and a Sister (female) slide in the template.')
    return found


def check_markers(slide, label):
    for token in (TOKEN_EN, TOKEN_AR):
        hits = sum(1 for sh in slide.shapes if sh.has_text_frame
                   for p in sh.text_frame.paragraphs if token in ''.join(r.text for r in p.runs))
        if hits != 1:
            sys.exit(f'Template {label} slide: marker "{token}" found {hits} time(s); expected exactly 1.')


def fill_marker(slide, token, value, notes, who):
    """Replaces the marker (which may be split across runs) and shrinks the text if it likely won't fit."""
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            runs = p.runs
            full = ''.join(r.text for r in runs)
            if token not in full:
                continue
            first = runs[0]
            first.text = full.replace(token, value)
            for r in runs[1:]:
                r._r.getparent().remove(r._r)
            rpr = first._r.find('{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
            if rpr is not None:
                rpr.attrib.pop('err', None)      # drop the "spelling error" flag the placeholder carried

            base = first.font.size.pt if first.font.size else DEFAULT_SIZE_PT[token]
            usable = (sh.width - inset_emu(sh, 'lIns') - inset_emu(sh, 'rIns')) / 12700
            need = text_width_pt(token, value, base)
            if need > usable:
                scale = max(MIN_SCALE, usable / need)
                new = math.floor(base * scale * 2) / 2
                first.font.size = Pt(new)
                notes.append(f'{who}: {token} "{value}" needs {need:.0f}pt but the box allows {usable:.0f}pt — '
                             f'shrunk {base:g}pt → {new:g}pt.')
                if usable / need < MIN_SCALE:
                    notes.append(f'{who}: {token} "{value}" is very long even at {new:g}pt — CHECK this certificate.')
            return
    raise RuntimeError(f'marker {token} not found on slide')


def inset_emu(shape, attr):
    """Left/right text padding of a text box (PowerPoint's default 0.1 in unless the box sets it)."""
    body = shape._element.txBody.find('{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
    v = body.get(attr) if body is not None else None
    return int(v) if v is not None else DEFAULT_INSET_EMU


_english_font = None
def text_width_pt(token, value, size_pt):
    """Width of `value` at `size_pt`. Measured for English when the font file exists; estimated otherwise."""
    global _english_font
    if token == TOKEN_EN and ENGLISH_FONT_FILE.exists():
        if _english_font is None:
            _english_font = ImageFont.truetype(str(ENGLISH_FONT_FILE), 200)
        return _english_font.getlength(value) * size_pt / 200
    return len(value) * size_pt * AVG_CHAR_EM[token]


def delete_slide(prs, slide):
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        if prs.part.related_part(sld_id.rId) is slide.part:
            prs.part.drop_rel(sld_id.rId)
            sld_id_lst.remove(sld_id)
            return


# ─── build ─────────────────────────────────────────────────────────────────────

def cmd_build(args):
    config   = load_config()
    csv_path = setting(args, 'csv', config)
    template = setting(args, 'template', config)
    paths    = Paths(setting(args, 'output_dir', config), csv_path)
    for label, p in (('CSV', csv_path), ('Template', template)):
        if not p.exists():
            sys.exit(f'{label} file not found: {p}')

    students, problems = read_students(csv_path)
    if problems:
        print('PROBLEMS in the CSV — fix these first:')
        for p in problems: print('  -', p)
        sys.exit(1)
    if not students:
        sys.exit('The CSV has no students.')
    if args.limit:
        students = students[:args.limit]
    students.sort(key=lambda s: s['gender'] != 'male')   # males first, then females; CSV order kept within each

    mapping_notes = []
    if args.urdu_to_arabic:
        for s in students:
            mapped = ''.join(URDU_TO_ARABIC.get(c, c) for c in s['ar'])
            if mapped != s['ar']:
                mapping_notes.append(f'{s["id"]} {s["en"]}: Arabic name changed {s["ar"]} → {mapped} (--urdu-to-arabic).')
                s['ar'] = mapped

    prs = Presentation(str(template))
    templates = find_gender_slides(prs)
    for g, s in templates.items():
        check_markers(s, g)

    notes = []
    for s in students:
        src = templates[s['gender']]
        new = prs.slides.add_slide(src.slide_layout)
        for el in list(new.shapes._spTree):          # drop anything add_slide put there
            if el.tag.endswith('}sp') or el.tag.endswith('}pic'):
                new.shapes._spTree.remove(el)
        for sh in src.shapes:
            new.shapes._spTree.append(copy.deepcopy(sh._element))
        who = f'{s["id"]} {s["en"]}'
        fill_marker(new, TOKEN_EN, s['en'], notes, who)
        fill_marker(new, TOKEN_AR, s['ar'], notes, who)

    for src in templates.values():
        delete_slide(prs, src)

    paths.out_dir.mkdir(parents=True, exist_ok=True)
    prs.save(str(paths.combined_pptx))
    with open(paths.manifest, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['Page', 'Student ID', 'Gender', 'Student Name', 'Arabic Name', 'File'])
        for i, s in enumerate(students, start=1):
            w.writerow([i, s['id'], s['gender'].capitalize(), s['en'], s['ar'], pdf_name(s['id'], s['en'])])

    print(f'Template: {template}\nStudents: {csv_path}')
    print(f'Built {len(students)} certificate(s): {paths.combined_pptx}')
    print(f'Manifest: {paths.manifest}')
    odd = arabic_warnings(students)
    for n in mapping_notes + notes:
        print('  NOTE:', n)
    for s in odd:
        letters = ','.join(dict.fromkeys(c for c in s['ar'] if c not in ARABIC_OK))   # each distinct letter once, in order
        print(f'  Found non-Arabic letter(s) {letters} in {s["ar"]}')
    males = sum(1 for s in students if s['gender'] == 'male')
    print('\nSUMMARY')
    print(f'  Total rows:                       {len(students)}')
    print(f'  Male rows:                        {males}')
    print(f'  Female rows:                      {len(students) - males}')
    print(f'  Names with non-Arabic letters:    {len(odd)}')
    print(f'\nNEXT: open {paths.combined_pptx} in PowerPoint → File → Print → PDF → Save as PDF →\n'
          f'      save it as {paths.combined_pdf}, then run:  make_certificates.py split')


# ─── split ─────────────────────────────────────────────────────────────────────

def pdf_name(sid, name):
    safe = re.sub(r'[\\/:*?"<>|]+', '', name).strip()
    safe = re.sub(r'\s+', ' ', safe)
    return f'{sid}_{safe}.pdf'


def cmd_split(args):
    from pypdf import PdfReader, PdfWriter
    config = load_config()
    paths = Paths(setting(args, 'output_dir', config), setting(args, 'csv', config))
    if not paths.manifest.exists():
        sys.exit(f'{paths.manifest} not found — run "build" first (same --output-dir).')
    pdf_path = Path(args.pdf) if args.pdf else paths.combined_pdf
    if not pdf_path.exists():
        sys.exit(f'{pdf_path} not found. Print the deck to PDF from PowerPoint and save it there.')

    with open(paths.manifest, encoding='utf-8-sig', newline='') as f:
        rows = list(csv.DictReader(f))
    reader = PdfReader(str(pdf_path))
    if len(reader.pages) != len(rows):
        sys.exit(f'Page count mismatch: the PDF has {len(reader.pages)} page(s) but the manifest lists '
                 f'{len(rows)} student(s). Nothing written. (Were all slides printed, one per page?)')

    paths.pdf_dir.mkdir(parents=True, exist_ok=True)
    clash = [r['File'] for r in rows if (paths.pdf_dir / r['File']).exists()]
    if clash and not args.overwrite:
        sys.exit(f'{len(clash)} file(s) already exist in {paths.pdf_dir} (e.g. {clash[0]}). '
                 f'Use --overwrite to replace them.')

    for page, r in zip(reader.pages, rows):
        w = PdfWriter()
        w.add_page(page)
        with open(paths.pdf_dir / r['File'], 'wb') as out:
            w.write(out)
    print(f'Wrote {len(rows)} PDF(s) to {paths.pdf_dir}/')


# ─── main ──────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest='cmd', required=True)

    b = sub.add_parser('build', help='create the combined .pptx and manifest from the CSV')
    b.add_argument('--csv', help='student CSV (overrides "csv" in make_certificates.txt)')
    b.add_argument('--template', help='template .pptx (overrides "template" in make_certificates.txt)')
    b.add_argument('--output-dir', dest='output_dir', help='output folder (overrides "output_dir" in make_certificates.txt)')
    b.add_argument('--limit', type=int, help='only the first N students (for a quick test)')
    b.add_argument('--urdu-to-arabic', action='store_true',
                   help='replace Urdu letters (ٹ ھ ں …) in Arabic names with the standard Arabic letters, '
                        'so they join properly in the template font. CHANGES the spelling on the certificate.')
    b.set_defaults(fn=cmd_build)

    s = sub.add_parser('split', help='split the printed PDF into one file per student')
    s.add_argument('--pdf', help='the printed PDF (default: <output_dir>/<csv name>.pdf)')
    s.add_argument('--csv', help='the CSV used for build (decides the file names; overrides "csv" in make_certificates.txt)')
    s.add_argument('--output-dir', dest='output_dir', help='output folder (overrides "output_dir" in make_certificates.txt)')
    s.add_argument('--overwrite', action='store_true')
    s.set_defaults(fn=cmd_split)

    args = ap.parse_args()
    args.fn(args)


if __name__ == '__main__':
    main()
