#!/usr/bin/env python3
"""Build publications.bib, docs/index.html, update Colonius.tex nocites,
from publications.xlsx (sheet: Publications).
"""

import os
import re
from collections import defaultdict
from datetime import datetime

import openpyxl

# ── Paths ─────────────────────────────────────────────────────────────────────

XLSX_FILE   = "publications.xlsx"
SHEET_NAME  = "Publications"
TEX_FILE    = "Colonius.tex"
BIB_ROOT    = "colonius.bib"          # used by Colonius.tex
OUT_DIR     = "docs"
BIB_DOCS    = os.path.join(OUT_DIR, "publications.bib")   # served by GitHub Pages
HTML_OUT    = os.path.join(OUT_DIR, "index.html")

# ── BibTeX field mapping ───────────────────────────────────────────────────────
# Hardcoded from the Publications sheet keyword rows.
# Each entry is (bibtex_field_name, spreadsheet_column_name).
# 'year', 'author', 'title' are common to all types and handled separately.

BIBTEX_FIELDS: dict[str, list[tuple[str, str]]] = {
    "article": [
        ("journal",   "Publication Title"),
        ("publisher", "Publisher"),
        ("volume",    "Volume"),
        ("number",    "Issue/Number"),
        ("pages",     "Page Range"),
        ("url",       "Persistent URL"),
        ("doi",       "DOI"),
        ("issn",      "ISSN"),
    ],
    "inproceedings": [
        ("booktitle", "Publication Title"),
        ("publisher", "Publisher"),
        ("pages",     "Page Range"),
        ("url",       "Persistent URL"),
        ("doi",       "DOI"),
    ],
    "incollection": [
        ("booktitle", "Publication Title"),
        ("editor",    "Editor"),
        ("publisher", "Publisher"),
        ("pages",     "Page Range"),
    ],
    "misc": [
        ("howpublished", "Publication Title"),
        ("url",          "Persistent URL"),
        ("note",         "Submitted to"),
    ],
    "phdthesis": [
        ("school", "School"),
        ("url",    "Persistent URL"),
    ],
}

# ── .tex refsection → entry type mapping ──────────────────────────────────────

SECTION_ENTRY_TYPE: dict[str, str] = {
    "Submitted Articles":           "misc",
    "Journal Articles":             "article",
    "Conference Papers":            "inproceedings",
    "Book Chapters":                "incollection",
    "Doctoral students (as advisor)": "phdthesis",
}

# ── Display labels ────────────────────────────────────────────────────────────

ENTRY_TYPE_LABELS: dict[str, str] = {
    "article":       "Journal Article",
    "inproceedings": "Conference Paper",
    "phdthesis":     "PhD Thesis",
    "mastersthesis": "Master's Thesis",
    "misc":          "Preprint / Other",
    "book":          "Book",
    "incollection":  "Book Chapter",
    "techreport":    "Technical Report",
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def field(row: dict, key: str) -> str:
    """Return stripped field value, or empty string if missing/None."""
    v = row.get(key)
    if v is None:
        return ""
    # openpyxl may return datetime objects for date cells
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def parse_authors(raw: str) -> list[str]:
    # Strip stray LaTeX braces that may be left over in the spreadsheet data
    cleaned = raw.replace("{", "").replace("}", "")
    return [a.strip() for a in cleaned.split("/") if a.strip()]


def authors_to_bibtex(authors: list[str]) -> str:
    return " and ".join(authors)


def format_author_display(authors: list[str]) -> str:
    def flip(a):
        parts = [p.strip() for p in a.split(",", 1)]
        return f"{parts[1]} {parts[0]}" if len(parts) == 2 else a
    return ", ".join(flip(a) for a in authors)


# ── Diacritical → LaTeX mapping ───────────────────────────────────────────────
# Covers the most common accented/special characters in academic author names
# and titles.  Expand as needed.

DIACRITICALS: list[tuple[str, str]] = [
    # Double-width / composed characters first to avoid partial matches
    ("ä", r'{\"a}'), ("ë", r'{\"e}'), ("ï", r'{\"i}'),
    ("ö", r'{\"o}'), ("ü", r'{\"u}'), ("ÿ", r'{\"y}'),
    ("Ä", r'{\"A}'), ("Ë", r'{\"E}'), ("Ï", r'{\"I}'),
    ("Ö", r'{\"O}'), ("Ü", r'{\"U}'),
    ("á", r"{\'a}"), ("é", r"{\'e}"), ("í", r"{\'i}"),
    ("ó", r"{\'o}"), ("ú", r"{\'u}"), ("ý", r"{\'y}"),
    ("Á", r"{\'A}"), ("É", r"{\'E}"), ("Í", r"{\'I}"),
    ("Ó", r"{\'O}"), ("Ú", r"{\'U}"), ("Ý", r"{\'Y}"),
    ("à", r"{\`a}"), ("è", r"{\`e}"), ("ì", r"{\`i}"),
    ("ò", r"{\`o}"), ("ù", r"{\`u}"),
    ("À", r"{\`A}"), ("È", r"{\`E}"), ("Ì", r"{\`I}"),
    ("Ò", r"{\`O}"), ("Ù", r"{\`U}"),
    ("â", r"{\^a}"), ("ê", r"{\^e}"), ("î", r"{\^i}"),
    ("ô", r"{\^o}"), ("û", r"{\^u}"),
    ("Â", r"{\^A}"), ("Ê", r"{\^E}"), ("Î", r"{\^I}"),
    ("Ô", r"{\^O}"), ("Û", r"{\^U}"),
    ("ã", r"{\~a}"), ("ñ", r"{\~n}"), ("õ", r"{\~o}"),
    ("Ã", r"{\~A}"), ("Ñ", r"{\~N}"), ("Õ", r"{\~O}"),
    ("ç", r"{\c{c}}"), ("Ç", r"{\c{C}}"),
    ("ø", r"{\o}"),   ("Ø", r"{\O}"),
    ("å", r"{\aa}"),  ("Å", r"{\AA}"),
    ("æ", r"{\ae}"),  ("Æ", r"{\AE}"),
    ("œ", r"{\oe}"),  ("Œ", r"{\OE}"),
    ("ß", r"{\ss}"),
    ("š", r"{\v{s}}"), ("Š", r"{\v{S}}"),
    ("č", r"{\v{c}}"), ("Č", r"{\v{C}}"),
    ("ž", r"{\v{z}}"), ("Ž", r"{\v{Z}}"),
    ("ř", r"{\v{r}}"), ("Ř", r"{\v{R}}"),
]


def to_latex_chars(text: str) -> str:
    """Replace Unicode diacriticals with LaTeX escape sequences."""
    for char, latex in DIACRITICALS:
        text = text.replace(char, latex)
    return text


def protect_caps(title: str) -> str:
    """Protect capitalised words in a BibTeX title string.

    Assumes the spreadsheet uses sentence case: only the first word and genuine
    proper nouns / acronyms carry capitals.  Any word (after the very first)
    that starts with a capital, or contains an uppercase letter after position 0,
    gets its leading capital wrapped in {X} so BibTeX styles cannot downcase it.

    Subtitles after ':' are treated as a fresh sentence (their first word is
    also left unprotected).
    """
    # Split on whitespace, preserving tokens
    tokens = re.split(r'(\s+)', title)
    first_word = True        # first word of title
    after_colon = False      # first word of a subtitle

    result = []
    for tok in tokens:
        if re.match(r'\s+', tok):
            result.append(tok)
            continue

        # Detect subtitle boundary: token ends with ':'  or is ':'
        is_sentence_start = first_word or after_colon
        after_colon = tok.rstrip().endswith(':')
        first_word = False

        if is_sentence_start:
            # Leave the first letter of each sentence unprotected
            result.append(tok)
            continue

        # Protect any word that has at least one uppercase letter
        def protect_letter(m):
            ch = m.group(0)
            return '{' + ch + '}' if ch.isupper() else ch

        protected = re.sub(r'[A-Za-z]', protect_letter, tok)
        result.append(protected)

    return "".join(result)


# ── Load spreadsheet ──────────────────────────────────────────────────────────

def load_xlsx(path: str, sheet: str) -> list[dict]:
    """Load publication rows, skipping all preamble rows.

    Preamble rows are identified by their first cell being:
      - 'CV builder'
      - starting with '~'  (the \nocite summary rows)
      - a BibTeX entry type keyword ('article', 'inproceedings', etc.)
        whose second cell is a BibTeX field name like 'key'
    The real header row has 'Entry type' in the first cell.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    rows_iter = ws.iter_rows(values_only=True)

    headers = None
    for raw_row in rows_iter:
        first = str(raw_row[0]).strip() if raw_row[0] is not None else ""
        if first == "Entry type":
            headers = [str(h).strip() if h is not None else "" for h in raw_row]
            break

    if headers is None:
        raise ValueError(f"Could not find 'Entry type' header row in sheet '{sheet}'")

    rows = []
    for raw_row in rows_iter:
        record = {
            headers[i]: (str(v).strip() if v is not None and not hasattr(v, "strftime")
                         else (v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else ""))
            for i, v in enumerate(raw_row)
        }
        if any(record.values()):
            rows.append(record)
    wb.close()
    return rows


# ── BibTeX generation ─────────────────────────────────────────────────────────

def make_bibtex_entry(row: dict) -> str:
    etype = field(row, "Entry type").lower()
    tag   = field(row, "Tag")
    if not tag:
        return ""

    authors = parse_authors(field(row, "Author"))

    try:
        year_val = str(int(float(field(row, "Year"))))
    except ValueError:
        year_val = field(row, "Year")

    fields_out: list[tuple[str, str]] = [
        ("year",   year_val),
        ("author", to_latex_chars(authors_to_bibtex(authors))),
        ("title",  "{" + protect_caps(to_latex_chars(field(row, "Title").replace("{", "").replace("}", ""))) + "}"),
    ]

    for bib_key, col_name in BIBTEX_FIELDS.get(etype, []):
        val = field(row, col_name)
        if val:
            if bib_key in ("journal", "booktitle", "howpublished", "school",
                           "publisher", "editor"):
                val = "{" + val + "}"
            fields_out.append((bib_key, val))

    lines = [f"@{etype}{{{tag},"]
    for k, v in fields_out:
        lines.append(f"  {k} = {{{v}}},")
    lines[-1] = lines[-1].rstrip(",")
    lines.append("}")
    return "\n".join(lines)


# ── Sort key ──────────────────────────────────────────────────────────────────

def entry_sort_key(row: dict):
    try:
        year = int(float(field(row, "Year")))
    except ValueError:
        year = 0
    first_author = field(row, "Author").split("/")[0].split(",")[0].lower()
    return (-year, first_author)


# ── HTML generation ───────────────────────────────────────────────────────────

def build_html(pubs_by_year: dict, date_str: str) -> str:
    rows_html = []
    for year in sorted(pubs_by_year.keys(), reverse=True):
        rows_html.append(
            f'<tr class="year-header"><td colspan="2"><h2>{year}</h2></td></tr>'
        )
        for row in pubs_by_year[year]:
            authors  = parse_authors(field(row, "Author"))
            author_str = format_author_display(authors)
            title  = field(row, "Title")
            etype  = field(row, "Entry type").lower()
            type_label = ENTRY_TYPE_LABELS.get(etype, etype)

            venue_parts = []
            pub       = field(row, "Publication Title")
            publisher = field(row, "Publisher")
            school    = field(row, "School")
            if pub:
                venue_parts.append(f"<em>{pub}</em>")
            elif publisher and etype == "inproceedings":
                venue_parts.append(f"<em>{publisher}</em>")
            elif school:
                venue_parts.append(f"<em>{school}</em>")
            if field(row, "Volume"):
                venue_parts.append(f"vol.&nbsp;{field(row, 'Volume')}")
            if field(row, "Page Range"):
                venue_parts.append(field(row, "Page Range"))
            venue_str = ", ".join(venue_parts)

            doi = field(row, "DOI")
            url = field(row, "Persistent URL")
            links = []
            if doi:
                links.append(f'<a href="https://doi.org/{doi}" target="_blank">DOI</a>')
            if url:
                links.append(f'<a href="{url}" target="_blank">Link</a>')
            link_str = " &nbsp;|&nbsp; ".join(links)

            tag = field(row, "Tag")
            rows_html.append(f"""
      <tr class="pub-row" id="{tag}">
        <td class="pub-type"><span class="badge badge-{etype}">{type_label}</span></td>
        <td class="pub-detail">
          <div class="pub-title">{title}</div>
          <div class="pub-authors">{author_str}</div>
          {"<div class='pub-venue'>" + venue_str + "</div>" if venue_str else ""}
          {"<div class='pub-links'>" + link_str + "</div>" if link_str else ""}
        </td>
      </tr>""")

    rows_joined = "\n".join(rows_html)
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Colonius Group Publications</title>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      max-width: 900px; margin: 0 auto; padding: 2rem 1rem;
      color: #222; background: #fafafa;
    }}
    h1 {{ font-size: 1.8rem; margin-bottom: 0.25rem; }}
    .subtitle {{ color: #666; margin-bottom: 2rem; font-size: 0.95rem; }}
    .toolbar {{ margin-bottom: 1.5rem; display: flex; gap: 1rem; flex-wrap: wrap; align-items: center; }}
    .toolbar input {{
      padding: 0.4rem 0.7rem; border: 1px solid #ccc; border-radius: 4px;
      font-size: 0.9rem; flex: 1; min-width: 200px;
    }}
    .toolbar a.dl-btn {{
      padding: 0.4rem 1rem; background: #0066cc; color: #fff;
      text-decoration: none; border-radius: 4px; font-size: 0.9rem; white-space: nowrap;
    }}
    .toolbar a.dl-btn:hover {{ background: #0055aa; }}
    table {{ width: 100%; border-collapse: collapse; }}
    tr.year-header td {{ padding: 1.5rem 0 0.25rem; }}
    tr.year-header h2 {{ margin: 0; font-size: 1.3rem; border-bottom: 2px solid #0066cc; padding-bottom: 4px; }}
    tr.pub-row {{ border-bottom: 1px solid #e8e8e8; }}
    td.pub-type {{ width: 140px; vertical-align: top; padding: 0.75rem 0.5rem 0.75rem 0; }}
    td.pub-detail {{ padding: 0.75rem 0; vertical-align: top; }}
    .pub-title {{ font-weight: 600; margin-bottom: 0.2rem; }}
    .pub-authors {{ color: #444; font-size: 0.9rem; margin-bottom: 0.2rem; }}
    .pub-venue {{ color: #555; font-size: 0.9rem; margin-bottom: 0.2rem; }}
    .pub-links {{ font-size: 0.85rem; margin-top: 0.3rem; }}
    .pub-links a {{ color: #0066cc; text-decoration: none; }}
    .pub-links a:hover {{ text-decoration: underline; }}
    .badge {{
      display: inline-block; padding: 2px 7px; border-radius: 3px;
      font-size: 0.75rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.03em;
    }}
    .badge-article       {{ background: #dbeafe; color: #1e40af; }}
    .badge-inproceedings {{ background: #dcfce7; color: #166534; }}
    .badge-phdthesis     {{ background: #fef9c3; color: #854d0e; }}
    .badge-mastersthesis {{ background: #fef9c3; color: #854d0e; }}
    .badge-misc          {{ background: #f3e8ff; color: #6b21a8; }}
    .badge-book          {{ background: #ffe4e6; color: #9f1239; }}
    .badge-incollection  {{ background: #ffedd5; color: #9a3412; }}
    tr.hidden {{ display: none; }}
    .updated {{ color: #999; font-size: 0.8rem; margin-top: 2rem; }}
  </style>
</head>
<body>
  <h1>Colonius Group Publications</h1>
  <p class="subtitle">California Institute of Technology &nbsp;&middot;&nbsp; Mechanical and Civil Engineering</p>
  <div class="toolbar">
    <input type="text" id="search" placeholder="Filter by title, author, journal…" oninput="filterPubs()" />
    <a class="dl-btn" href="publications.bib" download>Download .bib</a>
    <a class="dl-btn" href="cv.pdf" target="_blank">CV (PDF)</a>
  </div>
  <table id="pub-table">
    <tbody>
{rows_joined}
    </tbody>
  </table>
  <p class="updated">Last updated: {date_str}</p>
  <script>
    function filterPubs() {{
      const q = document.getElementById('search').value.toLowerCase();
      document.querySelectorAll('tr.pub-row').forEach(tr => {{
        tr.classList.toggle('hidden', q && !tr.textContent.toLowerCase().includes(q));
      }});
      document.querySelectorAll('tr.year-header').forEach(tr => {{
        let next = tr.nextElementSibling;
        let anyVisible = false;
        while (next && next.classList.contains('pub-row')) {{
          if (!next.classList.contains('hidden')) anyVisible = true;
          next = next.nextElementSibling;
        }}
        tr.classList.toggle('hidden', !anyVisible);
      }});
    }}
  </script>
</body>
</html>
"""


# ── Update Colonius.tex nocite commands ───────────────────────────────────────

def update_tex_nocites(tex_path: str, nocites: dict[str, str], date_str: str):
    """Replace the active \\nocite{...} line in each refsection of tex_path.

    nocites maps section title → comma-separated cite keys.
    Inserts a % AUTO-GENERATED <date> comment on the line before the new nocite.
    If the previous run left an AUTO-GENERATED comment, it is replaced too.
    """
    with open(tex_path, encoding="utf-8") as f:
        lines = f.readlines()

    in_refsection   = False
    section_title   = None
    nocite_idx      = None   # index of the active \nocite line
    auto_gen_idx    = None   # index of a preceding AUTO-GENERATED comment

    replacements: list[tuple[int, int, str]] = []  # (start, end_exclusive, new_text)

    for i, line in enumerate(lines):
        stripped = line.strip()

        if r"\begin{refsection}" in line:
            in_refsection = True
            section_title = None
            nocite_idx    = None
            auto_gen_idx  = None

        elif r"\end{refsection}" in line:
            if section_title and nocite_idx is not None and section_title in nocites:
                keys     = nocites[section_title]
                new_text = (f"% AUTO-GENERATED {date_str}\n"
                            f"\\nocite{{{keys}}}\n")
                start = auto_gen_idx if auto_gen_idx is not None else nocite_idx
                replacements.append((start, nocite_idx + 1, new_text))
            in_refsection = True   # stays True across \hrule etc — reset properly
            in_refsection = False
            section_title = None
            nocite_idx    = None
            auto_gen_idx  = None

        elif in_refsection:
            m = re.search(r'\\printbibliography\[.*?title=\{([^}]+)\}', line)
            if m:
                section_title = m.group(1)

            # Track AUTO-GENERATED comment immediately before a \nocite
            if stripped.startswith("% AUTO-GENERATED"):
                auto_gen_idx = i
            elif stripped.startswith(r"\nocite{") and nocite_idx is None:
                # Only use auto_gen_idx if it was the immediately preceding non-blank line
                if auto_gen_idx is not None:
                    # Check there are no non-blank, non-comment lines between them
                    between = [l.strip() for l in lines[auto_gen_idx + 1 : i]
                               if l.strip() and not l.strip().startswith("%")]
                    if between:
                        auto_gen_idx = None
                nocite_idx = i
            elif stripped and not stripped.startswith("%"):
                # A real content line resets the auto_gen tracker
                if nocite_idx is None:
                    auto_gen_idx = None

    # Apply replacements in reverse order to preserve line indices
    for start, end, new_text in sorted(replacements, reverse=True):
        lines[start:end] = [new_text]

    with open(tex_path, "w", encoding="utf-8") as f:
        f.writelines(lines)

    print(f"Updated {len(replacements)} \\nocite commands in {tex_path}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    date_str = datetime.now().strftime("%Y-%m-%d")

    rows = load_xlsx(XLSX_FILE, SHEET_NAME)
    rows.sort(key=entry_sort_key)
    print(f"Loaded {len(rows)} entries from {XLSX_FILE}")

    # ── BibTeX ────────────────────────────────────────────────────────────────
    bib_entries = [e for r in rows if (e := make_bibtex_entry(r))]
    bib_content = "\n\n".join(bib_entries) + "\n"
    for bib_path in (BIB_ROOT, BIB_DOCS):
        with open(bib_path, "w", encoding="utf-8") as f:
            f.write(bib_content)
    print(f"Wrote {len(bib_entries)} entries to {BIB_ROOT} and {BIB_DOCS}")

    # ── HTML ──────────────────────────────────────────────────────────────────
    pubs_by_year: dict[int, list] = defaultdict(list)
    for r in rows:
        try:
            year = int(float(field(r, "Year")))
        except ValueError:
            year = 0
        pubs_by_year[year].append(r)

    html = build_html(pubs_by_year, date_str)
    with open(HTML_OUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Wrote {HTML_OUT}")

    # ── Colonius.tex nocites ──────────────────────────────────────────────────
    # Build per-section key lists (year-descending order preserved by sort above)
    nocites: dict[str, str] = {}
    for section_title, etype in SECTION_ENTRY_TYPE.items():
        keys = [field(r, "Tag") for r in rows
                if field(r, "Entry type").lower() == etype and field(r, "Tag")]
        if keys:
            nocites[section_title] = ",".join(keys)

    update_tex_nocites(TEX_FILE, nocites, date_str)


if __name__ == "__main__":
    main()
