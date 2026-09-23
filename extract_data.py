# Extracts the summary tables from Seldi_CMS_2025.xlsx into assets/js/data.js
# for the static GitHub Pages dashboard. One entry per worksheet.
import json
import os
import openpyxl

SRC = r"C:\Users\Boris\Downloads\Seldi_CMS_2025.xlsx"
OUT = os.path.join(os.path.dirname(__file__), "assets", "js", "data.js")

YEARS = [2001, 2002, 2014, 2016, 2019, 2021, 2023, 2025]

# Per-sheet metadata cell coordinates (title / insight / footnotes)
META = {
    "1.Pressure": dict(
        key="pressure", page="pressure.html",
        label="Pressure",
        nav="Corruption Pressure",
        title_cell="E1", insight_cell="L1", note_cells=["L2"],
    ),
    "2.Involvm": dict(
        key="involvement", page="involvement.html",
        label="Involvement",
        nav="Involvement in Corruption",
        title_cell="E1", insight_cell="K1", note_cells=["K2"],
    ),
    "3.Acceptability-Tolerance": dict(
        key="acceptability", page="acceptability.html",
        label="Acceptability",
        nav="Acceptability / Tolerance",
        title_cell="E1", subtitle_cell="E2", insight_cell="M2", note_cells=["M3"],
    ),
    "4.Likelihood": dict(
        key="likelihood", page="likelihood.html",
        label="Likelihood",
        nav="Likelihood of Pressure",
        title_cell="D1", insight_cell="J1", note_cells=["D2", "J2"],
    ),
    "5.Susceptibility": dict(
        key="susceptibility", page="susceptibility.html",
        label="Susceptibility",
        nav="Susceptibility",
        title_cell="F1", insight_cell="L1", note_cells=["F2", "F3"],
    ),
    "6.Feasibility of policy resp": dict(
        key="feasibility", page="feasibility.html",
        label="Policy Response",
        nav="Feasibility of Policy Response",
        title_cell="F1", insight_cell=None, note_cells=[],
        # The sheet carries no insight cell; written to match the style of the
        # other five worksheets. Based on the 2025 column of this worksheet:
        # Western Balkan values run from Montenegro 27.0% to Albania 76.2%.
        insight_text=("In 2025, between a quarter and three quarters of the "
                      "public believes corruption cannot be substantially "
                      "reduced"),
    ),
}


def num(v):
    """Cell value -> float or None."""
    if isinstance(v, (int, float)):
        return round(float(v), 6)
    return None  # empty string / None / text


def extract():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    out = []
    for sheet_name, meta in META.items():
        ws = wb[sheet_name]

        # Locate the header row containing 'Country'
        hdr_row = hdr_country_col = None
        for row in ws.iter_rows(min_row=1, max_row=10):
            for c in row:
                if c.value == "Country":
                    hdr_row, hdr_country_col = c.row, c.column
                    break
            if hdr_row:
                break
        assert hdr_row, sheet_name

        # Year columns from the header row
        year_cols = {}
        for c in ws[hdr_row]:
            try:
                y = int(c.value)
            except (TypeError, ValueError):
                continue
            if y in YEARS:
                year_cols[y] = c.column
        assert sorted(year_cols) == YEARS, (sheet_name, sorted(year_cols))

        # Data rows: country cell holds a string. The SPSS frequency tables
        # further down reuse country-name / label cells in the same column, so
        # guard against them: labels from those tables are excluded, a repeated
        # name ends the scan, and every reported value must be a proportion
        # in [0, 1] (the tables below hold raw frequencies/percentages).
        STOP = {"country", "valid", "missing", "total", "cumulative percent"}
        countries, seen = [], set()
        for r in range(hdr_row + 1, ws.max_row + 1):
            name = ws.cell(row=r, column=hdr_country_col).value
            if not isinstance(name, str) or not name.strip():
                continue
            name = name.strip()
            if name.lower() in STOP or name in seen:
                break
            values = [num(ws.cell(row=r, column=col).value) for col in
                      (year_cols[y] for y in YEARS)]
            known = [v for v in values if v is not None]
            if not known or any(v < 0 or v > 1 for v in known):
                continue
            seen.add(name)
            countries.append({"country": name, "values": values})

        def txt(coord):
            if not coord:
                return None
            v = ws[coord].value
            return str(v).strip() if v not in (None, "") else None

        notes = [n for n in (txt(c) for c in meta["note_cells"]) if n]
        entry = {
            "key": meta["key"],
            "page": meta["page"],
            "nav": meta["nav"],
            "label": meta["label"],
            "sheet": sheet_name,
            "title": txt(meta["title_cell"]),
            "insight": (meta.get("insight_text")
                        or txt(meta.get("insight_cell"))),
            "notes": notes,
            "years": YEARS,
            "countries": countries,
        }
        out.append(entry)
        print(f"{sheet_name}: {len(countries)} countries x {len(YEARS)} years")

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("// Generated by extract_data.py from Seldi_CMS_2025.xlsx. Do not edit by hand.\n")
        f.write("window.SELDI = ")
        json.dump(out, f, ensure_ascii=False, indent=2)
        f.write(";\n")
    print("wrote", OUT)


if __name__ == "__main__":
    extract()
