import os
import re
import xlwings as xw

# --- Paths ---
folder_path = r"C:\Users\TGonzalezGomez\Documents\Extracted_RecipesForAmericanADmiralClub"
template_path = r"C:\Users\TGonzalezGomez\OneDrive - SODEXO\Documents\Master Recipe Development Templates delted.xlsx"
output_path = "SodynamicAmericanSummerAdmiral Club.xlsx"

EXPAND_PROCEDURE = True   
MAX_PROC_VISIBLE = 12     


HEADER_PATTERN = re.compile(
    r'^\s*(recipe\s*name|yield|procedure|ingredients?|chef\s*notes|chef(?:\s*created\s*by)?|created by|chef name)\s*:?\s*(.*)$',
    re.IGNORECASE
)

CAPTURE_BOUNDARY = re.compile(r'^\s*==+\s*(start|end)\s*capture', re.IGNORECASE)

def sanitize_sheet_name(name: str, existing: set) -> str:
    base = re.sub(r'[\\/*?:\[\]]', '', name)[:31] or "Sheet"
    cand = base
    i = 1
    while cand in existing:
        suffix = f"_{i}"
        cand = (base[:31 - len(suffix)] + suffix)
        i += 1
    existing.add(cand)
    return cand

def _strip_bullet(s: str) -> str:
    s = re.sub(r'^\s*[-•]\s*', '', s)                                 # bullets
    s = re.sub(r'^\s*(?:step\s*)?\d+[:.)]\s*', '', s, flags=re.IGNORECASE)  # numbering
    return s.strip()

def parse_recipe(text_lines):
    """
    Flexible parser for .txt like:
      Recipe Name:
      Yield:
      Chef Created By:
      Procedure:

    Each section continues until the next header or a '== Start/End Capture' line.
    """
    title = ""
    recipe_yield = ""
    chef_lines = []
    procedure_lines = []

    # Pass 1: grab any same-line values
    for raw in text_lines:
        line = raw.rstrip("\r\n")
        m = HEADER_PATTERN.match(line)
        if not m:
            continue
        key, rest = m.group(1).lower(), m.group(2).strip()
        if key == "recipe name" and rest and not title:
            title = rest
        elif key in ("chef", "chef created by", "created by", "chef name"):
            if rest:
                chef_lines.append(_strip_bullet(rest))
        elif key == "yield" and rest and not recipe_yield:
            recipe_yield = rest

    # Pass 2: block capture with smart stop
    section = None
    buf = []

    def flush(sec):
        nonlocal title, recipe_yield, chef_lines, procedure_lines
        content = [s.rstrip() for s in buf if s.strip()]
        if not content:
            return
        if sec == "recipe name" and not title:
            title = content[0]
        elif sec == "yield" and not recipe_yield:
            recipe_yield = " ".join(content)
        elif sec == "procedure" and not procedure_lines:
            for ln in content:
                step = _strip_bullet(ln)
                if step:
                    procedure_lines.append(step)
        elif sec in ("chef", "chef created by", "created by", "chef name"):
            for ln in content:
                nm = _strip_bullet(ln)
                if nm:
                    chef_lines.append(nm)

    for raw in text_lines:
        line = raw.rstrip("\r\n")

        # Section boundary markers
        if CAPTURE_BOUNDARY.match(line):
            if section is not None:
                flush(section)
                section = None
                buf = []
            continue

        # New header?
        m = HEADER_PATTERN.match(line)
        if m:
            if section is not None:
                flush(section)
            section = m.group(1).lower()
            rest = m.group(2).strip()
            buf = [rest] if rest else []
            continue

        # Continue appending to current section
        if section in {
            "recipe name", "yield", "procedure",
            "chef", "chef created by", "created by", "chef name",
            "chef notes", "ingredients", "ingredient"
        }:
            buf.append(line)

    # Flush last section
    if section is not None:
        flush(section)

    # Compose outputs
    title = title or "Untitled"
    # de-dupe while preserving order
    chef_lines = [c for i, c in enumerate(chef_lines) if c and c.lower() not in {x.lower() for x in chef_lines[:i]}]
    chef = "; ".join(chef_lines) or "Chef blas/ Commissary ATL"
    recipe_yield = recipe_yield or "No yield found."
    procedure = "\n".join(procedure_lines) if procedure_lines else "No procedure found."
    return title, recipe_yield, chef, procedure

# --- Main ---
txt_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".txt")]

app = xw.App(visible=False)
app.display_alerts = False
app.screen_updating = False
try:
    wb = app.books.open(template_path)
    template_sheet = wb.sheets[0]
    used_names = {s.name for s in wb.sheets}

    for txt_file in txt_files:
        print(f"Processing file: {txt_file}")
        file_path = os.path.join(folder_path, txt_file)
        with open(file_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()

        title, recipe_yield, chef, procedure = parse_recipe(lines)

        safe_title = sanitize_sheet_name(title, used_names)
        print(f"Creating sheet for: {safe_title}")
        new_sheet = template_sheet.copy(after=wb.sheets[-1])
        new_sheet.name = safe_title

        # --- Header / meta ---
        new_sheet.range("E3:N3").value = title
        new_sheet.range("E1:J1").value = chef
        new_sheet.range("E2:J2").value = chef

        # --- Procedure: primary block M6:N17 (up to 12 lines), optional overflow below
        proc_lines = [ln.strip() for ln in procedure.splitlines() if ln.strip()]
        rows = MAX_PROC_VISIBLE

        # Clear M6:N17 before writing
        new_sheet.range("M6:N17").value = [["", ""] for _ in range(rows)]

        # Write first 12 (or fewer), duplicated across M and N per your template
        first_chunk = proc_lines[:rows]
        block = [[first_chunk[i] if i < len(first_chunk) else "",
                  first_chunk[i] if i < len(first_chunk) else ""] for i in range(rows)]
        new_sheet.range("M6:N17").value = block

        # Overflow handling
        if len(proc_lines) > rows:
            print(f"[WARN] {safe_title}: Procedure has {len(proc_lines)} lines; wrote first {rows} to M6:N17.")
            if EXPAND_PROCEDURE:
                # Continue writing from row 18 downward in column M (single column to avoid breaking layout)
                start_row = 18
                rest = proc_lines[rows:]
                new_sheet.range(f"M{start_row}:M{start_row + len(rest) - 1}").value = [[s] for s in rest]

        # --- Yield ---
        new_sheet.range("N34").value = recipe_yield

    # Remove template sheet and save
    template_sheet.delete()
    wb.save(output_path)
    print(f"Excel file created and saved as '{output_path}' with {len(txt_files)} sheets.")
finally:
    app.quit()
