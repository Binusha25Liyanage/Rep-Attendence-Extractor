"""
Lakmee SFA - Rep Attendance Image Maker (desktop app)
-----------------------------------------------------
Run it, click "Choose Excel file", pick the columns you want in the picture,
click "Generate". You get cropped PNG image(s) plus a PDF.

Install once:   python -m pip install openpyxl reportlab pillow
(tkinter is already included with normal Windows Python.)
"""
import os, sys, json, datetime
import openpyxl
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# =====================================================================
#  Fonts (Calibri on Windows, same look as Excel)
# =====================================================================
_FONT_DIRS = [r"C:\Windows\Fonts",
              os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Microsoft\Windows\Fonts"),
              "/usr/share/fonts/truetype/crosextra", "/usr/share/fonts/truetype/liberation"]

def _find(names):
    for d in _FONT_DIRS:
        for n in names:
            p = os.path.join(d, n)
            if os.path.exists(p):
                return p
    return None

_REG = _find(["calibri.ttf", "Carlito-Regular.ttf", "arial.ttf", "LiberationSans-Regular.ttf"])
_BLD = _find(["calibrib.ttf", "Carlito-Bold.ttf", "arialbd.ttf", "LiberationSans-Bold.ttf"])
if _REG and _BLD:
    pdfmetrics.registerFont(TTFont("F", _REG)); pdfmetrics.registerFont(TTFont("F-B", _BLD))
    FONT, FONT_B = "F", "F-B"
else:
    FONT, FONT_B = "Helvetica", "Helvetica-Bold"

def _pil_font(bold, size):
    path = _BLD if bold else _REG
    try:
        return ImageFont.truetype(path, size) if path else ImageFont.load_default(size)
    except Exception:
        return ImageFont.load_default()

# =====================================================================
#  Read the SFA Excel
# =====================================================================
def _fmt(v):
    if v is None or v == "":
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, datetime.time):
        return v.strftime("%H:%M:%S")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def read_excel(path):
    """Returns dict(labels=[...], header_of={label: shown header}, records=[{label: text}], date=datetime|None)."""
    ws = openpyxl.load_workbook(path, data_only=True).active
    header_row = None
    for row in ws.iter_rows(min_row=1, max_row=40):
        vals = [str(c.value).strip().upper() for c in row if c.value is not None]
        if "TERRITORY" in vals and "LOGGED IN" in vals:
            header_row = row[0].row
            break
    if header_row is None:
        raise ValueError("Could not find the header row (TERRITORY / LOGGED IN). Is this the RepAttendance export?")

    ncol = ws.max_column
    # group names sit one row above the headers (merged cells) -> forward fill
    groups, cur = {}, ""
    for c in range(1, ncol + 1):
        g = ws.cell(header_row - 1, c).value if header_row > 1 else None
        if g:
            cur = str(g).strip()
        groups[c] = cur

    names = {c: str(ws.cell(header_row, c).value).strip() for c in range(1, ncol + 1)
             if ws.cell(header_row, c).value not in (None, "")}
    count = {}
    for n in names.values():
        count[n] = count.get(n, 0) + 1

    labels, header_of, col_of = [], {}, {}
    for c, n in names.items():
        label = n if count[n] == 1 else f"{n} ({groups[c] or 'col ' + str(c)})"
        while label in col_of:
            label += "'"
        labels.append(label); header_of[label] = n; col_of[label] = c

    def first(name):
        return next(l for l in labels if header_of[l] == name)
    in_l, out_l, terr_l = first("LOGGED IN"), first("LOGGED OUT"), first("TERRITORY")
    date_l = next((l for l in labels if header_of[l] == "LOGGED IN DATE"), None)

    records, date = [], None
    for r in range(header_row + 1, ws.max_row + 1):
        if not ws.cell(r, col_of[terr_l]).value:
            continue
        rec = {l: _fmt(ws.cell(r, col_of[l]).value) for l in labels}
        records.append(rec)
        if date is None and date_l:
            d = ws.cell(r, col_of[date_l]).value
            try:
                date = d if isinstance(d, (datetime.date, datetime.datetime)) else datetime.datetime.strptime(str(d)[:10], "%Y-%m-%d")
            except Exception:
                pass
    return dict(labels=labels, header_of=header_of, records=records, date=date,
                in_l=in_l, out_l=out_l, terr_l=terr_l)

# =====================================================================
#  Table geometry (points) -> used by both PDF and PNG
# =====================================================================
FS, ROW_H, PAD, CELL_PAD, MAX_CHARS = 11.5, 21, 1.5, 8, 45

def _clip(s):
    return s if len(s) <= MAX_CHARS else s[:MAX_CHARS - 1] + "…"

def table_cells(title, headers, rows):
    rows = [[_clip(v) if v else "-" for v in r] for r in rows]
    sw = pdfmetrics.stringWidth
    widths = [max(80, max([sw(h, FONT_B, FS)] + [sw(r[i], FONT, FS) for r in rows]) + 2 * CELL_PAD)
              for i, h in enumerate(headers)]
    need = sw(title, FONT_B, FS) + 2 * CELL_PAD
    if sum(widths) < need:
        widths[-1] += need - sum(widths)
    total = sum(widths)
    items = [(PAD, PAD, total, ROW_H, title, "c", True)]
    def add_row(vals, bold, y):
        x = PAD
        for i, v in enumerate(vals):
            items.append((x, y, widths[i], ROW_H, v, "l" if i == 0 else "c", bold))
            x += widths[i]
    y = PAD + ROW_H
    add_row(headers, True, y)
    for r in rows:
        y += ROW_H
        add_row(r, False, y)
    return items, total + 2 * PAD, ROW_H * (len(rows) + 2) + 2 * PAD

def draw_pdf_page(c, title, headers, rows):
    items, W, H = table_cells(title, headers, rows)
    c.setPageSize((W, H)); c.setLineWidth(0.8); c.setStrokeColorRGB(0, 0, 0)
    for x, y, w, h, text, align, bold in items:
        yb = H - y - h
        c.rect(x, yb, w, h)
        c.setFont(FONT_B if bold else FONT, FS)
        if align == "l":
            c.drawString(x + 4, yb + 6.5, text)
        else:
            c.drawCentredString(x + w / 2, yb + 6.5, text)
    c.showPage()

def draw_png(path, title, headers, rows, dpi=300):
    items, W, H = table_cells(title, headers, rows)
    k = dpi / 72.0
    im = Image.new("RGB", (int(W * k), int(H * k)), "white")
    d = ImageDraw.Draw(im)
    fb, fr = _pil_font(True, int(FS * k)), _pil_font(False, int(FS * k))
    lw = max(2, int(0.8 * k))
    for x, y, w, h, text, align, bold in items:
        d.rectangle([x * k, y * k, (x + w) * k, (y + h) * k], outline="black", width=lw)
        f = fb if bold else fr
        if align == "l":
            d.text(((x + 4) * k, (y + h / 2) * k), text, font=f, fill="black", anchor="lm")
        else:
            d.text(((x + w / 2) * k, (y + h / 2) * k), text, font=f, fill="black", anchor="mm")
    im.save(path, dpi=(dpi, dpi))

# =====================================================================
#  Build the outputs
# =====================================================================
MODES = {
    "Two images: not logged out + logged out": "split",
    "One image: only reps NOT logged out": "not_out",
    "One image: only reps who logged out": "out",
    "One image: all reps who logged in": "all",
}
SORTS = ["Same order as Excel", "Territory A-Z", "Login time"]

def generate(data, chosen, mode, sort, title_tpl, png_dirs, pdf_dir):
    """chosen = column labels in output order; png_dirs = folder per image (1st, 2nd); pdf_dir = folder for the PDF."""
    headers = [data["header_of"][l] for l in chosen]
    recs = [r for r in data["records"] if r[data["in_l"]]]            # only reps who logged in
    if sort == "Territory A-Z":
        recs.sort(key=lambda r: r[data["terr_l"]].lower())
    elif sort == "Login time":
        recs.sort(key=lambda r: r[data["in_l"]])
    not_out = [r for r in recs if not r[data["out_l"]]]
    done = [r for r in recs if r[data["out_l"]]]
    sets = {"split": [("not_logged_out", not_out), ("logged_out", done)],
            "not_out": [("not_logged_out", not_out)],
            "out": [("logged_out", done)],
            "all": [("all", recs)]}[mode]
    sets = [(n, s) for n, s in sets if s]
    if not sets:
        raise ValueError("No reps match this filter, nothing to draw.")
    date_txt = data["date"].strftime("%m/%d/%Y") if data["date"] else ""
    title = title_tpl.replace("{date}", date_txt).strip()

    os.makedirs(pdf_dir, exist_ok=True)
    pdf_path = os.path.join(pdf_dir, "rep_attendance.pdf")
    c = canvas.Canvas(pdf_path)
    pngs = []
    for idx, (name, subset) in enumerate(sets):
        table = [[r[l] for l in chosen] for r in subset]
        draw_pdf_page(c, title, headers, table)
        folder = png_dirs[min(idx, len(png_dirs) - 1)]
        os.makedirs(folder, exist_ok=True)
        p = os.path.join(folder, f"rep_attendance_{name}.png")
        draw_png(p, title, headers, table)
        pngs.append((name, len(subset), p))
    c.save()
    return pngs, pdf_path

SETTINGS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rep_attendance_settings.json")

def load_settings():
    try:
        with open(SETTINGS, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_settings(d):
    try:
        with open(SETTINGS, "w", encoding="utf-8") as f:
            json.dump(d, f, indent=2)
    except Exception:
        pass

# =====================================================================
#  The window
# =====================================================================
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from PIL import ImageTk

    root = tk.Tk()
    root.title("Rep Attendance Image Maker - Lakmee")
    root.geometry("1120x680")
    root.minsize(980, 620)
    state = {"data": None, "photos": []}

    top = ttk.Frame(root, padding=10); top.pack(fill="x")
    ttk.Button(top, text="1. Choose Excel file…", command=lambda: choose_file()).pack(side="left")
    file_var = tk.StringVar(value="No file chosen")
    ttk.Label(top, textvariable=file_var, foreground="#555").pack(side="left", padx=10)

    body = ttk.Frame(root, padding=(10, 0)); body.pack(fill="both", expand=True)
    left = ttk.Frame(body); left.pack(side="left", fill="both", expand=True)
    right = ttk.LabelFrame(body, text="Preview", padding=6); right.pack(side="right", fill="both", padx=(10, 0))

    # ---- column pickers ----
    ttk.Label(left, text="2. Choose the columns for the image", font=("Segoe UI", 10, "bold")).pack(anchor="w")
    pick = ttk.Frame(left); pick.pack(fill="both", expand=True, pady=4)
    def make_list(parent, title):
        f = ttk.Frame(parent); ttk.Label(f, text=title).pack(anchor="w")
        inner = ttk.Frame(f); inner.pack(fill="both", expand=True)
        lb = tk.Listbox(inner, selectmode="extended", exportselection=False, height=14, width=30)
        sb = ttk.Scrollbar(inner, command=lb.yview); lb.config(yscrollcommand=sb.set)
        lb.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
        return f, lb
    f1, avail = make_list(pick, "Available columns (double-click to add)")
    f2, chosen = make_list(pick, "Columns in the image (top = left-most)")
    f1.pack(side="left", fill="both", expand=True)
    mid = ttk.Frame(pick); mid.pack(side="left", padx=6)
    f2.pack(side="left", fill="both", expand=True)

    def add():
        have = list(chosen.get(0, "end"))
        for i in avail.curselection():
            v = avail.get(i)
            if v not in have:
                chosen.insert("end", v)
    def remove():
        for i in reversed(chosen.curselection()):
            chosen.delete(i)
    def move(delta):
        items = list(chosen.get(0, "end"))
        sel = set(chosen.curselection())
        if not sel:
            return
        for i in (sorted(sel) if delta < 0 else sorted(sel, reverse=True)):
            j = i + delta
            if 0 <= j < len(items) and j not in sel:
                items[i], items[j] = items[j], items[i]
                sel.remove(i); sel.add(j)
        chosen.delete(0, "end")
        for v in items:
            chosen.insert("end", v)
        for j in sel:
            chosen.selection_set(j)
    for text, cmd in (("Add  >", add), ("<  Remove", remove), ("Move up ↑", lambda: move(-1)), ("Move down ↓", lambda: move(1))):
        ttk.Button(mid, text=text, width=12, command=cmd).pack(pady=3)
    avail.bind("<Double-Button-1>", lambda e: add())
    chosen.bind("<Double-Button-1>", lambda e: remove())

    # ---- options ----
    opt = ttk.LabelFrame(left, text="3. Options", padding=8); opt.pack(fill="x", pady=6)
    ttk.Label(opt, text="Rows to show").grid(row=0, column=0, sticky="w")
    mode_var = tk.StringVar(value=list(MODES)[0])
    ttk.Combobox(opt, textvariable=mode_var, values=list(MODES), state="readonly", width=44).grid(row=0, column=1, sticky="w", padx=6, pady=2)
    ttk.Label(opt, text="Sort by").grid(row=1, column=0, sticky="w")
    sort_var = tk.StringVar(value=SORTS[0])
    ttk.Combobox(opt, textvariable=sort_var, values=SORTS, state="readonly", width=44).grid(row=1, column=1, sticky="w", padx=6, pady=2)
    ttk.Label(opt, text="Title ({date} = report date)").grid(row=2, column=0, sticky="w")
    title_var = tk.StringVar(value="LOGGED IN DATE {date}")
    ttk.Entry(opt, textvariable=title_var, width=47).grid(row=2, column=1, sticky="w", padx=6, pady=2)
    saved = load_settings()
    out1, out2, outp = (tk.StringVar(value=saved.get(k, "")) for k in ("img1", "img2", "pdf"))
    lab1, lab2 = tk.StringVar(), tk.StringVar()
    def browse(var):
        d = filedialog.askdirectory(title="Choose a folder", initialdir=var.get() or None)
        if d: var.set(d)
    for r, (lv, var) in enumerate(((lab1, out1), (lab2, out2), (tk.StringVar(value="PDF saved in"), outp)), start=3):
        ttk.Label(opt, textvariable=lv).grid(row=r, column=0, sticky="w")
        ttk.Entry(opt, textvariable=var, width=47).grid(row=r, column=1, sticky="w", padx=6, pady=2)
        ttk.Button(opt, text="Browse…", command=lambda v=var: browse(v)).grid(row=r, column=2)
    def update_labels(*_):
        m = MODES[mode_var.get()]
        lab1.set({"split": "Image 1 (not logged out) saved in", "not_out": "Image (not logged out) saved in",
                  "out": "Image (logged out) saved in", "all": "Image (all reps) saved in"}[m])
        lab2.set("Image 2 (logged out) saved in" if m == "split" else "(only used for the two-image option)")
    mode_var.trace_add("write", update_labels); update_labels()

    act = ttk.Frame(left); act.pack(fill="x", pady=4)
    ttk.Button(act, text="4. Generate images + PDF", command=lambda: do_generate()).pack(side="left")
    ttk.Button(act, text="Open the folders", command=lambda: open_folder()).pack(side="left", padx=8)
    status = tk.StringVar(); ttk.Label(left, textvariable=status, foreground="#0a6").pack(anchor="w")

    nb = ttk.Notebook(right); nb.pack(fill="both", expand=True)
    ttk.Label(right, text="Your image appears here after Generate.", foreground="#777").pack()

    # ---- actions ----
    def choose_file():
        p = filedialog.askopenfilename(title="Choose the RepAttendance Excel file",
                                       filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if not p: return
        try:
            data = read_excel(p)
        except Exception as e:
            messagebox.showerror("Could not read file", str(e)); return
        state["data"] = data
        file_var.set(f"{os.path.basename(p)}   ({len(data['records'])} rows)")
        for v in (out1, out2, outp):
            if not v.get():
                v.set(os.path.dirname(p))      # first time only: default to the Excel's folder
        avail.delete(0, "end"); chosen.delete(0, "end")
        for l in data["labels"]:
            avail.insert("end", l)
        for l in (data["terr_l"], data["in_l"], data["out_l"]):
            chosen.insert("end", l)
        status.set("Columns loaded. Adjust the list on the right, then Generate.")

    def do_generate():
        if not state["data"]:
            messagebox.showinfo("No file", "Choose an Excel file first."); return
        cols = list(chosen.get(0, "end"))
        if not cols:
            messagebox.showinfo("No columns", "Add at least one column."); return
        try:
            f1 = out1.get() or os.path.dirname(os.path.abspath(__file__))
            f2 = out2.get() or f1
            fp = outp.get() or f1
            pngs, pdf = generate(state["data"], cols, MODES[mode_var.get()], sort_var.get(),
                                 title_var.get(), [f1, f2], fp)
            save_settings({"img1": f1, "img2": f2, "pdf": fp})
        except Exception as e:
            messagebox.showerror("Could not generate", str(e)); return
        for t in nb.tabs(): nb.forget(t)
        state["photos"].clear()
        for name, n, p in pngs:
            im = Image.open(p); im.thumbnail((460, 520))
            ph = ImageTk.PhotoImage(im); state["photos"].append(ph)
            tab = ttk.Frame(nb); ttk.Label(tab, image=ph).pack(padx=4, pady=4)
            nb.add(tab, text=f"{name.replace('_', ' ')} ({n})")
        status.set("Done. Saved: " + "  |  ".join([os.path.basename(p) for _, _, p in pngs] + ["rep_attendance.pdf"]))

    def open_folder():
        for p in dict.fromkeys([out1.get(), out2.get(), outp.get()]):     # each distinct folder once
            if p and os.path.isdir(p) and sys.platform == "win32":
                os.startfile(p)

    root.mainloop()

if __name__ == "__main__":
    run_gui()
