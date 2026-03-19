
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime
from fpdf import FPDF
import tempfile
import webbrowser
import re
import textwrap
import sys
from openpyxl.styles import Alignment

EXCLUDE_SUPPLIER_KEYWORDS = [
    "bama",
    "samson",
    "odd langdalen",
    "nærbakst",     # dekker "Nærbakst" og "Nærbakst AS"
    "naerbakst",    # hvis noen ganger uten ø
    "bakehuset",
    "baker klausen",
]

def filter_excluded_suppliers(df: pd.DataFrame) -> pd.DataFrame:
    """Fjerner rader med suppliers vi ikke vil ha med (case-insensitiv substring match)."""
    if df is None or df.empty:
        return df
    if "Supplier" not in df.columns:
        return df

    s = df["Supplier"].fillna("").astype(str).str.strip().str.casefold()
    mask_exclude = s.apply(lambda x: any(k in x for k in EXCLUDE_SUPPLIER_KEYWORDS))
    return df.loc[~mask_exclude].copy()

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # PyInstaller
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))  # mappen der .py ligger
    return os.path.join(base_path, relative_path)

class RoutePDF(FPDF):
    def __init__(self, listetype="Brødliste", dato="ukjent dato", krs_mode=False):
        super().__init__(orientation="L")
        self.listetype = listetype
        self.dato = dato
        self.krs_mode = krs_mode
        
        font_path = resource_path("DejaVuSans.ttf")
        
        self.add_font("DejaVu", "", font_path)
        self.add_font("DejaVu", "B", font_path)
        self.set_font("DejaVu", "", 10)  # standard

    def header(self):
        try:
            self.image(resource_path("logo.gif"), x=10, y=8, w=40)
            self.ln(16)
        except RuntimeError:
            pass
        self.set_font("DejaVu", "B", 14)
        self.cell(0, 10, self.route_title, new_x="LMARGIN", new_y="NEXT", align='L')
        self.ln(4)


    def footer(self):
        self.set_y(-15)
        self.set_font("DejaVu", "", 8)
        date_str = datetime.today().strftime('%d.%m.%Y %H:%M')
        self.cell(0, 10, f"{date_str} - Side {self.page_no()}", align='C')

    def add_route_table(self, df):
        df = df.fillna("")
        
        # Skjul Route-kolonnen der ruten allerede vises i tittelen
        if getattr(self, "listetype", "") in ("Brødliste", "Fryseliste", "Plukkliste"):
            for col in ("Route", "Route nickname"):
                if col in df.columns:
                    df = df.drop(columns=[col])
                    break
        
        columns = df.columns.tolist()
        page_width = self.w - 2 * self.l_margin
        font_size = 9
        line_height = 4

        # --- Dynamisk kolonnebredde (beregnet per gruppe/rute) ---
        MM_PER_CHAR = 1.65   # DejaVu Sans 9pt, målt med PIL
        PADDING = 3.0
        FIXED_COLS = {"Order ID", "Quantity", "Missing", "AA"}
        WRAPPABLE = {"Product Name", "Customer", "Department",
                     "Route nickname", "Route", "Position", "Supplier"}

        self.set_font("DejaVu", "B", font_size)

        # Faste kolonner: akkurat bred nok for bredeste verdi i denne gruppen
        fixed_widths = {}
        for col in columns:
            if col in FIXED_COLS:
                header_w = self.get_string_width(col) + PADDING
                max_val_w = max(
                    (self.get_string_width(str(v)) + PADDING for v in df[col]),
                    default=0.0,
                )
                fixed_widths[col] = max(header_w, max_val_w)

        fixed_total = sum(fixed_widths.values())
        remaining = page_width - fixed_total

        # Wrappable kolonner:
        # 1. Finn naturlig cap = bredeste enkeltlinje-verdi i denne gruppen (ikke hele teksten)
        #    Dette er maks vi noensinne trenger — mer plass hjelper ikke.
        # 2. Finn min_w = bredeste enkeltord (textwrap bryter aldri midt i ord)
        # 3. Fordel remaining proporsjonalt, klemt mellom [min_w, natural_cap]
        wrap_weights = {
            "Product Name":   3.5,
            "Customer":       5.5,
            "Department":     1.0,
            "Position":       1.8,
            "Supplier":       1.5,
            "Route nickname": 1.5,
            "Route":          1.5,
        }
        active_wrap = [c for c in columns if c in WRAPPABLE]
        weight_total = sum(wrap_weights.get(c, 1.0) for c in active_wrap)

        wrap_widths = {}
        natural_caps = {}
        min_widths_wrap = {}
        for col in active_wrap:
            vals = df[col].astype(str).tolist()

            # Minstebredde = bredeste enkeltord (textwrap bryter aldri midt i ord)
            all_words = []
            for v in vals:
                all_words.extend(v.split())
            min_w = max(
                (self.get_string_width(w) + PADDING for w in all_words),
                default=PADDING,
            ) if all_words else PADDING
            min_w = max(min_w, self.get_string_width(col) + PADDING)
            min_widths_wrap[col] = min_w

            # Natural cap = bredeste enkeltverdi i gruppen (ikke bruttoverdi)
            natural_cap = max(
                (self.get_string_width(v) + PADDING for v in vals),
                default=min_w,
            )
            natural_caps[col] = max(natural_cap, min_w)

        # Fase 1: start med forholdsmessig fordeling, klemt til [min_w, natural_cap]
        for col in active_wrap:
            proportional = (wrap_weights.get(col, 1.0) / weight_total) * remaining
            wrap_widths[col] = min(max(proportional, min_widths_wrap[col]),
                                   natural_caps[col])

        # Fase 2: redistribuer ubrukt plass.
        # Prioriter kolonner som fortsatt overflower (verdi > kolonnebredde).
        # Gi dem plass fra kolonner der innholdet allerede passer (verdi <= bredde).
        for _pass in range(10):
            total_used = sum(wrap_widths.values())
            leftover = remaining - total_used
            if abs(leftover) < 0.1:
                break

            # Del kolonner i «trenger mer» og «kan gi»
            needs_more = [c for c in active_wrap
                          if natural_caps[c] > wrap_widths[c] + 0.1]
            can_shrink = [c for c in active_wrap
                          if wrap_widths[c] > min_widths_wrap[c] + 0.1
                          and natural_caps[c] <= wrap_widths[c] + 0.1]

            if leftover > 0.1 and needs_more:
                # Gi ubrukt plass til kolonner som trenger mer
                grow_w = sum(wrap_weights.get(c, 1.0) for c in needs_more)
                for col in needs_more:
                    extra = (wrap_weights.get(col, 1.0) / grow_w) * leftover
                    wrap_widths[col] = min(wrap_widths[col] + extra,
                                           natural_caps[col])
            elif leftover < -0.1 and can_shrink:
                # Skaler ned kolonner som ikke trenger plassen de har fått
                shrink_w = sum(wrap_weights.get(c, 1.0) for c in can_shrink)
                for col in can_shrink:
                    give_back = (wrap_weights.get(col, 1.0) / shrink_w) * abs(leftover)
                    wrap_widths[col] = max(wrap_widths[col] - give_back,
                                           min_widths_wrap[col])
            else:
                break

        # Endelig skaler ned hvis total fortsatt overstiger siden
        wrap_total = sum(wrap_widths.values())
        if wrap_total > remaining + 0.5:
            scale = remaining / wrap_total
            wrap_widths = {c: w * scale for c, w in wrap_widths.items()}

        # Bygg endelig liste i kolonnerekke
        col_widths = []
        for col in columns:
            if col in FIXED_COLS:
                col_widths.append(fixed_widths[col])
            else:
                col_widths.append(wrap_widths.get(col, remaining))


        # Header bakgrunnsfarge
        if self.listetype == "Brødliste":
            self.set_fill_color(245, 222, 179)
        elif self.listetype == "Fryseliste":
            self.set_fill_color(173, 216, 230)
        elif self.listetype in ["Fryseliste Plukk"]:
            self.set_fill_color(173, 216, 230)
        elif self.listetype == "Bamaliste Plukk":
            self.set_fill_color(186, 255, 201)
        elif self.listetype == "Gehebliste Plukk":
            self.set_fill_color(245, 222, 179)
        elif self.listetype == "Plukkliste":
            self.set_fill_color(255, 75, 62)
        else:
            self.set_fill_color(220, 220, 220)

        # Skriv header
        self.set_font("DejaVu", "B", font_size)
        for i, col in enumerate(columns):
            align = 'C' if col in ("Quantity", "Missing") else 'L'
            self.cell(col_widths[i], line_height * 2, str(col), border=1, align=align, fill=True)
        self.ln()

        for _, row in df.iterrows():
            # Forbered tekst og linjebryting
            cell_lines = []
            max_lines = 1
            for i, val in enumerate(row):
                text = str(val).replace("\n", " ").replace("\r", " ")
                text = text.replace("Samson Bakeri AS", "Samson")
                column_name = columns[i]
                if column_name in WRAPPABLE:
                    # Bryt kun når teksten faktisk ikke får plass (1.65mm/tegn målt)
                    wrap_width = max(int(col_widths[i] / MM_PER_CHAR), 6)
                    lines = textwrap.wrap(text, width=wrap_width)
                    if not lines:
                        lines = [""]
                else:
                    lines = [text]
                cell_lines.append(lines)
                max_lines = max(max_lines, len(lines))
            row_height = max_lines * line_height

            # ------ Sjekk page break FØR cellene ------
            if self.get_y() + row_height > self.page_break_trigger:
                self.add_page()
                self.set_font("DejaVu", "B", font_size)
                for i, col in enumerate(columns):
                    align = 'C' if col in ("Quantity", "Missing") else 'L'
                    self.cell(col_widths[i], line_height * 2, str(col), border=1, align=align, fill=True)
                self.ln()
            # -------------------------------------------

            # Hent startposisjon for raden en gang
            
            x_start = self.get_x()
            y_start = self.get_y()

            # --- Kristiansand: ev. bakgrunnsfarge for hele raden ---
            if getattr(self, "krs_mode", False):
                # Finn relevante verdier i denne raden
                colmap = {columns[i]: (cell_lines[i][0] if cell_lines[i] else "") for i in range(len(columns))}
                product_val = str(colmap.get("Product Name", "")).casefold()
                supplier_val = str(colmap.get("Supplier", "")).casefold()

                # Prioritet: barnehage (rød) > geheb (blå)
                fill_rgb = None
                if "barnehage" in product_val:
                    fill_rgb = (255, 230, 230)   # svak rødfarge
                elif supplier_val == "geheb" or "geheb" in supplier_val:
                    fill_rgb = (230, 240, 255)   # svak blåfarge

                if fill_rgb:
                    self.set_fill_color(*fill_rgb)
                    # Fyll hele raden før vi tegner rammer/tekst
                    self.rect(x_start, y_start, sum(col_widths), row_height, style='F')
            

            # Tegn ramme rundt hele raden (på toppen av evt. fyll)
            for i in range(len(columns)):
                x = x_start + sum(col_widths[:i])
                y = y_start
                self.rect(x, y, col_widths[i], row_height)

            # Skriv rad linje for linje
            self.set_font("DejaVu", "", font_size)
            for line_no in range(max_lines):
                self.set_x(x_start)
                for i, lines in enumerate(cell_lines):
                    text = lines[line_no] if line_no < len(lines) else ""
                    align = 'C' if columns[i] in ("Quantity", "Missing") else 'L'
                    self.cell(col_widths[i], line_height, text, border=0, align=align)
                self.ln(line_height)
                
    def add_route_summary(self, df):
    #Viser et sammendrag per rute: sum Quantity per Product Name.
    #Skrives under rute-tabellen. Hopper til ny side ved behov.
    
        if "Product Name" not in df.columns or "Quantity" not in df.columns:
            return  # ingenting å oppsummere

        # Sørg for numeriske antall
        qty = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
        tmp = df.copy()
        tmp["__qty__"] = qty

        summary = (
            tmp.groupby("Product Name", dropna=False)["__qty__"]
                .sum()
                .reset_index()
                .sort_values("Product Name")
        )

        # Litt luft før sammendraget
        self.ln(2)

        page_width = self.w - 2 * self.l_margin
        # 75 % til navn, 25 % til antall
        name_w = page_width * 0.75
        qty_w  = page_width * 0.25
        line_h = 5
        fs_hdr = 10
        fs_row = 9

        # Tittel
        title = "Sammendrag - antall per type"
        if self.get_y() + line_h*2 > self.page_break_trigger:
            self.add_page()

        self.set_font("DejaVu", "B", fs_hdr)
        self.set_fill_color(235, 235, 235)
        self.cell(name_w + qty_w, line_h + 1, title, border=1, new_x="LMARGIN", new_y="NEXT", fill=True, align='L')

        # Header for minitabell
        if self.get_y() + line_h*2 > self.page_break_trigger:
            self.add_page()
            self.set_font("DejaVu", "B", fs_hdr)
            self.set_fill_color(235, 235, 235)
            self.cell(name_w + qty_w, line_h + 1, title, border=1, new_x="LMARGIN", new_y="NEXT", fill=True, align='L')

        self.set_font("DejaVu", "B", fs_row)
        self.set_fill_color(245, 245, 245)
        self.cell(name_w, line_h, "Product Name", border=1, fill=True)
        self.cell(qty_w,  line_h, "Sum Quantity", border=1, new_x="LMARGIN", new_y="NEXT", align='C', fill=True)

        # Rader
        self.set_font("DejaVu", "", fs_row)
        for _, r in summary.iterrows():
            if self.get_y() + line_h > self.page_break_trigger:
                self.add_page()
                self.set_font("DejaVu", "B", fs_hdr)
                self.set_fill_color(235, 235, 235)
                self.cell(name_w + qty_w, line_h + 1, title, border=1, new_x="LMARGIN", new_y="NEXT", fill=True, align='L')
                self.set_font("DejaVu", "B", fs_row)
                self.set_fill_color(245, 245, 245)
                self.cell(name_w, line_h, "Product Name", border=1, fill=True)
                self.cell(qty_w,  line_h, "Sum Quantity", border=1, new_x="LMARGIN", new_y="NEXT", align='C', fill=True)
                self.set_font("DejaVu", "", fs_row)

            self.cell(name_w, line_h, str(r["Product Name"]), border=1)
            self.cell(qty_w,  line_h, f'{int(r["__qty__"]):,}'.replace(',', ' '), border=1, new_x="LMARGIN", new_y="NEXT", align='C')


def format_and_sort_excel(file_path, listetype=None):
    try:
        df = pd.read_excel(file_path, sheet_name=0, engine="openpyxl")

        # Filter for listetyper som skal ekskludere suppliers
        if listetype in ["Plukkliste", "Sammendragsliste Plukk"]:
            df = filter_excluded_suppliers(df)
            
        # Ekstra filter: fjern alle linjer der Storage er Freezer (-22)
        if listetype in ["Plukkliste", "Sammendragsliste Plukk"]:
            if "Storage" in df.columns:
                storage = df["Storage"].fillna("").astype(str).str.strip()
                df = df[storage != "Freezer (-22)"].copy()

        # Plukkliste: fjern ordrelinjer til kunden "TRY AS"
        if listetype == "Plukkliste":
            if "Customer" in df.columns:
                df = df[df["Customer"].fillna("").astype(str).str.strip().str.casefold() != "try as"].copy()
        
        # Lag en numerisk hjelpekolonne for rute-rekkefølge (så 2 < 10)
        has_route_ord = "Route ordering" in df.columns
        if has_route_ord:
            df["__route_order__"] = pd.to_numeric(df["Route ordering"], errors="coerce")

        # --- Sortering ---
        # Nå også plukklister: Route nickname -> Route ordering (numerisk) -> Customer -> Department
        if listetype in ["Bamaliste Plukk", "Fryseliste Plukk", "Gehebliste Plukk"]:
            sort_keys = ["Route nickname"]
            sort_keys.append("__route_order__" if has_route_ord else "Route ordering")
            sort_keys += ["Customer", "Department"]
        else:
            # Brødliste / Fryseliste (samme logikk – behold stopprekkefølgen)
            sort_keys = ["Route nickname"]
            sort_keys.append("__route_order__" if has_route_ord else "Route ordering")
            sort_keys += ["Customer", "Department"]

        active_keys = [k for k in sort_keys if k in df.columns]
        df_sorted = df.sort_values(by=active_keys, ascending=True, na_position="last")

        # Rydd opp hjelpekolonner fra visning
        if "Route ordering" in df_sorted.columns:
            df_sorted = df_sorted.drop(columns=["Route ordering"])
        if "__route_order__" in df_sorted.columns:
            df_sorted = df_sorted.drop(columns=["__route_order__"])

        # Hvilke kolonner skal med videre (Route skjules senere i add_route_table for Brød/Fryse)
        if listetype == "Plukkliste":
            columns_to_keep = [
                "Order ID",
                "Quantity",
                "Product Name",
                "Position",
                "Customer",
                "Department",
                "Route nickname",
                "Accept alternatives",
            ]
        else:
            columns_to_keep = [
                "Order ID",
                "Quantity",
                "Product Name",
                "Supplier",
                "Customer",
                "Department",
                "Route nickname",
                "Accept alternatives",
            ]
        if listetype == "Bamaliste Plukk" and "Supplier" in columns_to_keep:
            columns_to_keep.remove("Supplier")

        df_sorted = df_sorted[[c for c in columns_to_keep if c in df_sorted.columns]]

        # Legg til tom "Missing"-kolonne etter "Quantity" for Plukkliste
        if listetype == "Plukkliste" and "Quantity" in df_sorted.columns:
            qty_idx = df_sorted.columns.tolist().index("Quantity")
            df_sorted.insert(qty_idx + 1, "Missing", "")

        # Visningsnavn
        df_sorted = df_sorted.rename(columns={
            "Accept alternatives": "AA",
            "Route nickname": "Route",
        })

        return df_sorted
    except Exception as e:
        messagebox.showerror("Feil", f"Kunne ikke lese eller formatere Excel-fil: {str(e)}")
        return None


def save_to_pdf(grouped_data, output_path, listetype="Brødliste", dato="ukjent dato", krs_mode=False):
    if listetype == ("Sammendragsliste Brød", "Sammendragsliste Plukk"):
        # grouped_data er en DataFrame (df_sorted)
        df_all = grouped_data.fillna("")
        pdf = RoutePDF(listetype=listetype, dato=dato, krs_mode=krs_mode)
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.route_title = f"{listetype} til {dato}"
        pdf.add_page()

        # En liten nøytral headerfarge (valgfritt)
        # pdf.set_fill_color(220, 220, 220)

        # Skriv sammendrag for hver rute etter hverandre
        for route, group in df_all.groupby("Route"):
            if pdf.get_y() + 30 > pdf.page_break_trigger:
                pdf.add_page()
            pdf.set_font("DejaVu", "B", 11)
            pdf.cell(0, 6, f"Route: {route}", new_x="LMARGIN", new_y="NEXT")
            pdf.add_route_summary(group)
            pdf.ln(4)

        pdf.output(output_path)
        return
    # Sjekk listetype
    if listetype in ["Brødliste", "Fryseliste", "Plukkliste"]:
        grouped = grouped_data
        pdf = RoutePDF(listetype=listetype, dato=dato, krs_mode=krs_mode)
        pdf.set_auto_page_break(auto=True, margin=15)
        for route, group in grouped:
            pdf.route_title = f"{listetype} til {dato} - Route: {route}"
            pdf.add_page()
            pdf.add_route_table(group)
            if listetype in ["Brødliste", "Plukkliste"]:
                pdf.add_route_summary(group)
        pdf.output(output_path)
    else:
        # Plukk-lister: EN tabell med hele dataen, ikke grupper!
        pdf = RoutePDF(listetype=listetype, dato=dato, krs_mode=krs_mode)
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.route_title = f"{listetype} til {dato}"
        pdf.add_page()
        pdf.add_route_table(grouped_data)
        pdf.output(output_path)

def save_to_excel(grouped_data, output_path):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for route, group in grouped_data:
            safe_name = str(route)[:31]
            group.to_excel(writer, sheet_name=safe_name, index=False)
            sheet = writer.sheets[safe_name]
            for idx, cell in enumerate(sheet[1], start=1):
                if cell.value == "Quantity":
                    for row in sheet.iter_rows(min_col=idx, max_col=idx, min_row=2, max_row=sheet.max_row):
                        for c in row:
                            c.alignment = Alignment(horizontal="center")
                    break

def print_via_pdf(grouped_data, listetype="Brødliste", dato="ukjent dato", krs_mode=False):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        save_to_pdf(grouped_data, temp_pdf.name, listetype, dato, krs_mode=krs_mode)
        webbrowser.open(temp_pdf.name)

def main():
    # Opprett et midlertidig rot-vindu kun for fildialogen
    root = tk.Tk()
    root.withdraw() 
    
    file_path = filedialog.askopenfilename(title="Velg Excel-fil", filetypes=[("Excel files", "*.xlsx")])
    
    # Lukk det midlertidige vinduet når det er ferdig
    root.destroy() 

    if not file_path:
        messagebox.showinfo("Avbrutt", "Ingen fil valgt.")
        return

    # --- Nå lager vi det ENESTE hovedvinduet for programmet ---
    gui = tk.Tk()
    gui.title("Excel Formatter - Velg handling")
    # Sørg for at vinduet dukker opp i forgrunnen
    gui.lift()
    gui.attributes('-topmost', True)
    gui.after(0, lambda: gui.attributes('-topmost', False))
    gui.focus_force()
    
    # Opprett variabelen ETTER at hovedvinduet er laget
    listetype = tk.StringVar(value="Brødliste")

    # Prøv å finne dato fra filnavnet
    base_name = os.path.basename(file_path)
    match = re.search(r"\d{4}-\d{2}-\d{2}", base_name)
    if match:
        date_part = match.group(0)
    else:
        date_part = "ukjent dato"

    frame_type = tk.Frame(gui)
    frame_type.pack(pady=10)

    label_type = tk.Label(frame_type, text="Velg listetype:")
    label_type.pack(side=tk.LEFT)
    
    # Kristiansand-modus (farge-regler + regnbuebanner)
    krs_var = tk.BooleanVar(value=False)

    # Regnbuebanner (skjules/vises)
    banner_frame = tk.Frame(gui)
    # Bygg smale striper (regnbue)
    rainbow = ["#FF0000", "#FF7F00", "#FFFF00", "#00FF00", "#0000FF", "#4B0082", "#8B00FF"]
    for c in rainbow:
        stripe = tk.Frame(banner_frame, bg=c, height=6)
        stripe.pack(fill=tk.X)

    def apply_krs_theme(on):
        if on:
            # vis banner øverst
            banner_frame.pack(fill=tk.X, side=tk.TOP)
        else:
            banner_frame.pack_forget()

    # Toggle-funksjon koblet til checkbutton
    def on_toggle_krs():
        apply_krs_theme(krs_var.get())

    # Knapp for å skru av/på Kristiansand-modus
    krs_btn = tk.Checkbutton(gui, text="Kristiansand", variable=krs_var, command=on_toggle_krs)
    krs_btn.pack(pady=(0, 6))

    # Disse vil nå fungere korrekt da de tilhører 'gui' og 'listetype'
    listetyper = ["Brødliste", "Fryseliste", "Plukkliste", "Sammendragsliste Plukk", "Sammendragsliste Brød", "Fryseliste Plukk", "Bamaliste Plukk", "Gehebliste Plukk"]
    listetype = tk.StringVar(value=listetyper[0])
    dropdown = tk.OptionMenu(frame_type, listetype, *listetyper)
    dropdown.pack(side=tk.LEFT, padx=5)

    label_file = tk.Label(gui, text=f"Valgt fil:\n{os.path.basename(file_path)}", fg="blue")
    label_file.pack(pady=(10, 5))
    label = tk.Label(gui, text="Velg hva du vil gjøre med denne filen:")
    label.pack(pady=5)
    
    def on_print():
        lt = listetype.get()
        df_sorted = format_and_sort_excel(file_path, lt)
        if df_sorted is None or df_sorted.empty:
            return
        if "Route" not in df_sorted.columns:
            messagebox.showerror("Feil", "Filen mangler kolonnen 'Route'.")
            return
        if lt in ["Brødliste", "Fryseliste", "Plukkliste"]:
            grouped = df_sorted.groupby("Route")
            print_via_pdf(grouped, lt, date_part, krs_mode=krs_var.get())
        elif lt == "Sammendragsliste Brød":
            print_via_pdf(df_sorted, lt, date_part, krs_mode=krs_var.get())
        else:
            print_via_pdf(df_sorted, lt, date_part, krs_mode=krs_var.get())
        messagebox.showinfo("Ferdig", "PDF generert og åpnet for utskrift.")
        gui.destroy()



    def on_pdf():
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not output_path:
            return
        lt = listetype.get()
        df_sorted = format_and_sort_excel(file_path, lt)
        if df_sorted is None or df_sorted.empty:
            return
        if "Route" not in df_sorted.columns:
            messagebox.showerror("Feil", "Filen mangler kolonnen 'Route'.")
            return
        if lt in ["Brødliste", "Fryseliste"]:
            grouped = df_sorted.groupby("Route")
            save_to_pdf(grouped, output_path, lt, date_part, krs_mode=krs_var.get())
        elif lt == "Sammendragsliste Brød":
            save_to_pdf(df_sorted, output_path, lt, date_part, krs_mode=krs_var.get())    
        else:
            save_to_pdf(df_sorted, output_path, lt, date_part, krs_mode=krs_var.get())
        messagebox.showinfo("Ferdig", f"PDF lagret:\n{output_path}")
        gui.destroy()



    def on_excel():
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not output_path:
            return
        lt = listetype.get()
        df_sorted = format_and_sort_excel(file_path, lt)
        if df_sorted is None or df_sorted.empty:
            return
        if "Route" not in df_sorted.columns:
            messagebox.showerror("Feil", "Filen mangler kolonnen 'Route'.")
            return
        if lt in ["Brødliste", "Fryseliste"]:
            grouped = df_sorted.groupby("Route")
            save_to_excel(grouped, output_path)
        else:
            # Plukklister: Lag Excel av hele tabellen
            df_sorted.to_excel(output_path, index=False)
        messagebox.showinfo("Ferdig", f"Excel lagret:\n{output_path}")
        gui.destroy()


    btn_print = tk.Button(gui, text="🖨️ Skriv ut (via PDF)", command=on_print, width=40)
    btn_print.pack(pady=5)
    btn_pdf = tk.Button(gui, text="📄 Lagre som PDF", command=on_pdf, width=40)
    btn_pdf.pack(pady=5)
    btn_excel = tk.Button(gui, text="💾 Lagre som ny Excel-fil", command=on_excel, width=40)
    btn_excel.pack(pady=5)


    gui.mainloop()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nAvbrutt med CTRL+C. Programmet avsluttes.")

