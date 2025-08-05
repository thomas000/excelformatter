
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

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class RoutePDF(FPDF):
    def __init__(self, listetype="Brødliste", dato="ukjent dato"):
        super().__init__(orientation="L")
        self.listetype = listetype
        self.dato = dato

    def header(self):
        # Prøver å vise logo øverst til venstre (juster w=40 ved behov)
        try:
            self.image(resource_path("logo.gif"), x=10, y=8, w=40)

            self.ln(16)  # Gir luft etter logoen – juster om nødvendig
        except RuntimeError:
            pass  # Om logo mangler, hopper vi over den

        self.set_font("Helvetica", 'B', 14)
        self.cell(0, 10, self.route_title, ln=True, align='L')
        self.ln(4)


    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        date_str = datetime.today().strftime('%d.%m.%Y')
        self.cell(0, 10, f"{date_str} - Side {self.page_no()}", align='C')

    def add_route_table(self, df):
        df = df.fillna("")
        columns = df.columns.tolist()
        page_width = self.w - 2 * self.l_margin

        # Kolonneforhold
        ratios = {
            "Order ID": 1,
            "Quantity": 0.8,
            "Product Name": 4.7,
            "Supplier": 1.1,
            "Customer": 3.7,
            "Department": 1.5,
            "Route nickname": 1,
            "AA": 0.5,
        }
        active_ratios = {col: ratios.get(col, 1) for col in columns}
        total_ratio = sum(active_ratios.values())
        col_widths = [(active_ratios[col] / total_ratio) * page_width for col in columns]

        font_size = 9
        line_height = 4
        self.set_font("Helvetica", 'B', font_size)

        # Header bakgrunnsfarge
        if self.listetype == "Brødliste":
            self.set_fill_color(245, 222, 179)  # Lysebrun
        elif self.listetype == "Fryseliste":
            self.set_fill_color(173, 216, 230)  # Lyseblå
        elif self.listetype == "Fryseliste Plukk":
            self.set_fill_color(173, 216, 230)  # Samme lyseblå som Fryseliste
        elif self.listetype == "Bamaliste Plukk":
            self.set_fill_color(186, 255, 201)  # Lysegrønn
        else:
            self.set_fill_color(220, 220, 220)  # Default grå

        # Skriv header
        for i, col in enumerate(columns):
            align = 'C' if col == "Quantity" else 'L'
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
                if column_name in ["Customer", "Product Name", "Department", "Route nickname"]:
                    wrap_width = int(col_widths[i] // (font_size * 0.25))
                    wrap_width = max(wrap_width, 12)
                    lines = textwrap.wrap(text, width=wrap_width)
                    if not lines:
                        lines = [""]
                else:
                    lines = [text]
                cell_lines.append(lines)
                max_lines = max(max_lines, len(lines))
            row_height = max_lines * line_height

            # ------ NYTT: Sjekk page break FØR cellene ------
            if self.get_y() + row_height > self.page_break_trigger:
                self.add_page()
                for i, col in enumerate(columns):
                    align = 'C' if col == "Quantity" else 'L'
                    self.cell(col_widths[i], line_height * 2, str(col), border=1, align=align, fill=True)
                self.ln()
            # ------------------------------------------------

            # Tegn ramme rundt hele raden (valgfritt, men ryddig)
            x_start = self.get_x()
            y_start = self.get_y()
            for i in range(len(columns)):
                x = x_start + sum(col_widths[:i])
                y = y_start
                self.rect(x, y, col_widths[i], row_height)

            # Skriv rad linje for linje
            for line_no in range(max_lines):
                self.set_x(x_start)
                for i, lines in enumerate(cell_lines):
                    text = lines[line_no] if line_no < len(lines) else ""
                    align = 'C' if columns[i] == "Quantity" else 'L'
                    self.cell(col_widths[i], line_height, text, border=0, align=align)
                self.ln(line_height)


def format_and_sort_excel(file_path, listetype=None):
    try:
        df = pd.read_excel(file_path, sheet_name=0, engine="openpyxl")
        df_sorted = df.sort_values(by=["Customer", "Route ordering", "Route nickname"], ascending=True)
        if "Route ordering" in df_sorted.columns:
            df_sorted = df_sorted.drop(columns=["Route ordering"])
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
        df_sorted = df_sorted[[col for col in columns_to_keep if col in df_sorted.columns]]

        # Endre kolonnenavn for bedre PDF-layout
        df_sorted = df_sorted.rename(columns={"Accept alternatives": "AA"})
        df_sorted = df_sorted.rename(columns={"Route nickname": "Route"})

        return df_sorted
    except Exception as e:
        messagebox.showerror("Feil", f"Kunne ikke lese eller formatere Excel-fil: {str(e)}")
        return None

def save_to_pdf(grouped_data, output_path, listetype="Brødliste", dato="ukjent dato"):
    # Sjekk listetype
    if listetype in ["Brødliste", "Fryseliste"]:
        grouped = grouped_data
        pdf = RoutePDF(listetype=listetype, dato=dato)
        pdf.set_auto_page_break(auto=True, margin=15)
        for route, group in grouped:
            pdf.route_title = f"{listetype} til {dato} - Route: {route}"
            pdf.add_page()
            pdf.add_route_table(group)
        pdf.output(output_path)
    else:
        # Plukk-lister: EN tabell med hele dataen, ikke grupper!
        pdf = RoutePDF(listetype=listetype, dato=dato)
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

def print_via_pdf(grouped_data, listetype="Brødliste", dato="ukjent dato"):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        save_to_pdf(grouped_data, temp_pdf.name, listetype, dato)
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

    # Disse vil nå fungere korrekt da de tilhører 'gui' og 'listetype'
    listetyper = ["Brødliste", "Fryseliste", "Fryseliste Plukk", "Bamaliste Plukk"]
    listetype = tk.StringVar(value=listetyper[0])
    dropdown = tk.OptionMenu(frame_type, listetype, *listetyper)
    dropdown.pack(side=tk.LEFT, padx=5)

    label_file = tk.Label(gui, text=f"Valgt fil:\n{os.path.basename(file_path)}", fg="blue")
    label_file.pack(pady=(10, 5))
    label = tk.Label(gui, text="Velg hva du vil gjøre med denne filen:")
    label.pack(pady=5)

    # Resten av funksjonen din for knapper (on_print, on_pdf, etc.) er den samme
    # ... (on_print, on_pdf, on_excel definisjoner) ...
    # ... (Button-definisjoner) ...

    # Du må definere funksjonene som bruker 'listetype' inne i main, 
    # eller passe på å sende 'listetype' som argument hvis du definerer dem utenfor.
    # Siden de allerede er definert her, er det i orden.
    
    def on_print():
        lt = listetype.get()
        df_sorted = format_and_sort_excel(file_path, lt)
        if df_sorted is None or df_sorted.empty:
            return
        if "Route" not in df_sorted.columns:
            messagebox.showerror("Feil", "Filen mangler kolonnen 'Route'.")
            return
        if lt in ["Brødliste", "Fryseliste"]:
            grouped = df_sorted.groupby("Route")
            print_via_pdf(grouped, lt, date_part)
        else:
            print_via_pdf(df_sorted, lt, date_part)
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
            save_to_pdf(grouped, output_path, lt, date_part)
        else:
            save_to_pdf(df_sorted, output_path, lt, date_part)
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

