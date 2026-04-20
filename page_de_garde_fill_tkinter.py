"""
page_de_garde_fill_tkinter.py
------------------------------
Generates a filled CTD dossier cover page (page de garde) Word document
for a selected product, using data pulled from database.xlsx , I am using database-sample.xlsx as a sample for confidentiality reasons.

Requirements:
    pip install openpyxl

File structure expected (same folder as this script):
    database-sample.xlsx          <- product database
    page_de_garde.docx     <- Word template
    output/                <- generated documents saved here (auto-created)
"""

import os
import re
import shutil
import zipfile
import datetime

import openpyxl
import tkinter as tk
from tkinter import ttk, messagebox

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "database-sample.xlsx")
TEMPLATE   = os.path.join(BASE_DIR, "page_de_garde.docx")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
TEMP_DIR   = os.path.join(BASE_DIR, "_temp_page_de_garde")

DOSSIER_TYPES = [
    "Dossier d'enregistrement",
    "Dossier de variation",
    "Dossier de renouvellement",
]


# ---------------------------------------------------------------------------
# Data layer
# ---------------------------------------------------------------------------

def load_products(excel_path: str) -> dict:
    """
    Return {product_name: [(DCI, forme_pharma, dosage, CNDT), ...]}
    """
    wb    = openpyxl.load_workbook(excel_path)
    sheet = wb.active
    products = {}

    for row in range(2, sheet.max_row + 1):
        name        = sheet[f"A{row}"].value
        dci         = sheet[f"B{row}"].value
        forme       = sheet[f"C{row}"].value
        dosage      = sheet[f"D{row}"].value
        cndt        = sheet[f"P{row}"].value

        name   = str(name).strip()   if name   else ""
        dci    = str(dci).strip()    if dci    else ""
        forme  = str(forme).strip()  if forme  else ""
        dosage = str(dosage).strip() if dosage else ""
        cndt   = str(cndt).strip()   if cndt   else ""

        if not name:
            continue

        if name not in products:
            products[name] = []
        products[name].append((dci, forme, dosage, cndt))

    return products


# ---------------------------------------------------------------------------
# Document generation
# ---------------------------------------------------------------------------

def generate_page_de_garde(
    name: str,
    dci: str,
    forme_pharmaceutique: str,
    dosage: str,
    cndt: str,
    dossier_type: str,
) -> str:
    """
    Fill the page_de_garde.docx template with product data and save it.
    Returns the output file path.
    """
    replacements = {
        "@1":   f"{name}®",
        "@2":   dci,
        "@3":   forme_pharmaceutique,
        "@4":   dosage,
        "@+22": cndt,
        "@+23": dossier_type,
        "@+24": datetime.datetime.now().strftime("%d/%m/%Y"),
    }

    safe_name   = re.sub(r'[\\/*?:"<>|]', "_", name)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_path = os.path.join(OUTPUT_DIR, f"page_de_garde_{safe_name}.docx")

    # Unpack → replace → repack
    if os.path.exists(TEMP_DIR):
        shutil.rmtree(TEMP_DIR)

    with zipfile.ZipFile(TEMPLATE, "r") as zin:
        zin.extractall(TEMP_DIR)

    for root_dir, _dirs, files in os.walk(TEMP_DIR):
        for filename in files:
            if filename.endswith(".xml") or filename.endswith(".rels"):
                file_path = os.path.join(root_dir, filename)
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
                for placeholder, value in replacements.items():
                    content = content.replace(placeholder, str(value))
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(content)

    if os.path.exists(output_path):
        os.remove(output_path)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for root_dir, _dirs, files in os.walk(TEMP_DIR):
            for filename in files:
                file_path = os.path.join(root_dir, filename)
                arcname   = os.path.relpath(file_path, TEMP_DIR)
                zout.write(file_path, arcname)

    shutil.rmtree(TEMP_DIR)
    return output_path


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class PageDeGardeApp:

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Générateur de pages de garde")
        self.root.geometry("440x260")
        self.root.resizable(False, False)

        for path, label in [(EXCEL_PATH, "database-sample.xlsx"), (TEMPLATE, "page_de_garde.docx")]:
            if not os.path.exists(path):
                messagebox.showerror(
                    "Fichier manquant",
                    f"{label} introuvable :\n{path}\n\n"
                    "Placez le fichier dans le même dossier que ce script."
                )
                root.destroy()
                return

        self.products = load_products(EXCEL_PATH)
        self._build_ui()

    # ------------------------------------------------------------------
    def _build_ui(self):
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Produit :", anchor="w").pack(fill="x", pady=(0, 2))
        self.product_var = tk.StringVar()
        ttk.Combobox(
            frame, textvariable=self.product_var,
            values=list(self.products.keys()), state="readonly"
        ).pack(fill="x")

        ttk.Label(frame, text="Type de dossier :", anchor="w").pack(fill="x", pady=(12, 2))
        self.dossier_var = tk.StringVar()
        ttk.Combobox(
            frame, textvariable=self.dossier_var,
            values=DOSSIER_TYPES, state="readonly"
        ).pack(fill="x")

        ttk.Button(
            frame, text="Générer la page de garde", command=self._generate
        ).pack(pady=20)

    # ------------------------------------------------------------------
    def _generate(self):
        name         = self.product_var.get()
        dossier_type = self.dossier_var.get()

        if not name or not dossier_type:
            messagebox.showwarning("Attention", "Veuillez choisir un produit et un type de dossier.")
            return

        variants = self.products[name]

        if len(variants) > 1:
            self._choose_variant(name, variants, dossier_type)
        else:
            dci, forme, dosage, cndt = variants[0]
            self._do_generate(name, dci, forme, dosage, cndt, dossier_type)

    # ------------------------------------------------------------------
    def _choose_variant(self, name: str, variants: list, dossier_type: str):
        win = tk.Toplevel(self.root)
        win.title("Choisir une variante")
        win.grab_set()

        ttk.Label(
            win,
            text="Ce produit a plusieurs dosages / conditionnements.\nChoisissez-en un :"
        ).pack(pady=8)

        listbox = tk.Listbox(win, height=len(variants), width=72)
        for dci, forme, dosage, cndt in variants:
            listbox.insert(tk.END, f"DCI: {dci} | Forme: {forme} | Dosage: {dosage} | CNDT: {cndt}")
        listbox.pack(padx=10, pady=4)

        def confirm():
            idx = listbox.curselection()
            if not idx:
                messagebox.showwarning("Attention", "Veuillez sélectionner une variante.")
                return
            dci, forme, dosage, cndt = variants[idx[0]]
            win.destroy()
            self._do_generate(name, dci, forme, dosage, cndt, dossier_type)

        ttk.Button(win, text="Valider", command=confirm).pack(pady=8)

    # ------------------------------------------------------------------
    def _do_generate(self, name, dci, forme, dosage, cndt, dossier_type):
        try:
            output_path = generate_page_de_garde(name, dci, forme, dosage, cndt, dossier_type)
            messagebox.showinfo("Succès", f"Document créé :\n{output_path}")
        except Exception as exc:
            messagebox.showerror("Erreur", f"Une erreur est survenue :\n{exc}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    root = tk.Tk()
    PageDeGardeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
