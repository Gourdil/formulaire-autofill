# formulaire-autofill

**Python desktop tools for automated regulatory document generation — Word + Excel**

---

## Context

At **Laboratoires BEKER** (pharmaceutical industry, Algeria), the regulatory affairs department managed 180+ Manufacturing Licences simultaneously. Every variation submission to the national authority (ANPP) required a precisely completed Word form — the *Formulaire de Demande de Modification* — containing product-specific regulatory data pulled from multiple sources.

Filling each form manually took ~30 minutes. With dozens of products in active submission cycles, this was a recurring bottleneck and a source of transcription errors.

This project automates that workflow entirely.

---

## What it does

Three standalone desktop applications, each built with **Python + Tkinter** and packaged as a `.exe`:

| Script | What it generates |
|---|---|
| `form_fill_tkinter.py` | *Formulaire de Demande de Modification* — mandatory form for every variation submission |
| `page_de_garde_fill_tkinter.py` | CTD dossier cover page (*page de garde*) — bilingual FR/EN |
| `etiquette_fill_tkinter.py` | Archive box label (*étiquette*) for physical dossier filing |

**How it works:**
1. User selects a product from a dropdown (populated from `database.xlsx`)
2. App auto-fills all product-specific fields in the Word template
3. User fills only the submission-specific fields (modification type, payment references, etc.)
4. One click → complete, ready-to-print Word document saved to `/output`

**Result:** ~30 minutes of manual work → under 1 minute. Zero transcription errors.

---

## File structure

```
formulaire-autofill/
│
├── form_fill_tkinter.py              ← Main app (modification request form)
├── page_de_garde_fill_tkinter.py     ← Cover page generator
├── etiquette_fill_tkinter.py         ← Archive label generator
│
├── database.xlsx                     ← Product database (not included — see below)
├── formulaire_de_demande_de_Modification.docx  ← Word template (not included)
├── page_de_garde.docx                ← Word template (not included)
├── boite_d_archive.docx              ← Word template (not included)
│
└── output/                           ← Generated documents saved here (auto-created)
```

> **Note:** The Word templates and `database.xlsx` contain proprietary company data and are not included in this repository. To use these tools, provide your own templates and database following the placeholder convention described below.

---

## Placeholder convention

The Word templates use `@` codes as placeholders. The apps replace them at runtime with values from the database:

| Placeholder | Field |
|---|---|
| `@1` | Nom commercial |
| `@2` | DCI (INN) |
| `@3` | Forme pharmaceutique |
| `@4` | Dosage |
| `@5` | Type de conditionnement |
| `@+12` | Code ATC |
| `@+13` | Classe pharmaco-thérapeutique |
| `@+14` | Indication(s) thérapeutique(s) |
| `@+16` | Raison sociale du fabricant de PA |
| `@+17` | Adresse du fabricant de PA |
| `@+20` | Numéro d'ordre de la DE |
| `@+21` | Numéro d'identification administrative |
| `@+22` | CNDT |
| `@+23` | Type de dossier |
| `@+24` | Date du jour (auto-filled) |

The manually entered fields (`@7`, `@8`, `@9`, `@+10`, `@+11`, `@+15`, `@+18`, `@+19`) are filled by the user in the GUI at runtime.

---

## Database structure (`database.xlsx`)

The app reads from the **active sheet**, starting at row 2. Expected columns:

| Column | Field |
|---|---|
| A | Nom commercial |
| B | DCI |
| C | Forme pharmaceutique |
| D | Dosage |
| E | Type de conditionnement |
| F | Prémix ou non |
| G | Code ATC |
| H | Classe pharmaco-thérapeutique |
| I | Indication(s) thérapeutique(s) |
| J | Taux d'intégration |
| K | Raison sociale fabricant PA |
| L | Adresse fabricant PA |
| M | Adresse fabricant prémix |
| N | Numéro d'ordre DE |
| O | Numéro d'identification administrative |
| P | CNDT |

---

## Installation

```bash
# Clone the repository
git clone https://github.com/Gourdil/formulaire-autofill.git
cd formulaire-autofill

# Install dependencies
pip install openpyxl python-docx
```

Then place your `database.xlsx` and Word templates in the same folder as the scripts.

---

## Run

```bash
python form_fill_tkinter.py
python page_de_garde_fill_tkinter.py
python etiquette_fill_tkinter.py
```

Each script is fully self-contained and independent.

---

## Requirements

- Python 3.10+
- `openpyxl`
- `python-docx` (required only by `form_fill_tkinter.py`)
- Tkinter (included with standard Python on Windows and macOS)

---

## Author

**Nadir Belkessam** — Regulatory Affairs Specialist | Word & Excel Automation  
Ottawa, ON, Canada  
[LinkedIn](https://linkedin.com/in/nadir-belkessam-491566a4) · [Fiverr](https://fiverr.com)

---

## License

Apache License 2.0 — see [LICENSE](LICENSE) for details.
