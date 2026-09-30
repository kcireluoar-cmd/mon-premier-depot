# -*- coding: utf-8 -*-
"""
Extracteur de dossier -> Suivi DAO/AMI
=======================================
Charge un dossier (PDF / DOCX / TXT), extrait le texte, demande a Claude
de remplir les champs d'une ligne de suivi, laisse l'utilisateur verifier,
puis ecrit la ligne directement dans le classeur Suivi_Dossiers_DAO_AMI.xlsx.

Auteur : genere pour Raoul Erick Nguy Ngondo (ETS RENSERVICES).
"""

import os
import re
import sys
import json
import datetime
import traceback

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import openpyxl

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    import docx as docxlib  # python-docx
except ImportError:
    docxlib = None

try:
    import anthropic
except ImportError:
    anthropic = None


# --------------------------------------------------------------------------- constantes

APP_NAME = "ExtracteurDAO"
CONFIG_DIR = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), APP_NAME)
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")

MODEL_ID = "claude-sonnet-5"  # modifiable dans le fichier config.json (cle "model")

SHEET_NAME = "Suivi DAO-AMI"
TABLE_NAME = "TabSuivi"

TYPES = [
    "AOI (Appel d'Offres International)", "AONO (Appel d'Offres National Ouvert)",
    "AONR (Appel d'Offres National Restreint)", "AMI (Avis à Manifestation d'Intérêt)",
    "Consultation restreinte", "Demande de Cotation (DC)", "Gré à gré", "Autre",
]
ENTITES = ["ETS RENSERVICES", "NEW DESIGN BUILTSHINE (NDB)", "Groupement NDB/COMEXA", "FEEC SARL", "Autre"]
DEPOSE = ["Oui", "Non", "En cours de préparation"]
MODES = ["Physique", "Numérique", "Physique + Numérique", "Sans objet"]
STATUTS = [
    "En préparation", "Déposé — en attente de résultat", "Retenu / Attributaire",
    "Non retenu", "Infructueux / Sans suite", "Annulé", "Abandonné (non déposé)",
]
PERSONNES = ["Raoul NGUY NGONDO", "KENMOU DZUKAM Aurélien", "Autre"]

FIELD_KEYS = [
    "ref", "intitule", "type", "mo", "bailleur", "entite", "personne",
    "datePub", "dateLimite", "depose", "dateDepot", "mode", "montant",
    "statut", "dateNotif", "resultat", "obs", "lien",
]
DATE_KEYS = {"datePub", "dateLimite", "dateDepot", "dateNotif"}
COL_LETTERS = ["B", "C", "D", "E", "F", "G", "R", "H", "I", "J", "K", "L", "M", "N", "P", "O", "Q"]
# correspond a l'ordre des colonnes du classeur (voir README) : B..S sauf N° (A)


# --------------------------------------------------------------------------- config

def load_config():
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# --------------------------------------------------------------------------- extraction texte

def extract_text(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".txt":
        with open(path, encoding="utf-8", errors="ignore") as f:
            return f.read()
    if ext == ".pdf":
        if not PdfReader:
            raise RuntimeError("Le module pypdf n'est pas installé.")
        reader = PdfReader(path)
        parts = []
        for page in reader.pages[:30]:
            try:
                parts.append(page.extract_text() or "")
            except Exception:
                pass
        return "\n\n".join(parts)
    if ext == ".docx":
        if not docxlib:
            raise RuntimeError("Le module python-docx n'est pas installé.")
        d = docxlib.Document(path)
        return "\n".join(p.text for p in d.paragraphs)
    raise RuntimeError("Format non pris en charge : " + ext)


# --------------------------------------------------------------------------- appel Claude

def build_prompt(text):
    return (
        "Tu remplis une ligne d'un registre de suivi des appels d'offres publics camerounais "
        "(DAO, AMI, AONO, AONR, lettres-commandes, consultations). "
        "Voici le texte extrait d'un dossier :\n\n\"\"\"\n" + text + "\n\"\"\"\n\n"
        "Réponds uniquement par un objet JSON (aucun texte autour) avec exactement ces clés. "
        "Si une information n'apparaît pas clairement dans le texte, mets une chaîne vide \"\" — "
        "n'invente jamais une valeur.\n\n"
        "- reference : la référence/numéro de l'appel d'offres, du marché ou de la lettre-commande\n"
        "- intitule : l'intitulé/objet complet de la consultation\n"
        "- type : une valeur EXACTE parmi [" + " | ".join(TYPES) + "], ou \"\" si indéterminable\n"
        "- maitre_ouvrage : le maître d'ouvrage / l'autorité contractante\n"
        "- bailleur : le bailleur de fonds si mentionné, sinon \"\"\n"
        "- entite : une valeur EXACTE parmi [" + " | ".join(ENTITES) + "] si le titulaire/soumissionnaire "
        "correspond, sinon \"\"\n"
        "- date_publication : format JJ/MM/AAAA si trouvée, sinon \"\"\n"
        "- date_limite : date limite de dépôt, format JJ/MM/AAAA, sinon \"\"\n"
        "- date_depot : date de dépôt, de signature ou de soumission effective si mentionnée, "
        "format JJ/MM/AAAA, sinon \"\"\n"
        "- mode_depot : une valeur EXACTE parmi [" + " | ".join(MODES) + "] si déductible, sinon \"\"\n"
        "- montant : le montant total de l'offre ou du marché en francs CFA, chiffres uniquement sans "
        "espaces ni lettres (ex. 29991375), sinon \"\"\n"
        "- statut : une valeur EXACTE parmi [" + " | ".join(STATUTS) + "] déduite du contenu "
        "(ex. un contrat déjà signé et notifié = \"Retenu / Attributaire\"), sinon \"\"\n"
        "- resultat : un résumé court de la décision ou du résultat si mentionné, sinon \"\"\n"
        "- date_notification : date de notification du résultat si mentionnée, format JJ/MM/AAAA, sinon \"\"\n"
        "- observations : une phrase courte notant un élément utile, sinon \"\""
    )


def call_claude(api_key, text, model=MODEL_ID):
    if not anthropic:
        raise RuntimeError("Le module anthropic n'est pas installé (pip install anthropic).")
    client = anthropic.Anthropic(api_key=api_key)
    resp = client.messages.create(
        model=model,
        max_tokens=1024,
        messages=[{"role": "user", "content": build_prompt(text[:9000])}],
    )
    raw = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
    m = re.search(r"\{.*\}", raw, re.S)
    if not m:
        raise RuntimeError("Réponse de l'IA illisible : " + raw[:300])
    data = json.loads(m.group(0))
    return {
        "ref": data.get("reference", ""), "intitule": data.get("intitule", ""),
        "type": data.get("type", ""), "mo": data.get("maitre_ouvrage", ""),
        "bailleur": data.get("bailleur", ""), "entite": data.get("entite", ""),
        "datePub": data.get("date_publication", ""), "dateLimite": data.get("date_limite", ""),
        "dateDepot": data.get("date_depot", ""), "mode": data.get("mode_depot", ""),
        "montant": data.get("montant", ""), "statut": data.get("statut", ""),
        "resultat": data.get("resultat", ""), "dateNotif": data.get("date_notification", ""),
        "obs": data.get("observations", ""),
    }


# --------------------------------------------------------------------------- ecriture Excel

def parse_date_fr(s):
    s = (s or "").strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if not m:
        return s or None
    d, mo, y = (int(x) for x in m.groups())
    try:
        return datetime.date(y, mo, d)
    except ValueError:
        return s


def find_target_row(ws, last_row):
    for r in range(2, last_row + 1):
        if ws.cell(row=r, column=2).value in (None, ""):
            return r, False
    return last_row + 1, True


def write_row_to_workbook(xlsx_path, values):
    """Ecrit `values` (dict, clefs = FIELD_KEYS) dans la premiere ligne libre
    du classeur. Retourne (numero_de_ligne, overflow: bool)."""
    wb = openpyxl.load_workbook(xlsx_path)
    if SHEET_NAME not in wb.sheetnames:
        raise RuntimeError(f"Feuille « {SHEET_NAME} » introuvable dans ce classeur.")
    ws = wb[SHEET_NAME]

    last_row = None
    try:
        tbl = ws.tables.get(TABLE_NAME)
        if tbl:
            m = re.search(r":\w+?(\d+)$", tbl.ref)
            if m:
                last_row = int(m.group(1))
    except Exception:
        pass
    if last_row is None:
        last_row = ws.max_row

    row, overflow = find_target_row(ws, last_row)

    if overflow:
        # on prolonge sous le tableau : on copie la mise en forme de la derniere
        # ligne et on numerote, mais on NE touche PAS au Tableau (ref), aux
        # validations ni aux formules du Tableau de bord, pour ne jamais
        # corrompre le classeur. Un avertissement est renvoye a l'appelant.
        for col_idx in range(1, 20):
            src = ws.cell(row=last_row, column=col_idx)
            dst = ws.cell(row=row, column=col_idx)
            dst.font = src.font.copy()
            dst.border = src.border.copy()
            dst.fill = src.fill.copy()
            dst.alignment = src.alignment.copy()
            dst.number_format = src.number_format
        ws.cell(row=row, column=1, value=row - 1)

    ordered_keys = ["ref", "intitule", "type", "mo", "bailleur", "entite", "personne",
                    "datePub", "dateLimite", "depose", "dateDepot", "mode", "montant",
                    "statut", "dateNotif", "resultat", "obs"]
    for col_letter, key in zip(COL_LETTERS, ordered_keys):
        v = values.get(key, "")
        if key in DATE_KEYS:
            v = parse_date_fr(v) if v else None
        elif key == "montant":
            v = int(v) if str(v).strip().isdigit() else (v or None)
        else:
            v = v if v not in ("", None) else None
        ws[f"{col_letter}{row}"] = v
    # colonne S = lien (derniere colonne), traitee a part car hors boucle ci-dessus
    lien = values.get("lien") or None
    ws[f"S{row}"] = lien

    wb.save(xlsx_path)
    return row, overflow


# --------------------------------------------------------------------------- interface graphique

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Extracteur de dossier — Suivi DAO/AMI")
        self.geometry("880x780")
        self.minsize(760, 640)

        self.cfg = load_config()
        self.current_file = None
        self.extracted_text = ""
        self.entries = {}

        self._build_menu()
        self._build_ui()

    # ---- menu

    def _build_menu(self):
        menubar = tk.Menu(self)
        settings = tk.Menu(menubar, tearoff=0)
        settings.add_command(label="Clé API Anthropic…", command=self.change_api_key)
        settings.add_command(label="Classeur de suivi…", command=self.choose_workbook)
        menubar.add_cascade(label="Réglages", menu=settings)
        self.config(menu=menubar)

    def change_api_key(self):
        key = simpledialog.askstring(
            "Clé API Anthropic",
            "Entrez votre clé API (console.anthropic.com → API Keys) :",
            show="*", parent=self,
        )
        if key:
            self.cfg["api_key"] = key.strip()
            save_config(self.cfg)
            messagebox.showinfo("Clé enregistrée", "La clé API a été enregistrée localement.")

    def choose_workbook(self):
        path = filedialog.askopenfilename(
            title="Choisir le classeur Suivi_Dossiers_DAO_AMI.xlsx",
            filetypes=[("Classeur Excel", "*.xlsx")],
        )
        if path:
            self.cfg["workbook_path"] = path
            save_config(self.cfg)
            self.workbook_label.config(text=os.path.basename(path))

    # ---- UI

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        top = ttk.LabelFrame(self, text="1. Charger le dossier")
        top.pack(fill="x", **pad)

        row1 = ttk.Frame(top)
        row1.pack(fill="x", padx=8, pady=6)
        ttk.Button(row1, text="Choisir un fichier (PDF / DOCX / TXT)", command=self.pick_file).pack(side="left")
        self.file_label = ttk.Label(row1, text="Aucun fichier sélectionné", foreground="#666")
        self.file_label.pack(side="left", padx=12)

        self.preview = tk.Text(top, height=7, wrap="word", state="disabled")
        self.preview.pack(fill="x", padx=8, pady=(0, 6))

        row2 = ttk.Frame(top)
        row2.pack(fill="x", padx=8, pady=(0, 8))
        self.extract_btn = ttk.Button(row2, text="Extraire les informations (IA)", command=self.run_extraction, state="disabled")
        self.extract_btn.pack(side="left")
        self.status_label = ttk.Label(row2, text="", foreground="#8a6d1f")
        self.status_label.pack(side="left", padx=12)

        form = ttk.LabelFrame(self, text="2. Ligne du registre")
        form.pack(fill="both", expand=True, **pad)

        canvas = tk.Canvas(form, highlightthickness=0)
        scrollbar = ttk.Scrollbar(form, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def add_field(parent, r, label, key, kind="entry", options=None, width=42):
            ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", padx=8, pady=4)
            if kind == "combo":
                var = tk.StringVar()
                w = ttk.Combobox(parent, textvariable=var, values=[""] + options, width=width, state="readonly")
                w.grid(row=r, column=1, sticky="we", padx=8, pady=4)
                self.entries[key] = var
            elif kind == "text":
                w = tk.Text(parent, height=2, width=width)
                w.grid(row=r, column=1, sticky="we", padx=8, pady=4)
                self.entries[key] = w
            else:
                var = tk.StringVar()
                w = ttk.Entry(parent, textvariable=var, width=width)
                w.grid(row=r, column=1, sticky="we", padx=8, pady=4)
                self.entries[key] = var

        inner.columnconfigure(1, weight=1)
        r = 0
        add_field(inner, r, "Référence AO/DAO/AMI", "ref"); r += 1
        add_field(inner, r, "Intitulé", "intitule", kind="text"); r += 1
        add_field(inner, r, "Type d'appel", "type", kind="combo", options=TYPES); r += 1
        add_field(inner, r, "Maître d'ouvrage", "mo"); r += 1
        add_field(inner, r, "Bailleur / Financement", "bailleur"); r += 1
        add_field(inner, r, "Entité soumissionnaire", "entite", kind="combo", options=ENTITES); r += 1
        add_field(inner, r, "Personne en charge", "personne", kind="combo", options=PERSONNES); r += 1
        add_field(inner, r, "Date de publication (JJ/MM/AAAA)", "datePub"); r += 1
        add_field(inner, r, "Date limite de dépôt (JJ/MM/AAAA)", "dateLimite"); r += 1
        add_field(inner, r, "Déposé ?", "depose", kind="combo", options=DEPOSE); r += 1
        add_field(inner, r, "Date de dépôt effective (JJ/MM/AAAA)", "dateDepot"); r += 1
        add_field(inner, r, "Mode de dépôt", "mode", kind="combo", options=MODES); r += 1
        add_field(inner, r, "Montant de l'offre (FCFA)", "montant"); r += 1
        add_field(inner, r, "Statut du dossier", "statut", kind="combo", options=STATUTS); r += 1
        add_field(inner, r, "Date de notification (JJ/MM/AAAA)", "dateNotif"); r += 1
        add_field(inner, r, "Résultat / Décision", "resultat", kind="text"); r += 1
        add_field(inner, r, "Observations / Motif", "obs", kind="text"); r += 1
        add_field(inner, r, "Lien / emplacement du dossier", "lien"); r += 1

        bottom = ttk.Frame(self)
        bottom.pack(fill="x", **pad)
        self.workbook_label = ttk.Label(
            bottom,
            text=os.path.basename(self.cfg.get("workbook_path", "")) or "Aucun classeur choisi",
            foreground="#666",
        )
        ttk.Button(bottom, text="Classeur…", command=self.choose_workbook).pack(side="left")
        self.workbook_label.pack(side="left", padx=8)
        ttk.Button(bottom, text="Ajouter cette ligne au classeur", command=self.add_to_workbook).pack(side="right")

        self.footer = ttk.Label(
            self,
            text="Un champ vide n'a pas été trouvé dans le document — rien n'est jamais deviné. Vérifiez avant d'ajouter.",
            foreground="#888", wraplength=840,
        )
        self.footer.pack(fill="x", padx=12, pady=(0, 10))

    # ---- actions

    def pick_file(self):
        path = filedialog.askopenfilename(
            title="Choisir un dossier",
            filetypes=[("Dossiers pris en charge", "*.pdf *.docx *.txt")],
        )
        if not path:
            return
        self.current_file = path
        self.file_label.config(text=os.path.basename(path))
        self.status_label.config(text="Lecture du fichier…")
        self.update_idletasks()
        try:
            self.extracted_text = extract_text(path).strip()
        except Exception as e:
            messagebox.showerror("Erreur de lecture", str(e))
            self.status_label.config(text="")
            return
        self.preview.config(state="normal")
        self.preview.delete("1.0", "end")
        if self.extracted_text:
            self.preview.insert("1.0", self.extracted_text[:3000])
            self.extract_btn.config(state="normal")
            self.status_label.config(text=f"{len(self.extracted_text)} caractères extraits.")
        else:
            self.preview.insert("1.0", "(aucun texte lisible — document scanné ? remplissez la ligne manuellement)")
            self.extract_btn.config(state="disabled")
            self.status_label.config(text="")
        self.preview.config(state="disabled")

    def run_extraction(self):
        if not self.extracted_text:
            return
        api_key = os.environ.get("ANTHROPIC_API_KEY") or self.cfg.get("api_key")
        if not api_key:
            self.change_api_key()
            api_key = self.cfg.get("api_key")
            if not api_key:
                return
        self.extract_btn.config(state="disabled")
        self.status_label.config(text="Analyse en cours…")
        self.update_idletasks()
        try:
            model = self.cfg.get("model", MODEL_ID)
            data = call_claude(api_key, self.extracted_text, model=model)
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Échec de l'analyse", str(e))
            self.status_label.config(text="")
            self.extract_btn.config(state="normal")
            return
        data["lien"] = os.path.basename(self.current_file) if self.current_file else ""
        self.apply_to_form(data)
        self.status_label.config(text="Informations extraites — vérifiez la ligne ci-dessous.")
        self.extract_btn.config(state="normal")

    def apply_to_form(self, data):
        for key, widget in self.entries.items():
            val = data.get(key, "")
            if isinstance(widget, tk.Text):
                widget.delete("1.0", "end")
                widget.insert("1.0", val or "")
            else:
                widget.set(val or "")

    def read_form(self):
        out = {}
        for key, widget in self.entries.items():
            if isinstance(widget, tk.Text):
                out[key] = widget.get("1.0", "end").strip()
            else:
                out[key] = widget.get().strip()
        return out

    def add_to_workbook(self):
        wb_path = self.cfg.get("workbook_path")
        if not wb_path or not os.path.exists(wb_path):
            messagebox.showwarning("Classeur manquant", "Choisissez d'abord votre classeur Suivi_Dossiers_DAO_AMI.xlsx (menu Réglages ou bouton « Classeur… »).")
            self.choose_workbook()
            return
        values = self.read_form()
        if not values.get("ref") and not values.get("intitule"):
            messagebox.showwarning("Ligne incomplète", "Renseignez au moins la référence ou l'intitulé avant d'ajouter la ligne.")
            return
        try:
            row, overflow = write_row_to_workbook(wb_path, values)
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Échec de l'écriture", str(e))
            return
        if overflow:
            messagebox.showinfo(
                "Ligne ajoutée",
                f"Ligne ajoutée en ligne {row}, au-delà du tableau d'origine (plus de ligne vierge disponible).\n\n"
                "Pensez à réétendre le tableau (Ctrl+T dans Excel) et à vérifier les formules du "
                "Tableau de bord, qui ne couvrent pas encore cette ligne automatiquement.",
            )
        else:
            messagebox.showinfo("Ligne ajoutée", f"Ligne ajoutée avec succès (ligne {row}) dans « {SHEET_NAME} ».")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
