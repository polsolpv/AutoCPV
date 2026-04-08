import datetime as dt
import html
import json
import os
import re
import threading
import tkinter as tk
import urllib.parse
import urllib.request
import webbrowser
from io import BytesIO
from dataclasses import asdict, dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import openpyxl
from PIL import Image, ImageTk
from pypdf import PdfReader, PdfWriter


APP_NAME = "AutoCPV"
APP_VERSION = "1.2"
DEFAULT_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSe4bZ7PPgQK66LOBgXMc5gCG11p02ueQXB9glD2i4mvivJtXQ/viewform"
SESSION_FILETYPES = [("AutoCPV Session", "*.autocpv.json"), ("JSON", "*.json")]
DEFAULT_FACEBOOK_SEARCH = "https://www.facebook.com/search/top?q="
LOGO_PATH = Path(r"C:\Users\solso\Documents\New project\assets\logo-trimmed.png")

FIELD_LABELS = {
    "localitat": "Localitat",
    "data": "Data",
    "categoria": "Categoria",
    "altres": "En cas d'altres, quina?",
    "nom": "Nom de l'activitat",
    "companyia": "Companyia, artista",
    "lloc": "Lloc",
    "llengua": "Llengua de l'activitat",
    "preu": "Preu",
    "regidoria": "Regidoria organitzadora",
    "publicitat": "Llengua de la publicitat",
    "font": "Font",
    "persona": "Persona que ha introduit les dades",
}

EXPECTED_HEADERS = {
    "Localitat": "localitat",
    "Data": "data",
    "Categoria": "categoria",
    "En cas d'altres, quina?": "altres",
    "Nom de l'activitat": "nom",
    "Companyia, artista": "companyia",
    "Lloc": "lloc",
    "Llengua de l'activitat": "llengua",
    "Preu": "preu",
    "Regidoria organitzadora": "regidoria",
    "Llengua de la publicitat": "publicitat",
    "Llengua de la publicitat ⚠️": "publicitat",
    "Font": "font",
    "Persona que ha introduit les dades": "persona",
    "Persona que ha introduït les dades": "persona",
}

FALLBACK_FIELD_IDS = {
    "localitat": "2074597570",
    "data": "1851780647",
    "categoria": "25855538",
    "altres": "30622121",
    "nom": "757267152",
    "companyia": "799207947",
    "lloc": "319720355",
    "llengua": "2015577411",
    "preu": "1027978183",
    "regidoria": "936300587",
    "publicitat": "744125401",
    "font": "54121299",
    "persona": "1706428244",
}

PERSON_OPTIONS = ["Arantxa", "Erika", "Maria", "Maria R.", "Pol"]
CATEGORIA_OPTIONS = [
    "Teatre",
    "Cinema",
    "Música",
    "Exposició",
    "Presentació llibre",
    "Lectura poemes",
    "Monòlegs",
    "Contacontes",
    "Taller de teatre",
    "Taller literari",
    "Dansa",
    "Taller/curs de dansa",
    "Activitats sobre patrimoni",
    "Lectura en veu alta",
    "Club de lectura",
    "Conferència",
    "Altres",
]
LLENGUA_ACTIVITY_OPTIONS = [
    "Valencià/català",
    "Espanyol",
    "Bilingüe",
    "Anglès",
    "Altres",
    "No hi ha informació",
]
REGIDORIA_OPTIONS = [
    "Cultura",
    "Joventut",
    "Igualtat",
    "Altres",
    "En col·laboració o organitzat per una entitat ciutadana",
]
PUBLICITAT_OPTIONS = [
    "Valencià/català",
    "Espanyol",
    "Bilingüe",
    "Anglès",
    "Altres",
]

FILTER_OPTIONS = {
    "Totes": "all",
    "Pendents": "pending",
    "Enviades": "sent",
    "Amb error": "error",
    "Per revisar": "invalid",
}

REQUIRED_FIELDS = [
    "localitat",
    "data",
    "categoria",
    "nom",
    "llengua",
    "regidoria",
    "publicitat",
    "font",
    "persona",
]

FIELD_OPTIONS = {
    "categoria": CATEGORIA_OPTIONS,
    "llengua": LLENGUA_ACTIVITY_OPTIONS,
    "regidoria": REGIDORIA_OPTIONS,
    "publicitat": PUBLICITAT_OPTIONS,
    "persona": PERSON_OPTIONS,
}

LONG_TEXT_FIELDS = {"nom", "companyia", "lloc", "font", "altres"}

COLORS = {
    "red": "#C8102E",
    "red_dark": "#9E0E24",
    "cream": "#F6F0E8",
    "paper": "#FFFDF9",
    "charcoal": "#1F1F1F",
    "muted": "#655D57",
    "line": "#E5D8CB",
    "rose": "#F7E4E8",
    "ok": "#1C7C54",
    "warning": "#C4831B",
    "warning_bg": "#FFF3D6",
    "error_bg": "#F9E2E4",
    "sent_bg": "#E5F4EC",
}


@dataclass
class Record:
    localitat: str = ""
    data: str = ""
    categoria: str = ""
    altres: str = ""
    nom: str = ""
    companyia: str = ""
    lloc: str = ""
    llengua: str = ""
    preu: str = ""
    regidoria: str = ""
    publicitat: str = ""
    font: str = ""
    persona: str = "Pol"
    status: str = "Pendent"
    status_detail: str = ""


def normalize_label(text: str) -> str:
    text = text.replace("\xa0", " ").strip()
    return re.sub(r"\s+", " ", text)


def find_header_row(worksheet):
    for row_idx in range(1, min(worksheet.max_row, 30) + 1):
        headers = [normalize_label(str(cell.value or "")) for cell in worksheet[row_idx]]
        matches = sum(1 for item in headers if item in EXPECTED_HEADERS)
        if matches >= 5:
            return row_idx, headers
    raise ValueError("No he trobat una fila de capçaleres recognoscible en l'Excel.")


def excel_date_to_iso(value) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    return text


def normalize_price(value) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, (int, float)):
        return str(int(value)) if float(value).is_integer() else str(value)
    text = str(value).strip().replace("€", "").replace(",", ".")
    return re.sub(r"\s+", "", text)


def is_valid_date(value: str) -> bool:
    return bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", value or ""))


def is_valid_price(value: str) -> bool:
    return bool(re.fullmatch(r"\d+(?:\.\d+)?", value or ""))


def load_excel_records(path: str, default_person: str, fallback_font: str):
    workbook = openpyxl.load_workbook(path, data_only=True)
    worksheet = workbook[workbook.sheetnames[0]]
    header_row, headers = find_header_row(worksheet)

    mapping = {}
    for idx, header in enumerate(headers):
        key = EXPECTED_HEADERS.get(header)
        if key:
            mapping[key] = idx

    records = []
    for row in worksheet.iter_rows(min_row=header_row + 1, values_only=True):
        if not any(value not in (None, "") for value in row):
            continue
        record = Record(persona=default_person.strip() or "Pol")
        for key, col_idx in mapping.items():
            if col_idx >= len(row):
                continue
            value = row[col_idx]
            if key == "data":
                setattr(record, key, excel_date_to_iso(value))
            elif key == "preu":
                setattr(record, key, normalize_price(value))
            else:
                setattr(record, key, "" if value is None else str(value).strip())
        if not record.font:
            record.font = fallback_font.strip()
        if not record.persona:
            record.persona = default_person.strip() or "Pol"
        records.append(record)
    return records


def extract_form_metadata(form_url: str):
    view_url = form_url.strip()
    if "formResponse" in view_url:
        view_url = view_url.replace("formResponse", "viewform")
    response_url = view_url.replace("viewform", "formResponse")

    request = urllib.request.Request(view_url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(request, timeout=30) as response:
        html_text = response.read().decode("utf-8", errors="ignore")

    field_ids = FALLBACK_FIELD_IDS.copy()
    for key, label in FIELD_LABELS.items():
        pattern = re.compile(rf'\[\d+,"{re.escape(label)}".*?\[\[(\d+),', re.DOTALL)
        match = pattern.search(html_text)
        if match:
            field_ids[key] = match.group(1)

    form_action_match = re.search(r'<form[^>]*action="([^"]+)"', html_text)
    if form_action_match:
        response_url = html.unescape(form_action_match.group(1))

    fbzx_match = re.search(r'name="fbzx"\s+value="([^"]+)"', html_text)
    partial_match = re.search(r'name="partialResponse"\s+value="([^"]+)"', html_text)
    page_history_match = re.search(r'name="pageHistory"\s+value="([^"]+)"', html_text)
    fvv_match = re.search(r'name="fvv"\s+value="([^"]+)"', html_text)

    return {
        "response_url": response_url,
        "fbzx": html.unescape(fbzx_match.group(1)) if fbzx_match else "",
        "partialResponse": html.unescape(partial_match.group(1)) if partial_match else "",
        "pageHistory": html.unescape(page_history_match.group(1)) if page_history_match else "0",
        "fvv": html.unescape(fvv_match.group(1)) if fvv_match else "1",
        "field_ids": field_ids,
    }


def build_payload(record: Record, metadata):
    field_ids = metadata["field_ids"]
    payload = {
        f'entry.{field_ids["localitat"]}': record.localitat,
        f'entry.{field_ids["data"]}_year': record.data[0:4],
        f'entry.{field_ids["data"]}_month': str(int(record.data[5:7])),
        f'entry.{field_ids["data"]}_day': str(int(record.data[8:10])),
        f'entry.{field_ids["categoria"]}': record.categoria,
        f'entry.{field_ids["nom"]}': record.nom,
        f'entry.{field_ids["companyia"]}': record.companyia,
        f'entry.{field_ids["lloc"]}': record.lloc,
        f'entry.{field_ids["llengua"]}': record.llengua,
        f'entry.{field_ids["regidoria"]}': record.regidoria,
        f'entry.{field_ids["publicitat"]}': record.publicitat,
        f'entry.{field_ids["font"]}': record.font,
        f'entry.{field_ids["persona"]}': record.persona,
        "fvv": metadata["fvv"] or "1",
        "pageHistory": metadata["pageHistory"] or "0",
        "fbzx": metadata["fbzx"],
        "submissionTimestamp": "-1",
    }
    if metadata["partialResponse"]:
        payload["partialResponse"] = metadata["partialResponse"]
    if record.altres:
        payload[f'entry.{field_ids["altres"]}'] = record.altres
    if record.preu != "":
        payload[f'entry.{field_ids["preu"]}'] = record.preu
    return payload


class FormFillerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("1540x980")
        self.root.minsize(1320, 780)
        self.root.configure(bg=COLORS["cream"])

        self.records = []
        self.form_metadata = None
        self.current_index = None
        self.visible_indices = []
        self.worker = None
        self.dirty = False
        self.suspend_dirty = False
        self.autosave_after_id = None
        self.logo_image = None
        self.last_deleted = None
        self.history_entries = []
        self.editor_widgets = {}
        self.help_popup = None

        self.excel_path_var = tk.StringVar()
        self.form_url_var = tk.StringVar(value=DEFAULT_FORM_URL)
        self.person_var = tk.StringVar(value="Pol")
        self.fallback_font_var = tk.StringVar(value="")
        self.status_filter_var = tk.StringVar(value="Totes")
        self.status_var = tk.StringVar(value="A punt.")
        self.validation_var = tk.StringVar(value="Sense validacions pendents.")
        self.editor_vars = {key: tk.StringVar() for key in FIELD_LABELS}
        self.split_pdf_path_var = tk.StringVar()
        self.split_output_dir_var = tk.StringVar()
        self.split_status_var = tk.StringVar(value="A punt per a dividir PDFs OCR.")
        self.split_summary_var = tk.StringVar(value="Encara no hi ha cap PDF carregat.")
        self.split_progress_var = tk.DoubleVar(value=0.0)
        self.split_result_cards = []

        self.configure_styles()
        self.build_ui()
        self.bind_editor_events()
        self.bind_shortcuts()
        self.show_splash()

    def configure_styles(self):
        style = ttk.Style(self.root)
        if "vista" in style.theme_names():
            style.theme_use("vista")

        style.configure("Root.TFrame", background=COLORS["cream"])
        style.configure("Card.TFrame", background=COLORS["paper"], relief="flat")
        style.configure("Header.TFrame", background=COLORS["red"])
        style.configure("HeaderTitle.TLabel", background=COLORS["red"], foreground="white", font=("Segoe UI Semibold", 22))
        style.configure("HeaderSub.TLabel", background=COLORS["red"], foreground="white", font=("Segoe UI", 10))
        style.configure("Section.TLabel", background=COLORS["paper"], foreground=COLORS["charcoal"], font=("Segoe UI Semibold", 10))
        style.configure("Body.TLabel", background=COLORS["paper"], foreground=COLORS["charcoal"], font=("Segoe UI", 10))
        style.configure("Status.TLabel", background=COLORS["cream"], foreground=COLORS["charcoal"], font=("Segoe UI Semibold", 10))
        style.configure("Neutral.TButton", font=("Segoe UI Semibold", 10))
        style.configure("Treeview", rowheight=28, font=("Segoe UI", 10), fieldbackground="white", background="white", foreground=COLORS["charcoal"])
        style.configure("Treeview.Heading", font=("Segoe UI Semibold", 10))
        style.map("Treeview", background=[("selected", COLORS["red"])], foreground=[("selected", "white")])

    def build_ui(self):
        header = ttk.Frame(self.root, style="Header.TFrame", padding=(20, 18))
        header.pack(fill="x")
        header_left = ttk.Frame(header, style="Header.TFrame")
        header_left.pack(side="left", fill="x", expand=True)
        header_button_row = ttk.Frame(header_left, style="Header.TFrame")
        header_button_row.pack(anchor="w", pady=(0, 10))
        self.make_help_button(header_button_row, "Tecles ràpides", self.show_shortcuts_help, "Mostra la llista d'accessos ràpids.").pack(side="left")
        ttk.Label(header_left, text=APP_NAME, style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(header_left, text=f"Versió {APP_VERSION}", style="HeaderSub.TLabel").pack(anchor="w", pady=(2, 0))

        notebook_wrap = ttk.Frame(self.root, style="Root.TFrame", padding=(16, 14, 16, 8))
        notebook_wrap.pack(fill="both", expand=True)
        self.notebook = ttk.Notebook(notebook_wrap)
        self.notebook.pack(fill="both", expand=True)
        self.main_tab = ttk.Frame(self.notebook, style="Root.TFrame")
        self.splitter_tab = ttk.Frame(self.notebook, style="Root.TFrame")
        self.notebook.add(self.main_tab, text="Formularis")
        self.notebook.add(self.splitter_tab, text="Divisor PDF OCR")

        top = ttk.Frame(self.main_tab, style="Root.TFrame", padding=(0, 0, 0, 8))
        top.pack(fill="x")
        top_card = ttk.Frame(top, style="Card.TFrame", padding=14)
        top_card.pack(fill="x")

        ttk.Label(top_card, text="Excel", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(top_card, textvariable=self.excel_path_var, width=92).grid(row=0, column=1, sticky="ew", padx=8, pady=4)
        self.make_help_button(top_card, "Buscar", self.pick_excel, "Selecciona un fitxer Excel del teu ordinador.").grid(row=0, column=2, padx=4)
        self.make_help_button(top_card, "Carregar Excel", self.load_excel, "Llig l'Excel i carrega les files a la taula.").grid(row=0, column=3, padx=4)
        self.make_help_button(top_card, "Guardar sessió", self.save_session, "Guarda l'estat actual per continuar després.").grid(row=0, column=4, padx=4)
        self.make_help_button(top_card, "Obrir sessió", self.load_session, "Recupera una sessió guardada anteriorment.").grid(row=0, column=5, padx=4)

        ttk.Label(top_card, text="Formulari", style="Section.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Entry(top_card, textvariable=self.form_url_var, width=92).grid(row=1, column=1, sticky="ew", padx=8, pady=4)
        self.make_help_button(top_card, "Llegir formulari", self.load_form, "Detecta els camps del Google Form automàticament.").grid(row=1, column=2, padx=4)
        self.make_help_button(top_card, "Previsualitzar enviament", self.preview_current_payload, "Mostra exactament què s'enviarà al formulari.").grid(row=1, column=3, padx=4)
        self.make_help_button(top_card, "Revisió ampla", self.open_review_mode, "Obri un editor més gran per revisar la fila.").grid(row=1, column=4, padx=4)

        ttk.Label(top_card, text="Persona", style="Section.TLabel").grid(row=2, column=0, sticky="w")
        ttk.Combobox(top_card, textvariable=self.person_var, values=PERSON_OPTIONS, width=16, state="readonly").grid(row=2, column=1, sticky="w", padx=8, pady=4)
        ttk.Label(top_card, text="Font per defecte", style="Section.TLabel").grid(row=2, column=1, sticky="e")
        ttk.Entry(top_card, textvariable=self.fallback_font_var, width=44).grid(row=2, column=2, sticky="ew", padx=8, pady=4)
        self.make_help_button(top_card, "Aplicar a totes", self.apply_fallback_font_to_all, "Copia la font per defecte a totes les files.").grid(row=2, column=3, padx=4)
        self.make_help_button(top_card, "Aplicar només a buides", self.apply_fallback_font_to_empty, "Ompli només les files que no tenen font.").grid(row=2, column=4, padx=4)
        top_card.columnconfigure(1, weight=1)

        controls = ttk.Frame(self.main_tab, style="Root.TFrame", padding=(0, 0, 0, 10))
        controls.pack(fill="x")
        controls_card = ttk.Frame(controls, style="Card.TFrame", padding=10)
        controls_card.pack(fill="x")
        self.make_help_button(controls_card, "Aplicar canvis", self.apply_current_record, "Guarda manualment els canvis de la fila actual.").pack(side="left", padx=4)
        self.make_help_button(controls_card, "Eliminar fila", self.delete_current_record, "Esborra completament la fila seleccionada.").pack(side="left", padx=4)
        self.make_help_button(controls_card, "Desfer eliminació", self.undo_delete_record, "Recupera l'última fila eliminada.").pack(side="left", padx=4)
        self.make_help_button(controls_card, "Buscar a Google", self.open_google_search, "Busca l'activitat actual a Google.").pack(side="left", padx=4)
        self.make_help_button(controls_card, "Buscar font", self.open_source_helper, "Ajuda a trobar una font si encara no en tens.").pack(side="left", padx=4)
        self.make_help_button(controls_card, "Obrir font", self.open_source, "Obri l'enllaç de la font guardada.").pack(side="left", padx=4)
        self.make_help_button(controls_card, "Enviar fila seleccionada", self.submit_selected, "Envia només la fila actual al formulari.").pack(side="right", padx=4)
        self.make_help_button(controls_card, "Enviar-les totes", self.submit_all, "Envia totes les files preparades al formulari.").pack(side="right", padx=4)

        body = ttk.PanedWindow(self.main_tab, orient="horizontal")
        body.pack(fill="both", expand=True, pady=(0, 10))

        left = ttk.Frame(body, style="Card.TFrame", padding=10)
        right = ttk.Frame(body, style="Card.TFrame", padding=16)
        body.add(left, weight=3)
        body.add(right, weight=2)

        filter_bar = ttk.Frame(left, style="Card.TFrame")
        filter_bar.pack(fill="x", pady=(0, 8))
        ttk.Label(filter_bar, text="Registres carregats", style="Section.TLabel").pack(side="left")
        ttk.Label(filter_bar, text="Filtre", style="Section.TLabel").pack(side="right", padx=(8, 4))
        filter_combo = ttk.Combobox(filter_bar, textvariable=self.status_filter_var, values=list(FILTER_OPTIONS), width=18, state="readonly")
        filter_combo.pack(side="right")
        filter_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_tree())

        columns = ("localitat", "data", "categoria", "nom", "preu", "font", "status")
        self.tree = ttk.Treeview(left, columns=columns, show="headings", height=26)
        headings = {
            "localitat": "Localitat",
            "data": "Data",
            "categoria": "Categoria",
            "nom": "Activitat",
            "preu": "Preu",
            "font": "Font",
            "status": "Estat",
        }
        widths = {
            "localitat": 110,
            "data": 100,
            "categoria": 150,
            "nom": 360,
            "preu": 80,
            "font": 260,
            "status": 120,
        }
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="w")
        self.tree.tag_configure("sent", background=COLORS["sent_bg"])
        self.tree.tag_configure("error", background=COLORS["error_bg"])
        self.tree.tag_configure("invalid", background=COLORS["warning_bg"])
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        scrollbar = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        ttk.Label(right, text="Editor de fila", style="Section.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        validation_label = tk.Label(
            right,
            textvariable=self.validation_var,
            bg=COLORS["warning_bg"],
            fg=COLORS["charcoal"],
            justify="left",
            anchor="w",
            wraplength=430,
            padx=10,
            pady=8,
        )
        validation_label.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        row_index = 2
        for key, label in FIELD_LABELS.items():
            ttk.Label(right, text=label, style="Body.TLabel").grid(row=row_index, column=0, sticky="nw", pady=5)
            widget = self.build_editor_widget(right, key)
            widget.grid(row=row_index, column=1, sticky="ew", pady=5)
            self.editor_widgets[key] = widget
            row_index += 1

        ttk.Label(right, text="Historial d'enviaments", style="Section.TLabel").grid(row=row_index, column=0, columnspan=2, sticky="w", pady=(12, 6))
        row_index += 1
        self.history_box = tk.Text(
            right,
            height=10,
            wrap="word",
            font=("Consolas", 10),
            bg="white",
            fg=COLORS["charcoal"],
            relief="flat",
            padx=10,
            pady=10,
        )
        self.history_box.grid(row=row_index, column=0, columnspan=2, sticky="nsew")
        self.history_box.configure(state="disabled")

        right.columnconfigure(1, weight=1)
        right.rowconfigure(row_index, weight=1)

        self.build_splitter_tab()

        status = ttk.Label(self.root, textvariable=self.status_var, style="Status.TLabel", padding=(16, 0, 16, 12))
        status.pack(fill="x")

    def build_editor_widget(self, parent, key):
        if key in FIELD_OPTIONS:
            return ttk.Combobox(parent, textvariable=self.editor_vars[key], values=FIELD_OPTIONS[key], state="readonly", width=48)
        width = 60 if key in {"nom", "companyia", "lloc", "font"} else 50
        return ttk.Entry(parent, textvariable=self.editor_vars[key], width=width)

    def make_help_button(self, parent, text, command, help_text):
        button = ttk.Button(parent, text=text, command=command, style="Neutral.TButton")
        button.bind("<Button-3>", lambda event, msg=help_text: self.show_button_help(event, msg))
        return button

    def build_splitter_tab(self):
        outer = ttk.Frame(self.splitter_tab, style="Root.TFrame", padding=(0, 0, 0, 10))
        outer.pack(fill="both", expand=True)

        hero_card = ttk.Frame(outer, style="Card.TFrame", padding=16)
        hero_card.pack(fill="x", pady=(0, 10))
        ttk.Label(hero_card, text="Divisor PDF OCR", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            hero_card,
            text="Divideix PDFs grans en parts de fins a 24 MB per poder-los compartir o pujar més fàcilment.",
            style="Body.TLabel",
            wraplength=1100,
            justify="left",
        ).pack(anchor="w", pady=(6, 0))

        top_card = ttk.Frame(outer, style="Card.TFrame", padding=16)
        top_card.pack(fill="x")

        ttk.Label(top_card, text="PDF OCR", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(top_card, textvariable=self.split_pdf_path_var, width=92).grid(row=0, column=1, sticky="ew", padx=8, pady=4)
        ttk.Button(top_card, text="Buscar PDF", command=self.pick_split_pdf, style="Neutral.TButton").grid(row=0, column=2, padx=4)

        ttk.Label(top_card, text="Carpeta d'eixida", style="Section.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Entry(top_card, textvariable=self.split_output_dir_var, width=92).grid(row=1, column=1, sticky="ew", padx=8, pady=4)
        ttk.Button(top_card, text="Buscar carpeta", command=self.pick_split_output_dir, style="Neutral.TButton").grid(row=1, column=2, padx=4)
        ttk.Button(top_card, text="Dividir a 24 MB", command=self.run_split_pdf, style="Neutral.TButton").grid(row=1, column=3, padx=4)
        top_card.columnconfigure(1, weight=1)

        drop_card = ttk.Frame(outer, style="Card.TFrame", padding=16)
        drop_card.pack(fill="x", pady=(10, 0))
        self.drop_zone = tk.Label(
            drop_card,
            text="Arrossega ací un PDF OCR o fes clic per a seleccionar-lo",
            bg=COLORS["rose"],
            fg=COLORS["charcoal"],
            font=("Segoe UI Semibold", 11),
            padx=20,
            pady=20,
            relief="flat",
            cursor="hand2",
        )
        self.drop_zone.pack(fill="x")
        self.drop_zone.bind("<Button-1>", lambda _event: self.pick_split_pdf())
        self.enable_drop_support()

        summary_card = ttk.Frame(outer, style="Card.TFrame", padding=16)
        summary_card.pack(fill="x", pady=(10, 0))
        ttk.Label(summary_card, text="Resum", style="Section.TLabel").pack(anchor="w")
        ttk.Label(summary_card, textvariable=self.split_summary_var, style="Body.TLabel", wraplength=1100, justify="left").pack(anchor="w", pady=(6, 10))
        self.split_progress = ttk.Progressbar(summary_card, maximum=100, variable=self.split_progress_var)
        self.split_progress.pack(fill="x")
        self.open_output_button = ttk.Button(summary_card, text="Obrir carpeta d'eixida", command=self.open_split_output_dir, style="Neutral.TButton")
        self.open_output_button.pack(anchor="e", pady=(10, 0))

        info_card = ttk.Frame(outer, style="Card.TFrame", padding=16)
        info_card.pack(fill="both", expand=True, pady=(10, 0))
        ttk.Label(info_card, text="Resultat del divisor", style="Section.TLabel").pack(anchor="w", pady=(0, 8))
        ttk.Label(
            info_card,
            text="La ferramenta divideix el PDF només per pes aproximat, sense interpretar el contingut. Si una pàgina sola supera el límit, es guardarà igual en una part pròpia.",
            style="Body.TLabel",
            wraplength=1100,
            justify="left",
        ).pack(anchor="w", pady=(0, 12))

        self.results_cards_frame = ttk.Frame(info_card, style="Card.TFrame")
        self.results_cards_frame.pack(fill="x", pady=(0, 12))

        self.split_log_box = tk.Text(
            info_card,
            wrap="word",
            font=("Consolas", 10),
            bg="white",
            fg=COLORS["charcoal"],
            relief="flat",
            padx=12,
            pady=12,
        )
        self.split_log_box.pack(fill="both", expand=True)
        self.split_log_box.configure(state="disabled")

        ttk.Label(outer, textvariable=self.split_status_var, style="Status.TLabel", padding=(0, 8, 0, 0)).pack(fill="x")

    def bind_editor_events(self):
        for variable in self.editor_vars.values():
            variable.trace_add("write", self.on_editor_change)

    def bind_shortcuts(self):
        self.root.bind("<Control-o>", lambda _event: self.pick_excel())
        self.root.bind("<Control-l>", lambda _event: self.load_excel())
        self.root.bind("<Control-f>", lambda _event: self.load_form())
        self.root.bind("<Control-s>", lambda _event: self.apply_current_record())
        self.root.bind("<Delete>", lambda _event: self.delete_current_record())
        self.root.bind("<Control-Return>", lambda _event: self.submit_selected())
        self.root.bind("<Control-Shift-Return>", lambda _event: self.submit_all())
        self.root.bind("<Control-g>", lambda _event: self.open_google_search())
        self.root.bind("<Control-u>", lambda _event: self.open_source())
        self.root.bind("<Control-d>", lambda _event: self.apply_fallback_font_to_all())
        self.root.bind("<Control-e>", lambda _event: self.focus_editor())
        self.root.bind("<Control-t>", lambda _event: self.tree.focus_set())
        self.root.bind("<Control-Up>", lambda _event: self.move_selection(-1))
        self.root.bind("<Control-Down>", lambda _event: self.move_selection(1))

    def show_splash(self):
        if not LOGO_PATH.exists():
            return
        splash = tk.Toplevel(self.root)
        splash.overrideredirect(True)
        splash.configure(bg=COLORS["cream"])
        splash.attributes("-topmost", True)
        self.root.withdraw()

        width, height = 340, 320
        x = self.root.winfo_screenwidth() // 2 - width // 2
        y = self.root.winfo_screenheight() // 2 - height // 2
        splash.geometry(f"{width}x{height}+{x}+{y}")

        frame = tk.Frame(splash, bg=COLORS["cream"], highlightbackground=COLORS["line"], highlightthickness=1)
        frame.pack(fill="both", expand=True)

        image = Image.open(LOGO_PATH).convert("RGBA")
        image.thumbnail((180, 180))
        self.logo_image = ImageTk.PhotoImage(image)

        tk.Label(frame, image=self.logo_image, bg=COLORS["cream"]).pack(pady=(42, 14))
        tk.Label(frame, text=APP_NAME, bg=COLORS["cream"], fg=COLORS["red"], font=("Segoe UI Semibold", 22)).pack()
        tk.Label(frame, text="Iniciant aplicació...", bg=COLORS["cream"], fg=COLORS["muted"], font=("Segoe UI", 10)).pack(pady=(8, 0))

        steps = [88, 90, 92, 94, 96, 98, 100]

        def animate(i=0):
            if i >= len(steps):
                splash.destroy()
                self.root.deiconify()
                self.root.lift()
                try:
                    self.root.attributes("-topmost", True)
                    self.root.after(120, lambda: self.root.attributes("-topmost", False))
                except tk.TclError:
                    pass
                self.root.focus_force()
                return
            try:
                splash.attributes("-alpha", steps[i] / 100)
            except tk.TclError:
                splash.destroy()
                self.root.deiconify()
                return
            splash.after(55, lambda: animate(i + 1))

        self.root.after(150, animate)

    def now_stamp(self):
        return dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    def append_history(self, text):
        line = f"[{self.now_stamp()}] {text}"
        self.history_entries.append(line)
        self.history_box.configure(state="normal")
        self.history_box.insert("end", line + "\n")
        self.history_box.see("end")
        self.history_box.configure(state="disabled")

    def set_status(self, text):
        self.status_var.set(text)

    def show_button_help(self, event, text):
        self.hide_button_help()
        popup = tk.Toplevel(self.root)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.configure(bg=COLORS["charcoal"])
        label = tk.Label(
            popup,
            text=text,
            bg=COLORS["charcoal"],
            fg="white",
            font=("Segoe UI", 9),
            padx=10,
            pady=6,
        )
        label.pack()
        popup.geometry(f"+{event.x_root + 8}+{event.y_root + 8}")
        self.help_popup = popup
        self.root.after(2200, self.hide_button_help)
        return "break"

    def hide_button_help(self):
        if self.help_popup is not None:
            try:
                self.help_popup.destroy()
            except tk.TclError:
                pass
            self.help_popup = None

    def set_split_status(self, text):
        self.split_status_var.set(text)

    def append_split_log(self, text):
        self.split_log_box.configure(state="normal")
        self.split_log_box.insert("end", text + "\n")
        self.split_log_box.see("end")
        self.split_log_box.configure(state="disabled")

    def set_split_progress(self, current, total):
        percent = 0 if total <= 0 else (current / total) * 100
        self.split_progress_var.set(percent)

    def clear_split_results(self):
        for widget in self.results_cards_frame.winfo_children():
            widget.destroy()

    def add_split_result_card(self, file_path: Path, pages_text: str):
        size_mb = file_path.stat().st_size / (1024 * 1024)
        card = tk.Frame(self.results_cards_frame, bg="white", highlightbackground=COLORS["line"], highlightthickness=1)
        card.pack(fill="x", pady=5)
        tk.Label(card, text=file_path.name, bg="white", fg=COLORS["red"], font=("Segoe UI Semibold", 11), anchor="w").pack(fill="x", padx=12, pady=(10, 2))
        tk.Label(card, text=f"{pages_text}  |  {size_mb:.2f} MB", bg="white", fg=COLORS["charcoal"], font=("Segoe UI", 10), anchor="w").pack(fill="x", padx=12, pady=(0, 10))

    def open_split_output_dir(self):
        path = self.split_output_dir_var.get().strip()
        if path and Path(path).exists():
            os.startfile(path)  # type: ignore[attr-defined]

    def enable_drop_support(self):
        try:
            self.root.tk.call("package", "require", "tkdnd")
            self.drop_zone.drop_target_register("DND_Files")
            self.drop_zone.dnd_bind("<<Drop>>", self.on_drop_pdf)
            self.drop_zone.configure(text="Arrossega ací un PDF OCR o fes clic per a seleccionar-lo")
        except Exception:
            self.drop_zone.configure(text="Fes clic ací per a seleccionar un PDF OCR")

    def on_drop_pdf(self, event):
        raw = event.data.strip()
        if raw.startswith("{") and raw.endswith("}"):
            raw = raw[1:-1]
        path = Path(raw)
        if path.exists():
            self.split_pdf_path_var.set(str(path))
            default_output = path.with_name(f"{path.stem}_parts")
            self.split_output_dir_var.set(str(default_output))
            self.split_summary_var.set(f"PDF carregat: {path.name} | {path.stat().st_size / (1024 * 1024):.2f} MB")
            self.set_split_status("PDF preparat per a dividir.")

    def open_splitter_tab(self):
        self.notebook.select(self.splitter_tab)

    def pick_split_pdf(self):
        path = filedialog.askopenfilename(title="Selecciona un PDF OCR", filetypes=[("PDF", "*.pdf"), ("Tots", "*.*")])
        if path:
            self.split_pdf_path_var.set(path)
            source = Path(path)
            default_output = source.with_name(f"{source.stem}_parts")
            self.split_output_dir_var.set(str(default_output))
            self.split_summary_var.set(f"PDF carregat: {source.name} | {source.stat().st_size / (1024 * 1024):.2f} MB")
            self.split_progress_var.set(0)
            self.set_split_status("PDF preparat per a dividir.")

    def pick_split_output_dir(self):
        path = filedialog.askdirectory(title="Selecciona carpeta d'eixida")
        if path:
            self.split_output_dir_var.set(path)

    def run_split_pdf(self):
        source_path = Path(self.split_pdf_path_var.get().strip())
        output_dir = Path(self.split_output_dir_var.get().strip()) if self.split_output_dir_var.get().strip() else None
        if not source_path.exists():
            messagebox.showwarning("Falta el PDF", "Selecciona un PDF OCR vàlid.")
            return
        if output_dir is None:
            output_dir = source_path.with_name(f"{source_path.stem}_parts")
            self.split_output_dir_var.set(str(output_dir))

        try:
            output_dir.mkdir(parents=True, exist_ok=True)
            self.clear_split_results()
            self.split_progress_var.set(0)
            created = self.split_pdf_by_size(source_path, output_dir, max_size_mb=24)
        except Exception as exc:
            messagebox.showerror("Error dividint el PDF", str(exc))
            self.set_split_status("No s'ha pogut dividir el PDF.")
            return

        self.split_log_box.configure(state="normal")
        self.split_log_box.delete("1.0", "end")
        self.split_log_box.configure(state="disabled")
        self.append_split_log(f"PDF original: {source_path}")
        self.append_split_log(f"Carpeta d'eixida: {output_dir}")
        for item, page_indexes in created:
            size_mb = item.stat().st_size / (1024 * 1024)
            self.append_split_log(f"- {item.name}: {size_mb:.2f} MB")
            pages_text = f"Pàgines {page_indexes[0] + 1}-{page_indexes[-1] + 1}" if len(page_indexes) > 1 else f"Pàgina {page_indexes[0] + 1}"
            self.add_split_result_card(item, pages_text)
        original_mb = source_path.stat().st_size / (1024 * 1024)
        self.split_summary_var.set(
            f"PDF original: {source_path.name} ({original_mb:.2f} MB) | Parts creades: {len(created)} | Límit per part: 24 MB"
        )
        self.split_progress_var.set(100)
        self.set_split_status(f"PDF dividit en {len(created)} parts.")
        self.notebook.select(self.splitter_tab)
        self.open_split_output_dir()

    def split_pdf_by_size(self, source_path: Path, output_dir: Path, max_size_mb=24):
        max_bytes = int(max_size_mb * 1024 * 1024)
        reader = PdfReader(str(source_path))
        created_files = []
        chunk_page_indexes = []
        part_number = 1

        def writer_size(page_indexes):
            writer = PdfWriter()
            for page_index in page_indexes:
                writer.add_page(reader.pages[page_index])
            buffer = BytesIO()
            writer.write(buffer)
            return buffer.getbuffer().nbytes

        total_pages = len(reader.pages)
        for page_index in range(total_pages):
            proposed = chunk_page_indexes + [page_index]
            proposed_size = writer_size(proposed)
            if chunk_page_indexes and proposed_size > max_bytes:
                created_files.append((self.save_pdf_chunk(reader, chunk_page_indexes, output_dir, source_path.stem, part_number), list(chunk_page_indexes)))
                part_number += 1
                chunk_page_indexes = [page_index]
            else:
                chunk_page_indexes = proposed
            self.set_split_progress(page_index + 1, total_pages)

        if chunk_page_indexes:
            created_files.append((self.save_pdf_chunk(reader, chunk_page_indexes, output_dir, source_path.stem, part_number), list(chunk_page_indexes)))

        return created_files

    def save_pdf_chunk(self, reader, page_indexes, output_dir: Path, stem: str, part_number: int):
        writer = PdfWriter()
        for page_index in page_indexes:
            writer.add_page(reader.pages[page_index])
        output_path = output_dir / f"{stem}_part_{part_number:02d}.pdf"
        with output_path.open("wb") as handle:
            writer.write(handle)
        return output_path

    def show_shortcuts_help(self):
        window = tk.Toplevel(self.root)
        window.title("Tecles ràpides")
        window.geometry("520x500")
        window.configure(bg=COLORS["paper"])

        frame = ttk.Frame(window, style="Card.TFrame", padding=18)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Tecles ràpides", style="Section.TLabel").pack(anchor="w", pady=(0, 10))

        shortcuts = [
            ("Ctrl + O", "Buscar Excel"),
            ("Ctrl + L", "Carregar Excel"),
            ("Ctrl + F", "Llegir formulari"),
            ("Ctrl + S", "Aplicar canvis de la fila"),
            ("Supr", "Eliminar fila seleccionada"),
            ("Ctrl + Enter", "Enviar fila seleccionada"),
            ("Ctrl + Shift + Enter", "Enviar totes les files"),
            ("Ctrl + G", "Buscar a Google"),
            ("Ctrl + U", "Obrir font"),
            ("Ctrl + D", "Aplicar font per defecte a totes"),
            ("Ctrl + E", "Portar el focus a l'editor"),
            ("Ctrl + T", "Portar el focus a la taula"),
            ("Ctrl + Arriba / Baix", "Canviar de fila"),
        ]

        box = tk.Text(
            frame,
            wrap="word",
            font=("Segoe UI", 10),
            bg="white",
            fg=COLORS["charcoal"],
            relief="flat",
            padx=12,
            pady=12,
        )
        box.pack(fill="both", expand=True)
        box.insert("1.0", "\n".join(f"{combo}: {desc}" for combo, desc in shortcuts))
        box.configure(state="disabled")

    def focus_editor(self):
        for key in ("nom", "font", "localitat", "data"):
            widget = self.editor_widgets.get(key)
            if widget is not None:
                widget.focus_set()
                return

    def move_selection(self, delta):
        items = self.tree.get_children()
        if not items:
            return
        selection = self.tree.selection()
        current = selection[0] if selection else items[0]
        try:
            position = items.index(current)
        except ValueError:
            position = 0
        target = items[max(0, min(len(items) - 1, position + delta))]
        self.tree.selection_set(target)
        self.tree.focus(target)
        self.tree.see(target)
        self.on_tree_select(None)

    def pick_excel(self):
        path = filedialog.askopenfilename(title="Selecciona un Excel", filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tots", "*.*")])
        if path:
            self.excel_path_var.set(path)

    def load_excel(self):
        path = self.excel_path_var.get().strip()
        if not path:
            messagebox.showwarning("Falta l'Excel", "Selecciona un arxiu Excel.")
            return
        try:
            self.records = load_excel_records(path, self.person_var.get(), self.fallback_font_var.get())
        except Exception as exc:
            messagebox.showerror("Error llegint l'Excel", str(exc))
            return
        self.current_index = None
        self.last_deleted = None
        self.refresh_tree()
        self.append_history(f"Excel carregat amb {len(self.records)} files.")
        self.set_status(f"Excel carregat: {len(self.records)} files.")

    def load_form(self):
        form_url = self.form_url_var.get().strip()
        if not form_url:
            messagebox.showwarning("Falta la URL", "Escriu la URL del formulari.")
            return

        def task():
            try:
                metadata = extract_form_metadata(form_url)
            except Exception as exc:
                self.root.after(0, lambda: messagebox.showerror("Error llegint el formulari", str(exc)))
                return
            self.form_metadata = metadata
            self.root.after(0, lambda: self.set_status("Formulari llegit. Camps detectats i preparat per a enviar."))
            self.root.after(0, lambda: self.append_history("Formulari llegit correctament."))

        threading.Thread(target=task, daemon=True).start()
        self.set_status("Llegint formulari...")

    def save_session(self):
        self.apply_current_record(silent=True)
        path = filedialog.asksaveasfilename(
            title="Guardar sessió",
            defaultextension=".autocpv.json",
            filetypes=SESSION_FILETYPES,
        )
        if not path:
            return
        payload = {
            "app": APP_NAME,
            "version": APP_VERSION,
            "saved_at": self.now_stamp(),
            "excel_path": self.excel_path_var.get().strip(),
            "form_url": self.form_url_var.get().strip(),
            "person": self.person_var.get().strip(),
            "fallback_font": self.fallback_font_var.get().strip(),
            "status_filter": self.status_filter_var.get().strip(),
            "records": [asdict(record) for record in self.records],
            "history": self.history_entries,
        }
        Path(path).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        self.append_history(f"Sessió guardada en {path}.")
        self.set_status("Sessió guardada correctament.")

    def load_session(self):
        path = filedialog.askopenfilename(title="Obrir sessió", filetypes=SESSION_FILETYPES)
        if not path:
            return
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        self.records = [Record(**item) for item in data.get("records", [])]
        self.excel_path_var.set(data.get("excel_path", ""))
        self.form_url_var.set(data.get("form_url", DEFAULT_FORM_URL))
        self.person_var.set(data.get("person", "Pol"))
        self.fallback_font_var.set(data.get("fallback_font", ""))
        self.status_filter_var.set(data.get("status_filter", "Totes") or "Totes")
        self.history_entries = list(data.get("history", []))
        self.reload_history_box()
        self.current_index = None
        self.last_deleted = None
        self.refresh_tree()
        self.append_history(f"Sessió reoberta des de {path}.")
        self.set_status("Sessió carregada correctament.")

    def reload_history_box(self):
        self.history_box.configure(state="normal")
        self.history_box.delete("1.0", "end")
        if self.history_entries:
            self.history_box.insert("end", "\n".join(self.history_entries) + "\n")
        self.history_box.configure(state="disabled")

    def apply_fallback_font_to_all(self):
        self.apply_fallback_font(mode="all")

    def apply_fallback_font_to_empty(self):
        self.apply_fallback_font(mode="empty")

    def apply_fallback_font(self, mode: str):
        fallback_font = self.fallback_font_var.get().strip()
        if not fallback_font:
            messagebox.showwarning("Falta la font", "Escriu una font per defecte abans d'aplicar-la.")
            return
        if not self.records:
            messagebox.showwarning("Sense files", "Carrega primer un Excel.")
            return

        changed = 0
        for record in self.records:
            if mode == "all" or not record.font.strip():
                record.font = fallback_font
                changed += 1

        if self.current_index is not None and changed:
            self.suspend_dirty = True
            try:
                self.editor_vars["font"].set(self.records[self.current_index].font)
            finally:
                self.suspend_dirty = False
            self.dirty = False

        self.refresh_tree()
        if mode == "all":
            self.set_status("Font per defecte aplicada a totes les files.")
            self.append_history("Font per defecte aplicada a totes les files.")
        else:
            self.set_status(f"Font per defecte aplicada a {changed} files buides.")
            self.append_history(f"Font per defecte aplicada a {changed} files buides.")

    def on_editor_change(self, *_args):
        if self.suspend_dirty or self.current_index is None:
            return
        self.dirty = True
        self.set_status("Guardant canvis automàticament...")
        if self.autosave_after_id is not None:
            self.root.after_cancel(self.autosave_after_id)
        self.autosave_after_id = self.root.after(350, self.autosave_current_record)

    def autosave_current_record(self):
        self.autosave_after_id = None
        self.apply_current_record(silent=True)
        self.set_status("Canvis guardats automàticament.")

    def refresh_tree(self, preserve_selection=True):
        selected = None
        if preserve_selection:
            current_selection = self.tree.selection()
            selected = current_selection[0] if current_selection else (str(self.current_index) if self.current_index is not None else None)

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.visible_indices = [idx for idx, record in enumerate(self.records) if self.record_matches_filter(record)]
        for idx in self.visible_indices:
            record = self.records[idx]
            self.tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(record.localitat, record.data, record.categoria, record.nom, record.preu, record.font, self.display_status(record)),
                tags=(self.record_tag(record),),
            )

        if not self.visible_indices:
            self.clear_editor()
            self.validation_var.set("No hi ha files visibles amb el filtre actual.")
            return

        if selected is None or selected not in self.tree.get_children():
            selected = str(self.visible_indices[0])
        self.tree.selection_set(selected)
        self.tree.focus(selected)
        self.tree.see(selected)
        self.populate_editor(int(selected))

    def record_matches_filter(self, record: Record) -> bool:
        mode = FILTER_OPTIONS.get(self.status_filter_var.get(), "all")
        tag = self.record_tag(record)
        if mode == "all":
            return True
        if mode == "sent":
            return tag == "sent"
        if mode == "error":
            return tag == "error"
        if mode == "invalid":
            return tag == "invalid"
        return tag not in {"sent", "error", "invalid"}

    def display_status(self, record: Record) -> str:
        if record.status == "Enviat":
            return "Enviat"
        if record.status == "Error":
            return "Error"
        errors = self.validate_record(record)
        if errors:
            return "Revisar"
        return "Pendent"

    def record_tag(self, record: Record) -> str:
        if record.status == "Enviat":
            return "sent"
        if record.status == "Error":
            return "error"
        if self.validate_record(record):
            return "invalid"
        return ""

    def populate_editor(self, index):
        self.current_index = index
        record = self.records[index]
        self.suspend_dirty = True
        try:
            for key in FIELD_LABELS:
                self.editor_vars[key].set(getattr(record, key))
        finally:
            self.suspend_dirty = False
        self.dirty = False
        self.update_validation_summary(index)

    def clear_editor(self):
        self.current_index = None
        self.suspend_dirty = True
        try:
            for key in FIELD_LABELS:
                self.editor_vars[key].set("")
            self.editor_vars["persona"].set(self.person_var.get().strip() or "Pol")
        finally:
            self.suspend_dirty = False
        self.dirty = False

    def apply_current_record(self, silent=False):
        if self.current_index is None:
            return True
        record = self.records[self.current_index]
        previous_values = {key: getattr(record, key) for key in FIELD_LABELS}
        for key in FIELD_LABELS:
            setattr(record, key, self.editor_vars[key].get().strip())
        changed = any(previous_values[key] != getattr(record, key) for key in FIELD_LABELS)
        if changed:
            record.status = "Pendent"
            record.status_detail = ""
        self.update_tree_item(self.current_index)
        self.update_validation_summary(self.current_index)
        self.dirty = False
        if not silent:
            self.set_status("Canvis aplicats correctament.")
        return True

    def update_tree_item(self, index):
        iid = str(index)
        if iid not in self.tree.get_children():
            return
        record = self.records[index]
        self.tree.item(
            iid,
            values=(record.localitat, record.data, record.categoria, record.nom, record.preu, record.font, self.display_status(record)),
            tags=(self.record_tag(record),),
        )

    def update_validation_summary(self, index):
        record = self.records[index]
        errors = self.validate_record(record)
        if errors:
            self.validation_var.set("Cal revisar:\n- " + "\n- ".join(errors[:6]))
        else:
            detail = record.status_detail or "Fila preparada per a enviar."
            self.validation_var.set(detail)

    def validate_record(self, record: Record):
        errors = []
        for key in REQUIRED_FIELDS:
            if not getattr(record, key).strip():
                errors.append(f"Falta {FIELD_LABELS[key].lower()}.")
        if record.categoria == "Altres" and not record.altres.strip():
            errors.append("Has omplit 'Categoria' amb 'Altres' però falta la descripció.")
        if record.data and not is_valid_date(record.data):
            errors.append("La data ha d'estar en format YYYY-MM-DD.")
        if record.preu and not is_valid_price(record.preu):
            errors.append("El preu ha de ser numèric sense símbols.")
        if record.font and not re.match(r"^https?://", record.font.strip()):
            errors.append("La font ha de començar per http:// o https://.")
        return errors

    def delete_current_record(self):
        if self.current_index is None or not self.records:
            messagebox.showinfo("Sense selecció", "Selecciona una fila per a eliminar.")
            return
        self.apply_current_record(silent=True)
        record = self.records[self.current_index]
        description = record.nom or record.localitat or f"fila {self.current_index + 1}"
        confirmed = messagebox.askyesno("Eliminar fila", f"Vols eliminar completament esta fila?\n\n{description}")
        if not confirmed:
            return

        deleted_index = self.current_index
        deleted_record = self.records.pop(deleted_index)
        self.last_deleted = (deleted_index, deleted_record)
        self.refresh_tree()
        self.append_history(f"Fila eliminada: {description}.")
        if self.records:
            self.set_status("Fila eliminada correctament. Pots desfer-la.")
        else:
            self.set_status("Fila eliminada. Ja no queden registres carregats.")

    def undo_delete_record(self):
        if not self.last_deleted:
            messagebox.showinfo("Sense canvis", "No hi ha cap eliminació recent per a desfer.")
            return
        index, record = self.last_deleted
        index = max(0, min(index, len(self.records)))
        self.records.insert(index, record)
        self.last_deleted = None
        self.refresh_tree()
        if str(index) in self.tree.get_children():
            self.tree.selection_set(str(index))
            self.tree.focus(str(index))
            self.populate_editor(index)
        self.append_history(f"Eliminació desfeta: {record.nom or record.localitat or 'fila'}")
        self.set_status("Fila recuperada correctament.")

    def on_tree_select(self, _event):
        selection = self.tree.selection()
        if not selection:
            return
        target = int(selection[0])
        if self.current_index is not None and target != self.current_index and self.dirty:
            self.apply_current_record(silent=True)
        self.populate_editor(target)

    def build_search_query(self, record: Record):
        return " ".join(part for part in [record.nom, record.companyia, record.localitat] if part).strip()

    def open_google_search(self):
        if self.current_index is None:
            return
        self.apply_current_record(silent=True)
        record = self.records[self.current_index]
        query = self.build_search_query(record)
        if query:
            webbrowser.open("https://www.google.com/search?q=" + urllib.parse.quote(query))

    def open_source_helper(self):
        if self.current_index is None:
            return
        self.apply_current_record(silent=True)
        record = self.records[self.current_index]
        if record.font.strip():
            webbrowser.open(record.font.strip())
            return
        query = self.build_search_query(record)
        if query:
            webbrowser.open("https://www.google.com/search?q=" + urllib.parse.quote(query))
            return
        fallback = record.localitat.strip() or "ajuntament cultura"
        webbrowser.open(DEFAULT_FACEBOOK_SEARCH + urllib.parse.quote(fallback))

    def open_source(self):
        if self.current_index is None:
            return
        self.apply_current_record(silent=True)
        source = self.records[self.current_index].font.strip()
        if source:
            webbrowser.open(source)
        else:
            messagebox.showinfo("Sense font", "Esta fila encara no té cap font.")

    def ensure_ready_to_submit(self):
        if not self.records:
            messagebox.showwarning("Sense files", "Carrega primer un Excel.")
            return False
        if not self.form_metadata:
            messagebox.showwarning("Sense formulari", "Prem abans en 'Llegir formulari'.")
            return False
        return True

    def invalid_indexes(self, indexes):
        invalid = []
        for idx in indexes:
            errors = self.validate_record(self.records[idx])
            if errors:
                invalid.append((idx, errors))
        return invalid

    def preview_current_payload(self):
        if not self.ensure_ready_to_submit():
            return
        if self.current_index is None:
            messagebox.showinfo("Sense selecció", "Selecciona una fila.")
            return
        self.apply_current_record(silent=True)
        errors = self.validate_record(self.records[self.current_index])
        if errors:
            messagebox.showwarning("Fila incompleta", "\n".join(errors))
            return

        payload = build_payload(self.records[self.current_index], self.form_metadata)
        window = tk.Toplevel(self.root)
        window.title("Previsualització d'enviament")
        window.geometry("900x680")
        window.configure(bg=COLORS["paper"])

        frame = ttk.Frame(window, style="Card.TFrame", padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Payload que s'enviarà al formulari", style="Section.TLabel").pack(anchor="w", pady=(0, 8))
        box = tk.Text(frame, wrap="word", font=("Consolas", 10), bg="white", fg=COLORS["charcoal"], relief="flat", padx=12, pady=12)
        box.pack(fill="both", expand=True)
        lines = [f"{key} = {value}" for key, value in payload.items()]
        box.insert("1.0", "\n".join(lines))
        box.configure(state="disabled")

    def submit_selected(self):
        if not self.ensure_ready_to_submit():
            return
        if self.current_index is None:
            return
        self.apply_current_record(silent=True)
        invalid = self.invalid_indexes([self.current_index])
        if invalid:
            messagebox.showwarning("Fila no preparada", "\n".join(invalid[0][1]))
            self.update_validation_summary(self.current_index)
            return
        self.run_submission([self.current_index])

    def submit_all(self):
        if not self.ensure_ready_to_submit():
            return
        self.apply_current_record(silent=True)
        indexes = list(range(len(self.records)))
        invalid = self.invalid_indexes(indexes)
        if invalid:
            lines = []
            for idx, errors in invalid[:8]:
                title = self.records[idx].nom or self.records[idx].localitat or f"fila {idx + 1}"
                lines.append(f"{idx + 1}. {title}: {errors[0]}")
            if len(invalid) > 8:
                lines.append(f"... i {len(invalid) - 8} files més.")
            messagebox.showwarning("Hi ha files per revisar", "\n".join(lines))
            self.refresh_tree()
            return
        self.run_submission(indexes)

    def run_submission(self, indexes):
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("En curs", "Ja hi ha un enviament en marxa.")
            return

        def task():
            ok = 0
            for idx in indexes:
                record = self.records[idx]
                try:
                    payload = build_payload(record, self.form_metadata)
                    body = urllib.parse.urlencode(payload).encode("utf-8")
                    request = urllib.request.Request(
                        self.form_metadata["response_url"],
                        data=body,
                        headers={"User-Agent": "Mozilla/5.0"},
                        method="POST",
                    )
                    with urllib.request.urlopen(request, timeout=30) as response:
                        if response.status != 200:
                            raise ValueError(f"Resposta HTTP {response.status}")
                    record.status = "Enviat"
                    record.status_detail = f"Enviat correctament el {self.now_stamp()}."
                    ok += 1
                    self.root.after(0, lambda idx=idx, name=(record.nom or record.localitat or f"fila {idx + 1}"): self.append_history(f"Enviada correctament: {name}."))
                except Exception as exc:
                    record.status = "Error"
                    record.status_detail = self.friendly_error_message(exc)
                    self.root.after(0, lambda idx=idx, detail=record.status_detail, name=(record.nom or record.localitat or f"fila {idx + 1}"): self.append_history(f"Error en {name}: {detail}"))
                self.root.after(0, self.refresh_tree)
                self.root.after(0, lambda idx=idx: self.reselect_index(idx))
            self.root.after(0, lambda: self.set_status(f"Enviament acabat: {ok}/{len(indexes)} correctes."))

        self.worker = threading.Thread(target=task, daemon=True)
        self.worker.start()
        self.set_status("Enviant formularis...")

    def reselect_index(self, index):
        iid = str(index)
        if iid in self.tree.get_children():
            self.tree.selection_set(iid)
            self.tree.focus(iid)
            self.tree.see(iid)
            self.populate_editor(index)

    def friendly_error_message(self, exc):
        text = str(exc)
        lowered = text.lower()
        if "timed out" in lowered or "timeout" in lowered:
            return "Temps d'espera esgotat en enviar el formulari."
        if "http" in lowered:
            return text
        if "urlopen error" in lowered:
            return "No s'ha pogut connectar amb internet."
        return text

    def open_review_mode(self):
        if self.current_index is None:
            messagebox.showinfo("Sense selecció", "Selecciona una fila abans d'obrir la revisió ampla.")
            return
        self.apply_current_record(silent=True)

        window = tk.Toplevel(self.root)
        window.title("Revisió ampla")
        window.geometry("1100x820")
        window.configure(bg=COLORS["cream"])

        outer = ttk.Frame(window, style="Card.TFrame", padding=18)
        outer.pack(fill="both", expand=True)

        temp_vars = {key: tk.StringVar(value=self.editor_vars[key].get()) for key in FIELD_LABELS}
        temp_texts = {}
        row_index = 0

        for key, label in FIELD_LABELS.items():
            ttk.Label(outer, text=label, style="Body.TLabel").grid(row=row_index, column=0, sticky="nw", pady=6, padx=(0, 10))
            if key in LONG_TEXT_FIELDS:
                box = tk.Text(outer, height=3 if key != "font" else 4, wrap="word", font=("Segoe UI", 10))
                box.insert("1.0", temp_vars[key].get())
                box.grid(row=row_index, column=1, sticky="ew", pady=6)
                temp_texts[key] = box
            elif key in FIELD_OPTIONS:
                ttk.Combobox(outer, textvariable=temp_vars[key], values=FIELD_OPTIONS[key], state="readonly", width=60).grid(row=row_index, column=1, sticky="ew", pady=6)
            else:
                ttk.Entry(outer, textvariable=temp_vars[key], width=70).grid(row=row_index, column=1, sticky="ew", pady=6)
            row_index += 1

        outer.columnconfigure(1, weight=1)

        button_bar = ttk.Frame(outer, style="Card.TFrame")
        button_bar.grid(row=row_index, column=0, columnspan=2, sticky="ew", pady=(16, 0))

        def save_and_close():
            for key, box in temp_texts.items():
                temp_vars[key].set(box.get("1.0", "end-1c").strip())
            self.suspend_dirty = True
            try:
                for key in FIELD_LABELS:
                    self.editor_vars[key].set(temp_vars[key].get())
            finally:
                self.suspend_dirty = False
            self.apply_current_record(silent=True)
            window.destroy()
            self.set_status("Canvis guardats des de la revisió ampla.")

        ttk.Button(button_bar, text="Guardar i tancar", command=save_and_close, style="Neutral.TButton").pack(side="right", padx=4)


def main():
    root = tk.Tk()
    FormFillerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
