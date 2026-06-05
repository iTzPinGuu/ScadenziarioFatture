#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scadenziario Commerciale Premium — VIDALOCA di Michela Vidale
Versione 5.6 - Caricamento Logo Online (URL) & Firma Interattiva
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import csv
import os
import webbrowser
import urllib.request
import io
from datetime import datetime, date, timedelta
import calendar
import re

# Tenta l'importazione di Pillow per la gestione del logo da URL
try:
    from PIL import Image, ImageTk
    PILLOW_OK = True
except ImportError:
    PILLOW_OK = False

try:
    import openpyxl
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "scadenziario.db")
# URL Remoto del Logo VIDALOCA fornito dall'utente
LOGO_URL = "https://www.truccoloangelo.com/wp-content/uploads/2026/05/Logo-Vidaloca.png"

# ─── DESIGN SYSTEM & PALETTE VIDALOCA ────────────────────────────────────────
C = {
    "bg":          "#fdfbf7",  # Sfondo panna chiarissimo caldo
    "surface":     "#ffffff",  # Superficie card
    "border":      "#e7e1d5",  # Bordi morbidi tonalità sabbia
    "text":        "#3d2d1f",  # Marrone cioccolato profondo (dai testi del logo)
    "muted":       "#857364",  # Marrone desaturato per testi secondari
    "accent":      "#c5a059",  # Oro / Bronzo luminoso (colore dominante logo)
    "accent_dark": "#8c6f34",  # Oro brunito per stati attivi
    "teal_accent": "#1e6b7b",  # Ottanio (dal riflesso dei bracciali nel logo)
    "scaduta":     "#fdf2f2",  # Rosso ultra-soft per anomalie
    "scaduta_fg":  "#991b1b",  
    "urgente":     "#fef3c7",  # Giallo ambra per imminenti
    "urgente_fg":  "#92400e",  
    "ok":          "#f0fdf4",  # Verde soft per chiusi
    "ok_fg":       "#166534",  
    "red":         "#ef4444",  
    "green":       "#10b981",  
}

STATI_PASSIVO   = ["Da pagare", "Parziale", "Pagata", "Insoluta"]
STATI_ATTIVO    = ["Da incassare", "Parziale", "Incassata", "Insoluta"]
METODI_PAG      = ["Bonifico", "RID/SDD", "Contanti", "Assegno", "Carta", "Riba", "Altro"]
TIPI_DOC        = ["Fattura", "Nota di credito", "Parcella", "Fattura differita", "Altro"]

def get_conn():
    return sqlite3.connect(DB_FILE)

def init_db():
    with get_conn() as conn:
        c = conn.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS aziende (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT UNIQUE NOT NULL,
                condizioni TEXT NOT NULL DEFAULT '30 Giorni'
            )
        """)
        c.execute("""
            CREATE TABLE IF NOT EXISTS config_condizioni (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT UNIQUE NOT NULL
            )
        """)
        c.execute("""
            CREATE TABLE IF NOT EXISTS fatture (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                tipo        TEXT NOT NULL DEFAULT 'passivo',
                numero      TEXT NOT NULL,
                anagrafica  TEXT NOT NULL,
                data_doc    TEXT,
                data_scad   TEXT,
                importo     REAL NOT NULL,
                pagato      REAL NOT NULL DEFAULT 0,
                stato       TEXT NOT NULL DEFAULT 'Da pagare',
                metodo      TEXT,
                tipo_doc    TEXT DEFAULT 'Fattura',
                note        TEXT,
                data_ins    TEXT DEFAULT (date('now'))
            )
        """)
        c.execute("""
            CREATE TABLE IF NOT EXISTS pagamenti (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                fattura_id  INTEGER NOT NULL,
                data_pag    TEXT NOT NULL,
                importo     REAL NOT NULL,
                metodo      TEXT,
                note        TEXT,
                FOREIGN KEY(fattura_id) REFERENCES fatture(id)
            )
        """)
        c.execute("CREATE INDEX IF NOT EXISTS idx_fatture_ricerca ON fatture (numero, anagrafica, tipo)")
        
        c.execute("SELECT COUNT(*) FROM config_condizioni")
        if c.fetchone()[0] == 0:
            for cond in ["30 Giorni", "15 Giorni", "60 Giorni", "90 Giorni", "Vista fattura", "Contanti", "Riba 30 gg", "Riba 60 gg"]:
                c.execute("INSERT OR IGNORE INTO config_condizioni (nome) VALUES (?)", (cond,))

def get_condizioni_pagamento():
    try:
        with get_conn() as conn:
            rows = conn.execute("SELECT nome FROM config_condizioni ORDER BY id ASC").fetchall()
            if rows: return [r[0] for r in rows]
    except Exception: pass
    return ["30 Giorni", "15 Giorni", "60 Giorni", "90 Giorni", "Vista fattura", "Contanti", "Riba 30 gg", "Riba 60 gg"]

# ─── UTILS FORMATTAZIONE ─────────────────────────────────────────────────────
def parse_date(val):
    if not val: return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d") if hasattr(val, "hour") else val.isoformat()
    s = str(val).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"):
        try: return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError: continue
    return s

def parse_importo(val):
    if val is None or val == "": return 0.0
    s = str(val).strip().replace(" ", "")
    if "," in s and "." in s: s = s.replace(".", "").replace(",", ".")
    else: s = s.replace(",", ".")
    try: return float(s)
    except ValueError: return 0.0

def fmt_date(iso):
    if not iso: return ""
    try: return datetime.strptime(iso, "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception: return iso

def fmt_num(v):
    try: return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception: return str(v)

def mappa_stato_fte(stato_fte, tipo):
    if not stato_fte: return "Da pagare" if tipo == "passivo" else "Da incassare"
    s = stato_fte.strip().lower()
    if s in ("pagato", "pagata", "incassata", "incassato"): return "Pagata" if tipo == "passivo" else "Incassata"
    if s == "parziale": return "Parziale"
    return "Da pagare" if tipo == "passivo" else "Da incassare"

def calcola_scadenza_automatico(c_db, nome_azienda, data_doc_iso):
    if not data_doc_iso: return ""
    res = c_db.execute("SELECT condizioni FROM aziende WHERE nome = ?", (nome_azienda,)).fetchone()
    cond = res[0] if res else "30 Giorni"
    if not res:
        try: c_db.execute("INSERT INTO aziende (nome, condizioni) VALUES (?, ?)", (nome_azienda, "30 Giorni"))
        except sqlite3.IntegrityError: pass
    
    giorni = 30
    if "vista fattura" in cond.lower() or "contanti" in cond.lower(): giorni = 0
    else:
        match = re.search(r'\d+', cond)
        if match: giorni = int(match.group())
    try:
        dt = datetime.strptime(data_doc_iso, "%Y-%m-%d").date()
        return (dt + timedelta(days=giorni)).strftime("%Y-%m-%d")
    except Exception: return ""

# ─── COMPONENTI INTERFACCIA ──────────────────────────────────────────────────
class DateEntry(tk.Frame):
    def __init__(self, master, **kw):
        super().__init__(master, bg=C["surface"])
        self._var = tk.StringVar()
        self._entry = tk.Entry(self, textvariable=self._var, width=12, font=("Segoe UI", 10), relief="solid", bd=1, highlightthickness=0, borderwidth=1, fg=C["text"])
        self._entry.pack(side="left", ipady=3)
        btn = tk.Button(self, text="📅", command=self._open_cal, relief="flat", bg=C["surface"], fg=C["accent"], cursor="hand2", font=("Segoe UI", 10))
        btn.pack(side="left", padx=(4, 0))
        self._entry.bind("<FocusOut>", self._on_focus_out)

    def _on_focus_out(self, e=None):
        v = self._var.get().strip()
        if not v: return
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                d = datetime.strptime(v, fmt)
                self._var.set(d.strftime("%d/%m/%Y"))
                return
            except ValueError: continue

    def _open_cal(self):
        top = tk.Toplevel(self)
        top.title("Seleziona Data")
        top.resizable(False, False)
        top.geometry(f"+{self._entry.winfo_rootx()}+{self._entry.winfo_rooty() + 30}")
        CalPicker(top, self._var, top).pack()
        top.grab_set()

    def get(self):
        v = self._var.get().strip()
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try: return datetime.strptime(v, fmt).strftime("%Y-%m-%d")
            except ValueError: continue
        return v

    def set(self, iso_str):
        if not iso_str: self._var.set("")
        else:
            try: self._var.set(datetime.strptime(iso_str, "%Y-%m-%d").strftime("%d/%m/%Y"))
            except ValueError: self._var.set(iso_str)

class CalPicker(tk.Frame):
    def __init__(self, master, var, top):
        super().__init__(master, bg=C["surface"], padx=10, pady=10)
        self._var = var
        self._top = top
        try: sel = datetime.strptime(var.get().strip(), "%d/%m/%Y").date()
        except Exception: sel = date.today()
        self._year, self._month = sel.year, sel.month
        self._build()

    def _build(self):
        for w in self.winfo_children(): w.destroy()
        nav = tk.Frame(self, bg=C["surface"])
        nav.pack(fill="x", pady=(0,8))
        tk.Button(nav, text="←", command=self._prev, relief="flat", bg=C["surface"], fg=C["text"], font=("Segoe UI", 10, "bold"), cursor="hand2").pack(side="left")
        tk.Label(nav, text=f"{calendar.month_name[self._month]} {self._year}".upper(), font=("Segoe UI", 9, "bold"), bg=C["surface"], fg=C["text"], width=16).pack(side="left", expand=True)
        tk.Button(nav, text="→", command=self._next, relief="flat", bg=C["surface"], fg=C["text"], font=("Segoe UI", 10, "bold"), cursor="hand2").pack(side="right")
        
        gf = tk.Frame(self, bg=C["surface"])
        gf.pack()
        today = date.today()
        for i, d in enumerate(["LU", "MA", "ME", "GI", "VE", "SA", "DO"]):
            tk.Label(gf, text=d, width=4, font=("Segoe UI", 8, "bold"), bg=C["surface"], fg=C["muted"]).grid(row=0, column=i, pady=4)
        for r, week in enumerate(calendar.monthcalendar(self._year, self._month), start=1):
            for c_idx, day in enumerate(week):
                if day == 0: tk.Label(gf, width=4, bg=C["surface"]).grid(row=r, column=c_idx)
                else:
                    d = date(self._year, self._month, day)
                    bg = C["accent"] if d == today else C["surface"]
                    fg = "white" if d == today else C["text"]
                    tk.Button(gf, text=str(day), width=4, bg=bg, fg=fg, relief="flat", cursor="hand2", font=("Segoe UI", 9),
                              command=lambda dd=d: self._pick(dd)).grid(row=r, column=c_idx, padx=1, pady=1)

    def _prev(self):
        if self._month == 1: self._month, self._year = 12, self._year - 1
        else: self._month -= 1
        self._build()

    def _next(self):
        if self._month == 12: self._month, self._year = 1, self._year + 1
        else: self._month += 1
        self._build()

    def _pick(self, d):
        self._var.set(d.strftime("%d/%m/%Y"))
        self._top.destroy()

class ImportoEntry(tk.Entry):
    def __init__(self, master, **kw):
        width = kw.pop("width", 14)
        super().__init__(master, font=("Segoe UI", 10), relief="solid", bd=1, width=width, highlightthickness=0, fg=C["text"], **kw)
        self.bind("<FocusOut>", self._format)

    def _format(self, e=None):
        v = self.get().strip().replace(" ", "")
        if not v: return
        try:
            f = float(v.replace(",", "."))
            self.delete(0, tk.END)
            self.insert(0, f"{f:.2f}".replace(".", ","))
        except ValueError: pass

    def get_float(self):
        v = self.get().strip().replace(" ", "").replace(",", ".")
        try: return float(v)
        except ValueError: return 0.0

    def set_float(self, value):
        self.delete(0, tk.END)
        if value is not None:
            self.insert(0, f"{float(value):.2f}".replace(".", ","))

# ─── MODULI E FINESTRE DI DIALOGO ────────────────────────────────────────────
class ConfigCondizioniWindow(tk.Toplevel):
    def __init__(self, master, on_close_callback=None):
        super().__init__(master)
        self.title("Impostazioni Condizioni")
        self.geometry("460x420")
        self.resizable(False, False)
        self.configure(bg=C["surface"])
        self._on_close_callback = on_close_callback
        self._build()
        self._load_condizioni()
        self.grab_set()

    def _build(self):
        hdr = tk.Frame(self, bg=C["bg"], padx=16, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Condizioni di Pagamento", font=("Segoe UI", 12, "bold"), bg=C["bg"], fg=C["text"]).pack(anchor="w")
        
        main_fr = tk.Frame(self, bg=C["surface"], padx=16, pady=16)
        main_fr.pack(fill="both", expand=True)
        
        input_fr = tk.Frame(main_fr, bg=C["surface"])
        input_fr.pack(fill="x", pady=(0, 12))
        
        tk.Label(input_fr, text="Nuova stringa automatica (es. Riba 120 gg):", font=("Segoe UI", 9), bg=C["surface"], fg=C["muted"]).pack(anchor="w")
        self.ent_cond = tk.Entry(input_fr, font=("Segoe UI", 10), relief="solid", bd=1)
        self.ent_cond.pack(side="left", fill="x", expand=True, padx=(0, 8), pady=4, ipady=2)
        
        tk.Button(input_fr, text="Aggiungi", command=self._add_condizione, bg=C["accent"], fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=12).pack(side="right", ipady=2)
        
        list_fr = tk.Frame(main_fr, bg=C["surface"])
        list_fr.pack(fill="both", expand=True)
        
        self.tree = ttk.Treeview(list_fr, columns=("Nome"), show="headings", height=8)
        self.tree.heading("Nome", text="Regole di dilazione registrate")
        self.tree.column("Nome", width=300)
        self.tree.pack(side="left", fill="both", expand=True)
        
        sb = ttk.Scrollbar(list_fr, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        
        btn_fr = tk.Frame(main_fr, bg=C["surface"])
        btn_fr.pack(fill="x", pady=(14, 0))
        tk.Button(btn_fr, text="Rimuovi Selezionata", command=self._delete_condizione, bg=C["scaduta"], fg=C["scaduta_fg"], font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4).pack(side="left")
        tk.Button(btn_fr, text="Salva e Chiudi", command=self._chiudi, bg=C["border"], fg=C["text"], font=("Segoe UI", 9, "bold"), relief="flat", padx=14, pady=4).pack(side="right")

    def _load_condizioni(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        with get_conn() as conn:
            for r in conn.execute("SELECT nome FROM config_condizioni ORDER BY id ASC").fetchall():
                self.tree.insert("", "end", values=r)

    def _add_condizione(self):
        nome = self.ent_cond.get().strip()
        if not nome: return
        try:
            with get_conn() as conn:
                conn.execute("INSERT INTO config_condizioni (nome) VALUES (?)", (nome,))
            self.ent_cond.delete(0, tk.END)
            self._load_condizioni()
        except sqlite3.IntegrityError:
            messagebox.showwarning("Attenzione", "Condizione già esistente.")

    def _delete_condizione(self):
        sel = self.tree.selection()
        if not sel: return
        nome = self.tree.item(sel[0], "values")[0]
        with get_conn() as conn:
            conn.execute("DELETE FROM config_condizioni WHERE nome=?", (nome,))
        self._load_condizioni()

    def _chiudi(self):
        if self._on_close_callback: self._on_close_callback()
        self.destroy()

class AziendeWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Anagrafiche Clienti / Fornitori")
        self.geometry("720x480") 
        self.resizable(False, False)
        self.configure(bg=C["surface"])
        self._build()
        self._load_aziende()
        self.grab_set()

    def _build(self):
        hdr = tk.Frame(self, bg=C["bg"], padx=16, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Scadenze di Default per Anagrafica", font=("Segoe UI", 12, "bold"), bg=C["bg"], fg=C["text"]).pack(anchor="w")
        
        main_fr = tk.Frame(self, bg=C["surface"], padx=16, pady=12)
        main_fr.pack(fill="both", expand=True)

        left_fr = tk.Frame(main_fr, bg=C["surface"])
        left_fr.pack(side="left", fill="both", expand=True, padx=(0, 12))

        self.tree = ttk.Treeview(left_fr, columns=("Nome", "Condizioni"), show="headings", height=12)
        self.tree.heading("Nome", text="Ragione Sociale")
        self.tree.heading("Condizioni", text="Dilazione Applicata")
        self.tree.column("Nome", width=240)
        self.tree.column("Condizioni", width=140)
        self.tree.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(left_fr, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

        right_fr = tk.Frame(main_fr, bg=C["surface"], bd=1, relief="solid", highlightthickness=0, borderwidth=1, padx=12, pady=12)
        right_fr.pack(side="right", fill="y", pady=2)

        tk.Label(right_fr, text="Scheda Anagrafica", font=("Segoe UI", 10, "bold"), bg=C["surface"], fg=C["text"]).pack(anchor="w", pady=(0,10))
        tk.Label(right_fr, text="Ragione Sociale:", font=("Segoe UI", 9), bg=C["surface"], fg=C["muted"]).pack(anchor="w")
        self.ent_nome = tk.Entry(right_fr, font=("Segoe UI", 10), relief="solid", bd=1, width=24)
        self.ent_nome.pack(anchor="w", pady=(2, 10), ipady=2)

        tk.Label(right_fr, text="Termini Pagamento:", font=("Segoe UI", 9), bg=C["surface"], fg=C["muted"]).pack(anchor="w")
        cond_values = get_condizioni_pagamento()
        self.cmb_cond = ttk.Combobox(right_fr, values=cond_values, width=22, state="readonly")
        if cond_values: self.cmb_cond.set(cond_values[0])
        self.cmb_cond.pack(anchor="w", pady=(2, 16))

        tk.Button(right_fr, text="Salva Modifiche", command=self._save_azienda, bg=C["accent"], fg="white", font=("Segoe UI", 9, "bold"), relief="flat", pady=6).pack(fill="x", pady=2)
        tk.Button(right_fr, text="Elimina", command=self._delete_azienda, bg=C["scaduta"], fg=C["scaduta_fg"], font=("Segoe UI", 9, "bold"), relief="flat", pady=6).pack(fill="x", pady=2)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

    def _load_aziende(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        with get_conn() as conn:
            for r in conn.execute("SELECT nome, condizioni FROM aziende ORDER BY nome ASC").fetchall():
                self.tree.insert("", "end", values=r)

    def _on_select(self, e):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0], "values")
        self.ent_nome.delete(0, tk.END)
        self.ent_nome.insert(0, vals[0])
        self.cmb_cond.set(vals[1])

    def _save_azienda(self):
        doc = self.ent_nome.get().strip()
        cond = self.cmb_cond.get()
        if not doc: return
        with get_conn() as conn:
            conn.execute("INSERT INTO aziende (nome, condizioni) VALUES (?, ?) ON CONFLICT(nome) DO UPDATE SET condizioni=?", (doc, cond, cond))
        self.ent_nome.delete(0, tk.END)
        self._load_aziende()

    def _delete_azienda(self):
        sel = self.tree.selection()
        if not sel: return
        nome = self.tree.item(sel[0], "values")[0]
        if messagebox.askyesno("Conferma", f"Rimuovere '{nome}' dalla rubrica?"):
            with get_conn() as conn:
                conn.execute("DELETE FROM aziende WHERE nome=?", (nome,))
            self._load_aziende()

class FatturaForm(tk.Toplevel):
    def __init__(self, master, tipo="passivo", fattura=None, on_save=None):
        super().__init__(master)
        self.title("Anagrafica Documento")
        self.resizable(False, False)
        self._tipo, self._fattura, self._on_save = tipo, fattura, on_save
        self._widgets = {}
        self._build()
        if fattura: self._populate(fattura)
        self.grab_set()

    def _build(self):
        self.configure(bg=C["surface"])
        hdr = tk.Frame(self, bg=C["bg"], padx=16, pady=12)
        hdr.pack(fill="x")
        lbl = "Nuova Fattura Passiva (Fornitore)" if self._tipo == "passivo" else "Nuova Fattura Attiva (Cliente)"
        if self._fattura: lbl = f"Modifica Documento ID #{self._fattura[0]}"
        tk.Label(hdr, text=lbl, font=("Segoe UI", 11, "bold"), bg=C["bg"], fg=C["text"]).pack(anchor="w")
        
        fr = tk.Frame(self, bg=C["surface"], padx=20, pady=12)
        fr.pack()
        
        stati = STATI_PASSIVO if self._tipo == "passivo" else STATI_ATTIVO
        fields = [("Numero Fattura", "numero"), ("Ragione Sociale", "anagrafica"), ("Tipo Documento", "tipo_doc"),
                  ("Data Emissione", "data_doc"), ("Scadenza Netta", "data_scad"), ("Importo Lordo €", "importo"),
                  ("Volume Regolato €", "pagato"), ("Stato Corrente", "stato"), ("Canale Preferito", "metodo"), ("Note / Tag", "note")]
                  
        for i, (lbl_t, key) in enumerate(fields):
            tk.Label(fr, text=lbl_t, font=("Segoe UI", 9), bg=C["surface"], fg=C["muted"]).grid(row=i, column=0, sticky="w", pady=4, padx=(0,12))
            if key in ("data_doc", "data_scad"): w = DateEntry(fr, bg=C["surface"])
            elif key in ("importo", "pagato"): w = ImportoEntry(fr)
            elif key == "stato":
                w = ttk.Combobox(fr, values=stati, width=22, state="readonly")
                w.set(stati[0])
                w.bind("<<ComboboxSelected>>", self._on_stato_changed)
            elif key == "metodo": w = ttk.Combobox(fr, values=METODI_PAG, width=22)
            elif key == "tipo_doc":
                w = ttk.Combobox(fr, values=TIPI_DOC, width=22, state="readonly")
                w.set("Fattura")
            else: w = tk.Entry(fr, font=("Segoe UI", 10), relief="solid", bd=1, width=24)
            w.grid(row=i, column=1, sticky="w", pady=4, ipady=1 if key not in ("data_doc","data_scad","stato","metodo","tipo_doc") else 0)
            self._widgets[key] = w

        bf = tk.Frame(self, bg=C["surface"], padx=20, pady=14)
        bf.pack(fill="x")
        tk.Button(bf, text="Salva Scheda", command=self._save, font=("Segoe UI", 9, "bold"), bg=C["accent"], fg="white", relief="flat", padx=16, pady=6).pack(side="right")
        tk.Button(bf, text="Esci", command=self.destroy, font=("Segoe UI", 9, "bold"), bg=C["border"], fg=C["text"], relief="flat", padx=16, pady=6).pack(side="right", padx=4)

    def _on_stato_changed(self, event=None):
        stato_sel = self._widgets["stato"].get()
        if stato_sel in ("Pagata", "Incassata"):
            totale = self._widgets["importo"].get_float()
            if totale > 0: self._widgets["pagato"].set_float(totale)

    def _populate(self, f):
        def _set(k, v):
            w = self._widgets[k]
            if isinstance(w, ImportoEntry): w.set_float(v)
            elif isinstance(w, DateEntry): w.set(v or "")
            elif isinstance(w, ttk.Combobox): w.set(v or "")
            else: w.delete(0, tk.END); w.insert(0, v or "")
        _set("numero", f[2]); _set("anagrafica", f[3]); _set("data_doc", f[4]); _set("data_scad", f[5])
        _set("importo", f[6]); _set("pagato", f[7]); _set("stato", f[8]); _set("metodo", f[9])
        _set("tipo_doc", f[10]); _set("note", f[11])

    def _save(self):
        num = self._widgets["numero"].get().strip()
        ana = self._widgets["anagrafica"].get().strip()
        if not num or not ana: return
        
        with get_conn() as conn:
            c = conn.cursor()
            vals = (self._tipo, num, ana, self._widgets["data_doc"].get(), self._widgets["data_scad"].get(),
                    self._widgets["importo"].get_float(), self._widgets["pagato"].get_float(), self._widgets["stato"].get(),
                    self._widgets["metodo"].get(), self._widgets["tipo_doc"].get(), self._widgets["note"].get().strip())
            if self._fattura:
                c.execute("""UPDATE fatture SET tipo=?, numero=?, anagrafica=?, data_doc=?, data_scad=?,
                           importo=?, pagato=?, stato=?, metodo=?, tipo_doc=?, note=? WHERE id=?""", vals + (self._fattura[0],))
            else:
                c.execute("""INSERT INTO fatture (tipo, numero, anagrafica, data_doc, data_scad, importo, pagato, stato, metodo, tipo_doc, note)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", vals)
                                  
        if self._on_save: self._on_save()
        self.destroy()

class ImportWindow(tk.Toplevel):
    def __init__(self, master, on_done=None):
        super().__init__(master)
        self.title("Modulo Importazione")
        self.geometry("450x260")
        self.resizable(False, False)
        self._on_done = on_done
        self.configure(bg=C["surface"])
        self._build()
        self.grab_set()

    def _build(self):
        hdr = tk.Frame(self, bg=C["bg"], padx=16, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Acquisizione Tracciati Esterni", fg=C["text"], bg=C["bg"], font=("Segoe UI", 11, "bold")).pack(anchor="w")
        
        fr = tk.Frame(self, bg=C["surface"], padx=20, pady=16)
        fr.pack(fill="both", expand=True)
        
        self._path_var = tk.StringVar()
        tk.Entry(fr, textvariable=self._path_var, width=30, relief="solid", bd=1).grid(row=0, column=0, padx=2, ipady=2)
        tk.Button(fr, text="Sfoglia...", command=self._browse, bg=C["border"], fg=C["text"], font=("Segoe UI", 9, "bold"), relief="flat", padx=10).grid(row=0, column=1, padx=4)
        
        self._tipo = tk.StringVar(value="passivo")
        ttk.Combobox(fr, textvariable=self._tipo, values=["passivo (fornitori)", "attivo (clienti)"], width=24, state="readonly").grid(row=1, column=0, sticky="w", pady=12)
        
        self._log = tk.Label(fr, text="", font=("Segoe UI", 9, "bold"), bg=C["surface"], fg=C["green"])
        self._log.grid(row=2, columnspan=2, pady=4, sticky="w")
        
        tk.Button(fr, text="Esegui Importazione", command=self._import, bg=C["accent"], fg="white", font=("Segoe UI", 10, "bold"), relief="flat", pady=6).grid(row=3, columnspan=2, sticky="ew")

    def _browse(self):
        p = filedialog.askopenfilename(filetypes=[("Excel/CSV Assets", "*.xlsx *.xls *.csv")])
        if p: self._path_var.set(p)

    def _import(self):
        path = self._path_var.get().strip()
        if not path or not os.path.exists(path): return
        t = "passivo" if "passivo" in self._tipo.get() else "attivo"
        ext = os.path.splitext(path)[1].lower()
        try:
            ins, sal = self._process_xlsx(path, t) if ext in (".xlsx", ".xls") else self._process_csv(path, t)
            self._log.config(text=f"Import completato. Inseriti: {ins} | Saltati: {sal}")
            if self._on_done: self._on_done()
        except Exception as e: messagebox.showerror("Errore Import", str(e))

    def _process_xlsx(self, path, t):
        if not OPENPYXL_OK: return 0, 0
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows: return 0, 0
        hd = [str(x).strip() if x else "" for x in rows[0]]
        ins = sal = 0
        with get_conn() as conn:
            c = conn.cursor()
            for r in rows[1:]:
                rec = {hd[i]: r[i] for i in range(min(len(hd), len(r)))}
                num = str(rec.get("Numero", "") or "").strip()
                ana = str(rec.get("Fornitore", rec.get("Cliente", "")) or "").strip()
                if not num or not ana: continue
                if c.execute("SELECT id FROM fatture WHERE numero=? AND anagrafica=? AND tipo=?", (num, ana, t)).fetchone():
                    sal += 1; continue
                dd = parse_date(rec.get("Data"))
                ds = calcola_scadenza_automatico(c, ana, dd)
                imp = parse_importo(rec.get("Totale"))
                td = str(rec.get("Tipo documento") or "Fattura").strip()
                if "differita" in td.lower(): td = "Fattura differita"
                st = mappa_stato_fte(rec.get("Stato FTE"), t)
                c.execute("INSERT INTO fatture (tipo,numero,anagrafica,data_doc,data_scad,importo,pagato,stato,tipo_doc) VALUES (?,?,?,?,?,?,?,?,?)",
                          (t, num, ana, dd, ds, imp, 0.0, st, td))
                ins += 1
        return ins, sal

    def _process_csv(self, path, t):
        ins = sal = 0
        with open(path, newline="", encoding="utf-8-sig") as f:
            rdr = csv.DictReader(f)
            with get_conn() as conn:
                c = conn.cursor()
                for rec in rdr:
                    num = str(rec.get("Numero", "") or "").strip()
                    ana = str(rec.get("Fornitore", rec.get("Cliente", "")) or "").strip()
                    if not num or not ana: continue
                    if c.execute("SELECT id FROM fatture WHERE numero=? AND anagrafica=? AND tipo=?", (num, ana, t)).fetchone():
                        sal += 1; continue
                    dd = parse_date(rec.get("Data"))
                    ds = calcola_scadenza_automatico(c, ana, dd)
                    imp = parse_importo(rec.get("Totale"))
                    td = str(rec.get("Tipo documento") or "Fattura").strip()
                    if "differita" in td.lower(): td = "Fattura differita"
                    st = mappa_stato_fte(rec.get("Stato FTE"), t)
                    c.execute("INSERT INTO fatture (tipo,numero,anagrafica,data_doc,data_scad,importo,pagato,stato,tipo_doc) VALUES (?,?,?,?,?,?,?,?,?)",
                              (t, num, ana, dd, ds, imp, 0.0, st, td))
                    ins += 1
        return ins, sal


# ─── APPLICAZIONE CORE (VIDALOCA STYLING) ────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Scadenziario Cash Flow 2026 — VIDALOCA")
        self.geometry("1340x800") 
        self.configure(bg=C["bg"])
        init_db()
        
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=28, background=C["surface"], fieldbackground=C["surface"], foreground=C["text"])
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background=C["border"], foreground=C["text"], relief="flat")
        style.map("Treeview", background=[("selected", C["accent"])], foreground=[("selected", "white")])
        
        self._view = "dashboard" 
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._load_data())
        self._sort_col, self._sort_rev = "data_scad", False
        self._current_tag_filter = "tutti"  
        
        self._build_sidebar()
        
        self.main_container = tk.Frame(self, bg=C["bg"])
        self.main_container.pack(side="left", fill="both", expand=True)
        
        self.dashboard_frame = tk.Frame(self.main_container, bg=C["bg"])
        self.table_frame = tk.Frame(self.main_container, bg=C["bg"])
        
        self._build_table_layout()
        self._set_view("dashboard")

    def _build_sidebar(self):
        # Sidebar con colore scuro coordinato ai testi principali del logo
        sb = tk.Frame(self, bg=C["text"], width=230) 
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)
        
        # Area Brand Logo (Posizionata in alto a sinistra nella Sidebar)
        brand_fr = tk.Frame(sb, bg=C["text"], pady=16)
        brand_fr.pack(fill="x")
        
        # Caricamento dinamico del logo dall'URL web fornito
        loaded_logo = False
        if PILLOW_OK:
            try:
                # Richiesta HTTP con User-Agent per evitare blocchi server standard
                req = urllib.request.Request(LOGO_URL, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=5) as response:
                    raw_data = response.read()
                
                img_data = io.BytesIO(raw_data)
                img = Image.open(img_data)
                # Adatta il logo in larghezza mantenendo le proporzioni corrette
                img.thumbnail((200, 110), Image.Resampling.LANCZOS)
                
                self._logo_img = ImageTk.PhotoImage(img)
                lbl_logo = tk.Label(brand_fr, image=self._logo_img, bg=C["text"])
                lbl_logo.pack()
                loaded_logo = True
            except Exception:
                # Se manca connessione o l'URL fallisce, passa silenziosamente al testo alternativo
                pass
                
        if not loaded_logo:
            # Rimpiazzo testuale elegante di backup
            tk.Label(brand_fr, text="VIDALOCA", font=("Segoe UI", 16, "bold"), bg=C["text"], fg=C["accent"]).pack()
            tk.Label(brand_fr, text="di Michela Vidale", font=("Segoe UI", 9, "italic"), bg=C["text"], fg=C["muted"]).pack(pady=(2,0))
        
        sep_top = tk.Frame(sb, bg=C["muted"], height=1)
        sep_top.pack(fill="x", padx=16, pady=(10, 14))

        self._nav_btns = {}
        navs = [
            ("dashboard", "🏠   Dashboard"), 
            ("passivo", "📉   Registro Passivo"),
            ("attivo", "📈   Registro Attivo"), 
            ("tutte", "📋   Tutti i Documenti")
        ]
        for key, text in navs:
            b = tk.Button(sb, text=text, font=("Segoe UI", 10), bg=C["text"], fg="#dcd1c4", anchor="w",
                          relief="flat", padx=20, pady=10, cursor="hand2", activebackground=C["accent"], activeforeground="white")
            b.config(command=lambda k=key: self._set_view(k))
            b.pack(fill="x", padx=8, pady=2)
            self._nav_btns[key] = b
            
        sep = tk.Frame(sb, bg=C["muted"], height=1)
        sep.pack(fill="x", padx=16, pady=16)
        
        tk.Button(sb, text="🏢   Rubrica Aziende", font=("Segoe UI", 9), bg=C["text"], fg="#dcd1c4", anchor="w", relief="flat", padx=20, pady=6, command=lambda: AziendeWindow(self)).pack(fill="x", padx=8)
        tk.Button(sb, text="📥   Import Gestionale", font=("Segoe UI", 9), bg=C["text"], fg="#dcd1c4", anchor="w", relief="flat", padx=20, pady=6, command=lambda: ImportWindow(self, self._refresh_current_view)).pack(fill="x", padx=8)
        tk.Button(sb, text="⚙️   Impostazioni", font=("Segoe UI", 9), bg=C["text"], fg="#dcd1c4", anchor="w", relief="flat", padx=20, pady=6, command=lambda: ConfigCondizioniWindow(self, on_close_callback=self._refresh_current_view)).pack(fill="x", padx=8)

        # Spazio flessibile per ancorare la firma sul fondo
        spacer = tk.Frame(sb, bg=C["text"])
        spacer.pack(fill="both", expand=True)

        # ─── FOOTER SIDEBAR CON LINK ATTIVO (In basso a sinistra) ───────────
        footer_fr = tk.Frame(sb, bg=C["text"], pady=12)
        footer_fr.pack(fill="x", side="bottom")
        
        lbl_f1 = tk.Label(footer_fr, text="Fatto con ❤️ da ", font=("Segoe UI", 8), bg=C["text"], fg="#bfae9e")
        lbl_f1.pack(side="left", padx=(16, 0))
        
        lbl_link = tk.Label(footer_fr, text="Giovanni Pio", font=("Segoe UI", 8, "bold", "underline"), bg=C["text"], fg=C["accent"], cursor="hand2")
        lbl_link.pack(side="left")
        
        # Interazioni mouse per il link
        lbl_link.bind("<Button-1>", lambda _: webbrowser.open("https://familiarigiovannipio.it"))
        lbl_link.bind("<Enter>", lambda _: lbl_link.config(fg="white"))
        lbl_link.bind("<Leave>", lambda _: lbl_link.config(fg=C["accent"]))

    def _build_table_layout(self):
        self._kpi_frame = tk.Frame(self.table_frame, bg=C["bg"])
        self._kpi_frame.pack(fill="x", padx=24, pady=(20, 10))
        
        tb = tk.Frame(self.table_frame, bg=C["bg"])
        tb.pack(fill="x", padx=24, pady=8)
        
        search_fr = tk.Frame(tb, bg="white", bd=1, relief="solid", highlightthickness=0)
        search_fr.pack(side="left", ipady=2)
        search_fr.config(highlightbackground=C["border"], highlightcolor=C["accent"])
        tk.Label(search_fr, text="  🔍  ", bg="white", fg=C["muted"]).pack(side="left")
        tk.Entry(search_fr, textvariable=self._search_var, font=("Segoe UI", 10), bg="white", relief="flat", bd=0, width=28).pack(side="left", padx=4)
        
        tk.Button(tb, text="💥  Elimina", command=self._delete_invoice, bg=C["surface"], fg=C["red"], font=("Segoe UI", 9, "bold"), relief="solid", borderwidth=1, padx=14, pady=5).pack(side="right", padx=3)
        tk.Button(tb, text="✏️  Modifica", command=self._edit_invoice, bg=C["surface"], fg=C["text"], font=("Segoe UI", 9, "bold"), relief="solid", borderwidth=1, padx=14, pady=5).pack(side="right", padx=3)
        tk.Button(tb, text="➕  Nuovo Documento", command=self._new_invoice, bg=C["accent"], fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=16, pady=6).pack(side="right", padx=3)

        f = tk.Frame(self.table_frame, bg="white", bd=1, relief="solid")
        f.pack(fill="both", expand=True, padx=24, pady=8)
        
        cols = ("id", "tipo", "numero", "anagrafica", "tipo_doc", "data_doc", "data_scad", "importo", "pagato", "residuo", "stato")
        self.tree = ttk.Treeview(f, columns=cols, show="headings", selectmode="browse")
        
        hd = {"id": "ID", "tipo": "Reg.", "numero": "N. Fattura", "anagrafica": "Azienda / Ragione Sociale",
              "tipo_doc": "Tipo Doc.", "data_doc": "Data Doc.", "data_scad": "Scadenza", "importo": "Totale Lordo",
              "pagato": "Pagato", "residuo": "Residuo Aperto", "stato": "Stato"}
        for c, text in hd.items():
            self.tree.heading(c, text=text, command=lambda _c=c: self._sort(_c))
            
        self.tree.column("id", width=50, anchor="center")
        self.tree.column("tipo", width=65, anchor="center")
        self.tree.column("numero", width=110)
        self.tree.column("anagrafica", width=280)
        self.tree.column("tipo_doc", width=120)
        self.tree.column("data_doc", width=100, anchor="center")
        self.tree.column("data_scad", width=100, anchor="center")
        self.tree.column("importo", width=105, anchor="e")
        self.tree.column("pagato", width=105, anchor="e")
        self.tree.column("residuo", width=105, anchor="e")
        self.tree.column("stato", width=105, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)
        
        sb = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        
        self.tree.tag_configure("scaduta", background=C["scaduta"], foreground=C["scaduta_fg"])
        self.tree.tag_configure("urgente", background=C["urgente"], foreground=C["urgente_fg"])
        self.tree.tag_configure("ok", background=C["ok"], foreground=C["ok_fg"])
        
        self.tree.bind("<Double-1>", lambda _: self._edit_invoice())

        # Legenda Filtri Attiva
        leg = tk.Frame(self.table_frame, bg=C["bg"])
        leg.pack(fill="x", padx=24, pady=10)
        
        tk.Label(leg, text="Filtra tabella per stato: ", font=("Segoe UI", 9, "italic"), bg=C["bg"], fg=C["muted"]).pack(side="left", padx=(0, 4))
        
        self.btn_f_scaduto = tk.Button(leg, text="  Scaduto  ", font=("Segoe UI", 8, "bold"), bg=C["scaduta"], fg=C["scaduta_fg"], bd=1, relief="solid", cursor="hand2", command=lambda: self._set_tag_filter("scaduta"))
        self.btn_f_scaduto.pack(side="left", padx=3, ipady=1)
        
        self.btn_f_urgente = tk.Button(leg, text="  In scadenza (7gg)  ", font=("Segoe UI", 8, "bold"), bg=C["urgente"], fg=C["urgente_fg"], bd=1, relief="solid", cursor="hand2", command=lambda: self._set_tag_filter("urgente"))
        self.btn_f_urgente.pack(side="left", padx=3, ipady=1)
        
        self.btn_f_ok = tk.Button(leg, text="  Chiuso / Saldato  ", font=("Segoe UI", 8, "bold"), bg=C["ok"], fg=C["ok_fg"], bd=1, relief="solid", cursor="hand2", command=lambda: self._set_tag_filter("ok"))
        self.btn_f_ok.pack(side="left", padx=3, ipady=1)
        
        self.btn_f_tutti = tk.Button(leg, text="  ❌ Mostra Tutti  ", font=("Segoe UI", 8, "bold"), bg=C["surface"], fg=C["text"], bd=1, relief="solid", cursor="hand2", command=lambda: self._set_tag_filter("tutti"))
        self.btn_f_tutti.pack(side="left", padx=(12, 3), ipady=1)

    def _set_tag_filter(self, tag_name):
        self._current_tag_filter = tag_name
        for b, name in [(self.btn_f_scaduto, "scaduta"), (self.btn_f_urgente, "urgente"), (self.btn_f_ok, "ok"), (self.btn_f_tutti, "tutti")]:
            if name == tag_name:
                b.config(relief="sunken", borderwidth=2)
            else:
                b.config(relief="solid", borderwidth=1)
        self._load_data()

    def _set_view(self, view_name):
        self._view = view_name
        self._current_tag_filter = "tutti" 
        for k, btn in self._nav_btns.items():
            if k == view_name:
                btn.config(bg=C["accent"], fg="white")
            else:
                btn.config(bg=C["text"], fg="#dcd1c4")
            
        if view_name == "dashboard":
            self.table_frame.pack_forget()
            self.dashboard_frame.pack(fill="both", expand=True)
            self._load_dashboard()
        else:
            self.dashboard_frame.pack_forget()
            self.table_frame.pack(fill="both", expand=True)
            for b in (self.btn_f_scaduto, self.btn_f_urgente, self.btn_f_ok): b.config(relief="solid", borderwidth=1)
            self.btn_f_tutti.config(relief="sunken", borderwidth=2)
            self._load_data()

    def _refresh_current_view(self):
        if self._view == "dashboard": self._load_dashboard()
        else: self._load_data()

    def _load_dashboard(self):
        for w in self.dashboard_frame.winfo_children(): w.destroy()
        
        header_fr = tk.Frame(self.dashboard_frame, bg=C["bg"], padx=24, pady=16)
        header_fr.pack(fill="x")
        tk.Label(header_fr, text="CRUSCOTTO DIREZIONALE FLUSSI DI CASSA — VIDALOCA", font=("Segoe UI", 14, "bold"), bg=C["bg"], fg=C["text"]).pack(anchor="w")
        tk.Label(header_fr, text="Analisi flussi e scadenze imminenti entro 7 giorni.", font=("Segoe UI", 9), bg=C["bg"], fg=C["muted"]).pack(anchor="w")
        
        today_str = date.today().strftime("%Y-%m-%d")
        prox_7_giorni = (date.today() + timedelta(days=7)).strftime("%Y-%m-%d")
        
        scaduto_fornitori = 0.0
        scaduto_clienti = 0.0
        in_scadenza_totale = 0.0
        
        with get_conn() as conn:
            tutte_non_saldate = conn.execute("SELECT tipo, importo, pagato, data_scad, tipo_doc FROM fatture WHERE stato NOT IN ('Pagata', 'Incassata')").fetchall()
            
        for tipo, importo, pagato, data_scad, tipo_doc in tutte_non_saldate:
            if tipo_doc == "Nota di credito": continue 
            residuo = importo - pagato
            if residuo <= 0: continue
            
            if data_scad and data_scad < today_str:
                if tipo == "passivo": scaduto_fornitori += residuo
                else: scaduto_clienti += residuo
            elif data_scad and data_scad <= prox_7_giorni:
                in_scadenza_totale += residuo

        kpi_container = tk.Frame(self.dashboard_frame, bg=C["bg"], padx=20)
        kpi_container.pack(fill="x", pady=6)
        
        box_data = [
            ("⚠️  SCADUTO PASSIVO (Fornitori)", f"{fmt_num(scaduto_fornitori)} €", C["scaduta_fg"] if scaduto_fornitori > 0 else C["text"]),
            ("💰  SCADUTO ATTIVO (Clienti)", f"{fmt_num(scaduto_clienti)} €", C["teal_accent"] if scaduto_clienti > 0 else C["text"]),
            ("📅  IN SCADENZA IMMINENTE (7gg)", f"{fmt_num(in_scadenza_totale)} €", C["accent_dark"])
        ]
        
        for titolo, valore, colore in box_data:
            box = tk.Frame(kpi_container, bg="white", bd=1, relief="solid", borderwidth=1, padx=16, pady=16)
            box.pack(side="left", expand=True, fill="x", padx=6)
            tk.Label(box, text=titolo, font=("Segoe UI", 9, "bold"), bg="white", fg=C["muted"]).pack(anchor="w")
            tk.Label(box, text=valore, font=("Segoe UI", 16, "bold"), bg="white", fg=colore).pack(anchor="w", pady=(6,0))

        action_title_fr = tk.Frame(self.dashboard_frame, bg=C["bg"], padx=24)
        action_title_fr.pack(fill="x", pady=(24, 6))
        
        tk.Label(action_title_fr, text="ATTENZIONI RICHIESTE (Scadute o Scadenze a 7 Giorni)", font=("Segoe UI", 11, "bold"), bg=C["bg"], fg=C["text"]).pack(side="left")
        lbl_info = tk.Label(action_title_fr, text=" (Doppio clic per modificare)", font=("Segoe UI", 9, "italic"), bg=C["bg"], fg=C["muted"])
        lbl_info.pack(side="left", pady=(1, 0))
        
        f_dash = tk.Frame(self.dashboard_frame, bg="white", bd=1, relief="solid")
        f_dash.pack(fill="both", expand=True, padx=24, pady=4)
        
        cols_dash = ("db_id", "tipo", "numero", "anagrafica", "tipo_doc", "data_scad", "residuo", "stato")
        self.tree_dash = ttk.Treeview(f_dash, columns=cols_dash, show="headings", selectmode="browse")
        
        hd_dash = {"db_id": "ID", "tipo": "Registro", "numero": "N. Documento", "anagrafica": "Ragione Sociale Azienda", "tipo_doc": "Tipo", "data_scad": "Scadenza", "residuo": "Residuo €", "stato": "Stato"}
        for cd, txtd in hd_dash.items():
            self.tree_dash.heading(cd, text=txtd)
            
        self.tree_dash.column("db_id", width=50, anchor="center")
        self.tree_dash.column("tipo", width=90, anchor="center")
        self.tree_dash.column("numero", width=120)
        self.tree_dash.column("anagrafica", width=360)
        self.tree_dash.column("tipo_doc", width=120)
        self.tree_dash.column("data_scad", width=110, anchor="center")
        self.tree_dash.column("residuo", width=140, anchor="e")
        self.tree_dash.column("stato", width=110, anchor="center")
        self.tree_dash.pack(side="left", fill="both", expand=True)
        
        sb_d = ttk.Scrollbar(f_dash, orient="vertical", command=self.tree_dash.yview)
        self.tree_dash.configure(yscrollcommand=sb_d.set)
        sb_d.pack(side="right", fill="y")
        
        self.tree_dash.tag_configure("scaduta", background=C["scaduta"], foreground=C["scaduta_fg"])
        self.tree_dash.tag_configure("urgente", background=C["urgente"], foreground=C["urgente_fg"])

        with get_conn() as conn:
            rows_dash = conn.execute("""SELECT id, tipo, numero, anagrafica, tipo_doc, data_scad, importo, pagato, stato 
                            FROM fatture WHERE stato NOT IN ('Pagata', 'Incassata') AND data_scad <= ? ORDER BY data_scad ASC""", (prox_7_giorni,)).fetchall()
            
        for r in rows_dash:
            fid, tp, num, ana, t_doc, d_scad, imp, pag, st = r
            res_val = imp - pag
            if res_val <= 0: continue
            tag = "scaduta" if d_scad and d_scad < today_str else "urgente"
            self.tree_dash.insert("", "end", values=(fid, tp.upper(), num, ana, t_doc, fmt_date(d_scad), fmt_num(res_val), st), tags=(tag,))

        self.tree_dash.bind("<Double-1>", lambda _: self._edit_invoice_from_dashboard())

    def _sort(self, col):
        if self._sort_col == col: self._sort_rev = not self._sort_rev
        else: self._sort_col, self._sort_rev = col, False
        self._load_data()

    def _load_data(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        query = "SELECT id, tipo, numero, anagrafica, data_doc, data_scad, importo, pagato, stato, metodo, tipo_doc, note FROM fatture WHERE 1=1"
        params = []
        
        if self._view in ("passivo", "attivo"):
            query += " AND tipo = ?"
            params.append(self._view)
            
        search = self._search_var.get().strip()
        if search:
            query += " AND (numero LIKE ? OR anagrafica LIKE ? OR note LIKE ?)"
            lk = f"%{search}%"
            params.extend([lk, lk, lk])
            
        if self._sort_col:
            direction = "DESC" if self._sort_rev else "ASC"
            query += f" ORDER BY {self._sort_col} {direction}"
            
        with get_conn() as conn:
            rows = conn.execute(query, params).fetchall()
            
        kpi_dovuto = 0.0
        kpi_scaduto = 0.0
        kpi_saldate = 0
        today_s = date.today().strftime("%Y-%m-%d")
        limit_u = (date.today() + timedelta(days=7)).strftime("%Y-%m-%d")
        
        for r in rows:
            fid, tipo, numero, anagrafica, data_doc, data_scad, importo, pagato, stato, metodo, tipo_doc, note = r
            residuo = max(0.0, importo - pagato)
            
            if tipo_doc != "Nota di credito":
                kpi_dovuto += residuo
                if stato in ("Pagata", "Incassata"): kpi_saldate += 1
                elif data_scad and data_scad < today_s: kpi_scaduto += residuo
                    
            tag = ""
            if stato in ("Pagata", "Incassata"): tag = "ok"
            elif data_scad and data_scad < today_s: tag = "scaduta"
            elif data_scad and data_scad <= limit_u: tag = "urgente"
            
            if self._current_tag_filter != "tutti" and tag != self._current_tag_filter:
                continue
                
            display = (fid, tipo.upper(), numero, anagrafica, tipo_doc, fmt_date(data_doc), fmt_date(data_scad), fmt_num(importo), fmt_num(pagato), fmt_num(residuo), stato)
            self.tree.insert("", "end", values=display, tags=(tag,))
            
        self._render_table_kpis(kpi_dovuto, kpi_scaduto, len(rows), kpi_saldate)

    def _render_table_kpis(self, dovuto, scaduto, totali, saldate):
        for w in self._kpi_frame.winfo_children(): w.destroy()
        lbl_txt = "Volume Aperto" if self._view == "tutte" else ("Esposizione Fornitori" if self._view == "passivo" else "Massa Incassi Attesa")
        
        kpis = [
            ("Elementi In Elenco", str(totali), C["muted"]),
            (lbl_txt, f"{fmt_num(dovuto)} €", C["teal_accent"]),
            ("Di cui Scaduto Effettivo ⚠️", f"{fmt_num(scaduto)} €", C["scaduta_fg"] if scaduto > 0 else C["muted"]),
            ("Partite Saldate", str(saldate), C["green"])
        ]
        for title, val, color in kpis:
            f = tk.Frame(self._kpi_frame, bg="white", bd=1, relief="solid", padx=14, pady=10)
            f.pack(side="left", expand=True, fill="x", padx=4)
            tk.Label(f, text=title, font=("Segoe UI", 8, "bold"), bg="white", fg=C["muted"]).pack(anchor="w")
            tk.Label(f, text=val, font=("Segoe UI", 13, "bold"), bg="white", fg=color).pack(anchor="w", pady=(4,0))

    def _get_selected_invoice(self, tree_widget):
        sel = tree_widget.selection()
        if not sel: return None
        fid = tree_widget.item(sel[0], "values")[0]
        with get_conn() as conn:
            return conn.execute("SELECT * FROM fatture WHERE id=?", (fid,)).fetchone()

    def _new_invoice(self):
        FatturaForm(self, tipo=self._view if self._view in ("passivo", "attivo") else "passivo", on_save=self._refresh_current_view)

    def _edit_invoice(self):
        f = self._get_selected_invoice(self.tree)
        if not f:
            messagebox.showwarning("Selezione mancante", "Scegli una riga dal registro per apportare modifiche.")
            return
        FatturaForm(self, tipo=f[1], fattura=f, on_save=self._refresh_current_view)

    def _edit_invoice_from_dashboard(self):
        f = self._get_selected_invoice(self.tree_dash)
        if not f: return
        FatturaForm(self, tipo=f[1], fattura=f, on_save=self._refresh_current_view)

    def _delete_invoice(self):
        f = self._get_selected_invoice(self.tree)
        if not f:
            messagebox.showwarning("Selezione mancante", "Seleziona la riga da cancellare.")
            return
        if messagebox.askyesno("Conferma", f"Eliminare la fattura n. {f[2]}?"):
            with get_conn() as conn:
                conn.execute("DELETE FROM pagamenti WHERE fattura_id=?", (f[0],))
                conn.execute("DELETE FROM fatture WHERE id=?", (f[0],))
            self._load_data()


if __name__ == "__main__":
    app = App()
    app.mainloop()
