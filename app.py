import os
import sqlite3
from datetime import datetime, date
import tkinter as tk
from tkinter import ttk, messagebox

DB_NAME = "presenze.db"

# ------------------ UTIL TIME ------------------
def parse_hhmm(s: str):
    s = (s or "").strip()
    if not s:
        return None
    try:
        parts = s.split(":")
        if len(parts) != 2:
            return None
        h = int(parts[0])
        m = int(parts[1])
        if h < 0 or h > 23 or m < 0 or m > 59:
            return None
        return h * 60 + m
    except:
        return None

def minutes_to_hhmm(mins: int):
    if mins is None:
        return "0:00"
    h = mins // 60
    m = mins % 60
    return f"{h}:{m:02d}"

def calc_work_minutes(entrata, uscita_pausa, rientro, uscita):
    e = parse_hhmm(entrata)
    u = parse_hhmm(uscita)
    if e is None or u is None:
        return 0

    up = parse_hhmm(uscita_pausa)
    rp = parse_hhmm(rientro)

    if up is None or rp is None:
        mins = u - e
        return max(0, mins)

    mins = (up - e) + (u - rp)
    return max(0, mins)

def month_bounds(yyyy_mm: str):
    y, m = map(int, yyyy_mm.split("-"))
    start = date(y, m, 1)
    if m == 12:
        end = date(y + 1, 1, 1)
    else:
        end = date(y, m + 1, 1)
    return start.isoformat(), end.isoformat()

def year_bounds(yyyy: int):
    start = date(yyyy, 1, 1).isoformat()
    end = date(yyyy + 1, 1, 1).isoformat()
    return start, end

# ------------------ DB ------------------
def db_connect():
    return sqlite3.connect(DB_NAME)

def db_init():
    con = db_connect()
    cur = con.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS dipendenti(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            cognome TEXT NOT NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS presenze(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            dipendente_id INTEGER NOT NULL,
            data TEXT NOT NULL,
            tipo TEXT NOT NULL DEFAULT 'Lavoro', -- Lavoro/Ferie/Malattia/Permesso
            entrata TEXT,
            uscita_pausa TEXT,
            rientro TEXT,
            uscita TEXT,
            minuti_lavorati INTEGER NOT NULL DEFAULT 0,
            note TEXT,
            FOREIGN KEY(dipendente_id) REFERENCES dipendenti(id)
        )
    """)

    con.commit()
    con.close()

# ---- Dipendenti ----
def db_add_employee(nome, cognome):
    con = db_connect()
    cur = con.cursor()
    cur.execute("INSERT INTO dipendenti(nome,cognome) VALUES(?,?)", (nome, cognome))
    con.commit()
    con.close()

def db_update_employee(emp_id, nome, cognome):
    con = db_connect()
    cur = con.cursor()
    cur.execute("UPDATE dipendenti SET nome=?, cognome=? WHERE id=?", (nome, cognome, emp_id))
    con.commit()
    con.close()

def db_delete_employee(emp_id):
    con = db_connect()
    cur = con.cursor()
    cur.execute("DELETE FROM presenze WHERE dipendente_id=?", (emp_id,))
    cur.execute("DELETE FROM dipendenti WHERE id=?", (emp_id,))
    con.commit()
    con.close()

def db_list_employees():
    con = db_connect()
    cur = con.cursor()
    cur.execute("SELECT id, nome, cognome FROM dipendenti ORDER BY cognome, nome")
    rows = cur.fetchall()
    con.close()
    return rows

# ---- Presenze ----
def db_add_presence(dip_id, data_str, tipo, entrata, uscita_pausa, rientro, uscita, minuti, note):
    con = db_connect()
    cur = con.cursor()
    cur.execute("""
        INSERT INTO presenze(dipendente_id, data, tipo, entrata, uscita_pausa, rientro, uscita, minuti_lavorati, note)
        VALUES(?,?,?,?,?,?,?,?,?)
    """, (dip_id, data_str, tipo, entrata, uscita_pausa, rientro, uscita, minuti, note))
    con.commit()
    con.close()

def db_list_presences(dip_id, yyyy_mm):
    start, end = month_bounds(yyyy_mm)
    con = db_connect()
    cur = con.cursor()
    cur.execute("""
        SELECT id, data, tipo, entrata, uscita_pausa, rientro, uscita, minuti_lavorati, note
        FROM presenze
        WHERE dipendente_id = ?
          AND data >= ?
          AND data < ?
        ORDER BY data
    """, (dip_id, start, end))
    rows = cur.fetchall()
    con.close()
    return rows

def db_month_stats(dip_id, yyyy_mm):
    start, end = month_bounds(yyyy_mm)
    con = db_connect()
    cur = con.cursor()
    cur.execute("""
        SELECT
            COALESCE(SUM(minuti_lavorati),0) as tot,
            COALESCE(SUM(CASE WHEN minuti_lavorati>0 THEN 1 ELSE 0 END),0) as giorni_lavorati,
            COALESCE(AVG(CASE WHEN minuti_lavorati>0 THEN minuti_lavorati END),0) as media,
            COALESCE(SUM(CASE WHEN tipo='Ferie' THEN 1 ELSE 0 END),0) as ferie
        FROM presenze
        WHERE dipendente_id = ?
          AND data >= ?
          AND data < ?
    """, (dip_id, start, end))
    tot, giorni, media, ferie = cur.fetchone()
    con.close()
    return int(tot), int(giorni), int(round(media)), int(ferie)

def db_year_ferie(dip_id, yyyy: int):
    start, end = year_bounds(yyyy)
    con = db_connect()
    cur = con.cursor()
    cur.execute("""
        SELECT COALESCE(SUM(CASE WHEN tipo='Ferie' THEN 1 ELSE 0 END),0)
        FROM presenze
        WHERE dipendente_id = ?
          AND data >= ?
          AND data < ?
    """, (dip_id, start, end))
    ferie = cur.fetchone()[0]
    con.close()
    return int(ferie)

# ------------------ UI ------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Presenze ENG")
        self.geometry("1060x640")
        self.minsize(980, 580)

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("SubHeader.TLabel", font=("Segoe UI", 10))
        style.configure("Card.TLabelframe", padding=10)
        style.configure("Card.TLabelframe.Label", font=("Segoe UI", 10, "bold"))

        self._set_icon()

        self.emp_map = {}
        self.emp_reverse = {}
        self.emp_id_selected = None

        self._build_ui()
        self.refresh_employees()

    def _assets_path(self, filename):
        return os.path.join(os.path.dirname(__file__), "assets", filename)

    def _set_icon(self):
        ico_path = self._assets_path("icon.ico")
        png_path = self._assets_path("icon.png")

        try:
            if os.path.exists(ico_path):
                self.iconbitmap(ico_path)
        except:
            pass

        try:
            if os.path.exists(png_path):
                img = tk.PhotoImage(file=png_path)
                self.iconphoto(True, img)
                self._icon_ref = img
        except:
            pass

    def _load_logo_small(self, target_max_height=80):
        """
        Carica logo.png e lo riduce automaticamente con subsample()
        finché l'altezza è <= target_max_height.
        """
        path = self._assets_path("logo.png")
        if not os.path.exists(path):
            return None

        try:
            img = tk.PhotoImage(file=path)
        except:
            return None

        # riduzione progressiva (1,2,3,4,5...) finché rientra in altezza
        factor = 1
        h = img.height()
        while h / factor > target_max_height and factor < 10:
            factor += 1

        if factor > 1:
            img = img.subsample(factor, factor)

        return img

    def _build_ui(self):
        # HEADER (compatto) con logo in alto a destra
        header = ttk.Frame(self)
        header.pack(fill="x", padx=12, pady=(10, 6))

        left = ttk.Frame(header)
        left.pack(side="left", fill="x", expand=True)

        ttk.Label(left, text="Presenze ENG", style="Header.TLabel").pack(anchor="w")
        ttk.Label(left, text="Gestione dipendenti, presenze, ore e ferie", style="SubHeader.TLabel").pack(anchor="w")

        # Logo a destra, piccolo
        self._logo_ref = self._load_logo_small(target_max_height=160)
        if self._logo_ref is not None:
            ttk.Label(header, image=self._logo_ref).pack(side="right", padx=(10, 0))

        # NOTEBOOK
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.tab_emp = ttk.Frame(notebook)
        self.tab_pres = ttk.Frame(notebook)

        notebook.add(self.tab_emp, text="Dipendenti")
        notebook.add(self.tab_pres, text="Presenze")

        self._build_employees_tab()
        self._build_presences_tab()

    # -------- Dipendenti TAB --------
    def _build_employees_tab(self):
        main = ttk.Frame(self.tab_emp)
        main.pack(fill="both", expand=True, padx=8, pady=8)

        main.columnconfigure(0, weight=2)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        frm_list = ttk.Labelframe(main, text="Elenco dipendenti", style="Card.TLabelframe")
        frm_list.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        self.lst_emp = tk.Listbox(frm_list, height=18)
        self.lst_emp.pack(fill="both", expand=True, padx=6, pady=6)
        self.lst_emp.bind("<<ListboxSelect>>", self.on_emp_select)

        frm_edit = ttk.Labelframe(main, text="Gestione dipendente", style="Card.TLabelframe")
        frm_edit.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        ttk.Label(frm_edit, text="Nome").grid(row=0, column=0, sticky="w", padx=6, pady=(6, 2))
        self.ent_nome = ttk.Entry(frm_edit, width=26)
        self.ent_nome.grid(row=1, column=0, padx=6, pady=2)

        ttk.Label(frm_edit, text="Cognome").grid(row=2, column=0, sticky="w", padx=6, pady=(8, 2))
        self.ent_cognome = ttk.Entry(frm_edit, width=26)
        self.ent_cognome.grid(row=3, column=0, padx=6, pady=2)

        btns = ttk.Frame(frm_edit)
        btns.grid(row=4, column=0, padx=6, pady=10, sticky="ew")

        ttk.Button(btns, text="Aggiungi", command=self.add_employee).pack(fill="x", pady=3)
        ttk.Button(btns, text="Modifica", command=self.update_employee).pack(fill="x", pady=3)
        ttk.Button(btns, text="Elimina", command=self.delete_employee).pack(fill="x", pady=3)

        hint = ttk.Label(frm_edit, text="Seleziona un dipendente per modificarlo/eliminarlo.")
        hint.grid(row=5, column=0, padx=6, pady=(8, 6), sticky="w")

    def refresh_employees(self):
        self.lst_emp.delete(0, tk.END)
        employees = db_list_employees()

        self.emp_map.clear()
        self.emp_reverse.clear()

        labels = []
        for emp_id, nome, cognome in employees:
            label = f"{cognome} {nome}"
            labels.append(label)
            self.emp_map[label] = emp_id
            self.emp_reverse[emp_id] = (nome, cognome)
            self.lst_emp.insert(tk.END, f"{emp_id} - {label}")

        self.cmb_emp["values"] = labels
        if labels and not self.cmb_emp.get():
            self.cmb_emp.current(0)

    def on_emp_select(self, _evt=None):
        sel = self.lst_emp.curselection()
        if not sel:
            return
        text = self.lst_emp.get(sel[0])
        emp_id = int(text.split(" - ")[0])
        self.emp_id_selected = emp_id
        nome, cognome = self.emp_reverse.get(emp_id, ("", ""))
        self.ent_nome.delete(0, tk.END)
        self.ent_nome.insert(0, nome)
        self.ent_cognome.delete(0, tk.END)
        self.ent_cognome.insert(0, cognome)

    def add_employee(self):
        nome = self.ent_nome.get().strip()
        cognome = self.ent_cognome.get().strip()
        if not nome or not cognome:
            messagebox.showwarning("Errore", "Inserisci nome e cognome.")
            return
        db_add_employee(nome, cognome)
        self.emp_id_selected = None
        self.refresh_employees()
        messagebox.showinfo("OK", "Dipendente aggiunto.")

    def update_employee(self):
        if not self.emp_id_selected:
            messagebox.showwarning("Errore", "Seleziona un dipendente da modificare.")
            return
        nome = self.ent_nome.get().strip()
        cognome = self.ent_cognome.get().strip()
        if not nome or not cognome:
            messagebox.showwarning("Errore", "Inserisci nome e cognome.")
            return
        db_update_employee(self.emp_id_selected, nome, cognome)
        self.refresh_employees()
        messagebox.showinfo("OK", "Dipendente aggiornato.")

    def delete_employee(self):
        if not self.emp_id_selected:
            messagebox.showwarning("Errore", "Seleziona un dipendente da eliminare.")
            return
        if not messagebox.askyesno("Conferma", "Eliminare dipendente e tutte le sue presenze?"):
            return
        db_delete_employee(self.emp_id_selected)
        self.emp_id_selected = None
        self.ent_nome.delete(0, tk.END)
        self.ent_cognome.delete(0, tk.END)
        self.refresh_employees()
        messagebox.showinfo("OK", "Dipendente eliminato.")

    # -------- Presenze TAB --------
    def _build_presences_tab(self):
        top = ttk.Labelframe(self.tab_pres, text="Selezione", style="Card.TLabelframe")
        top.pack(fill="x", padx=8, pady=8)

        ttk.Label(top, text="Dipendente").grid(row=0, column=0, padx=6, pady=6, sticky="w")
        self.cmb_emp = ttk.Combobox(top, width=35, state="readonly")
        self.cmb_emp.grid(row=0, column=1, padx=6, pady=6)

        ttk.Label(top, text="Mese (YYYY-MM)").grid(row=0, column=2, padx=6, pady=6, sticky="w")
        self.ent_month = ttk.Entry(top, width=12)
        self.ent_month.grid(row=0, column=3, padx=6, pady=6)
        self.ent_month.insert(0, datetime.now().strftime("%Y-%m"))

        ttk.Button(top, text="Carica mese", command=self.load_month).grid(row=0, column=4, padx=10, pady=6)

        frm_ins = ttk.Labelframe(self.tab_pres, text="Inserisci presenza", style="Card.TLabelframe")
        frm_ins.pack(fill="x", padx=8, pady=8)

        ttk.Label(frm_ins, text="Data (YYYY-MM-DD)").grid(row=0, column=0, padx=6, pady=4, sticky="w")
        ttk.Label(frm_ins, text="Tipo").grid(row=0, column=1, padx=6, pady=4, sticky="w")
        ttk.Label(frm_ins, text="Entrata").grid(row=0, column=2, padx=6, pady=4, sticky="w")
        ttk.Label(frm_ins, text="Uscita pausa").grid(row=0, column=3, padx=6, pady=4, sticky="w")
        ttk.Label(frm_ins, text="Rientro").grid(row=0, column=4, padx=6, pady=4, sticky="w")
        ttk.Label(frm_ins, text="Uscita").grid(row=0, column=5, padx=6, pady=4, sticky="w")
        ttk.Label(frm_ins, text="Note").grid(row=0, column=6, padx=6, pady=4, sticky="w")

        self.ent_date = ttk.Entry(frm_ins, width=14)
        self.cmb_tipo = ttk.Combobox(frm_ins, width=12, state="readonly",
                                     values=["Lavoro", "Ferie", "Malattia", "Permesso"])
        self.cmb_tipo.set("Lavoro")

        self.ent_in = ttk.Entry(frm_ins, width=10)
        self.ent_up = ttk.Entry(frm_ins, width=10)
        self.ent_rp = ttk.Entry(frm_ins, width=10)
        self.ent_out = ttk.Entry(frm_ins, width=10)
        self.ent_note = ttk.Entry(frm_ins, width=30)

        self.ent_date.grid(row=1, column=0, padx=6, pady=6)
        self.cmb_tipo.grid(row=1, column=1, padx=6, pady=6)
        self.ent_in.grid(row=1, column=2, padx=6, pady=6)
        self.ent_up.grid(row=1, column=3, padx=6, pady=6)
        self.ent_rp.grid(row=1, column=4, padx=6, pady=6)
        self.ent_out.grid(row=1, column=5, padx=6, pady=6)
        self.ent_note.grid(row=1, column=6, padx=6, pady=6)

        self.ent_date.insert(0, datetime.now().strftime("%Y-%m-%d"))

        ttk.Button(frm_ins, text="Salva presenza", command=self.add_presence).grid(row=1, column=7, padx=10, pady=6)

        frm_tbl = ttk.Labelframe(self.tab_pres, text="Presenze del mese", style="Card.TLabelframe")
        frm_tbl.pack(fill="both", expand=True, padx=8, pady=8)

        cols = ("data", "tipo", "entrata", "uscita_pausa", "rientro", "uscita", "totale", "note")
        self.tree = ttk.Treeview(frm_tbl, columns=cols, show="headings", height=12)
        widths = {"data": 110, "tipo": 90, "entrata": 85, "uscita_pausa": 95, "rientro": 85, "uscita": 85, "totale": 85, "note": 320}
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=widths.get(c, 100), anchor="w")
        self.tree.pack(fill="both", expand=True, padx=6, pady=6)

        frm_stats = ttk.Labelframe(self.tab_pres, text="Statistiche mese", style="Card.TLabelframe")
        frm_stats.pack(fill="x", padx=8, pady=8)

        self.lbl_tot = ttk.Label(frm_stats, text="Totale: 0:00", font=("Segoe UI", 11, "bold"))
        self.lbl_giorni = ttk.Label(frm_stats, text="Giorni lavorati: 0", font=("Segoe UI", 11, "bold"))
        self.lbl_media = ttk.Label(frm_stats, text="Media: 0:00", font=("Segoe UI", 11, "bold"))
        self.lbl_ferie_mese = ttk.Label(frm_stats, text="Ferie mese: 0", font=("Segoe UI", 11, "bold"))
        self.lbl_ferie_anno = ttk.Label(frm_stats, text="Ferie anno: 0", font=("Segoe UI", 11, "bold"))

        self.lbl_tot.pack(side="left", padx=12, pady=8)
        self.lbl_giorni.pack(side="left", padx=12, pady=8)
        self.lbl_media.pack(side="left", padx=12, pady=8)
        self.lbl_ferie_mese.pack(side="left", padx=12, pady=8)
        self.lbl_ferie_anno.pack(side="left", padx=12, pady=8)

        hint = ttk.Label(frm_stats, text="Orari: H:MM (es. 8:30). Se Ferie/Malattia/Permesso puoi lasciare gli orari vuoti.")
        hint.pack(side="left", padx=12)

    def selected_employee_id(self):
        label = self.cmb_emp.get().strip()
        if not label or label not in self.emp_map:
            return None
        return self.emp_map[label]

    def add_presence(self):
        dip_id = self.selected_employee_id()
        if dip_id is None:
            messagebox.showwarning("Errore", "Seleziona un dipendente.")
            return

        data_str = self.ent_date.get().strip()
        try:
            datetime.strptime(data_str, "%Y-%m-%d")
        except:
            messagebox.showwarning("Errore", "Data non valida. Usa YYYY-MM-DD.")
            return

        tipo = self.cmb_tipo.get().strip() or "Lavoro"
        entrata = self.ent_in.get().strip()
        uscita_pausa = self.ent_up.get().strip()
        rientro = self.ent_rp.get().strip()
        uscita = self.ent_out.get().strip()
        note = self.ent_note.get().strip()

        minuti = 0 if tipo != "Lavoro" else calc_work_minutes(entrata, uscita_pausa, rientro, uscita)

        db_add_presence(dip_id, data_str, tipo, entrata, uscita_pausa, rientro, uscita, minuti, note)

        self.load_month()
        messagebox.showinfo("OK", f"Presenza salvata. Totale: {minutes_to_hhmm(minuti)}")

    def load_month(self):
        dip_id = self.selected_employee_id()
        if dip_id is None:
            return

        yyyy_mm = self.ent_month.get().strip()
        try:
            datetime.strptime(yyyy_mm + "-01", "%Y-%m-%d")
        except:
            messagebox.showwarning("Errore", "Mese non valido. Usa YYYY-MM (es. 2026-01).")
            return

        for row in self.tree.get_children():
            self.tree.delete(row)

        rows = db_list_presences(dip_id, yyyy_mm)
        for _pid, d, tipo, e, up, rp, u, mins, note in rows:
            self.tree.insert("", tk.END, values=(d, tipo, e, up, rp, u, minutes_to_hhmm(mins), note or ""))

        tot, giorni, media, ferie_mese = db_month_stats(dip_id, yyyy_mm)
        year = int(yyyy_mm.split("-")[0])
        ferie_anno = db_year_ferie(dip_id, year)

        self.lbl_tot.config(text=f"Totale: {minutes_to_hhmm(tot)}")
        self.lbl_giorni.config(text=f"Giorni lavorati: {giorni}")
        self.lbl_media.config(text=f"Media: {minutes_to_hhmm(media)}")
        self.lbl_ferie_mese.config(text=f"Ferie mese: {ferie_mese}")
        self.lbl_ferie_anno.config(text=f"Ferie anno: {ferie_anno}")


if __name__ == "__main__":
    db_init()
    App().mainloop()