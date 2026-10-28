import copy
import json
import os
import queue
import re
import shutil
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import win32com.client as win32
import pythoncom


# ======================
# CONFIGURATION DES MODES
# ======================
# Valeurs d'origine. La configuration modifiée depuis le logiciel
# (bouton « Configurer… ») est enregistrée dans modes.json à côté de l'exe.
# SOURCE_CELLS : {cellule du .xls source : cellule de l'onglet cible}
DEFAULT_MODES = {
    "HubView-xT": {
        "TARGET_SHEET": "Test final (1_4)",
        "TARGET_EXT": ".xlsm",
        "SOURCE_CELLS": {
            "J9": "C16",
            "J10": "C17",
            "J20": "C22",
            "J21": "C24",
            "J31": "C26",
            "J32": "C28",
            "J40": "C33",
            "J41": "C34",
        },
    },
    "NVR": {
        "TARGET_SHEET": "Test final (1_4)",
        "TARGET_EXT": ".xlsm",
        "SOURCE_CELLS": {
            "J9": "C16",
            "J10": "C17",
            "J20": "C22",
            "J21": "C24",
            "J31": "C26",
            "J32": "C28",
            "J40": "C33",
            "J41": "C34",
        },
    },
    "HubView-MF19": {
        "TARGET_SHEET": "Test Final(1_3)",
        "TARGET_EXT": ".xlsx",
        "SOURCE_CELLS": {
            "J9": "C18",
            "J10": "C19",
            "J20": "C24",
            "J21": "C26",
            "J31": "C28",
            "J32": "C30",
            "J40": "C35",
            "J41": "C36",
        },
    },
    # Groupe-1 = Alimentation, Groupe-2 = Ethernet.
    # Rigidité : la source enchaîne G1+, G2+, G2-, G1- alors que le PV
    # attend Alim +, Alim -, Ethernet +, Ethernet -.
    # J42 (test 100 V [1]...[8]) n'a pas de case dans l'onglet 1/4.
    "Poste de Locution": {
        "TARGET_SHEET": "Test final (1_4)",
        "TARGET_EXT": ".xlsm",
        "SOURCE_CELLS": {
            "J9": "C16",
            "J10": "C17",
            "J20": "C22",
            "J23": "C24",
            "J21": "C26",
            "J22": "C28",
            "J31": "C33",
            "J32": "C34",
        },
    },
}

TARGET_EXTS = (".xlsm", ".xlsx")
CELL_RE = re.compile(r"^[A-Z]{1,3}[1-9][0-9]*$")


def app_dir():
    # Exe PyInstaller : dossier de l'exe (et non le dossier temporaire _MEIxxxx)
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_FILE = os.path.join(app_dir(), "modes.json")


def validate_modes(modes):
    """
    Retourne la liste des erreurs (vide si la configuration est valide)
    """
    if not isinstance(modes, dict) or not modes:
        return ["Aucun type de copie défini"]

    errors = []
    for name, cfg in modes.items():
        prefix = f"« {name} » : "
        if not str(name).strip():
            errors.append("Un type de copie n'a pas de nom")
        if not isinstance(cfg, dict):
            errors.append(prefix + "configuration invalide")
            continue
        if not str(cfg.get("TARGET_SHEET", "")).strip():
            errors.append(prefix + "onglet cible vide")
        if cfg.get("TARGET_EXT") not in TARGET_EXTS:
            errors.append(prefix + f"extension cible invalide ({cfg.get('TARGET_EXT')})")

        cells = cfg.get("SOURCE_CELLS")
        if not isinstance(cells, dict) or not cells:
            errors.append(prefix + "aucune cellule à copier")
            continue
        for src, dst in cells.items():
            if not CELL_RE.match(str(src)) or not CELL_RE.match(str(dst)):
                errors.append(prefix + f"cellule invalide : {src} → {dst}")
        targets = list(cells.values())
        dups = sorted({str(d) for d in targets if targets.count(d) > 1})
        if dups:
            errors.append(prefix + f"cellule cible utilisée plusieurs fois : {', '.join(dups)}")
    return errors


def load_modes():
    """
    Retourne (modes, avertissement ou None)
    """
    if not os.path.exists(CONFIG_FILE):
        return copy.deepcopy(DEFAULT_MODES), None
    try:
        with open(CONFIG_FILE, encoding="utf-8") as f:
            modes = json.load(f)
        errors = validate_modes(modes)
    except (OSError, ValueError) as e:
        errors = [str(e)]
    if errors:
        return copy.deepcopy(DEFAULT_MODES), (
            f"Le fichier {CONFIG_FILE} est invalide, la configuration d'origine est utilisée.\n\n"
            + "\n".join(errors[:10])
        )
    return modes, None


def save_modes(modes):
    # Écriture dans un fichier temporaire puis remplacement : un plantage
    # pendant l'écriture ne laisse pas un modes.json à moitié écrit
    tmp = CONFIG_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(modes, f, ensure_ascii=False, indent=2)
    os.replace(tmp, CONFIG_FILE)


# ======================
# OUTILS
# ======================
def extract_id(filename):
    """
    Extrait le dernier identifiant xxxxxx-yyyyyyy-zzzzz avant l'extension
    """
    name = os.path.splitext(filename)[0]
    matches = re.findall(r"\d{6}-\d{7}-\d{5}", name)
    return matches[-1] if matches else None


def list_by_id(folder, ext):
    """
    Retourne ({id: fichier}, {id: [fichiers]}) : les fichiers uniques et les
    identifiants présents plusieurs fois (qu'on refuse de traiter).
    """
    found = {}
    for f in sorted(os.listdir(folder)):
        # "~$..." : fichier de verrouillage créé par Excel quand le classeur est ouvert
        if f.startswith("~$") or not f.lower().endswith(ext):
            continue
        file_id = extract_id(f)
        if file_id:
            found.setdefault(file_id, []).append(f)

    unique = {k: v[0] for k, v in found.items() if len(v) == 1}
    duplicates = {k: v for k, v in found.items() if len(v) > 1}
    return unique, duplicates


def clean_value(value):
    """
    "I = 132 nA" -> "132 nA" ; une valeur vide reste vide, 0 reste 0
    """
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value)
    if "=" in text:
        text = text.split("=", 1)[1]
    return text.strip()


def find_sheet(workbook, sheet_name):
    wanted = sheet_name.replace(" ", "")
    for sheet in workbook.Worksheets:
        if sheet.Name.replace(" ", "") == wanted:
            return sheet
    return None


def copy_one(excel, src_xls, src_target, dest_target, config):
    shutil.copy(src_target, dest_target)

    wb_xls = wb_target = None
    saved = False
    try:
        # Ouverture en lecture seule (UpdateLinks=0, ReadOnly=True)
        wb_xls = excel.Workbooks.Open(src_xls, 0, True)
        ws_xls = wb_xls.Worksheets(1)
        values = {src: clean_value(ws_xls.Range(src).Value) for src in config["SOURCE_CELLS"]}
        wb_xls.Close(False)
        wb_xls = None

        wb_target = excel.Workbooks.Open(dest_target)
        ws_target = find_sheet(wb_target, config["TARGET_SHEET"])
        if ws_target is None:
            raise ValueError(f"Onglet '{config['TARGET_SHEET']}' introuvable")

        for src, dst in config["SOURCE_CELLS"].items():
            ws_target.Range(dst).Value = values[src]

        wb_target.Save()
        saved = True
        wb_target.Close(False)
        wb_target = None

    except Exception:
        for wb in (wb_xls, wb_target):
            if wb is not None:
                try:
                    wb.Close(False)
                except Exception:
                    pass
        # Ne pas laisser dans la destination une copie vierge qu'on prendrait pour un PV rempli
        if not saved:
            try:
                os.remove(dest_target)
            except OSError:
                pass
        raise


# ======================
# APPLICATION
# ======================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Copie Diélectrie → PV de Test")
        self.geometry("860x700")
        self.resizable(False, False)

        self.xls_dir = tk.StringVar()
        self.target_dir = tk.StringVar()
        self.dest_dir = tk.StringVar()

        self.modes, config_warning = load_modes()
        self.mode = tk.StringVar(value=next(iter(self.modes)))

        self.progress_var = tk.IntVar()
        self.progress_label = tk.StringVar(value="0 / 0")
        self.stop_requested = False

        # Tkinter n'est pas thread-safe : le thread de traitement envoie ses
        # messages dans cette file, lue par la boucle Tk (poll_queue)
        self.queue = queue.Queue()
        self.worker = None
        self.closing = False

        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.poll_id = self.after(100, self.poll_queue)

        if config_warning:
            self.after_idle(lambda: messagebox.showwarning("Configuration", config_warning, parent=self))

    def destroy(self):
        self.after_cancel(self.poll_id)
        super().destroy()

    # ---------- UI ----------

    def create_widgets(self):
        pad = {"padx": 10, "pady": 5}

        tk.Label(self, text="Type de copie").pack(anchor="w", **pad)
        frame_mode = tk.Frame(self)
        frame_mode.pack(anchor="w", padx=10)
        self.mode_combo = ttk.Combobox(
            frame_mode,
            values=list(self.modes.keys()),
            textvariable=self.mode,
            state="readonly",
            width=30,
        )
        self.mode_combo.pack(side="left")
        tk.Button(frame_mode, text="Configurer…", command=self.open_config).pack(side="left", padx=10)

        tk.Label(self, text="Dossier fichiers source (.xls)").pack(anchor="w", **pad)
        self.create_selector(self.xls_dir)

        tk.Label(self, text="Dossier fichiers cibles (.xlsm / .xlsx)").pack(anchor="w", **pad)
        self.create_selector(self.target_dir)

        tk.Label(self, text="Dossier de destination").pack(anchor="w", **pad)
        self.create_selector(self.dest_dir)

        frame_btn = tk.Frame(self)
        frame_btn.pack(pady=10)
        self.btn_start = tk.Button(frame_btn, text="Lancer", width=15, command=self.start_processing)
        self.btn_start.pack(side="left", padx=10)
        tk.Button(frame_btn, text="Annuler", width=15, command=self.cancel_processing).pack(side="left", padx=10)

        frame = tk.Frame(self)
        frame.pack(fill="x", padx=10)

        self.progress = ttk.Progressbar(
            frame, orient="horizontal", length=680,
            mode="determinate", variable=self.progress_var
        )
        self.progress.pack(side="left", padx=(0, 10))
        tk.Label(frame, textvariable=self.progress_label).pack(side="left")

        self.log = scrolledtext.ScrolledText(
            self, width=105, height=22, state="disabled"
        )
        self.log.pack(padx=10, pady=10)

    def create_selector(self, var):
        frame = tk.Frame(self)
        frame.pack(fill="x", padx=10)
        tk.Entry(frame, textvariable=var).pack(side="left", fill="x", expand=True)
        tk.Button(frame, text="Parcourir",
                  command=lambda: self.browse(var)).pack(side="right")

    def browse(self, var):
        folder = filedialog.askdirectory(initialdir=var.get() or None)
        if folder:  # "Annuler" dans la boîte de dialogue ne doit pas vider le champ
            var.set(folder)

    # ---------- LOG ----------

    def write_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def post_log(self, text):
        self.queue.put(("log", text))

    def poll_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                if msg[0] == "log":
                    self.write_log(msg[1])
                elif msg[0] == "progress":
                    _, idx, total = msg
                    self.progress_var.set(idx)
                    self.progress_label.set(f"{idx} / {total}")
                elif msg[0] == "done":
                    self.worker = None
                    self.btn_start.configure(state="normal")
                    if self.closing:
                        self.destroy()
                        return
        except queue.Empty:
            pass
        self.poll_id = self.after(100, self.poll_queue)

    # ---------- CONFIGURATION ----------

    def open_config(self):
        ConfigEditor(self)

    def apply_modes(self, modes):
        self.modes = modes
        self.mode_combo.configure(values=list(modes.keys()))
        if self.mode.get() not in modes:
            self.mode.set(next(iter(modes)))
        self.write_log(f"[INFO] Configuration enregistrée : {CONFIG_FILE}")

    # ---------- CONTROLE ----------

    def is_running(self):
        return self.worker is not None and self.worker.is_alive()

    def cancel_processing(self):
        if not self.is_running():
            return
        self.stop_requested = True
        self.write_log("\n[INFO] Annulation demandée (arrêt après le fichier en cours)")

    def on_close(self):
        if not self.is_running():
            self.destroy()
            return
        # Fermer brutalement laisserait un EXCEL.EXE invisible tourner en arrière-plan
        self.closing = True
        self.cancel_processing()
        self.write_log("[INFO] La fenêtre se fermera à la fin du fichier en cours")

    def start_processing(self):
        if self.is_running():
            return

        xls_dir, target_dir, dest_dir = self.xls_dir.get(), self.target_dir.get(), self.dest_dir.get()
        if not all([xls_dir, target_dir, dest_dir]):
            messagebox.showerror("Erreur", "Tous les dossiers doivent être sélectionnés")
            return
        for d in (xls_dir, target_dir, dest_dir):
            if not os.path.isdir(d):
                messagebox.showerror("Erreur", f"Dossier introuvable :\n{d}")
                return
        if os.path.normcase(os.path.abspath(target_dir)) == os.path.normcase(os.path.abspath(dest_dir)):
            messagebox.showerror("Erreur", "Le dossier de destination doit être différent du dossier des fichiers cibles")
            return

        mode_name = self.mode.get()
        if mode_name not in self.modes:
            messagebox.showerror("Erreur", "Sélectionne un type de copie")
            return
        # Copie : le traitement n'est pas affecté si la config est modifiée pendant qu'il tourne
        config = copy.deepcopy(self.modes[mode_name])

        xls_files, xls_dup = list_by_id(xls_dir, ".xls")
        target_files, target_dup = list_by_id(target_dir, config["TARGET_EXT"])

        self.write_log(f"\n=== Mode : {mode_name} ===")
        self.write_log(f"[DEBUG] XLS trouvés    : {len(xls_files)}")
        self.write_log(f"[DEBUG] Cibles trouvées: {len(target_files)}")
        for label, dup in (("source", xls_dup), ("cible", target_dup)):
            for file_id, files in dup.items():
                self.write_log(f"[ATTENTION] ID {file_id} en double ({label}), ignoré : {', '.join(files)}")

        ids = sorted(i for i in xls_files if i in target_files)
        self.write_log(f"[DEBUG] IDs communs    : {ids}")

        if not ids:
            self.write_log("[ERREUR] Aucun couple source / cible trouvé")
            return

        pairs = [
            (
                common_id,
                os.path.join(xls_dir, xls_files[common_id]),
                os.path.join(target_dir, target_files[common_id]),
                os.path.join(dest_dir, target_files[common_id]),
            )
            for common_id in ids
        ]

        existing = [p for p in pairs if os.path.exists(p[3])]
        if existing and not messagebox.askyesno(
            "Confirmation",
            f"{len(existing)} fichier(s) existe(nt) déjà dans le dossier de destination "
            "et seront écrasés.\n\nContinuer ?",
        ):
            return

        self.stop_requested = False
        self.progress_var.set(0)
        self.progress["maximum"] = len(pairs)
        self.progress_label.set(f"0 / {len(pairs)}")
        self.btn_start.configure(state="disabled")

        self.worker = threading.Thread(
            target=self.process_files, args=(pairs, config, mode_name), daemon=True
        )
        self.worker.start()

    # ---------- TRAITEMENT (thread) ----------
    # Ne touche jamais aux widgets directement : passe par post_log / self.queue

    def process_files(self, pairs, config, mode_name):
        pythoncom.CoInitialize()
        excel = None
        ok = errors = 0

        try:
            self.post_log(f"=== Début du traitement ({mode_name}) ===")

            # DispatchEx : instance Excel dédiée. Dispatch se raccrocherait à un
            # Excel déjà ouvert par l'utilisateur, que Quit() fermerait ensuite.
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            total = len(pairs)
            for idx, (common_id, src_xls, src_target, dest_target) in enumerate(pairs, start=1):
                if self.stop_requested:
                    self.post_log("[INFO] Traitement annulé")
                    break

                self.post_log(f"\n[Traitement {idx}/{total}] ID: {common_id}")

                try:
                    copy_one(excel, src_xls, src_target, dest_target, config)
                    self.post_log("[OK] Fichier traité")
                    ok += 1
                except Exception as e:
                    self.post_log(f"[ERREUR] {common_id} : {e}")
                    errors += 1

                self.queue.put(("progress", idx, total))

            self.post_log(f"\n=== Traitement terminé : {ok} OK, {errors} erreur(s) ===")

        except Exception as e:
            # Ex. Excel absent ou impossible à lancer : sans ce bloc, l'erreur
            # disparaissait (pas de console dans l'exe)
            self.post_log(f"[ERREUR] Traitement interrompu : {e}")

        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()
            self.queue.put(("done",))


# ======================
# EDITEUR DE CONFIGURATION
# ======================
class ConfigEditor(tk.Toplevel):
    """
    Edition des types de copie. Travaille sur une copie de la configuration :
    rien n'est modifié tant qu'on n'a pas cliqué sur « Enregistrer ».
    """

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Configuration des types de copie")
        self.resizable(False, False)
        self.transient(app)

        self.modes = copy.deepcopy(app.modes)
        self.current = None  # nom du type affiché dans le formulaire

        self.name_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.ext_var = tk.StringVar()
        self.src_var = tk.StringVar()
        self.dst_var = tk.StringVar()

        self.create_widgets()
        self.refresh_list(select=app.mode.get())

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.grab_set()

    # ---------- UI ----------

    def create_widgets(self):
        left = tk.Frame(self)
        left.grid(row=0, column=0, sticky="ns", padx=10, pady=10)

        tk.Label(left, text="Types de copie").pack(anchor="w")
        self.listbox = tk.Listbox(left, width=28, height=17, exportselection=False)
        self.listbox.pack()
        self.listbox.bind("<<ListboxSelect>>", self.on_select)
        tk.Button(left, text="Nouveau", command=self.new_mode).pack(fill="x", pady=(5, 0))
        tk.Button(left, text="Dupliquer", command=self.duplicate_mode).pack(fill="x", pady=(5, 0))
        tk.Button(left, text="Supprimer", command=self.delete_mode).pack(fill="x", pady=(5, 0))

        right = tk.Frame(self)
        right.grid(row=0, column=1, sticky="n", padx=10, pady=10)

        tk.Label(right, text="Nom").grid(row=0, column=0, sticky="w")
        self.name_entry = tk.Entry(right, textvariable=self.name_var, width=40)
        self.name_entry.grid(row=0, column=1, sticky="ew", pady=2)
        tk.Label(right, text="Onglet cible").grid(row=1, column=0, sticky="w")
        tk.Entry(right, textvariable=self.sheet_var, width=40).grid(row=1, column=1, sticky="ew", pady=2)
        tk.Label(right, text="Extension cible").grid(row=2, column=0, sticky="w")
        ttk.Combobox(right, textvariable=self.ext_var, values=TARGET_EXTS,
                     state="readonly", width=8).grid(row=2, column=1, sticky="w", pady=2)

        tk.Label(right, text="Cellules à copier (source .xls → onglet cible)").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(10, 2))

        frame_tree = tk.Frame(right)
        frame_tree.grid(row=4, column=0, columnspan=2, sticky="w")
        self.tree = ttk.Treeview(frame_tree, columns=("src", "dst"), show="headings",
                                 height=9, selectmode="browse")
        self.tree.heading("src", text="Cellule source")
        self.tree.heading("dst", text="Cellule cible")
        self.tree.column("src", width=170, anchor="center")
        self.tree.column("dst", width=170, anchor="center")
        self.tree.pack(side="left")
        scroll = ttk.Scrollbar(frame_tree, orient="vertical", command=self.tree.yview)
        scroll.pack(side="left", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)

        frame_row = tk.Frame(right)
        frame_row.grid(row=5, column=0, columnspan=2, sticky="w", pady=5)
        tk.Label(frame_row, text="Source").pack(side="left")
        self.src_entry = tk.Entry(frame_row, textvariable=self.src_var, width=7)
        self.src_entry.pack(side="left", padx=(2, 8))
        tk.Label(frame_row, text="→ Cible").pack(side="left")
        dst_entry = tk.Entry(frame_row, textvariable=self.dst_var, width=7)
        dst_entry.pack(side="left", padx=(2, 8))
        dst_entry.bind("<Return>", lambda e: self.add_row())
        tk.Button(frame_row, text="Ajouter / Modifier", command=self.add_row).pack(side="left", padx=2)
        tk.Button(frame_row, text="Supprimer la ligne", command=self.delete_row).pack(side="left", padx=2)

        bottom = tk.Frame(self)
        bottom.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 10))
        tk.Button(bottom, text="Restaurer les valeurs d'origine", command=self.restore_defaults).pack(side="left")
        tk.Button(bottom, text="Enregistrer", width=12, command=self.save).pack(side="right")
        tk.Button(bottom, text="Annuler", width=12, command=self.destroy).pack(side="right", padx=5)

    def refresh_list(self, select=None):
        names = list(self.modes.keys())
        self.listbox.delete(0, "end")
        for name in names:
            self.listbox.insert("end", name)
        index = names.index(select) if select in names else 0
        self.listbox.selection_set(index)
        self.listbox.see(index)
        self.show(names[index])

    def show(self, name):
        self.current = name
        cfg = self.modes[name]
        self.name_var.set(name)
        self.sheet_var.set(cfg.get("TARGET_SHEET", ""))
        self.ext_var.set(cfg.get("TARGET_EXT", TARGET_EXTS[0]))
        self.tree.delete(*self.tree.get_children())
        for src, dst in cfg.get("SOURCE_CELLS", {}).items():
            self.tree.insert("", "end", values=(src, dst))
        self.src_var.set("")
        self.dst_var.set("")

    def store_form(self):
        """
        Reporte le formulaire dans self.modes. Retourne False si le nom est invalide.
        """
        if self.current is None:
            return True
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror("Erreur", "Le nom du type de copie est vide", parent=self)
            return False
        if name != self.current and name in self.modes:
            messagebox.showerror("Erreur", f"Un type « {name} » existe déjà", parent=self)
            return False

        cfg = {
            "TARGET_SHEET": self.sheet_var.get().strip(),
            "TARGET_EXT": self.ext_var.get(),
            "SOURCE_CELLS": {self.tree.set(i, "src"): self.tree.set(i, "dst") for i in self.tree.get_children()},
        }
        # Reconstruit le dict pour garder l'ordre en cas de renommage
        self.modes = {(name if k == self.current else k): (cfg if k == self.current else v)
                      for k, v in self.modes.items()}
        self.current = name
        return True

    # ---------- TYPES ----------

    def on_select(self, event=None):
        selection = self.listbox.curselection()
        if not selection:
            return
        name = self.listbox.get(selection[0])
        if name == self.current:
            return
        if not self.store_form():
            # On reste sur le type en cours sans effacer ce qui a été saisi
            self.listbox.selection_clear(0, "end")
            self.listbox.selection_set(list(self.modes).index(self.current))
            return
        self.refresh_list(select=name)

    def unique_name(self, base):
        name, n = base, 2
        while name in self.modes:
            name, n = f"{base} {n}", n + 1
        return name

    def new_mode(self):
        if not self.store_form():
            return
        name = self.unique_name("Nouveau type")
        self.modes[name] = {"TARGET_SHEET": "", "TARGET_EXT": TARGET_EXTS[0], "SOURCE_CELLS": {}}
        self.refresh_list(select=name)
        self.name_entry.focus_set()
        self.name_entry.select_range(0, "end")

    def duplicate_mode(self):
        if not self.store_form():
            return
        name = self.unique_name(f"{self.current} (copie)")
        self.modes[name] = copy.deepcopy(self.modes[self.current])
        self.refresh_list(select=name)
        self.name_entry.focus_set()
        self.name_entry.select_range(0, "end")

    def delete_mode(self):
        if len(self.modes) == 1:
            messagebox.showerror("Erreur", "Il faut garder au moins un type de copie", parent=self)
            return
        if not messagebox.askyesno("Confirmation", f"Supprimer le type « {self.current} » ?", parent=self):
            return
        del self.modes[self.current]
        self.current = None
        self.refresh_list()

    def restore_defaults(self):
        if not messagebox.askyesno(
            "Confirmation",
            "Remplacer toute la configuration par les valeurs d'origine ?\n"
            "Les types ajoutés ou modifiés seront perdus après « Enregistrer ».",
            parent=self,
        ):
            return
        self.modes = copy.deepcopy(DEFAULT_MODES)
        self.current = None
        self.refresh_list()

    # ---------- CELLULES ----------

    def on_row_select(self, event=None):
        selection = self.tree.selection()
        if selection:
            self.src_var.set(self.tree.set(selection[0], "src"))
            self.dst_var.set(self.tree.set(selection[0], "dst"))

    def add_row(self):
        src = self.src_var.get().strip().upper().replace("$", "")
        dst = self.dst_var.get().strip().upper().replace("$", "")
        for label, cell in (("source", src), ("cible", dst)):
            if not CELL_RE.match(cell):
                messagebox.showerror("Erreur", f"Cellule {label} invalide : « {cell} » (exemple : J9)", parent=self)
                return

        rows = {self.tree.set(i, "src"): i for i in self.tree.get_children()}
        for s, item in rows.items():
            if s != src and self.tree.set(item, "dst") == dst:
                messagebox.showerror("Erreur", f"La cellule cible {dst} reçoit déjà {s}", parent=self)
                return

        # Une cellule source ne peut aller qu'à une seule cible : on remplace la ligne existante
        if src in rows:
            self.tree.item(rows[src], values=(src, dst))
            item = rows[src]
        else:
            item = self.tree.insert("", "end", values=(src, dst))
        self.tree.selection_set(item)
        self.tree.see(item)
        self.src_var.set("")
        self.dst_var.set("")
        self.src_entry.focus_set()

    def delete_row(self):
        for item in self.tree.selection():
            self.tree.delete(item)
        self.src_var.set("")
        self.dst_var.set("")

    # ---------- ENREGISTREMENT ----------

    def save(self):
        if not self.store_form():
            return
        errors = validate_modes(self.modes)
        if errors:
            messagebox.showerror("Configuration invalide", "\n".join(errors), parent=self)
            return
        try:
            save_modes(self.modes)
        except OSError as e:
            messagebox.showerror(
                "Erreur",
                f"Impossible d'enregistrer {CONFIG_FILE} :\n{e}\n\n"
                "Place l'exe dans un dossier où tu as les droits d'écriture.",
                parent=self,
            )
            return
        self.app.apply_modes(self.modes)
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
