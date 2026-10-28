import os
import queue
import re
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import win32com.client as win32
import pythoncom


# ======================
# CONFIGURATION DES MODES
# ======================
# SOURCE_CELLS : {cellule du .xls source : cellule de l'onglet cible}
MODES = {
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
        "TARGET_EXT": ".xlsx",
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
        self.mode = tk.StringVar(value="HubView-xT")

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

    def destroy(self):
        self.after_cancel(self.poll_id)
        super().destroy()

    # ---------- UI ----------

    def create_widgets(self):
        pad = {"padx": 10, "pady": 5}

        tk.Label(self, text="Type de copie").pack(anchor="w", **pad)
        ttk.Combobox(
            self,
            values=list(MODES.keys()),
            textvariable=self.mode,
            state="readonly",
            width=30,
        ).pack(anchor="w", padx=10)

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
        config = MODES[mode_name]

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


if __name__ == "__main__":
    App().mainloop()
