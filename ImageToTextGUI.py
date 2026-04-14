import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import threading
import asyncio
import pyperclip
from concurrent.futures import ProcessPoolExecutor
import sys
import multiprocessing
from ImageToText import get_config_path

# --- UI Configuration ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

APP_VERSION = "1.1"
GITHUB_VERSION_URL = "https://raw.githubusercontent.com/SimonB-cyber/text_extractor/main/version.txt"
GITHUB_RELEASE_URL = "https://github.com/SimonB-cyber/text_extractor/releases/latest/download/TextExtraktPro_Setup_v{version}.exe"

class OCRExtractorGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Premium OCR & Text Extraktor Pro")
        self.geometry("1000x800")
        
        self.files_to_process = []
        self.processing_results = []
        self.executor = None
        self.abort_event = threading.Event()
        self.logs = []
        self.log_window = None

        self.ai_config = {}
        self.ai_config_file = get_config_path("ai_config.json")
        if os.path.exists(self.ai_config_file):
            import json
            try:
                with open(self.ai_config_file, 'r', encoding='utf-8') as f:
                    self.ai_config = json.load(f)
            except: pass
            
        self.setup_ui()
        self.log("GUI gestartet. Lade Engines im Hintergrund...")
        
        # Start background task to load heavy stuff (Tesseract, AI, Modules)
        threading.Thread(target=self.bg_init_engines, daemon=True).start()

    def bg_init_engines(self):
        try:
            self.after(0, lambda: self.status_label.configure(text="⚠️ Lade Text-Engine im Hintergrund... (Das kann kurz dauern)", text_color="#fbbf24"))
            self.after(0, lambda: self.start_button.configure(state="disabled"))
            
            from ImageToText import get_tesseract_config
            self.log("System-Check: Suche nach Tesseract-Installation...")
            t_path, status = get_tesseract_config()
            
            if status == "INSTALL_FAILED":
                self.log("ERROR: Auto-Installation fehlgeschlagen.")
                self.after(0, lambda: messagebox.showwarning("Tesseract Setup", "Tesseract konnte nicht automatisch gefunden werden.\nBitte wähle es in den Einstellungen aus."))
                self.after(0, self.open_settings_window)
            elif status == "WANT_PROMPT" or not t_path:
                self.log("Information: Fordere Benutzer zur Pfadauswahl auf.")
                self.after(0, self.open_settings_window)
            else:
                self.log(f"Tesseract bereit: {t_path}")
            
            # Check for Updates
            self.check_for_updates()
            
            # Re-Enable start button once imports are fully loaded!
            self.after(0, lambda: self.status_label.configure(text="✅ Alle Engines geladen! Drei Schritte zum Text: Dateien laden -> Modus wählen -> Starten", text_color="#34d399"))
            self.after(0, lambda: self.start_button.configure(state="normal"))
        except Exception as e:
            self.log(f"Fehler beim Laden im Hintergrund: {e}")

    def check_for_updates(self):
        try:
            import urllib.request
            self.log("Suche nach Updates...")
            with urllib.request.urlopen(GITHUB_VERSION_URL, timeout=5) as r:
                latest = r.read().decode('utf-8').strip()
            
            def version_tuple(v):
                return tuple(int(x) for x in v.split('.'))
            
            if version_tuple(latest) > version_tuple(APP_VERSION):
                self.log(f"Update verfügbar: v{latest} (aktuell: v{APP_VERSION})")
                self.after(0, lambda: self._prompt_update(latest))
            else:
                self.log(f"Programm ist aktuell (v{APP_VERSION}).")
        except Exception as e:
            self.log(f"Update-Check fehlgeschlagen (kein Internet?): {e}")

    def _prompt_update(self, new_version):
        if messagebox.askyesno(
            "🔔 Update verfügbar!",
            f"Version {new_version} ist verfügbar!\nAktuell installiert: {APP_VERSION}\n\nJetzt herunterladen und installieren?"
        ):
            self.log(f"Download von v{new_version} gestartet...")
            threading.Thread(target=self._download_and_install_update, args=(new_version,), daemon=True).start()

    def _download_and_install_update(self, new_version):
        import urllib.request, tempfile, subprocess, os
        url = GITHUB_RELEASE_URL.format(version=new_version)
        try:
            tmp = os.path.join(tempfile.gettempdir(), f"TextExtraktPro_Update_v{new_version}.exe")
            self.log(f"Lade herunter: {url}")
            urllib.request.urlretrieve(url, tmp)
            self.log(f"Download fertig! Starte Installer...")
            subprocess.Popen([tmp], shell=True)
            self.after(0, self.destroy)
        except Exception as e:
            self.log(f"Update-Download fehlgeschlagen: {e}")
            self.after(0, lambda: messagebox.showerror("Update fehlgeschlagen", f"Download gescheitert:\n{e}\n\nBitte manuell auf GitHub herunterladen."))

    def log(self, message):
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        full_msg = f"[{timestamp}] {message}"
        self.logs.append(full_msg)
        print(full_msg) # Standard output as backup
        
        if self.log_window and self.log_window.winfo_exists():
            self.log_window.update_log(full_msg)
            
        if hasattr(self, 'console_box'):
            try:
                self.console_box.configure(state="normal")
                self.console_box.insert("end", full_msg + "\n")
                self.console_box.see("end")
                self.console_box.configure(state="disabled")
            except: pass

    def open_settings_window(self):
        self.log("Öffne Einstellungen...")
        set_win = ctk.CTkToplevel(self)
        set_win.title("Einstellungen (Tesseract & KI)")
        set_win.geometry("550x450")
        set_win.attributes('-topmost', True)
        
        # Tesseract
        ctk.CTkLabel(set_win, text="Tesseract.exe Pfad:", font=ctk.CTkFont(weight="bold")).pack(pady=(15,0), padx=20, anchor="w")
        t_frame = ctk.CTkFrame(set_win, fg_color="transparent")
        t_frame.pack(fill="x", padx=20, pady=5)
        t_entry = ctk.CTkEntry(t_frame)
        t_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        
        from ImageToText import TESS_PATH
        if TESS_PATH: t_entry.insert(0, TESS_PATH)
        
        def pick_tess():
            p = filedialog.askopenfilename(title="Tesseract suchen", filetypes=[("Exe", "*.exe")])
            if p:
                t_entry.delete(0, 'end')
                t_entry.insert(0, p)
        ctk.CTkButton(t_frame, text="Suchen", width=60, command=pick_tess).pack(side="left")
        
        # AI Config
        ctk.CTkLabel(set_win, text="KI-Modell API (Optional für Korrektur):", font=ctk.CTkFont(weight="bold")).pack(pady=(20,0), padx=20, anchor="w")
        
        ctk.CTkLabel(set_win, text="API URL (z.B. http://localhost:11434/v1/chat/completions):").pack(padx=20, anchor="w")
        url_entry = ctk.CTkEntry(set_win)
        url_entry.pack(fill="x", padx=20, pady=2)
        url_entry.insert(0, self.ai_config.get('api_url', 'http://localhost:11434/v1/chat/completions'))
        
        ctk.CTkLabel(set_win, text="Modell Name (z.B. llama3, gpt-3.5-turbo):").pack(padx=20, anchor="w")
        mod_entry = ctk.CTkEntry(set_win)
        mod_entry.pack(fill="x", padx=20, pady=2)
        mod_entry.insert(0, self.ai_config.get('model', 'llama3'))
        
        ctk.CTkLabel(set_win, text="API Key (bei lokaler KI oft egal, sonst 'sk-...'):").pack(padx=20, anchor="w")
        key_entry = ctk.CTkEntry(set_win)
        key_entry.pack(fill="x", padx=20, pady=2)
        key_entry.insert(0, self.ai_config.get('api_key', 'sk-local'))
        
        def save_and_close():
            import ImageToText
            p = t_entry.get()
            if p and os.path.exists(p):
                ImageToText.TESS_PATH = p
                config_file = os.path.join(os.path.dirname(os.path.abspath(ImageToText.__file__)), "tesseract_path.txt")
                with open(config_file, 'w', encoding='utf-8') as f: f.write(p)
                from ImageToText import get_tesseract_config
                get_tesseract_config()
                
            self.ai_config = {
                "api_url": url_entry.get().strip(),
                "model": mod_entry.get().strip(),
                "api_key": key_entry.get().strip()
            }
            import json
            with open(self.ai_config_file, 'w', encoding='utf-8') as f:
                json.dump(self.ai_config, f, indent=4)
                
            self.log("Einstellungen gespeichert.")
            set_win.destroy()
            
        ctk.CTkButton(set_win, text="Speichern & Schließen", fg_color="#10b981", hover_color="#059669", command=save_and_close).pack(pady=20)

    def setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # 1. Header
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="nsew")
        
        self.title_label = ctk.CTkLabel(self.header_frame, text="📚 Text-Extraktor Pro", font=ctk.CTkFont(size=28, weight="bold"), text_color="#38bdf8")
        self.title_label.pack(side="left", anchor="w")
        
        self.log_btn = ctk.CTkButton(self.header_frame, text="📋 Logs", width=80, height=32, fg_color="#334155", hover_color="#475569", command=self.open_log_window)
        self.log_btn.pack(side="right", padx=5)

        self.settings_btn = ctk.CTkButton(self.header_frame, text="⚙️", width=40, height=32, fg_color="#334155", hover_color="#475569", command=self.open_settings_window)
        self.settings_btn.pack(side="right", padx=5)

        # 2. Control Buttons Area
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        
        self.pick_button = ctk.CTkButton(self.button_frame, text="+ Dateien wählen", font=ctk.CTkFont(weight="bold"), fg_color="#1e293b", hover_color="#334155", command=self.pick_files, height=45)
        self.pick_button.pack(side="left", padx=(0, 10))
        
        self.clear_button = ctk.CTkButton(self.button_frame, text="Liste leeren", fg_color="#ef4444", hover_color="#dc2626", command=self.clear_list, height=45, width=100)
        self.clear_button.pack(side="left", padx=(0, 20))

        # Language Selection
        self.lang_label = ctk.CTkLabel(self.button_frame, text="Sprache:", font=ctk.CTkFont(size=12))
        self.lang_label.pack(side="left", padx=(0, 5))
        self.lang_menu = ctk.CTkComboBox(self.button_frame, values=["deu", "eng", "fra", "spa", "ita"], width=80)
        self.lang_menu.set("deu")
        self.lang_menu.pack(side="left", padx=(0, 20))

        # Table Mode Switch
        self.table_switch = ctk.CTkSwitch(self.button_frame, text="Tabellen-Modus", font=ctk.CTkFont(size=12))
        self.table_switch.pack(side="left", padx=10)

        # AI Switch
        self.ai_switch = ctk.CTkSwitch(self.button_frame, text="🧠 KI-Korrektur", font=ctk.CTkFont(size=12), progress_color="#8b5cf6")
        self.ai_switch.pack(side="left", padx=10)
        self.ai_switch.select()  # Standardmäßig einschalten

        # 3. File List (Scrollable)
        self.scroll_frame = ctk.CTkScrollableFrame(self, label_text="Warteschlange", label_font=ctk.CTkFont(weight="bold"))
        self.scroll_frame.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        self.file_items = []

        # 4. Status & Progress
        self.progress_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.progress_frame.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="nsew")
        
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame, mode="determinate", height=12)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", pady=10)
        
        self.status_label = ctk.CTkLabel(self.progress_frame, text="Drei Schritte zum Text: Dateien laden -> Modus wählen -> Starten", text_color="#94a3b8")
        self.status_label.pack()

        self.console_box = ctk.CTkTextbox(self.progress_frame, height=120, font=ctk.CTkFont(family="Consolas", size=12), text_color="#a1a1aa", fg_color="#18181b")
        self.console_box.pack(fill="x", pady=(10, 0))
        self.console_box.configure(state="disabled")

        # 5. Action Buttons (Bottom)
        self.action_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.action_frame.grid(row=4, column=0, padx=20, pady=(0, 30), sticky="nsew")
        
        self.start_button = ctk.CTkButton(self.action_frame, text="🚀 VERARBEITUNG STARTEN", font=ctk.CTkFont(size=16, weight="bold"), fg_color="#38bdf8", hover_color="#0ea5e9", text_color="white", height=55, command=self.start_processing_thread)
        self.start_button.pack(side="left", expand=True, fill="x", padx=(0, 10))
        
        self.abort_button = ctk.CTkButton(self.action_frame, text="🛑 Stopp", font=ctk.CTkFont(weight="bold"), fg_color="#475569", state="disabled", height=55, width=120, command=self.abort_processing)
        self.abort_button.pack(side="left", padx=(0, 10))

        self.copy_button = ctk.CTkButton(self.action_frame, text="📋 Kopieren", font=ctk.CTkFont(weight="bold"), height=55, width=120, state="disabled", command=self.copy_all)
        self.copy_button.pack(side="left")

    def open_log_window(self):
        if self.log_window is None or not self.log_window.winfo_exists():
            self.log_window = LogWindow(self, self.logs)
        else:
            self.log_window.focus()

    def pick_files(self):
        files = filedialog.askopenfilenames(
            title="Wähle Bilder oder PDFs",
            filetypes=[("Extrakt-Dateien", "*.jpg *.jpeg *.png *.pdf *.bmp *.tiff")]
        )
        if files:
            for f in files:
                if f not in self.files_to_process:
                    self.files_to_process.append(f)
                    self.add_file_to_ui(f)
            self.log(f"{len(files)} neue Dateien hinzugefügt.")
            self.status_label.configure(text=f"{len(self.files_to_process)} Dateien in der Warteschlange.")

    def add_file_to_ui(self, path):
        item = ctk.CTkLabel(self.scroll_frame, text=f"⏳ {os.path.basename(path)}", anchor="w", font=ctk.CTkFont(size=13))
        item.pack(fill="x", padx=5, pady=2)
        self.file_items.append({"path": path, "label": item})

    def clear_list(self):
        self.files_to_process = []
        for item in self.file_items:
            item["label"].destroy()
        self.file_items = []
        self.processing_results = []
        self.copy_button.configure(state="disabled")
        self.progress_bar.set(0)
        self.status_label.configure(text="Warteschlange geleert.")
        self.log("Benutzer hat die Liste geleert.")

    def update_ui_success(self, path, page_num):
        for item in self.file_items:
            if item["path"] == path:
                status = f"Seite {page_num}" if page_num != 999999 else "Seite ?"
                item["label"].configure(text=f"✅ {os.path.basename(path)} — {status}", text_color="#38bdf8")
                break

    def start_processing_thread(self):
        if not self.files_to_process:
            messagebox.showwarning("Hinweis", "Bitte erst Dateien auswählen!")
            return
        
        self.start_button.configure(state="disabled", text="⚡ Verarbeite...")
        self.abort_button.configure(state="normal", fg_color="#ef4444")
        self.pick_button.configure(state="disabled")
        self.clear_button.configure(state="disabled")
        
        self.abort_event.clear()
        self.log(f"Starte Prozess-Pipeline für {len(self.files_to_process)} Dateien...")
        
        thread = threading.Thread(target=self.run_processing, daemon=True)
        thread.start()

    def abort_processing(self):
        self.log("Abbruch-Signal erhalten. Beende nach aktueller Datei...")
        self.abort_event.set()
        self.abort_button.configure(state="disabled", text="Beende...")

    def run_processing(self):
        self.processing_results = []
        total = len(self.files_to_process)
        
        selected_lang = self.lang_menu.get()
        table_mode = self.table_switch.get() == 1
        ai_mode = self.ai_switch.get() == 1
        
        self.log(f"OCR-Einstellungen: Sprache={selected_lang}, TabellenModus={table_mode}, KI-Korrektur={ai_mode}")
        
        from ImageToText import verarbeite_datei
        import concurrent.futures
        
        results = []
        with ProcessPoolExecutor() as ex:
            self.log("Arbeiter-Pool hochgefahren. Verteile Dateien...")
            futures_map = {ex.submit(verarbeite_datei, p, file_index=idx, tabelle=table_mode, sprache=selected_lang, ki_korrektur=ai_mode, ai_config=self.ai_config): p for idx, p in enumerate(self.files_to_process)}
            processed = 0
            
            for future in concurrent.futures.as_completed(futures_map):
                if self.abort_event.is_set():
                    self.log("Verarbeitung durch Benutzer abgebrochen.")
                    break
                    
                path = futures_map[future]
                try:
                    res = future.result()
                    pfad, seite, returned_text, logs = res
                    results.append((pfad, seite, returned_text))
                    
                    for log_msg in logs:
                        self.log(f"-> {log_msg}")
                        
                    processed += 1
                    self.after(0, self.update_progress, processed, total, res)
                    self.log(f"Erfolg: {os.path.basename(pfad)} (Seite {seite if seite != 999999 else '?'})")
                except Exception as e:
                    self.log(f"FEHLER bei {os.path.basename(path)}: {str(e)}")
                    processed += 1
                    self.after(0, self.update_progress, processed, total, (path, 999999, f"ERROR: {str(e)}", []))
        
        results.sort(key=lambda x: x[1])
        self.processing_results = results
        
        if results:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            self.output_file = os.path.join(desktop, "gesammelter_text.txt")
            with open(self.output_file, "a", encoding="utf-8") as f:
                for p, s, t in results:
                    f.write(t + "\n\n")
        
        self.after(0, self.finish_processing)

    def update_progress(self, current, total, last_result):
        pfad, seite, text, logs = last_result
        self.progress_bar.set(current / total)
        self.status_label.configure(text=f"Fortschritt: {current} / {total} Dateien verarbeitet...")
        self.update_ui_success(pfad, seite)

    def finish_processing(self):
        if self.abort_event.is_set():
            self.status_label.configure(text="⚠️ Verarbeitung wurde vorzeitig abgebrochen.", text_color="#fbbf24")
            self.start_button.configure(text="🚀 WEITERMACHEN")
        else:
            self.status_label.configure(text="✅ Fertig! Alles wurde gesichert.", text_color="#34d399")
            self.start_button.configure(text="🚀 NEUE VERARBEITUNG")
            messagebox.showinfo("Abschluss", f"Verarbeitung erfolgreich beendet!\nDer Text wurde auf deinem Desktop gespeichert:\n{self.output_file}")

        self.start_button.configure(state="normal", fg_color="#38bdf8")
        self.abort_button.configure(state="disabled", text="🛑 Stopp", fg_color="#475569")
        self.pick_button.configure(state="normal")
        self.clear_button.configure(state="normal")
        self.copy_button.configure(state="normal")
        self.log("Batch-Sitzung beendet.")

    def copy_all(self):
        if not self.processing_results: return
        full_text = "\n\n".join([r[2] for r in self.processing_results])
        pyperclip.copy(full_text)
        self.log("Alle Ergebnisse in die Zwischenablage kopiert.")
        self.status_label.configure(text="📋 Text wurde in die Zwischenablage kopiert!")

class LogWindow(ctk.CTkToplevel):
    def __init__(self, master, current_logs):
        super().__init__(master)
        self.title("System-Logs & Live-Meldungen")
        self.geometry("600x400")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.textbox = ctk.CTkTextbox(self, font=ctk.CTkFont(family="Consolas", size=11))
        self.textbox.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        # Initial fill
        for l in current_logs:
            self.textbox.insert("end", l + "\n")
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

    def update_log(self, msg):
        self.textbox.configure(state="normal")
        self.textbox.insert("end", msg + "\n")
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = OCRExtractorGUI()
    app.mainloop()
