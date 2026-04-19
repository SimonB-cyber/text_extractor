import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import threading
import asyncio
import pyperclip
import sys
import multiprocessing
import datetime
import json
import time
import re
import urllib.request
import tempfile
import subprocess
import shlex
import shutil
import winreg
import ctypes
from ctypes import wintypes
from concurrent.futures import ProcessPoolExecutor
import importlib.util
import concurrent.futures
import traceback
import g4f

# --- Windows Silent Mode Fix ---
if os.name == 'nt':
    CREATE_NO_WINDOW = 0x08000000
    _orig_popen = subprocess.Popen
    def _silent_popen(*args, **kwargs):
        kwargs['creationflags'] = CREATE_NO_WINDOW | kwargs.get('creationflags', 0)
        return _orig_popen(*args, **kwargs)
    subprocess.Popen = _silent_popen

# --- GLOBAL CONFIG & HELPER ---
APP_VERSION = "1.2"
GITHUB_VERSION_URL = "https://raw.githubusercontent.com/SimonB-cyber/text_extractor/main/version.txt"
GITHUB_RELEASE_URL = "https://github.com/SimonB-cyber/text_extractor/releases/latest/download/TextExtraktPro_Setup_v{version}.exe"

TESS_PATH = None
TESSDATA_DIR = None

def get_config_path(filename):
    """Gibt einen schreibbaren Pfad im Benutzer-AppData-Verzeichnis zurück."""
    app_data = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "TextExtraktPro")
    os.makedirs(app_data, exist_ok=True)
    return os.path.join(app_data, filename)

def get_short_path_name(long_name):
    """Gets the short 8.3 path name of a given long path (Windows only)"""
    try:
        _GetShortPathNameW = ctypes.windll.kernel32.GetShortPathNameW
        _GetShortPathNameW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
        _GetShortPathNameW.restype = wintypes.DWORD
        output_buf_size = 4096
        output_buf = ctypes.create_unicode_buffer(output_buf_size)
        needed = _GetShortPathNameW(long_name, output_buf, output_buf_size)
        if 0 < needed < output_buf_size:
            return output_buf.value
    except:
        pass
    return long_name

def install_python_packages():
    """Prüft und installiert alle fehlenden Python-Module"""
    packages = {'pdfplumber': 'pdfplumber', 'PIL': 'pillow', 'pytesseract': 'pytesseract', 'pyperclip': 'pyperclip', 'customtkinter': 'customtkinter', 'g4f': 'g4f', 'symspellpy': 'symspellpy'}
    missing = []
    for imp_name, pip_name in packages.items():
        if importlib.util.find_spec(imp_name) is None:
            missing.append(pip_name)
    if missing:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', *missing, '--quiet'])

# Initial call to ensure modules are ready
install_python_packages()
import pdfplumber
import pytesseract
from PIL import Image, ImageEnhance, ImageOps

def auto_find_tesseract():
    """Hilfsfunktion zur automatischen Suche von Tesseract an Standardorten"""
    def find_in_registry():
        results = []
        for root in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
            for subkey in [r"SOFTWARE\Tesseract-OCR", r"SOFTWARE\WOW6432Node\Tesseract-OCR"]:
                try:
                    with winreg.OpenKey(root, subkey) as key:
                        install_path, _ = winreg.QueryValueEx(key, "InstallPath")
                        exe = os.path.join(install_path, "tesseract.exe")
                        if os.path.exists(exe): results.append(exe)
                except: continue
        return results[0] if results else None

    reg_exe = find_in_registry()
    if reg_exe: return reg_exe
    
    path_exe = shutil.which("tesseract")
    if path_exe: return path_exe
    
    possible_paths = [
        os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "Tesseract-OCR", "tesseract.exe"),
        os.path.join(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"), "Tesseract-OCR", "tesseract.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Tesseract-OCR", "tesseract.exe"),
        r"C:\Tesseract-OCR\tesseract.exe"
    ]
    for p in possible_paths:
        if os.path.exists(p): return p
    return None

def ensure_lang(l, target_dir):
    os.makedirs(target_dir, exist_ok=True)
    lang_file = os.path.join(target_dir, f"{l}.traineddata")
    if not os.path.exists(lang_file):
        url = f"https://github.com/tesseract-ocr/tessdata/raw/main/{l}.traineddata"
        try:
            print(f"Lade Sprachdatei herunter: {url} -> {lang_file}")
            urllib.request.urlretrieve(url, lang_file)
        except Exception as e:
            print(f"Fehler beim Download der Sprachdatei: {e}")
            return False
    return True

def setup_tesseract(lang='deu'):
    config_file = get_config_path("tesseract_path.txt")
    user_tessdata = get_config_path("tessdata")
    
    # 1. Config Check
    if os.path.exists(config_file):
        with open(config_file, 'r', encoding='utf-8') as f:
            p = f.read().strip().strip('"')
            if os.path.exists(p):
                pytesseract.pytesseract.tesseract_cmd = p
                exe_dir = os.path.dirname(p)
                system_tessdata = os.path.join(exe_dir, "tessdata")
                
                if os.path.exists(os.path.join(system_tessdata, f"{lang}.traineddata")):
                    return p, get_short_path_name(system_tessdata)
                elif ensure_lang(lang, user_tessdata):
                    return p, get_short_path_name(user_tessdata)

    # 2. Auto-Find
    exe_path = auto_find_tesseract()
    if not exe_path:
        # Versuche winget install
        try:
            subprocess.check_call(["winget", "install", "-e", "--id", "UB-Mannheim.TesseractOCR", "--accept-package-agreements", "--accept-source-agreements", "--silent"])
            time.sleep(2)
            exe_path = auto_find_tesseract()
        except: pass

    if exe_path:
        pytesseract.pytesseract.tesseract_cmd = exe_path
        with open(config_file, 'w', encoding='utf-8') as f: f.write(exe_path)
        
        exe_dir = os.path.dirname(exe_path)
        system_tessdata = os.path.join(exe_dir, "tessdata")
        
        # Check system tessdata first
        if os.path.exists(os.path.join(system_tessdata, f"{lang}.traineddata")):
            return exe_path, get_short_path_name(system_tessdata)
            
        # Fallback to user tessdata
        if ensure_lang(lang, user_tessdata):
            return exe_path, get_short_path_name(user_tessdata)
            
        # Last resort fallback to system anyway (might still fail)
        return exe_path, get_short_path_name(system_tessdata)
    
    return None, "WANT_PROMPT"

def get_tesseract_config(lang='deu'):
    global TESS_PATH, TESSDATA_DIR
    TESS_PATH, TESSDATA_DIR = setup_tesseract(lang=lang)
    return TESS_PATH, TESSDATA_DIR

def finde_seitenzahl_text_fallback(text):
    zeilen = [z.strip() for z in text.strip().split('\n') if z.strip()]
    if not zeilen: return -1
    for zeile in zeilen[:6] + zeilen[-6:]:
        # Sehr striktes Muster: Nur die Zahl, evtl. mit kleinen Symbolen drumherum
        treffer = re.search(r'^[~\s]*(\d+)[~\s]*$', zeile)
        if treffer: return int(treffer.group(1))
        # Muster für "Seite X" am Anfang/Ende
        treffer = re.search(r'^(?:Seite|S\.|Page|Pg\.)\s*(\d+)\s*$', zeile, re.IGNORECASE)
        if treffer: return int(treffer.group(1))
    return -1

def finde_seitenzahl_bild(bild):
    try:
        pytesseract.pytesseract.tesseract_cmd = TESS_PATH
        config = f'--tessdata-dir {TESSDATA_DIR}'
        w, h = bild.size
        data = pytesseract.image_to_data(bild, lang='deu', config=config, output_type=pytesseract.Output.DICT)
        
        candidates = []
        for i in range(len(data['text'])):
            txt = data['text'][i].strip()
            conf = int(data['conf'][i])
            
            if txt.isdigit() and 1 <= int(txt) <= 2000 and conf > 85:
                top = data['top'][i]
                # Nur akzeptieren wenn im oberen oder unteren Zehntel des Bildes
                if top < (h * 0.12) or top > (h * 0.88):
                    candidates.append((conf, int(txt)))
        
        if candidates:
            # Nimm den Kandidaten mit der höchsten Konfidenz
            candidates.sort(reverse=True)
            return candidates[0][1]
    except: pass
    return -1

def symspell_fallback(text):
    from symspellpy import SymSpell, Verbosity
    dict_path = get_config_path("de-100k.txt")
    if not os.path.exists(dict_path):
        url = "https://raw.githubusercontent.com/wolfgarbe/SymSpell/master/SymSpell/de-100k.txt"
        try: urllib.request.urlretrieve(url, dict_path)
        except: return text
    sym_spell = SymSpell(max_dictionary_edit_distance=2, prefix_length=7)
    sym_spell.load_dictionary(dict_path, term_index=0, count_index=1, separator=" ")
    words = re.findall(r'\S+|\s+', text)
    cleaned = []
    for w in words:
        if w.isspace() or len(w) < 3:
            cleaned.append(w)
            continue
        clean_w = re.sub(r'[^\w]', '', w)
        suggestions = sym_spell.lookup(clean_w, Verbosity.CLOSEST, max_edit_distance=2)
        if suggestions: cleaned.append(w.replace(clean_w, suggestions[0].term))
        else: cleaned.append(w)
    return "".join(cleaned)

def ai_clean_text(text, ai_config, logs_list):
    if not text.strip(): return text
    api_url = ai_config.get('api_url', '')
    api_key = ai_config.get('api_key', '')
    model = ai_config.get('model', 'gpt-4o-mini')
    sys_prompt = "Repariere den OCR-Text, behalte Inhalt und Absätze bei. Korrigiere Tippfehler und Worttrennungen. Antworte NUR mit dem bereinigten Text."
    
    # Try custom API first
    if api_url and "localhost" not in api_url:
        headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
        data = {"model": model, "messages": [{"role": "system", "content": sys_prompt}, {"role": "user", "content": text}], "temperature": 0.1}
        try:
            req = urllib.request.Request(api_url, data=json.dumps(data).encode('utf-8'), headers=headers, method='POST')
            with urllib.request.urlopen(req, timeout=30) as res:
                result = json.loads(res.read().decode('utf-8'))
                return result['choices'][0]['message']['content'].strip()
        except: pass
    
    # Fallback/Primary: g4f (GPT4Free)
    try:
        # Dynamisch nach funktionierenden Providern suchen
        working_providers = [p for p in g4f.Provider.__providers__ if p.working]
        
        response = None
        for provider in working_providers:
            if not provider.needs_auth:
                try:
                    # Verwende gpt-3.5-turbo als sichereren Fallback für viele Provider
                    model_to_use = "gpt-3.5-turbo" if provider.__name__ == "ApiAirforce" else "gpt-4"
                    
                    response = g4f.ChatCompletion.create(
                        model=model_to_use,
                        provider=provider,
                        messages=[{"role": "system", "content": sys_prompt}, {"role": "user", "content": text}],
                    )
                    if response and len(response) > 10:
                        # Filter für bekannte Fehlermeldungen von Providern
                        res_lower = response.lower()
                        if "model does not exist" in res_lower or "api key" in res_lower or "limit reached" in res_lower:
                            continue
                        return response.strip()
                except: continue
            
    except Exception as e:
        if logs_list: logs_list.append(f"AI Error (g4f): {e}")
    
    return symspell_fallback(text)

def ai_sort_pages(results):
    """Sortiert die Ergebnisse logisch basierend auf dem Textfluss zwischen den Seiten."""
    if len(results) < 2: return results
    
    # Extrahiere Snippets für die Sortierung
    snippets = []
    for idx, (pfad, sz, text, logs, *rest) in enumerate(results):
        cleaned_text = text.strip()
        start = cleaned_text[:400].replace('\n', ' ')
        end = cleaned_text[-400:].replace('\n', ' ')
        snippets.append(f"Seite {idx}: [START: {start}] ... [ENDE: {end}]")
    
    prompt = "Hier sind Textfragmente von verschiedenen Seiten eines Dokuments. Ordne sie in die logisch richtige Reihenfolge, sodass der Textfluss stimmt. Antworte NUR mit einer Komma-getrennten Liste der Seiten-Indizes (z.B. 3,0,2,...).\n\n" + "\n".join(snippets)
    
    try:
        working_providers = [p for p in g4f.Provider.__providers__ if p.working and not p.needs_auth]
        for provider in working_providers:
            try:
                response = g4f.ChatCompletion.create(
                    model="gpt-3.5-turbo", # Sichererer Standard
                    provider=provider,
                    messages=[{"role": "user", "content": prompt}],
                )
                if not response or len(response) < 1: continue
                # Filter für bekannte Fehlermeldungen von Providern
                res_lower = response.lower()
                if "model does not exist" in res_lower or "api key" in res_lower:
                    continue
                    
                # Extrahiere Indizes
                match = re.search(r'([\d\s,]+)', response)
                if match:
                    indices = [int(i.strip()) for i in match.group(1).split(',') if i.strip().isdigit()]
                    if len(set(indices)) == len(results): # Validiere Vollständigkeit
                        return [results[i] for i in indices]
            except: continue
    except: pass
    
    return sorted(results, key=lambda x: x[1])

def verarbeite_datei(pfad, file_index=0, tabelle=False, sprache='deu', ki_korrektur=False, ai_config=None):
    logs_list = []
    try:
        logs_list.append(f"Starte Verarbeitung für {os.path.basename(pfad)}...")
        t_path, t_data = get_tesseract_config(lang=sprache)
        if not t_path: 
            return pfad, 100000 + file_index, "FEHLER: Tesseract-Pfad nicht konfiguriert.", logs_list
        
        if t_data == "MISSING_DATA": 
            return pfad, 100000 + file_index, f"FEHLER: Tesseract-Sprachdatei ({sprache}) fehlt.", logs_list

        pytesseract.pytesseract.tesseract_cmd = t_path
        fname = os.path.basename(pfad)
        ext = fname.lower().split('.')[-1]
        res_txt, sz = "", 999999
        
        # OCR Config
        config = f'--tessdata-dir {t_data} --psm 6'
        if tabelle: config += ' -c preserve_interword_spaces=1'

        if ext == 'pdf':
            logs_list.append("Extrahiere PDF-Text...")
            with pdfplumber.open(pfad) as pdf:
                for s in pdf.pages:
                    t = s.extract_text()
                    if t: res_txt += t + "\n"
            sz = finde_seitenzahl_text_fallback(res_txt)
        elif ext in ['png', 'jpg', 'jpeg', 'bmp', 'tiff']:
            logs_list.append(f"Führe OCR auf Bild aus (Tesseract-Data: {t_data})...")
            b = Image.open(pfad)
            prep_img = ImageEnhance.Contrast(b.convert('L')).enhance(2.0)
            res_txt = pytesseract.image_to_string(prep_img, lang=sprache, config=config)
            sz = finde_seitenzahl_bild(b)
        
        final_text = res_txt.strip()
        if ki_korrektur and final_text:
            logs_list.append("Bereinige Text mit KI/Offline-Wörterbuch...")
            final_text = ai_clean_text(final_text, ai_config, logs_list)
            
        # Fallback-Index für Sortierung beibehalten, aber sz kann -1 sein
        sort_sz = sz if sz != -1 else 100000 + file_index
        logs_list.append(f"Datei verarbeitet. Seitenzahl erkannt: {sz}")
        return pfad, sort_sz, final_text, logs_list, sz
    except Exception as e:
        err_msg = f"FEHLER: {str(e)}\n{traceback.format_exc()}"
        logs_list.append(err_msg)
        return pfad, 100000 + file_index, f"FEHLER: {str(e)}", logs_list

# --- GUI CLASS ---
class OCRExtractorGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Premium OCR & Text Extraktor Pro")
        self.geometry("1000x800")
        self.files_to_process = []
        self.processing_results = []
        self.abort_event = threading.Event()
        self.logs = []
        self.log_window = None
        self.ai_config = {}
        self.ai_config_file = get_config_path("ai_config.json")
        if os.path.exists(self.ai_config_file):
            try:
                with open(self.ai_config_file, 'r', encoding='utf-8') as f:
                    self.ai_config = json.load(f)
            except: pass
        self.setup_ui()
        self.log("GUI gestartet. Lade Engines im Hintergrund...")
        threading.Thread(target=self.bg_init_engines, daemon=True).start()

    def bg_init_engines(self):
        try:
            self.after(0, lambda: self.status_label.configure(text="⚠️ Lade Text-Engine...", text_color="#fbbf24"))
            self.after(0, lambda: self.start_button.configure(state="disabled"))
            self.log("System-Check: Suche nach Tesseract-Installation...")
            t_path, status = get_tesseract_config()
            if status == "WANT_PROMPT" or not t_path:
                self.after(0, self.open_settings_window)
            else:
                self.log(f"Tesseract bereit: {t_path}")
            self.check_for_updates()
            self.after(0, lambda: self.status_label.configure(text="✅ Alle Engines geladen!", text_color="#34d399"))
            self.after(0, lambda: self.start_button.configure(state="normal"))
        except Exception as e: self.log(f"Fehler: {e}")

    def check_for_updates(self):
        try:
            self.log("Suche nach Updates...")
            with urllib.request.urlopen(GITHUB_VERSION_URL, timeout=5) as r:
                latest = r.read().decode('utf-8').strip()
            if latest > APP_VERSION: self.after(0, lambda: self._prompt_update(latest))
            else: self.log(f"Programm ist aktuell (v{APP_VERSION}).")
        except: pass

    def log(self, message):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        full_msg = f"[{timestamp}] {message}"
        self.logs.append(full_msg)
        if self.log_window and self.log_window.winfo_exists(): self.log_window.update_log(full_msg)
        try:
            self.console_box.configure(state="normal")
            self.console_box.insert("end", full_msg + "\n")
            self.console_box.see("end")
            self.console_box.configure(state="disabled")
        except: pass

    def open_settings_window(self):
        set_win = ctk.CTkToplevel(self)
        set_win.title("Einstellungen")
        set_win.geometry("550x450")
        set_win.attributes('-topmost', True)
        ctk.CTkLabel(set_win, text="Tesseract.exe Pfad:", font=ctk.CTkFont(weight="bold")).pack(pady=10)
        t_entry = ctk.CTkEntry(set_win, width=400)
        t_entry.pack(pady=5)
        if TESS_PATH: t_entry.insert(0, TESS_PATH)
        def save():
            global TESS_PATH
            TESS_PATH = t_entry.get()
            with open(get_config_path("tesseract_path.txt"), 'w', encoding='utf-8') as f: f.write(TESS_PATH)
            get_tesseract_config()
            set_win.destroy()
        ctk.CTkButton(set_win, text="Speichern", command=save).pack(pady=20)

    def setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        ctk.CTkLabel(self.header_frame, text="📚 Text-Extraktor Pro", font=ctk.CTkFont(size=28, weight="bold"), text_color="#38bdf8").pack(side="left")
        ctk.CTkButton(self.header_frame, text="📋 Logs", command=self.open_log_window, width=80).pack(side="right", padx=5)
        ctk.CTkButton(self.header_frame, text="⚙️", command=self.open_settings_window, width=40).pack(side="right")

        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        ctk.CTkButton(self.button_frame, text="+ Dateien", command=self.pick_files).pack(side="left", padx=5)
        ctk.CTkButton(self.button_frame, text="Leeren", command=self.clear_list, fg_color="#ef4444").pack(side="left", padx=20)
        self.lang_menu = ctk.CTkComboBox(self.button_frame, values=["deu", "eng", "fra", "spa"], width=80)
        self.lang_menu.set("deu")
        self.lang_menu.pack(side="left")
        self.table_switch = ctk.CTkSwitch(self.button_frame, text="Tabelle")
        self.table_switch.pack(side="left", padx=10)
        self.ai_switch = ctk.CTkSwitch(self.button_frame, text="🧠 KI", progress_color="#8b5cf6")
        self.ai_switch.pack(side="left")
        self.ai_switch.select()

        self.scroll_frame = ctk.CTkScrollableFrame(self, label_text="Warteschlange")
        self.scroll_frame.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        self.file_items = []

        self.progress_bar = ctk.CTkProgressBar(self, height=12)
        self.progress_bar.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        self.progress_bar.set(0)
        self.status_label = ctk.CTkLabel(self, text="Bereit.")
        self.status_label.grid(row=4, column=0)
        self.console_box = ctk.CTkTextbox(self, height=100, font=("Consolas", 11))
        self.console_box.grid(row=5, column=0, padx=20, pady=10, sticky="nsew")

        self.action_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.action_frame.grid(row=6, column=0, padx=20, pady=20, sticky="nsew")
        self.start_button = ctk.CTkButton(self.action_frame, text="🚀 STARTEN", height=50, command=self.start_processing_thread)
        self.start_button.pack(side="left", expand=True, fill="x", padx=5)
        self.copy_button = ctk.CTkButton(self.action_frame, text="📋 Kopieren", height=50, state="disabled", command=self.copy_all)
        self.copy_button.pack(side="right")

    def open_log_window(self):
        if self.log_window is None or not self.log_window.winfo_exists(): self.log_window = LogWindow(self, self.logs)
        else: self.log_window.focus()

    def pick_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Bilder/PDF", "*.jpg *.jpeg *.png *.pdf")])
        for f in files:
            if f not in self.files_to_process:
                self.files_to_process.append(f)
                item = ctk.CTkLabel(self.scroll_frame, text=f"⏳ {os.path.basename(f)}", anchor="w")
                item.pack(fill="x", pady=2)
                self.file_items.append({"path": f, "label": item})
        self.status_label.configure(text=f"{len(self.files_to_process)} Dateien geladen.")

    def clear_list(self):
        for i in self.file_items: i["label"].destroy()
        self.files_to_process, self.file_items = [], []
        self.progress_bar.set(0)

    def start_processing_thread(self):
        self.start_button.configure(state="disabled")
        threading.Thread(target=self.run_processing, daemon=True).start()

    def run_processing(self):
        total = len(self.files_to_process)
        lang = self.lang_menu.get()
        table = self.table_switch.get() == 1
        ai = self.ai_switch.get() == 1
        results = []
        with ProcessPoolExecutor() as ex:
            futures = {ex.submit(verarbeite_datei, p, idx, table, lang, ai, self.ai_config): p for idx, p in enumerate(self.files_to_process)}
            for i, future in enumerate(concurrent.futures.as_completed(futures)):
                res = future.result()
                results.append(res)
                # Zeige Engine-Logs im GUI-Log an
                if len(res) > 3:
                    for engine_log in res[3]:
                        self.log(f"-> {engine_log}")
                
                is_err = res[2].startswith("FEHLER")
                self.after(0, self.update_ui, i+1, total, res, is_err)
        
        # KI-basierte Sortierung falls nötig
        if ai and len(results) > 1:
            self.log("Führe KI-basierte Seitensortierung durch...")
            old_order = [os.path.basename(r[0]) for r in results]
            results = ai_sort_pages(results)
            new_order = [os.path.basename(r[0]) for r in results]
            
            if old_order != new_order:
                self.log(f"✅ KI-Sortierung abgeschlossen! Reihenfolge angepasst.")
                self.log(f"-> Erste Seite: {new_order[0]}")
            else:
                self.log("ℹ️ KI-Sortierung abgeschlossen. Keine Änderungen an der Reihenfolge nötig.")
            
        full_text = "\n\n".join([r[2] for r in results])
        self.processing_results = results
        with open(os.path.join(os.path.expanduser("~"), "Desktop", "OCR_Output.txt"), "w", encoding="utf-8") as f: f.write(full_text)
        self.after(0, lambda: self.status_label.configure(text="✅ Fertig!"))
        self.after(0, lambda: self.start_button.configure(state="normal"))
        self.after(0, lambda: self.copy_button.configure(state="normal"))
        self.after(0, lambda: messagebox.showinfo("Prozess abgeschlossen", "Alle Dateien wurden erfolgreich verarbeitet, bereinigt und sortiert!\n\nDas Gesamtergebnis wurde als 'OCR_Output.txt' auf dem Desktop gespeichert."))

    def update_ui(self, current, total, res, is_err):
        self.progress_bar.set(current/total)
        for item in self.file_items:
            if item["path"] == res[0]:
                if is_err:
                    item["label"].configure(text=f"❌ {os.path.basename(res[0])} - FEHLER", text_color="#ef4444")
                else:
                    orig_sz = res[4] if len(res) > 4 else res[1]
                    if orig_sz == -1:
                        item["label"].configure(text=f"ℹ️ {os.path.basename(res[0])} - Seitenzahl nicht gefunden", text_color="#fbbf24")
                    else:
                        item["label"].configure(text=f"✅ {os.path.basename(res[0])} - S. {orig_sz}", text_color="#38bdf8")

    def copy_all(self):
        pyperclip.copy("\n\n".join([r[2] for r in self.processing_results]))
        self.status_label.configure(text="Kopiert!")

class LogWindow(ctk.CTkToplevel):
    def __init__(self, master, logs):
        super().__init__(master)
        self.title("Logs")
        self.geometry("600x400")
        self.txt = ctk.CTkTextbox(self)
        self.txt.pack(fill="both", expand=True, padx=10, pady=10)
        for l in logs: self.txt.insert("end", l + "\n")
        self.txt.configure(state="disabled")
    def update_log(self, m):
        self.txt.configure(state="normal")
        self.txt.insert("end", m + "\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")

if __name__ == "__main__":
    multiprocessing.freeze_support()
    OCRExtractorGUI().mainloop()
