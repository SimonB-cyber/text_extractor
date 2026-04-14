import sys
import subprocess
import importlib.util
import os
import shlex
import shutil
import urllib.request
import re
import json
from concurrent.futures import ProcessPoolExecutor

# --- Windows Silent Mode Fix ---
if os.name == 'nt':
    import subprocess
    # Diese Flagge verhindert, dass sich bei jedem Tesseract-Aufruf ein CMD-Fenster öffnet
    CREATE_NO_WINDOW = 0x08000000
    _orig_popen = subprocess.Popen
    def _silent_popen(*args, **kwargs):
        kwargs['creationflags'] = CREATE_NO_WINDOW | kwargs.get('creationflags', 0)
        return _orig_popen(*args, **kwargs)
    subprocess.Popen = _silent_popen

def install_python_packages():
    """1. Prüft und installiert alle fehlenden Python/Pip-Module"""
    packages = {'pdfplumber': 'pdfplumber', 'PIL': 'pillow', 'pytesseract': 'pytesseract', 'pyperclip': 'pyperclip', 'customtkinter': 'customtkinter', 'g4f': 'g4f', 'symspellpy': 'symspellpy'}
    missing = []
    
    for imp_name, pip_name in packages.items():
        if importlib.util.find_spec(imp_name) is None:
            missing.append(pip_name)
    if missing:
        print(f"[System-Check] Installiere Python-Module: {', '.join(missing)}")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', *missing, '--quiet'])

install_python_packages()

import pdfplumber
import pytesseract
import pyperclip
from PIL import Image, ImageEnhance, ImageOps

def auto_find_tesseract():
    """Hilfsfunktion zur automatischen Suche von Tesseract an Standardorten"""
    import winreg
    
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
    
    local_appdata = os.environ.get("LOCALAPPDATA", "")
    program_files = os.environ.get("ProgramFiles", "C:\\Program Files")
    program_files_x86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)")
    
    possible_paths = [
        os.path.join(program_files, "Tesseract-OCR", "tesseract.exe"),
        os.path.join(program_files_x86, "Tesseract-OCR", "tesseract.exe"),
        os.path.join(local_appdata, "Tesseract-OCR", "tesseract.exe"),
        os.path.join(os.environ.get("USERPROFILE", ""), "AppData", "Local", "Tesseract-OCR", "tesseract.exe"),
        r"C:\Tesseract-OCR\tesseract.exe"
    ]
    for p in possible_paths:
        if os.path.exists(p): return p
    return None

def setup_tesseract(lang='deu'):
    """Mehrstufige Einrichtung: Suche -> Install -> Info/Prompt"""
    skript_ordner = os.path.dirname(os.path.abspath(__file__))
    config_file = os.path.join(skript_ordner, "tesseract_path.txt")
    failed_flag = os.path.join(skript_ordner, "setup_failed.txt")
    
    # 1. Config Check
    if os.path.exists(config_file):
        with open(config_file, 'r', encoding='utf-8') as f:
            p = f.read().strip().strip('"')
            if os.path.exists(p): 
                pytesseract.pytesseract.tesseract_cmd = p
                return p, os.path.join(skript_ordner, "tessdata")

    # 2. Auto-Find Standard-Orte
    exe_path = auto_find_tesseract()
    
    # 3. Wenn nicht gefunden, wurde es schon probiert?
    if not exe_path:
        if os.path.exists(failed_flag):
            return None, "WANT_PROMPT" # Zweiter Lauf -> Frage den User
        
        # Erster Lauf -> Versuche Install
        try:
            subprocess.check_call(["winget", "install", "-e", "--id", "UB-Mannheim.TesseractOCR", "--accept-package-agreements", "--accept-source-agreements", "--silent"])
            import time
            time.sleep(3)
            exe_path = auto_find_tesseract()
        except: pass
        
        if not exe_path:
            # Install gescheitert -> Flag setzen und Info zurückgeben
            with open(failed_flag, 'w') as f: f.write("failed")
            return None, "INSTALL_FAILED"

    # Erfolg! Pfade initialisieren
    pytesseract.pytesseract.tesseract_cmd = exe_path
    with open(config_file, 'w', encoding='utf-8') as f: f.write(exe_path)
    if os.path.exists(failed_flag): os.remove(failed_flag)

    # WICHTIG: Prüfen, ob Tesseract sein eigenes tessdata-Verzeichnis hat (Standard-Installation)
    exe_dir = os.path.dirname(exe_path)
    system_tessdata = os.path.join(exe_dir, "tessdata")
    
    # Lokaler Ordner als Fallback
    local_tessdata = os.path.join(skript_ordner, "tessdata")
    os.makedirs(local_tessdata, exist_ok=True)

    # Strategie: Wenn im System-Ordner die Sprache fehlt, laden wir sie lokal herunter
    lang_found_system = os.path.exists(os.path.join(system_tessdata, f"{lang}.traineddata")) if os.path.exists(system_tessdata) else False
    
    if lang_found_system:
        # Tesseract findet seine eigenen Daten am besten selbst
        if "TESSDATA_PREFIX" in os.environ: del os.environ["TESSDATA_PREFIX"]
        return exe_path, system_tessdata
    else:
        # Wir müssen unsere lokalen Daten nutzen
        lang_file = os.path.join(local_tessdata, f"{lang}.traineddata")
        if not os.path.exists(lang_file):
            url = f"https://github.com/tesseract-ocr/tessdata/raw/main/{lang}.traineddata"
            try: urllib.request.urlretrieve(url, lang_file)
            except: pass
        
        os.environ["TESSDATA_PREFIX"] = local_tessdata
        return exe_path, local_tessdata



# Config variables
TESS_PATH = None
TESSDATA_DIR = None

def get_tesseract_config(lang='deu'):
    global TESS_PATH, TESSDATA_DIR
    TESS_PATH, TESSDATA_DIR = setup_tesseract(lang=lang)
    return TESS_PATH, TESSDATA_DIR


def split_drag_drop_paths(eingabe_str):
    if not eingabe_str.strip(): return []
    try:
        pfade_teile = shlex.split(eingabe_str, posix=False)
    except:
        pfade_teile = [s.strip().strip('"') for s in eingabe_str.split(' "')]
        
    gueltige_pfade = []
    for teil in pfade_teile:
        sauberer_pfad = teil.strip('"').strip("'")
        if os.path.exists(sauberer_pfad): gueltige_pfade.append(sauberer_pfad)
    return gueltige_pfade


def finde_seitenzahl_text_fallback(text):
    zeilen = [z.strip() for z in text.strip().split('\n') if z.strip()]
    if not zeilen: return 999999
    pruef_zeilen = zeilen[:6] + zeilen[-6:]
    for zeile in pruef_zeilen:
        treffer = re.search(r'^[-~_\|\.\s]*(\d+)[-~_\|\.\s]*$', zeile)
        if treffer: return int(treffer.group(1))
        
        treffer = re.search(r'(?:Seite|S\.|Page|Pg\.|P\.)\s*(\d+)\s*$', zeile, re.IGNORECASE)
        if treffer: return int(treffer.group(1))
        
        # User-Specific: "63 / Kapitel" oder "Kapitel / 63"
        treffer = re.search(r'^(\d+)\s*/\s*.+', zeile)
        if treffer: return int(treffer.group(1))
        
        treffer = re.search(r'.+\s*/\s*(\d+)$', zeile)
        if treffer: return int(treffer.group(1))
        
    return -1


def finde_seitenzahl_bild(bild):
    def score_kandidaten(data, w, h, region_pct_offset=0):
        zg = {}
        for i in range(len(data['text'])):
            txt = data['text'][i].strip()
            if not txt: continue
            top = data['top'][i]
            y_key = (top // 15) * 15
            if y_key not in zg: zg[y_key] = []
            zg[y_key].append({'text': txt, 'top': top, 'left': data['left'][i]})

        kand = []
        for y in zg:
            gruppe = sorted(zg[y], key=lambda x: x['left'])
            ganze = " ".join([g['text'] for g in gruppe])
            pct = max(0, min(1, (sum([g['top'] for g in gruppe]) / len(gruppe) + region_pct_offset) / h))
            if not (pct < 0.12 or pct > 0.88 or re.search(rf"(?:/|-|\||—|Seite|S\.|Page|P\.)", ganze, re.IGNORECASE)): continue
            for z_str in re.findall(r'\d+', ganze):
                z = int(z_str)
                if 1 <= z <= 1000:
                    score = 600 if min(pct, 1-pct) < 0.08 else 200
                    score += 1000 if 10 <= z <= 999 else -500
                    kand.append({'zahl': z, 'score': score})
        return kand

    try:
        pytesseract.pytesseract.tesseract_cmd = TESS_PATH
        os.environ["TESSDATA_PREFIX"] = TESSDATA_DIR
        w, h = bild.size
        data = pytesseract.image_to_data(bild, lang='deu', output_type=pytesseract.Output.DICT)
        kand = score_kandidaten(data, w, h)
        if kand:
            kand.sort(key=lambda x: x['score'], reverse=True)
            if kand[0]['score'] > 1200: return kand[0]['zahl']
        for r in ['top', 'bottom']:
            y_s = 0 if r == 'top' else int(h * 0.80)
            crop = bild.crop((0, y_s, w, int(h * 0.20) if r == 'top' else h)).convert('L')
            prep = ImageEnhance.Contrast(crop).enhance(2.0).resize((crop.width * 3, crop.height * 3), Image.Resampling.LANCZOS)
            data_d = pytesseract.image_to_data(prep, lang='deu', config='--psm 6', output_type=pytesseract.Output.DICT)
            kand.extend(score_kandidaten(data_d, w, h, region_pct_offset=y_s))
        if kand:
            kand.sort(key=lambda x: x['score'], reverse=True)
            if kand[0]['score'] > 300: return kand[0]['zahl']
    except: pass
    return 999999

def symspell_fallback(text):
    import os
    import urllib.request
    from symspellpy import SymSpell, Verbosity
    
    skript_ordner = os.path.dirname(os.path.abspath(__file__))
    dict_path = os.path.join(skript_ordner, "de-100k.txt")
    
    if not os.path.exists(dict_path):
        print("[System] Lade offline SymSpell Wörterbuch herunter...")
        try:
            url = "https://raw.githubusercontent.com/wolfgarbe/SymSpell/master/SymSpell/de-100k.txt"
            urllib.request.urlretrieve(url, dict_path)
        except Exception as e:
            return text + f"\n[SymSpell Offline-Dict laden fehlgeschlagen: {str(e)}]"

    sym_spell = SymSpell(max_dictionary_edit_distance=2, prefix_length=7)
    sym_spell.load_dictionary(dict_path, term_index=0, count_index=1, separator=" ")
    
    # Simple word by word replacement
    words = re.findall(r'\S+|\s+', text)
    cleaned = []
    for w in words:
        if w.isspace():
            cleaned.append(w)
            continue
            
        # Filter out extreme garbage (more special chars than letters/numbers)
        symbols = len(re.findall(r'[^a-zA-Z0-9äöüÄÖÜß\s.,!?:-]', w))
        if symbols > len(w) * 0.4:  # If 40% of word is special chars like %^°€
            continue # drop it entirely
            
        if len(w) < 3 or not any(c.isalpha() for c in w):
            cleaned.append(w)
            continue
            
        # Remove punctuation for lookup
        clean_w = re.sub(r'[^\w]', '', w)
        if not clean_w:
            cleaned.append(w)
            continue
            
        suggestions = sym_spell.lookup(clean_w, Verbosity.CLOSEST, max_edit_distance=2)
        if suggestions:
            # Re-apply non-letters if simple (just replace the core word)
            best = suggestions[0].term
            if w[0].isupper(): best = best.capitalize()
            res = w.replace(clean_w, best)
            cleaned.append(res)
        else:
            cleaned.append(w)
            
    return "".join(cleaned)


def ai_clean_text(text, ai_config, logs_list):
    if not text.strip(): return text
    
    api_url = ai_config.get('api_url', '') if ai_config else ""
    api_key = ai_config.get('api_key', '') if ai_config else ""
    model = ai_config.get('model', 'gpt-4o-mini') if ai_config else 'gpt-4o-mini'
    
    sys_prompt = "Du bist ein Experten-Assistent für OCR-Texte. Deine Aufgabe: Behebe Rechtschreibfehler, entferne Müll-Zeichen (Garbage, Fragmente von Linien), behalte aber den Sinn, Fokus und die Absätze PERFEKT bei. Erfinde keinen Text dazu. Antworte NUR ausnahmslos mit dem reparierten Text."
    
    # 1. Custom API given? (z.B. User hat eigene Ollama URL eingetragen!)
    if api_url and "api.openai.com" not in api_url:
        if "11434" in api_url and "v1" not in api_url:
            if not api_url.endswith("/api/generate"): 
                api_url = api_url.rstrip("/") + "/v1/chat/completions"
        
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key if api_key else 'sk-local'}"
        }
        data = {
            "model": model,
            "messages": [
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": text}
            ],
            "temperature": 0.1
        }
        import urllib.request
        import json
        try:
            req = urllib.request.Request(api_url, data=json.dumps(data).encode('utf-8'), headers=headers, method='POST')
            with urllib.request.urlopen(req, timeout=30) as response:
                result = json.loads(response.read().decode('utf-8'))
                if 'choices' in result and len(result['choices']) > 0:
                    cleaned = result['choices'][0]['message']['content'].strip()
                    if cleaned:
                        logs_list.append("Erfolgreich gereinigt über eigene Custom API.")
                        return cleaned
        except Exception as e:
            logs_list.append(f"Custom API fehlgeschlagen: {str(e)}")
            # Fall through

    # 2. GPT4Free (g4f) - Versuch der absolut kostenlosen Online-KI
    try:
        import g4f
        response = g4f.ChatCompletion.create(
            model=g4f.models.default,
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": text}
            ]
        )
        if response and isinstance(response, str):
            cleaned = response.strip()
            if cleaned:
                logs_list.append("Erfolgreich gereinigt über gratis g4f Cloud.")
                return cleaned
    except Exception as e:
        logs_list.append(f"g4f Cloud KI fehlgeschlagen: {str(e)}")
        # Fall through
        
    # 3. Offline Dictionary Fallback (SymSpell) -> 100% Local and robust!
    try:
        res = symspell_fallback(text)
        logs_list.append("Text mit lokalem Offline-Wörterbuch (Symspell) ausgebessert.")
        return res
    except Exception as e:
        logs_list.append(f"Offline-Korrektur fehlgeschlagen: {str(e)}")
        return text


def verarbeite_datei(pfad, file_index=0, tabelle=False, sprache='deu', ki_korrektur=False, ai_config=None):
    logs_list = []
    t_path, t_data = get_tesseract_config(lang=sprache)
    if not t_path: 
        logs_list.append("FEHLER: Tesseract nicht gefunden.")
        return pfad, 100000 + file_index, "FEHLER: Tesseract nicht gefunden.", logs_list
    
    pytesseract.pytesseract.tesseract_cmd = t_path
    os.environ["TESSDATA_PREFIX"] = t_data
    
    fname = os.path.basename(pfad)
    ext = fname.lower().split('.')[-1]
    res_txt, sz = "", 999999
    
    # OCR Config
    config = f'--psm 6'
    if tabelle:
        config = f'--psm 6 -c preserve_interword_spaces=1'
        
    try:
        name_no_ext = os.path.splitext(fname)[0]
        if name_no_ext.isdigit(): sz = int(name_no_ext)
        
        if ext == 'pdf':
            with pdfplumber.open(pfad) as pdf:
                for s in pdf.pages:
                    t = s.extract_text()
                    if t: res_txt += t + "\n"
            if sz == 999999: sz = finde_seitenzahl_text_fallback(res_txt)
        elif ext in ['png', 'jpg', 'jpeg', 'bmp', 'tiff']:
            b = Image.open(pfad)
            
            # --- BILD VORVERARBEITUNG FÜR BESSERE OCR ---
            prep_img = b.convert('L')
            enhancer = ImageEnhance.Contrast(prep_img)
            prep_img = enhancer.enhance(2.0)
            
            res_txt = pytesseract.image_to_string(prep_img, lang=sprache, config=config)
            if sz == 999999: sz = finde_seitenzahl_bild(b)
        
        final_text = res_txt.strip()
        if ki_korrektur:
            final_text = ai_clean_text(final_text, ai_config, logs_list)
            
        # Post-Processing Seitenzahl nach KI-bereinigtem Text
        if sz == 999999: 
            found = finde_seitenzahl_text_fallback(final_text)
            sz = found if found != -1 else (100000 + file_index)
            if found != -1: logs_list.append(f"Seitenzahl {found} im Text gefunden.")
            else: logs_list.append(f"Keine Seitenzahl gefunden -> Nutze Listen-Position {file_index}.")
            
        return pfad, sz, final_text, logs_list
    except Exception as e:
        logs_list.append(f"Engine-Fehler: {str(e)}")
        return pfad, 100000 + file_index, f"FEHLER: {str(e)}", logs_list



def main():
    print("\n" + "=" * 55)
    print(" 📚 TEXT-EXTRAKTOR - PARALLEL MODUS")
    ausgabe_datei = "gesammelter_text.txt"
    with open(ausgabe_datei, 'w', encoding='utf-8') as f: f.write("")
    while True:
        eingabe = input("\n👉 Bilder ziehen und ENTER (oder 'exit'): ")
        if eingabe.strip().lower() in ['exit', 'quit']: break
        files = split_drag_drop_paths(eingabe)
        if not files: continue
        results = []
        with ProcessPoolExecutor() as ex:
            for res in ex.map(verarbeite_datei, files):
                results.append(res)
                print(f" ✅ {os.path.basename(res[0])} -> Seite {res[1] if res[1] != 999999 else '?'}")
        results.sort(key=lambda x: x[1])
        with open(ausgabe_datei, 'a', encoding='utf-8') as f:
            for p, s, t in results: f.write(t + "\n\n")
        pyperclip.copy("\n\n".join([r[2] for r in results]))
        print(f"Fertig! Text in '{ausgabe_datei}' und Zwischenablage.")

if __name__ == "__main__":
    main()