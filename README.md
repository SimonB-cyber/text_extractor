# 📚 TextExtraktor Pro

[![Version](https://img.shields.io/badge/Version-1.1-blue.svg)](https://github.com/SimonB-cyber/text_extractor)
[![Python](https://img.shields.io/badge/Python-3.10%2B-green.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

**TextExtraktor Pro** ist eine hochperformante Desktop-Anwendung zur präzisen Textextraktion aus Bildern und PDFs. Durch die Kombination moderner OCR-Technologie mit KI-gestützter Fehlerkorrektur (via Llama3/OpenAI) bietet dieses Tool eine unvergleichliche Genauigkeit bei der Digitalisierung von Dokumenten.

![TextExtraktor Pro Mockup](https://raw.githubusercontent.com/SimonB-cyber/text_extractor/main/docs/gui_preview.png)

---

## 🌟 Highlights

- **🚀 Batch-Verarbeitung**: Verarbeite hunderte Dateien gleichzeitig dank effizientem Multiprocessing.
- **🧠 KI-Korrektur**: Optionale Integration von lokalen KI-Modellen (z.B. Ollama/Llama3) oder OpenAI, um OCR-Fehler automatisch zu korrigieren.
- **📊 Tabellen-Modus**: Spezialisierte Extraktionslogik für tabellarische Daten zur Erhaltung des Layouts.
- **🌍 Multi-Language Support**: Unterstützt Deutsch, Englisch, Französisch, Spanisch und Italienisch.
- **🎨 Modernes UI**: Ein elegantes, dunkles Interface basierend auf `CustomTkinter` für maximale Benutzerfreundlichkeit.
- **🔄 Auto-Update**: Bleiben Sie immer auf dem neuesten Stand mit dem integrierten GitHub-Update-System.

---

## 🚀 Installation

### Option 1: Installer (Empfohlen für Windows)
Lade die neueste `.exe` Datei unter [Releases](https://github.com/SimonB-cyber/text_extractor/releases) herunter und starte das Setup.

### Option 2: Aus dem Source Code
Wenn du Python installiert hast, kannst du das Repository klonen und die Abhängigkeiten installieren:

```bash
# Repository klonen
git clone https://github.com/SimonB-cyber/text_extractor.git
cd text_extractor

# Abhängigkeiten installieren
pip install -r requirements.txt
```

**Anforderungen:**
- Python 3.10+
- [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) (wird oft automatisch erkannt oder kann in den Einstellungen konfiguriert werden).

---

## 📖 Bedienungsanleitung

1. **Dateien wählen**: Ziehe Bilder oder PDFs einfach in das Fenster oder nutze den "+ Dateien wählen"-Button.
2. **Modus wählen**: Wähle die Sprache und entscheide, ob der **Tabellen-Modus** oder die **KI-Korrektur** aktiv sein soll.
3. **Starten**: Klicke auf den Raketen-Button. Die Ergebnisse werden automatisch auf deinem Desktop in der Datei `gesammelter_text.txt` gespeichert.

### ⚙️ KI-Integration (Optional)
Für die beste Qualität kannst du in den Einstellungen (`⚙️`) eine Verbindung zu einem lokalen Ollama-Server oder einem OpenAI-kompatiblen API-Endpunkt herstellen. Dies ermöglicht dem Programm, den extrahierten Text intelligent nachzubereiten.

---

## 🛠 Tech-Stack

- **Frontend**: [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) (Modernes UI Framework)
- **OCR Engine**: [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)
- **Parallelisierung**: Python `multiprocessing` & `concurrent.futures`
- **KI-Interface**: Lokale & Remote API Integration

---

## 🤝 Mitwirken
Beiträge sind herzlich willkommen! Wenn du Fehler findest oder neue Features vorschlagen möchtest, öffne bitte ein [Issue](https://github.com/SimonB-cyber/text_extractor/issues) oder sende einen Pull Request.

---

## 📄 Lizenz
Dieses Projekt steht unter der MIT-Lizenz. Siehe [LICENSE](LICENSE) für Details.

---
*Entwickelt von SimonB-cyber*