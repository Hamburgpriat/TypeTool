# Copyright Mario Wilhelm
import sys
import os
if hasattr(sys, 'frozen'):
    import ctypes
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

import pyperclip
import keyboard
import time
import pystray
from PIL import Image, ImageDraw, ImageFont
import threading
import tkinter as tk
import json
import subprocess
import logging
import signal
import atexit

# Logging konfigurieren
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('typetool.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Pfad zur Konfigurationsdatei
config_file = 'config.json'

# Standardkonfiguration
default_config = {
    'enter_key_enabled': False,  # Standardmäßig deaktiviert
    'hotkey': 'ctrl+b',
    'toggle_enter_hotkey': 'ctrl+alt+b',
    'typing_delay': 0.001,  # Schnellere Tippgeschwindigkeit
    'logging_enabled': True
}

# Funktion zum Laden der Konfiguration
def load_config():
    if os.path.exists(config_file):
        with open(config_file, 'r', encoding='utf-8') as file:
            loaded_config = json.load(file)
            # Merge mit default_config um neue Einstellungen hinzuzufügen
            merged_config = default_config.copy()
            merged_config.update(loaded_config)
            return merged_config
    else:
        return default_config

# Funktion zum Speichern der Konfiguration
def save_config(config):
    with open(config_file, 'w', encoding='utf-8') as file:
        json.dump(config, file, indent=2, ensure_ascii=False)

# Konfiguration laden
config = load_config()
press_enter = config.get('enter_key_enabled', False)
hotkey = config.get('hotkey', 'ctrl+b')
toggle_enter_hotkey = config.get('toggle_enter_hotkey', 'ctrl+alt+b')
typing_delay = config.get('typing_delay', 0.01)
tray_icon = None  # Globale Variable für das Tray-Icon
running = True

# Funktion zum Erstellen des Tray-Icons
def create_image():
    width = 64
    height = 64
    image = Image.new('RGB', (width, height), (255, 255, 255))
    dc = ImageDraw.Draw(image)
    font = ImageFont.truetype("arial", 14)
    text = "Type\nTool"
    text_bbox = dc.textbbox((0, 0), text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]
    text_x = (width - text_width) // 2
    text_y = (height - text_height) // 2
    dc.text((text_x, text_y), text, fill='black', font=font, align="center")
    return image

# Event zum Abbrechen des Tippvorgangs
stop_typing_event = threading.Event()

# Funktion zum Schreiben von Text (angepasst mit konfigurierbarer Geschwindigkeit)
def type_text(text):
    stop_typing_event.clear()
    logger.info(f"Tippe Text: {text[:50]}{'...' if len(text) > 50 else ''}")
    
    # Zeige Vorschau-Fenster an
    if config.get('show_preview_window', True):
        show_preview_window(text)
    
    for char in text:
        if stop_typing_event.is_set():
            break
        keyboard.write(char)
        time.sleep(typing_delay)  # Konfigurierbare Geschwindigkeit
    if press_enter and not stop_typing_event.is_set():
        keyboard.press_and_release('enter')
        logger.info("Enter-Taste gedrückt")

# Funktion zum Starten des Tippvorgangs in einem Thread
typing_thread = None

def toggle_typing():
    global typing_thread
    if typing_thread and typing_thread.is_alive():
        stop_typing_event.set()
        logger.info("Tippvorgang abgebrochen")
    else:
        stop_typing_event.clear()
        try:
            text = pyperclip.paste()
            if not text or not text.strip():
                logger.warning("Zwischenablage ist leer oder enthält nur Leerzeichen")
                show_popup("Zwischenablage ist leer!")
                return
            typing_thread = threading.Thread(target=type_text, args=(text,))
            typing_thread.start()
        except pyperclip.PyperclipException as e:
            logger.error(f"Fehler beim Zugriff auf die Zwischenablage: {e}")
            show_popup("Fehler beim Zugriff auf Zwischenablage!")

# Funktion zum Beenden des Programms
def cleanup_and_exit():
    global running
    running = False
    logger.info("Programm wird beendet...")
    
    # Entferne alle Hotkeys
    try:
        keyboard.unhook_all()
        logger.info("Alle Hotkeys entfernt")
    except Exception as e:
        logger.error(f"Fehler beim Entfernen der Hotkeys: {e}")
    
    # Beende Tray-Icon
    if tray_icon:
        tray_icon.stop()
    
    # Entferne Lock-Datei
    try:
        if os.path.exists('typetool.lock'):
            os.remove('typetool.lock')
            logger.info("Lock-Datei entfernt")
    except Exception as e:
        logger.error(f"Fehler beim Entfernen der Lock-Datei: {e}")
    
    logger.info("Programm beendet")
    sys.exit(0)

def on_quit(icon, item):
    cleanup_and_exit()

# Funktion zum Umschalten der Enter-Taste
def toggle_enter(icon=None, item=None):
    global press_enter
    press_enter = not press_enter
    config['enter_key_enabled'] = press_enter
    save_config(config)
    if tray_icon:
        update_menu(tray_icon)
    show_popup(f"Enter nach Text: {'An' if press_enter else 'Aus'}")
    logger.info(f"Enter-Taste umgeschaltet: {'An' if press_enter else 'Aus'}")

# Funktion zum Ändern der Hotkeys (ohne Neustart)
def change_hotkey(icon=None, item=None):
    def show_hotkey_window():
        def is_modifier_only(event):
            return event.keysym.lower() in ['shift_l', 'shift_r', 'control_l', 'control_r', 'alt_l', 'alt_r']

        def on_entry_hotkey(event):
            keys = []
            if event.state & 0x4:  # Control
                keys.append('ctrl')
            if event.state & 0x1:  # Shift
                keys.append('shift')
            if event.state & 0x20000:  # Alt
                keys.append('alt')
            if not is_modifier_only(event):
                keys.append(event.keysym.lower())
            hotkey_str = '+'.join(keys)
            new_hotkey.set(hotkey_str)
            return "break"

        def on_entry_toggle_enter_hotkey(event):
            keys = []
            if event.state & 0x4:
                keys.append('ctrl')
            if event.state & 0x1:
                keys.append('shift')
            if event.state & 0x20000:
                keys.append('alt')
            if not is_modifier_only(event):
                keys.append(event.keysym.lower())
            hotkey_str = '+'.join(keys)
            new_toggle_enter_hotkey.set(hotkey_str)
            return "break"

        def save_new_hotkeys():
            global hotkey, toggle_enter_hotkey
            new_hotkey_value = new_hotkey.get()
            new_toggle_enter_hotkey_value = new_toggle_enter_hotkey.get()
            if new_hotkey_value:
                keyboard.remove_hotkey(hotkey)
                hotkey = new_hotkey_value
                config['hotkey'] = hotkey
                keyboard.add_hotkey(hotkey, toggle_typing)
            if new_toggle_enter_hotkey_value:
                keyboard.remove_hotkey(toggle_enter_hotkey)
                toggle_enter_hotkey = new_toggle_enter_hotkey_value
                config['toggle_enter_hotkey'] = toggle_enter_hotkey
                keyboard.add_hotkey(toggle_enter_hotkey, toggle_enter)
            save_config(config)
            root.destroy()
            show_popup("Hotkeys erfolgreich geändert!")

        def reset_to_default():
            new_hotkey.set(default_config['hotkey'])
            new_toggle_enter_hotkey.set(default_config['toggle_enter_hotkey'])

        root = tk.Tk()
        root.title("Hotkeys ändern")
        root.geometry("400x340")
        root.configure(bg="#f0f0f0")

        label_info = tk.Label(root, text="Drücke im Feld die gewünschte Tastenkombination!", bg="#f0f0f0", font=("Helvetica", 10), fg="blue")
        label_info.pack(pady=5)

        label_hotkey = tk.Label(root, text="Neuer Hotkey:", bg="#f0f0f0", font=("Helvetica", 12))
        label_hotkey.pack(pady=10)

        new_hotkey = tk.StringVar(value=hotkey)
        entry_hotkey = tk.Entry(root, textvariable=new_hotkey, font=("Helvetica", 12))
        entry_hotkey.pack(pady=5)
        entry_hotkey.focus_set()
        entry_hotkey.bind('<KeyPress>', on_entry_hotkey)

        label_toggle_enter_hotkey = tk.Label(root, text="Hotkey für Enter drücken:", bg="#f0f0f0", font=("Helvetica", 12))
        label_toggle_enter_hotkey.pack(pady=10)

        new_toggle_enter_hotkey = tk.StringVar(value=toggle_enter_hotkey)
        entry_toggle_enter_hotkey = tk.Entry(root, textvariable=new_toggle_enter_hotkey, font=("Helvetica", 12))
        entry_toggle_enter_hotkey.pack(pady=5)
        entry_toggle_enter_hotkey.bind('<KeyPress>', on_entry_toggle_enter_hotkey)

        button_save = tk.Button(root, text="Speichern", command=save_new_hotkeys, font=("Helvetica", 12), bg="#4CAF50", fg="white")
        button_save.pack(pady=10)

        button_reset = tk.Button(root, text="Auf Standard zurücksetzen", command=reset_to_default, font=("Helvetica", 12), bg="#f44336", fg="white")
        button_reset.pack(pady=5)

        root.mainloop()

    threading.Thread(target=show_hotkey_window, daemon=True).start()

# Funktion zum Einstellen der Tippgeschwindigkeit
def change_typing_speed(icon=None, item=None):
    def show_speed_window():
        def save_speed():
            try:
                new_delay = float(speed_var.get())
                if 0.001 <= new_delay <= 1.0:
                    global typing_delay
                    typing_delay = new_delay
                    config['typing_delay'] = typing_delay
                    save_config(config)
                    show_popup(f"Tippgeschwindigkeit geändert: {typing_delay}s")
                    logger.info(f"Tippgeschwindigkeit geändert: {typing_delay}s")
                    root.destroy()
                else:
                    show_popup("Bitte einen Wert zwischen 0.001 und 1.0 eingeben!")
            except ValueError:
                show_popup("Bitte eine gültige Zahl eingeben!")

        root = tk.Tk()
        root.title("Tippgeschwindigkeit ändern")
        root.geometry("300x200")
        root.configure(bg="#f0f0f0")
        root.resizable(False, False)

        label = tk.Label(root, text="Tippgeschwindigkeit (Sekunden zwischen Zeichen):", 
                        bg="#f0f0f0", font=("Helvetica", 12))
        label.pack(pady=20)

        speed_var = tk.StringVar(value=str(typing_delay))
        entry = tk.Entry(root, textvariable=speed_var, font=("Helvetica", 12), width=10)
        entry.pack(pady=10)
        entry.focus_set()

        button_save = tk.Button(root, text="Speichern", command=save_speed, 
                               font=("Helvetica", 12), bg="#4CAF50", fg="white")
        button_save.pack(pady=10)

        root.bind('<Return>', lambda e: save_speed())
        root.bind('<Escape>', lambda e: root.destroy())

        root.mainloop()

    threading.Thread(target=show_speed_window, daemon=True).start()

# Funktion zum Umschalten des Loggings
def toggle_logging(icon=None, item=None):
    global config
    config['logging_enabled'] = not config.get('logging_enabled', True)
    save_config(config)
    
    # Aktualisiere Logging-Level
    if config['logging_enabled']:
        logging.getLogger().setLevel(logging.INFO)
        logger.info("Logging aktiviert")
    else:
        logging.getLogger().setLevel(logging.ERROR)
        logger.info("Logging deaktiviert")
    
    if tray_icon:
        update_menu(tray_icon)
    show_popup(f"Logging: {'An' if config['logging_enabled'] else 'Aus'}")

# Funktion zum Umschalten des Vorschau-Fensters
def toggle_preview(icon=None, item=None):
    global config
    config['show_preview_window'] = not config.get('show_preview_window', True)
    save_config(config)
    
    if tray_icon:
        update_menu(tray_icon)
    show_popup(f"Vorschau-Fenster: {'An' if config['show_preview_window'] else 'Aus'}")
    logger.info(f"Vorschau-Fenster umgeschaltet: {'An' if config['show_preview_window'] else 'Aus'}")

# Funktion zum Neustarten des Programms
def restart_program():
    python = sys.executable
    script = os.path.abspath(sys.argv[0])
    starter = os.path.join(os.path.dirname(script), "starter.py")
    subprocess.Popen([python, starter, script])
    sys.exit(0)

# Funktion zum Neustarten über das Tray-Menü
def on_restart(icon, item):
    # Tray-Icon sauber stoppen, dann Neustart auslösen
    icon.stop()
    restart_program()

# Funktion zum Aktualisieren des Tray-Menüs
def update_menu(icon):
    icon.menu = pystray.Menu(
        pystray.MenuItem("Vorschau-Fenster: " + ("An" if config.get('show_preview_window', True) else "Aus"), toggle_preview),
        pystray.MenuItem("Enter nach Text: " + ("An" if press_enter else "Aus"), toggle_enter),
        pystray.MenuItem("Hotkeys ändern", change_hotkey),
        pystray.MenuItem("Tippgeschwindigkeit ändern", change_typing_speed),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Neustarten", on_restart),
        pystray.MenuItem("Beenden", on_quit)
    )

# Funktion zum Anzeigen eines Popups
def show_popup(message):
    def create_popup():
        try:
            root = tk.Tk()
            root.overrideredirect(1)
            root.attributes("-topmost", True)
            root.geometry(f"+{root.winfo_screenwidth()-200}+{root.winfo_screenheight()-100}")
            label = tk.Label(root, text=message, bg="yellow", fg="black", font=("Helvetica", 12))
            label.pack()
            root.after(2000, root.destroy)
            root.mainloop()
        except Exception as e:
            logger.error(f"Fehler beim Erstellen des Popups: {e}")
    
    threading.Thread(target=create_popup, daemon=True).start()

# Funktion zum Anzeigen des Vorschau-Fensters
def show_preview_window(text):
    text_length = len(text)
    
    # Bei 50+ Zeichen nur Warnung in der Mitte anzeigen
    if text_length >= 50:
        show_warning_popup(text_length)
        return
    
    # Normale Vorschau für weniger als 50 Zeichen
    def create_preview():
        try:
            preview_root = tk.Tk()
            preview_root.title("")  # Kein Titel
            preview_root.overrideredirect(True)  # Kein Fensterrahmen
            preview_root.attributes("-topmost", True)
            preview_root.attributes("-alpha", 0.9)  # Leicht transparent
            preview_root.configure(bg="#2c2c2c")  # Dunkler Hintergrund
            
            # Dynamische Fenstergröße basierend auf Textlänge
            if text_length <= 30:
                width, height = 250, 70
            elif text_length <= 40:
                width, height = 300, 85
            else:
                width, height = 350, 100
            
            preview_root.geometry(f"{width}x{height}+20+20")  # Oben links
            
            # Text-Vorschau
            text_frame = tk.Frame(preview_root, bg="#2c2c2c")
            text_frame.pack(fill="both", expand=True, padx=8, pady=4)
            
            preview_label = tk.Label(text_frame, 
                                   text=text,
                                   bg="#2c2c2c", fg="#ffffff", font=("Consolas", 11),
                                   wraplength=width-20, justify="left", anchor="nw")
            preview_label.pack(fill="both", expand=True)
            
            # Zeichenanzahl (unauffällig)
            count_label = tk.Label(text_frame, 
                                 text=f"{text_length} Zeichen",
                                 bg="#2c2c2c", fg="#888888", font=("Segoe UI", 8))
            count_label.pack(anchor="se")
            
            # Automatisch schließen nach 4.5 Sekunden
            preview_root.after(4500, preview_root.destroy)
            
            # Schließen bei Klick
            preview_root.bind('<Button-1>', lambda e: preview_root.destroy())
            preview_root.bind('<FocusOut>', lambda e: preview_root.destroy())
            
            # Sanfte Animation (Fade-in)
            preview_root.attributes("-alpha", 0.0)
            preview_root.update()
            
            def fade_in():
                for i in range(10):
                    alpha = i * 0.1
                    preview_root.attributes("-alpha", alpha)
                    preview_root.update()
                    time.sleep(0.02)
            
            threading.Thread(target=fade_in, daemon=True).start()
            
            preview_root.mainloop()
        except Exception as e:
            logger.error(f"Fehler beim Erstellen des Vorschau-Fensters: {e}")
    
    threading.Thread(target=create_preview, daemon=True).start()

# Funktion zum Anzeigen der Warnung bei 50+ Zeichen
def show_warning_popup(text_length):
    def create_warning():
        try:
            warning_root = tk.Tk()
            warning_root.title("")
            warning_root.overrideredirect(True)
            warning_root.attributes("-topmost", True)
            warning_root.attributes("-alpha", 0.95)
            warning_root.configure(bg="#ff6b35")
            
            # Zentriere das Fenster
            warning_width, warning_height = 400, 80
            screen_width = warning_root.winfo_screenwidth()
            screen_height = warning_root.winfo_screenheight()
            x = (screen_width - warning_width) // 2
            y = (screen_height - warning_height) // 2
            warning_root.geometry(f"{warning_width}x{warning_height}+{x}+{y}")
            
            # Warnungstext
            warning_label = tk.Label(warning_root, 
                                   text=f"⚠ {text_length} Zeichen werden getippt!\nESC zum Abbrechen",
                                   bg="#ff6b35", fg="white", font=("Segoe UI", 12, "bold"),
                                   justify="center")
            warning_label.pack(expand=True)
            
            # ESC zum Abbrechen
            def on_escape(event):
                if event.keysym == 'Escape':
                    stop_typing_event.set()
                    warning_root.destroy()
                    show_popup("Tippvorgang abgebrochen!")
            
            warning_root.bind('<KeyPress>', on_escape)
            warning_root.focus_set()
            
            # Automatisch schließen nach 3 Sekunden
            warning_root.after(3000, warning_root.destroy)
            
            # Schließen bei Klick
            warning_root.bind('<Button-1>', lambda e: warning_root.destroy())
            
            # Sanfte Animation
            warning_root.attributes("-alpha", 0.0)
            warning_root.update()
            
            def fade_in():
                for i in range(15):
                    alpha = i * 0.067
                    warning_root.attributes("-alpha", alpha)
                    warning_root.update()
                    time.sleep(0.02)
            
            threading.Thread(target=fade_in, daemon=True).start()
            
            warning_root.mainloop()
        except Exception as e:
            logger.error(f"Fehler beim Erstellen der Warnung: {e}")
    
    threading.Thread(target=create_warning, daemon=True).start()

# Funktion zum Einrichten des Tray-Icons
def setup_tray():
    global tray_icon
    tray_icon = pystray.Icon("TypeTool")
    tray_icon.icon = create_image()
    tray_icon.title = f"TypeTool - {hotkey} zum Tippen, {toggle_enter_hotkey} für Enter"
    update_menu(tray_icon)
    tray_icon.run()

# Signal-Handler für sauberes Beenden
def signal_handler(signum, frame):
    logger.info(f"Signal {signum} empfangen, beende Programm...")
    cleanup_and_exit()

# Einzelinstanz-Überprüfung
def check_single_instance():
    try:
        import psutil
    except ImportError:
        logger.warning("psutil nicht verfügbar, Einzelinstanz-Schutz deaktiviert")
        return True
    
    lock_file = 'typetool.lock'
    
    try:
        # Prüfe ob Lock-Datei existiert
        if os.path.exists(lock_file):
            # Lese PID aus Lock-Datei
            with open(lock_file, 'r') as f:
                pid = int(f.read().strip())
            
            # Prüfe ob Prozess noch läuft
            try:
                process = psutil.Process(pid)
                if process.is_running() and 'TypeTool' in ' '.join(process.cmdline()):
                    logger.warning(f"TypeTool läuft bereits mit PID {pid}")
                    show_popup("TypeTool läuft bereits! Nur eine Instanz erlaubt.")
                    return False
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                # Prozess existiert nicht mehr, Lock-Datei ist veraltet
                os.remove(lock_file)
        
        # Erstelle Lock-Datei mit aktueller PID
        with open(lock_file, 'w') as f:
            f.write(str(os.getpid()))
        
        logger.info(f"Lock-Datei erstellt: {lock_file}")
        return True
        
    except Exception as e:
        logger.error(f"Fehler bei Einzelinstanz-Überprüfung: {e}")
        return True  # Bei Fehlern trotzdem starten

# Hauptprogramm
if __name__ == "__main__":
    try:
        # Einzelinstanz-Überprüfung
        if not check_single_instance():
            sys.exit(1)
        
        # Signal-Handler registrieren
        signal.signal(signal.SIGINT, signal_handler)
        signal.signal(signal.SIGTERM, signal_handler)
        
        # Cleanup bei Programmende registrieren
        atexit.register(cleanup_and_exit)
        
        logger.info("TypeTool wird gestartet...")
        logger.info(f"Konfiguration: Enter={press_enter}, Hotkey={hotkey}, Toggle={toggle_enter_hotkey}")
        
        # Hotkeys hinzufügen
        keyboard.add_hotkey('esc', lambda: stop_typing_event.set())
        keyboard.add_hotkey(hotkey, toggle_typing)
        keyboard.add_hotkey(toggle_enter_hotkey, toggle_enter)
        
        logger.info("Alle Hotkeys erfolgreich initialisiert")
        
        # Tray-Icon einrichten
        setup_tray()
        
    except KeyboardInterrupt:
        logger.info("Programm durch Benutzer beendet")
        cleanup_and_exit()
    except Exception as e:
        logger.error(f"Unerwarteter Fehler: {e}")
        show_popup(f"Unerwarteter Fehler: {e}")
        cleanup_and_exit()