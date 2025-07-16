
import sys
import os
import ctypes
import pyperclip
import keyboard
import time
import pystray
from PIL import Image, ImageDraw, ImageFont
import threading
import tkinter as tk
import json
import subprocess

# Pfad zur Konfigurationsdatei
config_file = 'config.json'

# Standardkonfiguration
default_config = {
    'enter_key_enabled': False,  # Standardmäßig deaktiviert
    'hotkey': 'ctrl+b',
    'toggle_enter_hotkey': 'ctrl+alt+b',
    'show_preview_window': True  # NEU: Vorschau-Fenster standardmäßig an
}

# Funktion zum Laden der Konfiguration
def load_config():
    if os.path.exists(config_file):
        with open(config_file, 'r') as file:
            return json.load(file)
    else:
        return default_config

# Funktion zum Speichern der Konfiguration
def save_config(config):
    with open(config_file, 'w') as file:
        json.dump(config, file)

# Konfiguration laden
config = load_config()
press_enter = config.get('enter_key_enabled', False)
hotkey = config.get('hotkey', 'ctrl+b')
toggle_enter_hotkey = config.get('toggle_enter_hotkey', 'ctrl+alt+b')
show_preview_window = config.get('show_preview_window', True)
tray_icon = None

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

# Funktion zum Anzeigen des Tipp-Text-Fensters
def show_typing_window(text):
    window = tk.Tk()
    window.title("TypeTool Vorschau")
    window.attributes("-topmost", True)
    window.resizable(False, False)
    window.overrideredirect(True)  # Kein Rahmen

    # Dynamische Breite und Höhe berechnen
    max_width = 500
    padding = 16
    font = ("Consolas", 11)
    label = tk.Label(window, text=text, font=font, bg="#ffffe0", fg="black", anchor="w", justify="left")
    label.pack(fill="both", expand=True, padx=8, pady=8)

    # Textgröße messen
    window.update_idletasks()
    req_width = min(label.winfo_reqwidth() + padding, max_width)
    req_height = label.winfo_reqheight() + padding

    # Fenstergröße setzen
    window.geometry(f"{req_width}x{req_height}+20+20")
    window.update()
    window.after(3000, window.destroy)  # Fenster nach 3 Sekunden schließen
    window.mainloop()  # <-- Hält das Fenster offen, bis es zerstört wird
    return window

# Funktion zum Schreiben von Text (angepasst)
def type_text(text):
    stop_typing_event.clear()
    typing_window_thread = None
    if show_preview_window:
        # Vorschau-Fenster in eigenem Thread starten
        def show_window():
            show_typing_window(text)
        typing_window_thread = threading.Thread(target=show_window, daemon=True)
        typing_window_thread.start()
    # Fokus vor dem Tippen merken
    user32 = ctypes.windll.user32
    prev_hwnd = user32.GetForegroundWindow()
    for char in text:
        if stop_typing_event.is_set():
            break
        keyboard.write(char)
        time.sleep(0.05)
    if press_enter and not stop_typing_event.is_set():
        keyboard.press_and_release('enter')
    # Nach dem Tippen Fokus zurückgeben
    if prev_hwnd:
        user32.SetForegroundWindow(prev_hwnd)

# Funktion zum Starten des Tippvorgangs in einem Thread
typing_thread = None

def toggle_typing():
    global typing_thread
    print("toggle_typing aufgerufen")  # Debug
    if typing_thread and typing_thread.is_alive():
        stop_typing_event.set()
    else:
        stop_typing_event.clear()
        update_clipboard_history()  # <--- HIER
        text = pyperclip.paste()
        if not text:
            print("Zwischenablage ist leer.")
            return
        # Hinweis bei mehr als 50 Zeichen
        if len(text) > 50:
            popup_thread = show_popup("Hinweis: Mehr als 50 Zeichen kopiert!\n(Tippvorgang kann mit ESC abgebrochen werden)")
            popup_thread.join()
            if stop_typing_event.is_set():
                return
        typing_thread = threading.Thread(target=type_text, args=(text,))
        typing_thread.start()

# Hotkey für ESC zum Abbrechen des Tippvorgangs
keyboard.add_hotkey('esc', lambda: stop_typing_event.set())

# Hotkey für das Tippen (Start/Stop)
keyboard.add_hotkey(hotkey, toggle_typing)

clipboard_history = []

def update_clipboard_history():
    text = pyperclip.paste()
    if not clipboard_history or clipboard_history[-1] != text:
        clipboard_history.append(text)
        # Maximal 10 Einträge merken
        if len(clipboard_history) > 10:
            clipboard_history.pop(0)

def get_second_clipboard_entry():
    # Holt die Clipboard-History per PowerShell (nur Windows 10+)
    try:
        # PowerShell-Befehl, um die letzten 2 Text-Einträge zu bekommen
        ps_command = r'''
        Add-Type -AssemblyName PresentationCore
        $history = Get-Clipboard -Format Text -Raw -TextFormatType UnicodeText -History
        if ($history.Count -ge 2) { $history[1] } else { "" }
        '''
        result = subprocess.run(
            ["powershell", "-Command", ps_command],
            capture_output=True, text=True, timeout=2
        )
        return result.stdout.strip()
    except Exception as e:
        print("Fehler beim Lesen der Clipboard-History:", e)
        return ""

def type_second_clipboard_entry():
    if len(clipboard_history) < 2:
        print("Kein zweiter Eintrag in der lokalen Zwischenablage-History gefunden.")
        return
    text = clipboard_history[-2]
    if not text:
        print("Kein Text im zweitneuesten Eintrag.")
        return
    # Hinweis bei mehr als 50 Zeichen
    if len(text) > 50:
        popup_thread = show_popup("Hinweis: Mehr als 50 Zeichen kopiert!\n(Tippvorgang kann mit ESC abgebrochen werden)")
        popup_thread.join()
        if stop_typing_event.is_set():
            return
    global typing_thread
    if typing_thread and typing_thread.is_alive():
        stop_typing_event.set()
    else:
        stop_typing_event.clear()
        typing_thread = threading.Thread(target=type_text, args=(text,))
        typing_thread.start()

# Hotkey für Strg + Leertaste (ctrl+space)
keyboard.add_hotkey('ctrl+space', type_second_clipboard_entry)

# Funktion zur Überwachung der Zwischenablage (angepasst)
def monitor_clipboard():
    print(f"Das Programm überwacht die Zwischenablage. Drücke {hotkey}, um den Text aus der Zwischenablage zu schreiben. Drücke STRG+C, um das Programm zu beenden.")
    while True:
        try:
            keyboard.wait(hotkey)
            time.sleep(0.1)
            toggle_typing()
            time.sleep(0.1)
        except pyperclip.PyperclipException as e:
            print("Fehler beim Zugriff auf die Zwischenablage:", e)
            break
        except Exception as e:
            print(f"Fehler: {e}")
            break
        except KeyboardInterrupt:
            print("Programm wird beendet.")
            break

# Funktion zum Beenden des Programms
def on_quit(icon, item):
    icon.stop()

# Funktion zum Umschalten der Enter-Taste
def toggle_enter(icon=None, item=None):
    global press_enter
    press_enter = not press_enter
    config['enter_key_enabled'] = press_enter
    save_config(config)
    if tray_icon:
        update_menu(tray_icon)
    show_popup(f"Enter nach Text: {'An' if press_enter else 'Aus'}")

# Funktion zum Umschalten des Vorschau-Fensters
def toggle_preview_window(icon=None, item=None):
    global show_preview_window
    show_preview_window = not show_preview_window
    config['show_preview_window'] = show_preview_window
    save_config(config)
    if tray_icon:
        update_menu(tray_icon)
    show_popup(f"Vorschau-Fenster: {'An' if show_preview_window else 'Aus'}")

# Funktion zum Ändern der Hotkeys und Neustarten des Programms
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
                keyboard.add_hotkey(hotkey, lambda: None)
            if new_toggle_enter_hotkey_value:
                keyboard.remove_hotkey(toggle_enter_hotkey)
                toggle_enter_hotkey = new_toggle_enter_hotkey_value
                config['toggle_enter_hotkey'] = toggle_enter_hotkey
                keyboard.add_hotkey(toggle_enter_hotkey, toggle_enter)
            save_config(config)
            root.destroy()
            threading.Thread(target=restart_program, daemon=True).start()

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
        pystray.MenuItem("Vorschau-Fenster: " + ("An" if show_preview_window else "Aus"), toggle_preview_window),
        pystray.MenuItem("Enter nach Text: " + ("An" if press_enter else "Aus"), toggle_enter),
        pystray.MenuItem("Hotkeys ändern", change_hotkey),
        pystray.MenuItem("Neustarten", on_restart),
        pystray.MenuItem("Beenden", on_quit)
    )

# Funktion zum Anzeigen eines Popups
def show_popup(message):
    def popup():
        root = tk.Tk()
        root.overrideredirect(1)
        root.attributes("-topmost", True)
        width, height = 350, 80
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f"{width}x{height}+{x}+{y}")
        label = tk.Label(root, text=message, bg="yellow", fg="black", font=("Helvetica", 12))
        label.pack(expand=True, fill="both")

        def on_esc(event=None):
            stop_typing_event.set()
            root.destroy()
        root.bind('<Escape>', on_esc)

        # Macht das Fenster modal für Tastatureingaben
        root.lift()
        root.focus_force()
        root.grab_set()
        # Fokus mehrfach setzen, um Windows-Besonderheiten zu umgehen
        root.after(50, lambda: root.focus_force())
        root.after(150, lambda: root.focus_force())

        root.after(3000, root.destroy)
        root.mainloop()
    t = threading.Thread(target=popup, daemon=True)
    t.start()
    return t

# Funktion zum Einrichten des Tray-Icons
def setup_tray():
    global tray_icon
    tray_icon = pystray.Icon("TypeTool")
    tray_icon.icon = create_image()
    tray_icon.title = "TypeTool"
    update_menu(tray_icon)
    tray_icon.run()

# Hauptprogramm
if __name__ == "__main__":
    keyboard.add_hotkey(toggle_enter_hotkey, toggle_enter)
    setup_tray()
