import sys
import os
import time
import subprocess
import psutil
import signal

def find_and_kill_old_processes():
    """Findet und beendet alte TypeTool-Prozesse"""
    current_pid = os.getpid()
    killed_count = 0
    
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            # Prüfe ob es ein Python-Prozess mit TypeTool ist
            if proc.info['name'] and 'python' in proc.info['name'].lower():
                cmdline = proc.info['cmdline']
                if cmdline and any('TypeTool' in arg for arg in cmdline):
                    if proc.info['pid'] != current_pid:
                        print(f"Beende alten TypeTool-Prozess: PID {proc.info['pid']}")
                        proc.terminate()
                        killed_count += 1
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue
    
    if killed_count > 0:
        time.sleep(1)  # Warten bis Prozesse beendet sind
        print(f"{killed_count} alte Prozesse beendet")

if len(sys.argv) < 2:
    print("Usage: starter.py <script>")
    sys.exit(1)

script = sys.argv[1]
python = sys.executable

# Beende alte Prozesse
find_and_kill_old_processes()

# Starte neuen Prozess
try:
    subprocess.Popen([python, script])
    print("TypeTool gestartet")
except Exception as e:
    print(f"Fehler beim Starten: {e}")
    sys.exit(1)