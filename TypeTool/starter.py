import sys
import os
import time
import subprocess

if len(sys.argv) < 2:
    print("Usage: starter.py <script>")
    sys.exit(1)

script = sys.argv[1]
python = sys.executable

time.sleep(0.7)  # Warten, bis der alte Prozess wirklich beendet ist
subprocess.Popen([python, script])