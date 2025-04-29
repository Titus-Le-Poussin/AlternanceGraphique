from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import os
import time
import subprocess

class FileChangeHandler(FileSystemEventHandler):
    def __init__(self, script_path, watched_file):
        self.script_path = script_path
        self.watched_file = os.path.abspath(watched_file)  # Chemin absolu du fichier surveillé
        self.process = None
        self.restart_script()

    def restart_script(self):
        if self.process:
            self.process.terminate()
        print(f"Lancement du script {self.script_path}...")
        self.process = subprocess.Popen(["python", self.script_path])

    def on_modified(self, event):
        if os.path.abspath(event.src_path) == self.watched_file:
            print(f"{self.watched_file} modifié, redémarrage du script...")
            self.restart_script()

if __name__ == "__main__":
    script_path = "Alernance.py"  # Chemin vers ton script Python
    watched_file = "Alternances.xlsx"  # Fichier à surveiller

    # Obtenir le répertoire du fichier surveillé
    watched_dir = os.path.dirname(os.path.abspath(watched_file))

    event_handler = FileChangeHandler(script_path, watched_file)
    observer = Observer()
    observer.schedule(event_handler, path=watched_dir, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()