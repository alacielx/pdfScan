import tkinter as tk
from tkinter import ttk, messagebox
import time
import threading
import requests
import os
import psutil
import sys
import subprocess

class DownloadApp:
    def __init__(self, root, download_url, file_name):
        self.root = root
        self.root.title("Updating pdfScan")

        # Get the screen width and height
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Calculate the x and y coordinates to center the window
        x = (screen_width - 300) // 2
        y = (screen_height - 70) // 2

        self.root.geometry(f"300x70+{x}+{y}")

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(padx=20, pady=20, fill='both', expand=True)  # Center the progress bar

        self.start_download(download_url, file_name)

    def start_download(self, download_url, file_name):
        self.progress_var.set(0)

        download_thread = threading.Thread(target=self.download_file, args=(download_url, file_name))
        download_thread.daemon = True  # Set the thread as a daemon thread
        download_thread.start()

    def download_file(self, download_url, file_name):
        response = requests.get(download_url, stream=True)
        total_size = int(response.headers.get('content-length', 0))

        with open(file_name, 'wb') as new_exe:
            downloaded_size = 0
            for chunk in response.iter_content(chunk_size=8192):
                new_exe.write(chunk)
                downloaded_size += len(chunk)
                progress = (downloaded_size / total_size) * 100
                self.progress_var.set(progress)
                self.root.update_idletasks()
        
        time.sleep(1)  # Give some time to see the progress
        self.root.withdraw()  # Hide the DownloadApp window
        self.show_completion_message()
        
    def show_completion_message(self):
        messagebox.showinfo("pdfScan Updater", "Updated pdfScan.exe\n\nPlease close this window and run again")
        self.root.quit()  # Close the application when the messagebox is closed


def download_file_with_progress(download_url, file_name):
    root = tk.Tk()
    app = DownloadApp(root, download_url, file_name)
    root.mainloop()

def is_process_running(process_name):
    process_list = [p.name() for p in psutil.process_iter(attrs=['pid', 'name'])]
    return process_name in process_list

while is_process_running('pdfScan.exe'):
    time.sleep(1)

repo_owner = 'alacielx'
repo_name = 'pdfScan'
repo_url = f'https://api.github.com/repos/{repo_owner}/{repo_name}/releases/latest'

exeName = 'pdfScan.exe'
newExeName = 'pdfScan_new.exe'

response = requests.get(repo_url)
if response.status_code == 200:
    release_data = response.json()
    for asset in release_data.get("assets", []):
        if asset["name"] == exeName:
            download_url = asset["browser_download_url"]

download_file_with_progress(download_url, newExeName)

# Rename from pdfScan_new.exe to pdfScan.exe
try:
    if os.path.exists(exeName):
        os.remove(exeName)
    os.rename(newExeName, exeName)
except:
    print("Could not rename file")
    
sys.exit()