import os
import subprocess
import time
import sys
import psutil
import requests

def is_process_running(process_name):
    process_list = [p.name() for p in psutil.process_iter(attrs=['pid', 'name'])]
    return process_name in process_list

while is_process_running('pdfScan.exe'):
    time.sleep(1)

repo_owner = 'alacielx'
repo_name = 'pdfScan'
repo_url = f'https://api.github.com/repos/{repo_owner}/{repo_name}/releases/latest'

exe_name = 'pdfScan.exe'
new_exe_name = 'pdfScan_new.exe'

response = requests.get(repo_url)
if response.status_code == 200:
    release_data = response.json()
    for asset in release_data.get("assets", []):
        if asset["name"] == exe_name:
            download_url = asset["browser_download_url"]

response = requests.get(download_url, stream=True)
with open(new_exe_name, 'wb') as new_exe:
    for chunk in response.iter_content(chunk_size=8192):
        new_exe.write(chunk)

# Rename from pdfScan_new.exe to pdfScan.exe
if os.path.exists(exe_name):
    os.remove(exe_name)
os.rename(new_exe_name, exe_name)

time.sleep(2)

# Launch updated pdfScan.exe
subprocess.Popen(['pdfScan.exe'])

sys.exit()
