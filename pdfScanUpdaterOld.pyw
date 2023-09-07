import os
import subprocess
import time
import sys
import psutil
import requests

sys.path.append(r'C:\Users\agarza\OneDrive - Arrow Glass Industries\Documents\Scripts\Test\functions')
from functions import *

def is_process_running(process_name):
    process_list = [p.name() for p in psutil.process_iter(attrs=['pid', 'name'])]
    return process_name in process_list

repoName = 'pdfScan'
executableName = repoName + '.exe'

while is_process_running(executableName):
    time.sleep(1)

repoOwner = 'alacielx'
repoUrl = f'https://api.github.com/repos/{repoOwner}/{repoName}/releases/latest'

exeName = 'pdfScan.exe'
newExeName = 'pdfScan_temp.exe'

response = requests.get(repoUrl)
if response.status_code == 200:
    release_data = response.json()
    for asset in release_data.get("assets", []):
        if asset["name"] == exeName:
            download_url = asset["browser_download_url"]

response = requests.get(download_url, stream=True)
with open(newExeName, 'wb') as newExe:
    for chunk in response.iter_content(chunk_size=8192):
        newExe.write(chunk)

# Rename from pdfScan_new.exe to pdfScan.exe
if os.path.exists(exeName):
    os.remove(exeName)
os.rename(newExeName, exeName)

time.sleep(2)

# Launch updated pdfScan.exe

# pdfScan = os.path.abspath('pdfScan.exe')
try:
    process = subprocess.Popen(['pdfScan.exe'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    time.sleep(4)
    print("pdfScan.exe started successfully.")
except Exception as e:
    print(f"Error: {e}")

sys.exit()
