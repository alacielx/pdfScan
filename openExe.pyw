import os
import subprocess

absPath = os.path.abspath('pdfScan.exe')

try:
    process = subprocess.Popen(['pdfScan.exe'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = process.communicate()
    print("Standard Output:")
    print(stdout.decode())
    print("Standard Error:")
    print(stderr.decode())
except Exception as e:
    print(f"Error: {e}")
