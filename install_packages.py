import os
import importlib
import urllib.request
import zipfile
import msvcrt
import subprocess
import pkg_resources
import hashlib

def install_packages(package_list):
    for package in package_list:
        try:
            importlib.import_module(package)
            print(f"{package} is already installed.")
            try:
                # Get the installed version of the package
                installed_version = pkg_resources.get_distribution(package).version
                
                # Install the latest version if available
                subprocess.check_call(['python', '-m', 'pip', 'install', '--upgrade', package])
                
                # Check if the package was actually updated
                updated_version = pkg_resources.get_distribution(package).version
                if updated_version != installed_version:
                    print(f"{package} has been successfully updated to version {updated_version}.")
                else:
                    print(f"{package} is already up to date.")
            except Exception as e:
                print(f"An error occurred during the installation or update of {package}: {str(e)}")
        except ImportError:
            print(f"{package} is not installed. Installing...")
            try:
                subprocess.check_call(['python', '-m', 'pip', 'install', package])
                print(f"{package} has been successfully installed.")
            except Exception as e:
                print(f"An error occurred during the installation of {package}: {str(e)}")


def download_zip(zips):
    for key, value in zips.items():
        output_dir = fr"C:\{key}"
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir, exist_ok=True)
                output_file = os.path.join(output_dir, fr"{key}.zip")
                urllib.request.urlretrieve(value, output_file)

                # Extract the downloaded ZIP file
                with zipfile.ZipFile(output_file, 'r') as zip_ref:
                    zip_ref.extractall(output_dir)
                
                os.remove(output_file)

                print(fr"{key} has been downloaded and extracted successfully.")
            except Exception as e:
                print(f"An error occurred during the download and extraction of {key}: {str(e)}")
        else:
            print(fr"{key} already installed")

def calculate_sha256(file_path):
    sha256_hash = hashlib.sha256()
    with open(file_path, "rb") as file:
        # Read and update hash string value in blocks of 4K
        for byte_block in iter(lambda: file.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest()

def update_pdfScan(current_folder):
    new_file_url = "https://github.com/alacielx/pdfScan/raw/main/pdfScan.pyw"
    new_file_path = os.path.join(current_folder, "pdfScan.pyw")
    
    try:
        # Check if the file already exists
        if os.path.exists(new_file_path):
            # Calculate the SHA256 checksum of the existing file
            existing_file_hash = calculate_sha256(new_file_path)
            
            # Download the new file
            with urllib.request.urlopen(new_file_url) as response:
                remote_content = response.read()
            
            # Calculate the SHA256 checksum of the remote file content
            remote_content_hash = hashlib.sha256(remote_content).hexdigest()
            
            # Compare the checksums
            if existing_file_hash == remote_content_hash:
                print("The file is already up to date. No need to update.")
                return

        # Download the new file
        urllib.request.urlretrieve(new_file_url, new_file_path)

        print("Successfully updated pdfScan script.")
    except Exception as e:
        print("An error occurred during the download of pdfScan:", str(e))

def main():
    # Packages to install and zips to download
    required_packages = ['pip', 'PyPDF2', 'pdf2image', 'Pillow','opencv_python', 'pytesseract', 'PyMuPDF']
    zip_downloads = {
        'poppler': "https://github.com/oschwartz10612/poppler-windows/releases/download/v23.07.0-0/Release-23.07.0-0.zip",
        'tesseract': "https://github.com/alacielx/Tesseract-Portable/raw/main/Tesseract.zip"
    }
    current_folder = os.getcwd()

    install_packages(required_packages)
    print("\n")

    download_zip(zip_downloads)
    print("\n")

    update_pdfScan(current_folder)
    print("\nPress any key to exit")

    while True:
        if msvcrt.kbhit():
            key = msvcrt.getch()
            break


    import os
    import urllib.request

if __name__ == "__main__":
    main()