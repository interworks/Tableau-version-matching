import sys
import os
import linecache 
import re
from zipfile import ZipFile
import subprocess
import ctypes  # An included library with Python install.   


def extract_version(versionline):
    """Extract the version number from the third line of the file"""
    
    m=re.match("[^a-zA-Z]*build (\d+)\..*",versionline)
    version=m.group(1)
    version=version[:4]+"."+version[-1]
    return version

def extract_twbx(file_path):
    """Extract the twbx file and return the third line"""
    os.mkdir("tmp")
    with ZipFile(file_path,"r") as twbx:
        twbx.extract(file_path[:-1],".\\tmp")
    versionline=linecache.getline(f"tmp\\{file_path[:-1]}",3)
    os.remove(f"tmp\\{file_path[:-1]}")
    os.rmdir("tmp")
    return versionline

def check_tableau_install():
    """Check if Tableau is installed on the computer"""
    if "Tableau" not in os.listdir("C:\\Program Files"):
        print("Tableau is not installed")
        exit()
    return True

def check_version_match(version):
    """Check if the extracted version number is present in the Tableau folder"""
    versions=list()
    for file in os.listdir("C:\\Program Files\\Tableau"):
        if "Tableau" in file and not "Prep" in file:
            m=re.match("Tableau (\d+\.\d)",file)
            versions.append(m.group(1))
    if version in versions:
        print("Version match found")
        return True
    else:
        print("Version not found, please download appropriate version from tableau.com/esdalt")
        ctypes.windll.user32.MessageBoxW(0, "Version not found", f"Tableau version not found: {version}. Download from tableau.com/esdalt", 1)
        return False

def open_file_with_tableau(version, file_path):
    """Open the file with the matching version of Tableau"""
    DETACHED_PROCESS=0x00000008
    cwd=os.getcwd()
    os.chdir(f"C:\\Program Files\\Tableau\\Tableau {version}\\bin")
    pid=subprocess.Popen(f'tableau.exe "{cwd}\{file_path}"',creationflags=DETACHED_PROCESS)

if __name__ == "__main__":
    file_path = sys.argv[1]
    if file_path.endswith("twb"):
        versionline=linecache.getline(file_path,3)
        version = extract_version(versionline)

    elif file_path.endswith("twbx"):
        versionline = extract_twbx(file_path)
        version = extract_version(versionline)
    else:
        raise NotImplementedError

    if check_tableau_install():
        if check_version_match(version):
            open_file_with_tableau(version, file_path)