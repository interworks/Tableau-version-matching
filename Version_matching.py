import sys
import os
import linecache 
import re
from zipfile import ZipFile
import subprocess
import ctypes


def extract_version(versionline):
    """Extract the version number from the third line of the file"""
    
    m=re.match("[^a-zA-Z]*build (\d+)\..*",versionline)
    version=m.group(1)
    version=version[:4]+"."+version[-1]
    return version

def extract_twbx(file_path):
    """Extract the twbx file and return the third line"""
    os.mkdir("tmp")
    try:
        print(file_path)
        m=re.match(r".*\\([^\\]+.twb)x$",file_path)
        file_name=m.group(1)
        with ZipFile(file_path,"r") as twbx:
            twbx.extract(file_name,".\\tmp")
        
        versionline=linecache.getline(f"tmp\\{file_name}",3)
        os.remove(f"tmp\\{file_name}")
    except:
        os.rmdir("tmp")
        raise FileNotFoundError
        
    os.rmdir("tmp")
    return versionline

def check_tableau_install():
    """Check if Tableau is installed on the computer"""
    if "Tableau" not in os.listdir("C:\\Program Files"):
        print("Tableau is not installed")
        exit()
    return True

def check_version_match(version, file_path):
    """Check if the extracted version number is present in the Tableau folder"""
    versions = get_installed_versions()
    if version in versions:
        print("Version match found")
        open_file_with_tableau(version, file_path)
    else:
        print("Version not found")
        nearest_version = get_nearest_version(version, versions)
        if nearest_version:
            choice = prompt_user(version, nearest_version)
            if choice == 'open':
                open_file_with_tableau(nearest_version, file_path)
            elif choice == 'download':
                #download_version()
                pass
        else:
            #download_version()
            pass


def get_installed_versions():
    """Get the list of installed versions of Tableau"""
    versions = []
    for file in os.listdir("C:\\Program Files\\Tableau"):
        if "Tableau" in file and not "Prep" in file:
            m = re.match("Tableau (\d+\.\d)", file)
            versions.append(m.group(1))
    return versions

def get_nearest_version(version, versions):
    """Get the nearest version greater than the extracted version"""
    versions.sort()
    for ver in versions:
        if version < ver:
            return ver
    return None

def prompt_user(version, nearest_version):
    """Prompt the user to choose between opening the file with the nearest version or downloading the appropriate version"""
    messageBox = ctypes.windll.user32.MessageBoxW
    result = messageBox(0, f"Tableau version {version} not found, the earliest version found is {nearest_version}. Do you want to open the file with this version?", "Version not found", 3)
    if result == 6:
        return 'open'
    else:
        messageBox(0, "Please download the appropriate version from tableau.com/esdalt", "Download the appropriate version", 1)
        return 'download'

def open_file_with_tableau(version, file_path):
    """Open the file with the matching version of Tableau"""
    DETACHED_PROCESS=0x00000008
    os.chdir(f"C:\\Program Files\\Tableau\\Tableau {version}\\bin")
    pid=subprocess.Popen(f'tableau.exe "{file_path}"',creationflags=DETACHED_PROCESS)

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
        check_version_match(version,file_path)