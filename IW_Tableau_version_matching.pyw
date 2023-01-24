import sys
import os
import linecache 
import re
from zipfile import ZipFile
import subprocess
import ctypes
from bs4 import BeautifulSoup
from urllib.request import Request, urlopen
import webbrowser

def extract_version(versionline):
    """Extract the version number from the third line of the file"""
    
    m=re.match("[^a-zA-Z]*build (\d+)\..*",versionline)
    version=m.group(1)
    version=version[:4]+"."+version[-1]
    return version

def extract_twbx(file_path):
    """Extract the twbx file and return the third line"""
    
    
    m=re.match(r".*\\([^\\]+.twb)x$",file_path)
    file_name=m.group(1)
    
    with ZipFile(file_path,"r") as twbx:
        with twbx.open(file_name) as file:
            file.__next__()
            file.__next__()
            versionline=file.__next__().decode("UTF-8")
    return versionline

def check_tableau_install():
    """Check if Tableau is installed on the computer"""
    if "Tableau" not in os.listdir("C:\\Program Files"):
        ctypes.windll.user32.MessageBoxW(0,"Tableau is not installed","Tableau is not installed",1)
        exit()
    return True

def check_version_match(version, file_path):
    """Check if the extracted version number is present in the Tableau folder"""
    versions = get_installed_versions()
    if version in versions:
        open_file_with_tableau(version, file_path)
    else:
        nearest_version = get_nearest_version(version, versions)
        if nearest_version:
            choice = prompt_user(version, nearest_version)
            if choice == 'open':
                open_file_with_tableau(nearest_version, file_path)
            elif choice == 'download':
                download_version(version)
    return None

def download_version(version):
    req = Request("https://www.tableau.com/support/releases")
    html_page = urlopen(req)
    soup = BeautifulSoup(html_page, "lxml")
    version_links = []
    for link in soup.findAll('a'):
        curr_link=str(link.get('href'))
        if curr_link.find(version)!=-1:
            version_links.append(curr_link)
    version_dict={}
    available_sub_versions=list()
    for link in version_links:
        m=re.match(".*/([\d\.]+)$",link)
        sub_version=m.group(1)
        if version==sub_version:
            sub_version+=".0"
        available_sub_versions.append(int(sub_version[7:]))
        version_dict[sub_version]=link
    if len(version_links)==0:
        ctypes.windll.user32.MessageBoxW(0, f"This version is not supported by Tableau, please find a compatible version manually.", "Version too old", 1)
    webbrowser.open(version_dict[version+"."+str(max(available_sub_versions))])
    
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
    result = messageBox(0, f"Tableau version {version} not found, the earliest version found is {nearest_version}. Do you want to open the file with this version? Pressing 'No' will open the download page for the appropriate version. Press 'Cancel' to abort.", "Version not found", 3)
    if result == 6:
        return 'open'
    elif result == 7:
        return 'download'
    else:
        return None

def open_file_with_tableau(version, file_path):
    """Open the file with the matching version of Tableau"""
    DETACHED_PROCESS=0x00000008
    os.chdir(f"C:\\Program Files\\Tableau\\Tableau {version}\\bin")
    pid=subprocess.Popen(f'tableau.exe "{file_path}"',creationflags=DETACHED_PROCESS)

if __name__ == "__main__":
    if len(sys.argv)==2:
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
    else:
        ctypes.windll.user32.MessageBoxW(0, f"Please follow the instructions on github/slack to associate this software with .twb and .twbx files.", "Please look at the instructions", 1)