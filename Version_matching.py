import sys
import os
import linecache 
import re
from zipfile import ZipFile
import subprocess
import ctypes  # An included library with Python install.   


if sys.argv[1].endswith("twb"):
    versionline=linecache.getline(sys.argv[1],3)
    m=re.match("[^a-zA-Z]*build (\d+)\..*",versionline)
    version=m.group(1)
    version=version[:4]+"."+version[-1]

elif sys.argv[1].endswith("twbx"):
    os.mkdir("tmp")
    with ZipFile(sys.argv[1],"r") as twbx:
        twbx.extract(sys.argv[1][:-1],".\\tmp")
    versionline=linecache.getline(f"tmp\\{sys.argv[1][:-1]}",3)
    m=re.match("[^a-zA-Z]*build (\d+)\..*",versionline)
    version=m.group(1)
    version=version[:4]+"."+version[-1]
    print(version)
    os.remove(f"tmp\\{sys.argv[1][:-1]}")
    os.rmdir("tmp")
else:
    raise NotImplementedError
if "Tableau" not in os.listdir("C:\\Program Files"):
    print("Tableau is not installed")
    exit()
versions=list()
for file in os.listdir("C:\\Program Files\\Tableau"):
    if "Tableau" in file and not "Prep" in file:
        m=re.match("Tableau (\d+\.\d)",file)
        versions.append(m.group(1))
if version in versions:
    print("Version match found")
    DETACHED_PROCESS=0x00000008
    cwd=os.getcwd()
    print(f'"{cwd}\{sys.argv[1]}"')
    os.chdir(f"C:\\Program Files\\Tableau\\Tableau {version}\\bin")
    pid=subprocess.Popen(f'tableau.exe "{cwd}\{sys.argv[1]}"',creationflags=DETACHED_PROCESS)
else:
    print("Version not found, please download appropriate version from tableau.com/esdalt")
    ctypes.windll.user32.MessageBoxW(0, "Version not found", f"Tableau version not found: {version}. Download from tableau.com/esdalt", 1)