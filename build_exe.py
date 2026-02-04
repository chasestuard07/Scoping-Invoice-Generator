import os
import sys
import subprocess
import shutil

# =========================
# CONFIG - EDIT IF NEEDED
# =========================
SCRIPT_NAME = "InvoiceGenerator.py"  # your main script
EXE_NAME = "InvoiceGenerator.exe"     # desired exe name
HEADER_JSON = "header.json"           # file to include
ICON_PATH = ""                        # optional: path to icon, e.g. "icon.ico"

# =========================
# CLEAN OLD BUILD FILES (optional)
# =========================
for folder in ["build", "dist"]:
    if os.path.exists(folder):
        print(f"Removing old folder: {folder}")
        shutil.rmtree(folder)

if os.path.exists(f"{os.path.splitext(SCRIPT_NAME)[0]}.spec"):
    os.remove(f"{os.path.splitext(SCRIPT_NAME)[0]}.spec")

# =========================
# BUILD COMMAND
# =========================
cmd = [
    sys.executable, "-m", "PyInstaller",
    "--onefile",             # single executable
    f"--name={os.path.splitext(EXE_NAME)[0]}", 
    SCRIPT_NAME
]

if ICON_PATH:
    cmd.append(f"--icon={ICON_PATH}")

# =========================
# RUN THE BUILD
# =========================
print("Running PyInstaller...")
print("Command:", " ".join(cmd))
try:
    subprocess.run(cmd, check=True)
    print("\nBuild finished!")
    print(f"Your exe is here: dist/{EXE_NAME}")
    input("Press Enter to exit...")
except subprocess.CalledProcessError as e:
    print("Build failed:", e)
    input("Press Enter to exit...")
