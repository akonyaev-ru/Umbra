import os
import sys
import subprocess

def build():
    import customtkinter
    ctk_path = os.path.dirname(customtkinter.__file__)
    
    print("Packaging source code with PyInstaller...")
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconsole",
        "--onefile",
        "--icon=icon_purple.ico",
        "--version-file=version_info.txt",
        "--add-data=icon_purple.ico;.",
        "--add-data=logo.png;.",
        "--add-data=Audiowide-Regular.ttf;.",
        f"--add-data={ctk_path};customtkinter/",
        "--add-data=models;models/",
        "--copy-metadata=natasha",
        "--copy-metadata=slovnet",
        "--copy-metadata=navec",
        "-y",
        "--copy-metadata=razdel",
        "--distpath=..",
        "--name=Umbra"
    ]

    # Explicitly add all external dependencies that would normally be discovered
    external_modules = [
        "customtkinter", "PIL", "psutil", "pythoncom", "win32com", "win32com.client",
        "win32clipboard", "docx", "razdel", "navec", "slovnet", "markdown_it",
        "pypdf", "tkinter.messagebox", "tkinter.filedialog"
    ]
    for mod in external_modules:
        cmd.append(f"--hidden-import={mod}")

    cmd.append("main.py")
    
    subprocess.run(cmd, check=True)
    print("Build complete.")

if __name__ == "__main__":
    build()
