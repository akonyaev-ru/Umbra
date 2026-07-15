import os
import sys
import subprocess

def build():
    import customtkinter
    ctk_path = os.path.dirname(customtkinter.__file__)
    
    print("Packaging source code with PyInstaller...")
    sep = os.pathsep
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconsole",
        "--onefile",
        "--icon=icon_purple.ico",
        "--version-file=version_info.txt",
        f"--add-data=icon_purple.ico{sep}.",
        f"--add-data=logo.png{sep}.",
        f"--add-data=Audiowide-Regular.ttf{sep}.",
        f"--add-data={ctk_path}{sep}customtkinter/",
        f"--add-data=models{sep}models/",
        "--copy-metadata=natasha",
        "--copy-metadata=slovnet",
        "--copy-metadata=navec",
        "-y",
        "--copy-metadata=razdel",
        "--distpath=..",
        "--name=Umbra"
    ]

    external_modules = [
        "customtkinter", "PIL", "psutil", "docx", "razdel", "navec", "slovnet", "markdown_it",
        "pypdf", "tkinter.messagebox", "tkinter.filedialog"
    ]
    if sys.platform == 'win32':
        external_modules.extend([
            "pythoncom", "win32com", "win32com.client", "win32clipboard"
        ])
    for mod in external_modules:
        cmd.append(f"--hidden-import={mod}")

    cmd.append("main.py")
    
    subprocess.run(cmd, check=True)
    print("Build complete.")

if __name__ == "__main__":
    build()
