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
        "--clean",
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
    if sys.platform == 'win32':
        cmd.extend(["--icon=icon_purple.ico", "--version-file=version_info.txt"])

    external_modules = [
        "customtkinter", "PIL", "docx", "razdel", "navec", "slovnet", "markdown_it",
        "pypdf", "pdfplumber", "tkinter.messagebox", "tkinter.filedialog"
    ]
    if sys.platform == 'win32':
        external_modules.extend([
            "win32clipboard", "windnd"
        ])
    for mod in external_modules:
        cmd.append(f"--hidden-import={mod}")

    cmd.append("main.py")
    
    env = os.environ.copy()
    env.setdefault("PYTHONHASHSEED", "0")
    if "SOURCE_DATE_EPOCH" not in env:
        try:
            repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            env["SOURCE_DATE_EPOCH"] = subprocess.check_output(
                ["git", "log", "-1", "--format=%ct"], cwd=repo_root,
                text=True, stderr=subprocess.DEVNULL,
            ).strip()
        except (OSError, subprocess.SubprocessError):
            env["SOURCE_DATE_EPOCH"] = "0"
    subprocess.run(cmd, check=True, env=env)
    print("Build complete.")

if __name__ == "__main__":
    build()
