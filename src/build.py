import os
import sys
import subprocess
import shutil

def build():
    import customtkinter
    ctk_path = os.path.dirname(customtkinter.__file__)
    
    print("Obfuscating source code with PyArmor...")
    # patch.py — одноразовый dev-скрипт правки gui.py, в поставку не идёт.
    py_files = [f for f in os.listdir()
                if f.endswith('.py') and f not in ('build.py', 'make_ico.py', 'patch.py')]
    
    # Run PyArmor to generate obfuscated code in obf_dist
    pyarmor_cmd = [sys.executable, "-m", "pyarmor.cli.core", "gen", "-O", "obf_dist"] + py_files
    # Wait, pyarmor module execution is python -m pyarmor. Actually, 'pyarmor' is installed as an executable.
    # Since we saw earlier 'pyarmor' package can't be executed directly, we should just run the pyarmor executable.
    pyarmor_exe = os.path.join(os.path.dirname(sys.executable), "Scripts", "pyarmor.exe")
    pyarmor_cmd = [pyarmor_exe, "gen", "-O", "obf_dist"] + py_files
    
    subprocess.run(pyarmor_cmd, check=True)
    
    # Find the pyarmor runtime package name (e.g. pyarmor_runtime_000000)
    obf_files = os.listdir("obf_dist")
    runtime_pkg = next((f for f in obf_files if f.startswith("pyarmor_runtime_")), None)
    
    print("Packaging obfuscated code with PyInstaller...")
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
        "--copy-metadata=razdel",
        "--distpath=..",
        "--name=Umbra",
        "--paths=obf_dist",
        f"--hidden-import={runtime_pkg}"
    ]

    # Explicitly add all internal modules since PyInstaller can't read their imports from obfuscated code
    internal_modules = [
        "gui", "nlp_engine", "doc_processor", "md_serializer", "md_merge",
        "file_detector", "clipboard_util", "pdf_reader",
    ]
    for mod in internal_modules:
        cmd.append(f"--hidden-import={mod}")

    # Explicitly add all external dependencies that would normally be discovered
    external_modules = [
        "customtkinter", "PIL", "psutil", "pythoncom", "win32com", "win32com.client",
        "win32clipboard", "docx", "razdel", "navec", "slovnet", "markdown_it",
        "pypdf", "tkinter.messagebox", "tkinter.filedialog"
    ]
    for mod in external_modules:
        cmd.append(f"--hidden-import={mod}")

    cmd.append(os.path.join("obf_dist", "main.py"))
    
    subprocess.run(cmd, check=True)
    
    print("Cleaning up obf_dist directory...")
    try:
        shutil.rmtree("obf_dist")
    except Exception as e:
        print(f"Warning: failed to remove obf_dist: {e}")

if __name__ == "__main__":
    build()
