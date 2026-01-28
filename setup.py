import sys
import os
from cx_Freeze import setup, Executable

# Adicione aqui qualquer arquivo ou pasta extra que o robô precise
files = [
    "anexos/",      # Garante que a pasta de anexos vá junto
    "contatos.xlsx" # Se você quiser que uma planilha modelo já vá instalada
]

build_exe_options = {
    "packages": [
        "customtkinter", 
        "pyautogui", 
        "pyperclip", 
        "pandas", 
        "selenium", 
        "webdriver_manager", 
        "smtplib", 
        "email",     # ESSENCIAL PARA ANEXOS
        "openpyxl"   # ESSENCIAL PARA SALVAR O STATUS NA PLANILHA
    ],
    "include_files": files,
    "excludes": ["tkinter.test", "unittest"]
}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name = "SistemaEnvioATI_V2",
    version = "2.0",
    description = "Automação Inteligente com Anexo e Filtro - ATI PE",
    options = {
        "build_exe": build_exe_options,
        "bdist_msi": {
            "upgrade_code": "{A1B2C3D4-E5F6-G7H8-I9J0-K1L2M3N4O5P6}",
            "initial_target_dir": r"%LocalAppData%\SistemaEnvioATI",
        }
    },
    executables = [Executable("robo_ati.py", base=base)]
)