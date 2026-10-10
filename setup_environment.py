"""
Otomatik Ortam Hazırlama ve Paket Kurulum Betiği (Offline / Intranet Destekli)
Bu betik, eksik paketleri tespit eder ve varsa 'wheels/' klasöründen
tamamen internetsiz (offline) olarak kurar.
"""

import sys
import os
import subprocess

REQUIRED_MODULES = [
    ("mcp", "mcp"),
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("uvicorn", "uvicorn"),
    ("starlette", "starlette"),
    ("sse_starlette", "sse-starlette"),
    ("pydantic", "pydantic"),
]

def check_missing_packages():
    missing = []
    for mod_name, pkg_name in REQUIRED_MODULES:
        try:
            __import__(mod_name)
        except ImportError:
            missing.append(pkg_name)
    return missing

def install_packages():
    missing = check_missing_packages()
    if not missing:
        print("[OK] Tum gerekli paketler kurulu ve calismaya hazir.")
        return True

    print(f"[!] Eksik paketler tespit edildi: {missing}")
    project_root = os.path.dirname(os.path.abspath(__file__))
    wheels_dir = os.path.join(project_root, "wheels")
    req_file = os.path.join(project_root, "requirements.txt")

    # 1. Oncelik: wheels/ klasorunden internetsiz (offline) kurulum
    if os.path.exists(wheels_dir) and os.listdir(wheels_dir):
        print(f"[*] 'wheels/' klasoru bulundu. Internetsiz (offline) kurulum baslatiliyor...")
        cmd = [
            sys.executable, "-m", "pip", "install",
            "--no-index",
            f"--find-links={wheels_dir}",
            "-r", req_file
        ]
        res = subprocess.run(cmd)
        if res.returncode == 0:
            print("[OK] Paketler 'wheels/' uzerinden basariyla kuruldu.")
            return True
        else:
            print("[UYARI] wheels uzerinden kurulum basarisiz oldu, standart pip denenecek...")

    # 2. Standart pip kurulumu (varsa internet veya yerel pip cache)
    print("[*] Standart pip ile kurulum baslatiliyor...")
    cmd = [sys.executable, "-m", "pip", "install", "-r", req_file]
    res = subprocess.run(cmd)
    if res.returncode == 0:
        print("[OK] Paketler basariyla kuruldu.")
        return True
    else:
        print("[HATA] Paket kurulumu tamamlanamadi.")
        return False

if __name__ == "__main__":
    success = install_packages()
    sys.exit(0 if success else 1)
