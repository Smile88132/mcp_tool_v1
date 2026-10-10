"""
Otomatik Ortam Hazırlama ve Paket Kurulum Betiği (Offline / Intranet / Anaconda Uyumlu)

Özellikler:
1. İlk önce standart 'pip install' dener, internetsiz ise otomatik olarak 'wheels/' klasörünü kullanır.
2. Anaconda ortamlarında pywin32 DLL yollarını (pywintypes / pythoncom) otomatik tespit edip onarır.
3. Ağ Web Servisi için kritik paketlerin (FastAPI, Uvicorn, Pandas, OpenpyXL) hazır olduğunu garanti eder.
"""

import sys
import os
import shutil
import subprocess

# Ağ Web Servisi için kesinlikle zorunlu olan temel paketler
CRITICAL_CHECKS = [
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("uvicorn", "uvicorn"),
    ("fastapi", "fastapi"),
    ("pydantic", "pydantic"),
]

# İsteğe bağlı / ek paketler
OPTIONAL_CHECKS = [
    ("starlette", "starlette"),
    ("sse_starlette", "sse-starlette"),
    ("opentelemetry", "opentelemetry-api"),
    ("mcp", "mcp"),
]

def fix_anaconda_pywin32_dlls():
    """Anaconda ortamlarında pywintypes ve pythoncom DLL dosyalarını win32 klasörüne kopyalar ve DLL yoluna ekler."""
    if sys.platform != "win32":
        return
    try:
        import site
        site_dirs = site.getsitepackages() if hasattr(site, "getsitepackages") else [os.path.join(sys.prefix, "Lib", "site-packages")]
        if hasattr(site, "getusersitepackages"):
            site_dirs.append(site.getusersitepackages())

        dll_dirs = [
            os.path.join(sys.prefix, "Lib", "site-packages", "pywin32_system32"),
            os.path.join(sys.prefix, "Library", "bin"),
            os.path.join(sys.prefix, "Scripts"),
        ]

        for s_dir in site_dirs:
            p32 = os.path.join(s_dir, "pywin32_system32")
            if os.path.exists(p32):
                dll_dirs.append(p32)
                # pywintypes ve pythoncom DLL'lerini doğrudan win32 klasörüne kopyala
                win32_dir = os.path.join(s_dir, "win32")
                if os.path.exists(win32_dir):
                    for fname in os.listdir(p32):
                        if fname.lower().endswith(".dll"):
                            src = os.path.join(p32, fname)
                            dst = os.path.join(win32_dir, fname)
                            if not os.path.exists(dst):
                                try:
                                    shutil.copy2(src, dst)
                                except Exception:
                                    pass

        for d in dll_dirs:
            if os.path.exists(d):
                try:
                    os.add_dll_directory(d)
                except Exception:
                    pass
                os.environ["PATH"] = d + os.pathsep + os.environ.get("PATH", "")
    except Exception:
        pass


def clean_corrupted_dist_info():
    """Önceki başarısız kurulumlardan kalan bozuk dist-info klasörlerini temizler."""
    try:
        import site
        site_dirs = site.getsitepackages() if hasattr(site, "getsitepackages") else []
        if hasattr(site, "getusersitepackages"):
            site_dirs.append(site.getusersitepackages())
        for s_dir in site_dirs:
            broken_path = os.path.join(s_dir, "pywin32-312.dist-info")
            if os.path.exists(broken_path):
                shutil.rmtree(broken_path, ignore_errors=True)
    except Exception:
        pass


def is_module_available(mod_name: str, verbose: bool = False) -> bool:
    try:
        __import__(mod_name)
        return True
    except Exception as e:
        if verbose:
            print(f"[-] '{mod_name}' modülü yüklenemedi: {e}")
        return False


def run_pip_install(packages, use_wheels=False, use_user=False):
    """pip install komutunu çalıştırır."""
    if not packages:
        return True

    project_root = os.path.dirname(os.path.abspath(__file__))
    wheels_dir = os.path.join(project_root, "wheels")

    cmd = [sys.executable, "-m", "pip", "install"]
    if use_wheels and os.path.exists(wheels_dir):
        cmd += ["--no-index", f"--find-links={wheels_dir}"]
    if use_user:
        cmd.append("--user")
    cmd += ["--no-warn-script-location"] + packages

    mode_info = f"{'WHEELS' if use_wheels else 'STANDART PIP'} | {'USER' if use_user else 'SYSTEM'}"
    print(f"[*] Kurulum deneniyor [{mode_info}]: {packages}")
    res = subprocess.run(cmd)
    return res.returncode == 0


def install_packages():
    # 1. DLL onarımı ve temizlik
    clean_corrupted_dist_info()
    fix_anaconda_pywin32_dlls()

    # Eksik paketleri tespit et
    missing_critical = [pkg for mod, pkg in CRITICAL_CHECKS if not is_module_available(mod)]
    missing_optional = [pkg for mod, pkg in OPTIONAL_CHECKS if not is_module_available(mod)]

    all_missing = missing_critical + missing_optional
    if not all_missing:
        print("[OK] Tüm gerekli paketler kurulu ve çalışmaya hazır.")
        return True

    print(f"[!] Eksik paketler tespit edildi: {all_missing}")

    # Sıralı Deneme Stratejisi:
    # 1. Deneme: Standart 'pip install' (Kullanıcının isteği: internet/ayna varsa ilk önce normal pip)
    # 2. Deneme: Standart 'pip install --user'
    # 3. Deneme: Yerel 'wheels/' klasörü ile çevrimdışı kurulum
    # 4. Deneme: Yerel 'wheels/' klasörü + '--user'
    attempts = [
        (False, False),  # Online / Standart
        (False, True),   # Online + --user
        (True, False),   # Offline wheels
        (True, True),    # Offline wheels + --user
    ]

    for use_wheels, use_user in attempts:
        # mcp paketini diğer paketlerle birlikte doğrudan kur veya --no-deps yap
        pkgs = [p for p in all_missing if p != "mcp"]
        if pkgs:
            run_pip_install(pkgs, use_wheels=use_wheels, use_user=use_user)

        if "mcp" in all_missing:
            run_pip_install(["mcp"], use_wheels=use_wheels, use_user=use_user)

        # Tekrar kontrol et
        clean_corrupted_dist_info()
        fix_anaconda_pywin32_dlls()

        still_missing_critical = [pkg for mod, pkg in CRITICAL_CHECKS if not is_module_available(mod)]
        if not still_missing_critical:
            break

    # Son Değerlendirme
    fix_anaconda_pywin32_dlls()
    final_missing_critical = [pkg for mod, pkg in CRITICAL_CHECKS if not is_module_available(mod, verbose=True)]

    if final_missing_critical:
        print(f"\n[HATA] Web servisi için zorunlu paketler yüklenemedi: {final_missing_critical}")
        return False

    final_missing_optional = [pkg for mod, pkg in OPTIONAL_CHECKS if not is_module_available(mod, verbose=False)]
    if final_missing_optional:
        print(f"\n[BILGI] Web Servisi hazır! Ekstra protokoller ({final_missing_optional}) atlandı.")
    else:
        print("[OK] Tüm paketler başarıyla doğrulandı!")

    return True


if __name__ == "__main__":
    ok = install_packages()
    sys.exit(0 if ok else 1)
