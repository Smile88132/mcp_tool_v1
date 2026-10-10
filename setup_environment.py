"""
Otomatik Ortam Hazırlama ve Paket Kurulum Betiği (Offline / Intranet / Anaconda Uyumlu)

Özel Durum Çözümleri:
1. Anaconda ortamlarında pywin32 (305.1) DLL'i (_win32sysloader.pyd) kilitli olduğu için
   pywin32 yükseltmesi yapılmaz, mcp ve tüm alt bağımlılıkları (opentelemetry, pyjwt vb.)
   güvenli bir sırada kurulur.
2. Yarım kalan kurulumlardan kalan bozuk pywin32-312.dist-info temizlenir.
3. Eksik paket durumunda hata sebebi açıkça ekrana basılır.
"""

import sys
import os
import shutil
import subprocess

# Temel modül kontrolleri
REQUIRED_CHECKS = [
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("uvicorn", "uvicorn"),
    ("starlette", "starlette"),
    ("sse_starlette", "sse-starlette"),
    ("pydantic", "pydantic"),
    ("opentelemetry", "opentelemetry-api"),
    ("mcp", "mcp"),
]

# mcp'nin pywin32 haricindeki tüm zorunlu bağımlılıkları
MCP_SUB_DEPENDENCIES = [
    "mcp-types",
    "opentelemetry-api",
    "pyjwt",
    "python-multipart",
    "anyio",
    "httpx2",
    "httpcore2",
    "pydantic",
    "sse-starlette",
    "uvicorn",
    "starlette",
    "jsonschema",
    "typing-extensions",
    "typing-inspection",
]

def clean_corrupted_dist_info():
    """Önceki başarısız kurulumlardan kalan bozuk pywin32 dist-info klasörünü temizler."""
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

def check_missing_packages(verbose: bool = False):
    missing = []
    for mod_name, pkg_name in REQUIRED_CHECKS:
        if not is_module_available(mod_name, verbose=verbose):
            missing.append(pkg_name)
    return missing

def is_pywin32_present():
    """pywin32 veya win32api sistemde zaten var mı?"""
    try:
        import win32api  # noqa
        return True
    except Exception:
        pass
    try:
        import win32file  # noqa
        return True
    except Exception:
        return False

def run_pip_command(args, use_user=False):
    cmd = [sys.executable, "-m", "pip"] + args
    if use_user and "--user" not in cmd:
        cmd.append("--user")
    cmd.append("--no-warn-script-location")
    print(f"[*] Çalıştırılıyor: {' '.join(cmd)}")
    return subprocess.run(cmd)

def install_packages():
    clean_corrupted_dist_info()

    missing = check_missing_packages(verbose=False)
    if not missing:
        print("[OK] Tüm gerekli paketler kurulu ve çalışmaya hazır.")
        return True

    print(f"[!] Eksik paketler tespit edildi: {missing}")
    project_root = os.path.dirname(os.path.abspath(__file__))
    wheels_dir = os.path.join(project_root, "wheels")
    has_wheels = os.path.exists(wheels_dir) and len(os.listdir(wheels_dir)) > 0
    pywin32_already_ok = is_pywin32_present()

    # Kurulacak paket listesini oluştur
    pkgs_to_install = []
    install_mcp_no_deps = False

    for pkg in missing:
        if pkg == "mcp":
            install_mcp_no_deps = True
            # mcp'nin çalışması için gereken tüm alt bağımlılıkları ekle
            for sub in MCP_SUB_DEPENDENCIES:
                sub_mod = sub.replace("-", "_")
                if not is_module_available(sub_mod, verbose=False):
                    if sub not in pkgs_to_install:
                        pkgs_to_install.append(sub)
        else:
            if pkg not in pkgs_to_install:
                pkgs_to_install.append(pkg)

    print(f"[*] Kurulacak öncelikli paketler: {pkgs_to_install}")

    for use_user_flag in [False, True]:
        # 1. Adım: Tüm alt bağımlılıkları ve eksik kütüphaneleri kur
        if pkgs_to_install:
            base_args = ["install"]
            if has_wheels:
                base_args += ["--no-index", f"--find-links={wheels_dir}"]
            base_args += pkgs_to_install
            run_pip_command(base_args, use_user=use_user_flag)

        # 2. Adım: mcp paketini --no-deps ile kur (pywin32'ye dokunmasın)
        if install_mcp_no_deps:
            mcp_args = ["install", "--no-deps"]
            if has_wheels:
                mcp_args += ["--no-index", f"--find-links={wheels_dir}"]
            mcp_args.append("mcp")
            run_pip_command(mcp_args, use_user=use_user_flag)

        # Kontrol et: Başarılı oldu mu?
        clean_corrupted_dist_info()
        still_missing = check_missing_packages(verbose=False)
        if not still_missing:
            print("[OK] Kurulum başarıyla tamamlandı. Tüm paketler doğrulandı!")
            return True

        if not use_user_flag:
            print("[*] Standart kurulumda eksik kaldı, '--user' parametresi deneniyor...")

    # Son kontrol (hataları ayrıntılı göster)
    print("\n[!] Son durum kontrol ediliyor...")
    final_missing = check_missing_packages(verbose=True)
    if not final_missing:
        print("[OK] Tüm paketler hazır.")
        return True
    else:
        print(f"\n[HATA] Aşağıdaki paketler yüklenemedi: {final_missing}")
        return False

if __name__ == "__main__":
    ok = install_packages()
    sys.exit(0 if ok else 1)
