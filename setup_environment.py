"""
Otomatik Ortam Hazırlama ve Paket Kurulum Betiği (Offline / Intranet / Anaconda Uyumlu)

Özel Durum Çözümü:
Anaconda ortamlarında 'pywin32' (305.1) kütüphanesinin DLL dosyaları (_win32sysloader.pyd)
Windows tarafından kilitlendiği için pip paket yükseltirken [WinError 5] Erişim Engellendi
(Access is denied) hatası verir. Bu betik:
1. Yalnızca EKSİK olan paketleri tespit edip kurar (zaten kurulu olanları ellemez).
2. Anaconda'da pywin32 varsa onu kaldırmaya çalışmaz.
3. Gerekirse '--user' bayrağı ile kullanıcı alanına güvenli kurulum yapar.
"""

import sys
import os
import subprocess

REQUIRED_CHECKS = [
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("uvicorn", "uvicorn"),
    ("starlette", "starlette"),
    ("sse_starlette", "sse-starlette"),
    ("pydantic", "pydantic"),
    ("mcp", "mcp"),
]

def is_module_available(mod_name: str) -> bool:
    try:
        __import__(mod_name)
        return True
    except Exception:
        return False

def check_missing_packages():
    missing = []
    for mod_name, pkg_name in REQUIRED_CHECKS:
        if not is_module_available(mod_name):
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
    print(f"[*] Calistiriliyor: {' '.join(cmd)}")
    return subprocess.run(cmd)

def install_packages():
    missing = check_missing_packages()
    if not missing:
        print("[OK] Tum gerekli paketler kurulu ve calismaya hazir.")
        return True

    print(f"[!] Eksik paketler tespit edildi: {missing}")
    project_root = os.path.dirname(os.path.abspath(__file__))
    wheels_dir = os.path.join(project_root, "wheels")
    has_wheels = os.path.exists(wheels_dir) and len(os.listdir(wheels_dir)) > 0
    pywin32_already_ok = is_pywin32_present()

    # Paketleri güvenli sıra ile yükleme
    # Eğer pywin32 zaten varsa, mcp kurulumunun pywin32'yi bozmasını engellemek için --no-deps mcp yapılır
    pkgs_to_install = []
    install_mcp_no_deps = False

    for pkg in missing:
        if pkg == "mcp" and pywin32_already_ok:
            install_mcp_no_deps = True
            # mcp'nin diğer alt bağımlılıklarını ekle
            for sub in ["mcp-types", "anyio", "httpx2", "httpcore2", "pydantic", "sse-starlette", "uvicorn", "starlette"]:
                if not is_module_available(sub.replace("-", "_")):
                    if sub not in pkgs_to_install:
                        pkgs_to_install.append(sub)
        else:
            if pkg not in pkgs_to_install:
                pkgs_to_install.append(pkg)

    print(f"[*] Kurulacak hedef paketler: {pkgs_to_install}")

    for use_user_flag in [False, True]:
        success = True

        # 1. Önce alt bağımlılıkları ve normal eksik paketleri kur
        if pkgs_to_install:
            base_args = ["install"]
            if has_wheels:
                base_args += ["--no-index", f"--find-links={wheels_dir}"]
            base_args += pkgs_to_install

            res = run_pip_command(base_args, use_user=use_user_flag)
            if res.returncode != 0:
                print(f"[UYARI] Kurulum adimi basarisiz oldu (use_user={use_user_flag}).")
                success = False

        # 2. Eğer mcp eksikse ve pywin32 zaten varsa --no-deps ile mcp kur
        if success and install_mcp_no_deps:
            mcp_args = ["install", "--no-deps"]
            if has_wheels:
                mcp_args += ["--no-index", f"--find-links={wheels_dir}"]
            mcp_args.append("mcp")

            res = run_pip_command(mcp_args, use_user=use_user_flag)
            if res.returncode != 0:
                success = False

        # Tekrar kontrol et: Tüm paketler geldi mi?
        still_missing = check_missing_packages()
        if not still_missing:
            print("[OK] Kurulum basariyla tamamlandi. Tum paketler dogrulandi!")
            return True

        if not use_user_flag:
            print("[*] Standart kurulumda erisim hatasi olmus olabilir, '--user' bayragi deneniyor...")

    # Son kontrol
    final_missing = check_missing_packages()
    if not final_missing:
        print("[OK] Tum paketler hazir.")
        return True
    else:
        print(f"[HATA] Asagidaki paketler yuklenemedi: {final_missing}")
        return False

if __name__ == "__main__":
    ok = install_packages()
    sys.exit(0 if ok else 1)
