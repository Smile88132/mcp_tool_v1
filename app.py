"""
Intranet RAG ve Web Erişimi için Venn Şeması Uygunsuzluk Analiz Sunucusu
Model Context Protocol (MCP) & HTTP REST API Server for Nonconformity Venn Analysis

Ortam Özellikleri:
- Tamamen ağdaki herkese açık (0.0.0.0:8000) standart HTTP REST API ve Web Dokümantasyonu (/docs).
- Karmaşık SSE gereksinimi olmadan her bilgisayardan, tarayıcıdan veya RAG istemcisinden doğrudan erişilebilir.
- Windows Anaconda pywin32 DLL yükleme sorunlarına karşı korumalı.
- İsteğe bağlı olarak yerel stdio MCP protokolünü de destekler.
"""

import os
import sys
import io
import json
import uuid
import argparse
import unicodedata
import time
from difflib import SequenceMatcher
from typing import Any, Dict, List, Optional, Set, Tuple
import pandas as pd
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==============================================================================
# Windows / Anaconda DLL Yükleme Düzeltmesi (win32api DLL load failed fix)
# ==============================================================================
if sys.platform == "win32":
    # Anaconda'daki pywintypes.dll ve pythoncom.dll dosyalarını DLL arama yoluna ekle
    possible_dll_dirs = [
        os.path.join(sys.prefix, "Lib", "site-packages", "pywin32_system32"),
        os.path.join(sys.prefix, "Library", "bin"),
        os.path.join(sys.prefix, "Scripts"),
    ]
    # Kullanıcı alanı (Roaming) desteği
    user_roaming = os.environ.get("APPDATA", "")
    if user_roaming:
        py_ver = f"Python{sys.version_info.major}{sys.version_info.minor}"
        possible_dll_dirs.append(os.path.join(user_roaming, "Python", py_ver, "site-packages", "pywin32_system32"))

    for d in possible_dll_dirs:
        if os.path.exists(d):
            try:
                os.add_dll_directory(d)
                os.environ["PATH"] = d + os.pathsep + os.environ.get("PATH", "")
            except Exception:
                pass


# ==============================================================================
# Metin Temizleme ve Sütun Eşleme Yardımcıları
# ==============================================================================

def _safe_log(msg: str) -> None:
    """Konsola log yazar; Windows kod sayfası (cp1254/cp857) hatalarında isteği asla çökertmez."""
    try:
        print(msg, flush=True)
    except Exception:
        try:
            print(msg.encode("ascii", "replace").decode("ascii"), flush=True)
        except Exception:
            pass


def normalize_str(text: Any) -> str:
    """Türkçe karakter ve boşluk toleranslı küçük harf normalizasyonu."""
    if text is None or pd.isna(text):
        return ""
    s = str(text).strip()
    replacements = {
        "İ": "i", "I": "i", "ı": "i",
        "Ğ": "g", "ğ": "g",
        "Ü": "u", "ü": "u",
        "Ş": "s", "ş": "s",
        "Ö": "o", "ö": "o",
        "Ç": "c", "ç": "c"
    }
    for k, v in replacements.items():
        s = s.replace(k, v)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.lower()


def clean_text_content(text: Any) -> str:
    """Excel hücrelerindeki '_x000D_', satır sonları ve çoklu boşlukları temizler."""
    if text is None or pd.isna(text):
        return ""
    s = str(text)
    s = s.replace("_x000D_", " ")
    s = s.replace("\r", " ").replace("\n", " ")
    return " ".join(s.split()).strip()


ACTIVITY_ALIASES = [
    "faaliyet", "faaliyetler", "faaliyet adı", "faaliyet_adi", "aktivite",
    "activity", "islem", "işlem", "il", "şehir", "sehir", "bolge", "teslimat", "scope"
]

PHASE_ALIASES = [
    "faz", "fazlar", "aşama", "asama", "test aşaması", "test_asamasi",
    "phase", "stage", "test_fazı", "test fazi"
]

ISSUE_ALIASES = [
    "açıklama", "aciklama", "uygunsuzluk", "uygunsuzluklar",
    "uygunsuzluk tanımı", "uygunsuzluk_tanimi", "hata", "problem",
    "bulgu", "issue", "nonconformity", "defect", "tanım", "tanim"
]


def resolve_columns(
    df: pd.DataFrame,
    custom_act: Optional[str] = None,
    custom_phase: Optional[str] = None,
    custom_issue: Optional[str] = None
) -> Tuple[str, str, str]:
    """
    Excel DataFrame içindeki Faaliyet, Faz ve Uygunsuzluk sütunlarını otomatik çözer.
    """
    available_cols = list(df.columns)
    available_norm = {normalize_str(c): c for c in available_cols}

    # 1. Uygunsuzluk Sütunu (Açıklama)
    issue_col = None
    if custom_issue:
        if custom_issue in available_cols:
            issue_col = custom_issue
        elif normalize_str(custom_issue) in available_norm:
            issue_col = available_norm[normalize_str(custom_issue)]
    if not issue_col:
        for alias in ISSUE_ALIASES:
            norm_a = normalize_str(alias)
            if norm_a in available_norm:
                issue_col = available_norm[norm_a]
                break
            for c_norm, c_orig in available_norm.items():
                if norm_a in c_norm:
                    issue_col = c_orig
                    break
            if issue_col:
                break
    if not issue_col:
        raise ValueError(
            f"Uygunsuzluk/Açıklama sütunu tespit edilemedi! "
            f"Excel dosyasındaki mevcut sütunlar: {available_cols}"
        )

    # 2. Faz Sütunu
    phase_col = None
    if custom_phase:
        if custom_phase in available_cols:
            phase_col = custom_phase
        elif normalize_str(custom_phase) in available_norm:
            phase_col = available_norm[normalize_str(custom_phase)]

    if not phase_col:
        for alias in PHASE_ALIASES:
            norm_a = normalize_str(alias)
            if norm_a in available_norm:
                phase_col = available_norm[norm_a]
                break

    # 3. Faaliyet Sütunu
    act_col = None
    if custom_act:
        if custom_act in available_cols:
            act_col = custom_act
        elif normalize_str(custom_act) in available_norm:
            act_col = available_norm[normalize_str(custom_act)]

    # Akıllı Tespit: 'Faz' sütunu yoksa ama 'Faaliyet' sütunu 'dry-run', 'skt' içeriyorsa
    if not phase_col:
        faaliyet_candidate = None
        for c in available_cols:
            if normalize_str(c) == "faaliyet":
                faaliyet_candidate = c
                break
        
        if faaliyet_candidate:
            sample_values = set(df[faaliyet_candidate].dropna().astype(str).str.strip().str.lower().unique())
            phase_indicators = {"dry-run", "dry run", "skt", "fkt", "kabul_muayene_testi", "kabul"}
            if sample_values & phase_indicators:
                phase_col = faaliyet_candidate
                if not act_col:
                    for scope_alias in ["il", "sehir", "bolge", "teslimat", "faaliyet turu"]:
                        for c_norm, c_orig in available_norm.items():
                            if scope_alias == c_norm and c_orig != phase_col:
                                act_col = c_orig
                                break
                        if act_col:
                            break

    if not act_col:
        for alias in ACTIVITY_ALIASES:
            norm_a = normalize_str(alias)
            if norm_a in available_norm and available_norm[norm_a] != phase_col:
                act_col = available_norm[norm_a]
                break
            if act_col:
                break

    if not act_col:
        remaining = [c for c in available_cols if c not in [phase_col, issue_col, "Id", "Tags"]]
        act_col = remaining[0] if remaining else available_cols[0]

    if not phase_col:
        raise ValueError(
            f"Faz (Aşama) sütunu bulunamadı! Beklenen alternatifler: {PHASE_ALIASES}. "
            f"Excel dosyasındaki mevcut sütunlar: {available_cols}"
        )

    return act_col, phase_col, issue_col


def read_excel_safe(file_path: Optional[Any] = None, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """Excel dosyasını, bayt verisini veya DataFrame'i güvenli şekilde okur. Belirtilmezse varsayılan dosyayı kullanır."""
    if isinstance(file_path, pd.DataFrame):
        return file_path.copy()

    if isinstance(file_path, bytes):
        bio = io.BytesIO(file_path)
        return pd.read_excel(bio, sheet_name=sheet_name) if sheet_name else pd.read_excel(bio)

    if hasattr(file_path, "read"):
        return pd.read_excel(file_path, sheet_name=sheet_name) if sheet_name else pd.read_excel(file_path)

    if not file_path or str(file_path).strip() == "":
        candidates = ["Uygunsuzluk.xlsx", "ornek_uygunsuzluklar.xlsx"]
        found = None
        for c in candidates:
            if os.path.exists(c):
                found = os.path.abspath(c)
                break
            script_c = os.path.join(os.path.dirname(os.path.abspath(__file__)), c)
            if os.path.exists(script_c):
                found = script_c
                break
        if found:
            file_path = found
        else:
            raise FileNotFoundError("Excel dosya yolu belirtilmedi ve varsayılan 'Uygunsuzluk.xlsx' dosyası bulunamadı.")

    file_str = str(file_path)
    if not os.path.isabs(file_str):
        file_str = os.path.abspath(file_str)

    if not os.path.exists(file_str):
        raise FileNotFoundError(f"Belirtilen dosya bulunamadı: '{file_str}'")

    if not os.path.isfile(file_str):
        raise ValueError(f"Belirtilen yol bir dosya değil: '{file_str}'")

    valid_extensions = (".xlsx", ".xls", ".xlsm")
    if not file_str.lower().endswith(valid_extensions):
        raise ValueError(f"Geçersiz dosya uzantısı. Desteklenen uzantılar: {valid_extensions}. Verilen: '{file_str}'")

    try:
        if sheet_name:
            df = pd.read_excel(file_str, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_str)
    except Exception as e:
        raise RuntimeError(f"Excel dosyası okunurken hata oluştu: {str(e)}")

    if df.empty:
        raise ValueError(f"Excel dosyası ('{file_str}') boş veri içeriyor.")

    return df



def categorize_phase(raw_phase: str) -> str:
    """Faz ismini standart etiketlere dönüştürür."""
    p = normalize_str(raw_phase)
    if "dry" in p:
        return "Dry-run"
    elif "skt" in p or "saha" in p:
        return "SKT"
    elif "fkt" in p or "fabrika" in p or "kabul" in p or "muayene" in p:
        return "FKT"
    return str(raw_phase).strip()


# ==============================================================================
# Venn Şeması Matematik Motoru
# ==============================================================================

def compute_3way_venn(a_name: str, set_a: Set[str],
                       b_name: str, set_b: Set[str],
                       c_name: str, set_c: Set[str]) -> Dict[str, Any]:
    only_a = set_a - (set_b | set_c)
    only_b = set_b - (set_a | set_c)
    only_c = set_c - (set_a | set_b)

    ab_only = (set_a & set_b) - set_c
    ac_only = (set_a & set_c) - set_b
    bc_only = (set_b & set_c) - set_a

    all_three = set_a & set_b & set_c
    total_union = set_a | set_b | set_c

    ab_total = set_a & set_b
    ac_total = set_a & set_c
    bc_total = set_b & set_c

    def jaccard(s1: Set[str], s2: Set[str]) -> float:
        union_len = len(s1 | s2)
        return round(len(s1 & s2) / union_len, 4) if union_len > 0 else 0.0

    return {
        "ozet_istatistikler": {
            f"toplam_benzersiz_{a_name}": len(set_a),
            f"toplam_benzersiz_{b_name}": len(set_b),
            f"toplam_benzersiz_{c_name}": len(set_c),
            "toplam_benzersiz_uygunsuzluk_sayisi": len(total_union),
        },
        "venn_bolgeleri": {
            f"sadece_{a_name}": {
                "adet": len(only_a),
                "ogeler": sorted(list(only_a)),
                "aciklama": f"Yalnızca {a_name} aşamasında tespit edilen uygunsuzluklar."
            },
            f"sadece_{b_name}": {
                "adet": len(only_b),
                "ogeler": sorted(list(only_b)),
                "aciklama": f"Yalnızca {b_name} aşamasında ortaya çıkan yeni uygunsuzluklar."
            },
            f"sadece_{c_name}": {
                "adet": len(only_c),
                "ogeler": sorted(list(only_c)),
                "aciklama": f"Yalnızca {c_name} aşamasında tespit edilen uygunsuzluklar."
            },
            f"ortak_{a_name}_ve_{b_name}_({c_name}_haric)": {
                "adet": len(ab_only),
                "ogeler": sorted(list(ab_only)),
                "aciklama": f"Hem {a_name} hem {b_name} aşamasında görülen uygunsuzluklar."
            },
            f"ortak_{a_name}_ve_{c_name}_({b_name}_haric)": {
                "adet": len(ac_only),
                "ogeler": sorted(list(ac_only)),
                "aciklama": f"Hem {a_name} hem {c_name} aşamasında görülen uygunsuzluklar."
            },
            f"ortak_{b_name}_ve_{c_name}_({a_name}_haric)": {
                "adet": len(bc_only),
                "ogeler": sorted(list(bc_only)),
                "aciklama": f"Hem {b_name} hem {c_name} aşamasında devam eden uygunsuzluklar."
            },
            "ucunde_ortak_kesisim": {
                "adet": len(all_three),
                "ogeler": sorted(list(all_three)),
                "aciklama": "Üç test aşamasında da kesintisiz tekrarlayan kronik uygunsuzluklar."
            }
        },
        "ikili_kesisim_toplamlari": {
            f"{a_name}_ve_{b_name}_toplam_kesisim": len(ab_total),
            f"{a_name}_ve_{c_name}_toplam_kesisim": len(ac_total),
            f"{b_name}_ve_{c_name}_toplam_kesisim": len(bc_total),
        },
        "jaccard_benzerlik_indeksleri": {
            f"{a_name}_ile_{b_name}": jaccard(set_a, set_b),
            f"{a_name}_ile_{c_name}": jaccard(set_a, set_c),
            f"{b_name}_ile_{c_name}": jaccard(set_b, set_c),
        }
    }


def compute_general_venn(sets_dict: Dict[str, Set[str]]) -> Dict[str, Any]:
    names = list(sets_dict.keys())
    total_union = set().union(*sets_dict.values()) if sets_dict else set()
    common_all = set.intersection(*sets_dict.values()) if sets_dict else set()

    individual_uniques: Dict[str, Dict[str, Any]] = {}
    for name, s in sets_dict.items():
        other_union = set().union(*[s_other for n_other, s_other in sets_dict.items() if n_other != name])
        exclusive = s - other_union
        individual_uniques[name] = {
            "toplam_adet": len(s),
            "sadece_bu_kumedeki_adet": len(exclusive),
            "sadece_bu_kumede_olanlar": sorted(list(exclusive))[:100]
        }

    pairwise_intersections: Dict[str, Any] = {}
    for i in range(len(names)):
        for j in range(i + 1, len(names)):
            n1, n2 = names[i], names[j]
            s1, s2 = sets_dict[n1], sets_dict[n2]
            inter = s1 & s2
            union = s1 | s2
            jaccard = round(len(inter) / len(union), 4) if len(union) > 0 else 0.0
            pairwise_intersections[f"{n1} <-> {n2}"] = {
                "kesisen_adet": len(inter),
                "jaccard_benzerlik": jaccard,
                "kesisen_ogeler": sorted(list(inter))
            }

    return {
        "kumeler": names,
        "toplam_benzersiz_oge_sayisi": len(total_union),
        "tumunde_ortak_olanlar": {
            "adet": len(common_all),
            "ogeler": sorted(list(common_all))
        },
        "kume_detaylari": individual_uniques,
        "ikili_kesisimler": pairwise_intersections
    }


# ==============================================================================
# ÇEKİRDEK İŞ MANTIĞI FONKSİYONLARI (Python Core API)
# ==============================================================================

def execute_analyze_venn(
    file_path: str,
    compare_by: str = "phases",
    activity_name: Optional[str] = None,
    sheet_name: Optional[str] = None,
    activity_column: Optional[str] = None,
    phase_column: Optional[str] = None,
    issue_column: Optional[str] = None
) -> Dict[str, Any]:
    try:
        df = read_excel_safe(file_path, sheet_name=sheet_name)
        act_col, phase_col, issue_col = resolve_columns(
            df, custom_act=activity_column, custom_phase=phase_column, custom_issue=issue_column
        )

        clean_df = df.dropna(subset=[issue_col]).copy()
        clean_df[act_col] = clean_df[act_col].apply(clean_text_content)
        clean_df[phase_col] = clean_df[phase_col].apply(clean_text_content)
        clean_df[issue_col] = clean_df[issue_col].apply(clean_text_content)
        clean_df = clean_df[clean_df[issue_col] != ""]

        compare_by_clean = compare_by.strip().lower()

        if compare_by_clean in ["phases", "faz", "fazlar"]:
            target_df = clean_df
            scope_info = f"Tüm Kapsam ({act_col} Tümü)"

            if activity_name and activity_name.strip():
                act_req = activity_name.strip()
                matches = target_df[target_df[act_col].str.casefold() == act_req.casefold()]
                if matches.empty:
                    available_acts = sorted(target_df[act_col].unique().tolist())
                    return {
                        "durum": "hata",
                        "mesaj": f"'{act_req}' isimli faaliyet/kapsam bulunamadı.",
                        "mevcut_secenekler": available_acts,
                        "mevcut_faaliyetler": available_acts,
                        "kapsam_sutunu": act_col
                    }
                target_df = matches
                scope_info = f"{act_col}: '{matches[act_col].iloc[0]}'"

            target_df = target_df.copy()
            target_df["KategorizeFaz"] = target_df[phase_col].apply(categorize_phase)

            dry_run_set = set(target_df[target_df["KategorizeFaz"] == "Dry-run"][issue_col].unique())
            skt_set = set(target_df[target_df["KategorizeFaz"] == "SKT"][issue_col].unique())
            fkt_set = set(target_df[target_df["KategorizeFaz"] == "FKT"][issue_col].unique())

            venn_data = compute_3way_venn("Dry-run", dry_run_set, "SKT", skt_set, "FKT", fkt_set)

            total_unique = venn_data["ozet_istatistikler"]["toplam_benzersiz_uygunsuzluk_sayisi"]
            ucunde_ortak_count = venn_data["venn_bolgeleri"]["ucunde_ortak_kesisim"]["adet"]
            sadece_dry_count = venn_data["venn_bolgeleri"]["sadece_Dry-run"]["adet"]

            rag_insight = (
                f"{scope_info} kapsamında incelenen {total_unique} benzersiz uygunsuzluktan "
                f"{sadece_dry_count} tanesi yalnızca Dry-run aşamasında görülüp sonraki aşamalarda çözülmüştür. "
                f"{ucunde_ortak_count} adet uygunsuzluk ise Dry-run, SKT ve FKT aşamalarının üçünde de tekrarlamış olup kronik kök neden analizi gerektirmektedir."
            )

            return {
                "durum": "basarili",
                "analiz_tipi": "fazlar_arasi_venn_semasi",
                "kapsam": scope_info,
                "rag_ozet_yorumu": rag_insight,
                "sutun_bilgileri": {
                    "faaliyet_veya_kapsam_sutunu": act_col,
                    "faz_sutunu": phase_col,
                    "uygunsuzluk_sutunu": issue_col
                },
                "venn_analizi": venn_data
            }

        elif compare_by_clean in ["activities", "faaliyet", "faaliyetler", "iller", "cities"]:
            activity_sets: Dict[str, Set[str]] = {}
            for act in clean_df[act_col].unique():
                act_str = str(act).strip()
                if act_str:
                    issues = set(clean_df[clean_df[act_col] == act][issue_col].unique())
                    activity_sets[act_str] = issues

            if len(activity_sets) < 2:
                return {
                    "durum": "uyari",
                    "mesaj": f"Karşılaştırma için en az 2 grup bulunmalı. Mevcut sayı: {len(activity_sets)}",
                    "faaliyetler": list(activity_sets.keys())
                }

            venn_data = compute_general_venn(activity_sets)

            return {
                "durum": "basarili",
                "analiz_tipi": "faaliyetler_arasi_kume_analizi",
                "karsilastirilan_sutun": act_col,
                "toplam_grup_sayisi": len(activity_sets),
                "sutun_bilgileri": {
                    "faaliyet_veya_kapsam_sutunu": act_col,
                    "faz_sutunu": phase_col,
                    "uygunsuzluk_sutunu": issue_col
                },
                "venn_analizi": venn_data
            }

        else:
            return {
                "durum": "hata",
                "mesaj": f"Geçersiz 'compare_by' değeri: '{compare_by}'. Lütfen 'phases' veya 'activities' kullanın."
            }

    except Exception as e:
        return {
            "durum": "hata",
            "hata_turu": type(e).__name__,
            "mesaj": str(e),
            "dosya_yolu": file_path
        }


def execute_list_activities_and_phases(file_path: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    try:
        df = read_excel_safe(file_path, sheet_name=sheet_name)
        act_col, phase_col, issue_col = resolve_columns(df)

        activities_summary = []
        for act in sorted(df[act_col].dropna().unique()):
            act_clean = clean_text_content(act)
            act_df = df[df[act_col].apply(clean_text_content) == act_clean]
            phases = [clean_text_content(p) for p in act_df[phase_col].dropna().unique()]
            issue_count = act_df[issue_col].dropna().nunique()
            total_records = len(act_df)
            activities_summary.append({
                "faaliyet_veya_kapsam": act_clean,
                "mevcut_fazlar": phases,
                "benzersiz_uygunsuzluk_sayisi": issue_count,
                "toplam_kayit_satiri": total_records
            })

        all_phases = sorted(list(set(clean_text_content(p) for p in df[phase_col].dropna().unique())))

        return {
            "durum": "basarili",
            "dosya_yolu": file_path,
            "toplam_satir_sayisi": len(df),
            "tespit_edilen_sutunlar": {
                "faaliyet_veya_kapsam": act_col,
                "faz": phase_col,
                "uygunsuzluk": issue_col
            },
            "mevcut_fazlar": all_phases,
            "faaliyetler": [{"faaliyet_adi": a["faaliyet_veya_kapsam"], **a} for a in activities_summary],
            "faaliyetler_veya_kapsamlar": activities_summary
        }
    except Exception as e:
        return {
            "durum": "hata",
            "hata_turu": type(e).__name__,
            "mesaj": str(e),
            "dosya_yolu": file_path
        }


def execute_preview_excel_data(file_path: str, max_rows: int = 5, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    try:
        df = read_excel_safe(file_path, sheet_name=sheet_name)
        act_col, phase_col, issue_col = resolve_columns(df)
        return {
            "durum": "basarili",
            "toplam_satir": len(df),
            "sutunlar": list(df.columns),
            "tespit_edilen_roller": {
                "faaliyet_veya_kapsam": act_col,
                "faz": phase_col,
                "uygunsuzluk": issue_col
            },
            "ornek_veriler": df.head(max_rows).to_dict(orient="records")
        }
    except Exception as e:
        return {
            "durum": "hata",
            "hata_turu": type(e).__name__,
            "mesaj": str(e),
            "dosya_yolu": file_path
        }


# ==============================================================================
# Standart MCP Tool Arayüzü (Birim Testleri ve MCP İstemcileri ile Tam Uyumlu)
# ==============================================================================

def analyze_venn_nonconformities(
    file_path: str,
    compare_by: str = "phases",
    activity_name: Optional[str] = None,
    sheet_name: Optional[str] = None,
    activity_column: Optional[str] = None,
    phase_column: Optional[str] = None,
    issue_column: Optional[str] = None
) -> str:
    res = execute_analyze_venn(
        file_path=file_path,
        compare_by=compare_by,
        activity_name=activity_name,
        sheet_name=sheet_name,
        activity_column=activity_column,
        phase_column=phase_column,
        issue_column=issue_column
    )
    return json.dumps(res, ensure_ascii=False, indent=2)


def list_activities_and_phases(file_path: str, sheet_name: Optional[str] = None) -> str:
    res = execute_list_activities_and_phases(file_path, sheet_name=sheet_name)
    return json.dumps(res, ensure_ascii=False, indent=2)


def preview_excel_data(file_path: str, max_rows: int = 5, sheet_name: Optional[str] = None) -> str:
    res = execute_preview_excel_data(file_path, max_rows=max_rows, sheet_name=sheet_name)
    return json.dumps(res, ensure_ascii=False, indent=2)


def calculate_text_similarity(s1: str, s2: str) -> float:
    """
    İki metin arasındaki benzerlik puanını (0.0 - 1.0) hesaplar.
    Levenshtein SequenceMatcher + Kelime Sırası Bağımsız Token Sort + Jaccard token kesişimi.
    """
    if not s1 or not s2:
        return 0.0
    n1 = normalize_str(s1)
    n2 = normalize_str(s2)
    if n1 == n2:
        return 1.0

    seq_ratio = SequenceMatcher(None, n1, n2).ratio()
    if seq_ratio >= 0.98:
        return 1.0

    tokens1 = n1.split()
    tokens2 = n2.split()
    sorted1 = " ".join(sorted(tokens1))
    sorted2 = " ".join(sorted(tokens2))
    token_ratio = SequenceMatcher(None, sorted1, sorted2).ratio()

    set1, set2 = set(tokens1), set(tokens2)
    union_len = len(set1 | set2)
    jaccard_tokens = len(set1 & set2) / union_len if union_len > 0 else 0.0

    combined = max(seq_ratio, token_ratio) * 0.7 + jaccard_tokens * 0.3
    return round(min(max(combined, 0.0), 1.0), 4)


def compute_fuzzy_venn(
    items_a: List[str],
    items_b: List[str],
    threshold: float = 0.80
) -> Dict[str, Any]:
    """
    İki metin listesini satır bazlı karşılaştırır.
    Eğer threshold >= 1.0 ise tam eşleşme (exact match),
    threshold < 1.0 ise fuzzy distance matching yapar.
    """
    clean_a = [clean_text_content(x) for x in items_a if clean_text_content(x)]
    clean_b = [clean_text_content(x) for x in items_b if clean_text_content(x)]

    unique_a = list(dict.fromkeys(clean_a))
    unique_b = list(dict.fromkeys(clean_b))

    matched_pairs: List[Dict[str, Any]] = []
    used_b_indices: Set[int] = set()

    # 1. Aşama: Tam normalleştirilmiş eşleşmeler
    norm_b_map: Dict[str, List[int]] = {}
    for idx_b, b in enumerate(unique_b):
        nb = normalize_str(b)
        if nb not in norm_b_map:
            norm_b_map[nb] = []
        norm_b_map[nb].append(idx_b)

    candidates_a = []
    for a in unique_a:
        na = normalize_str(a)
        if na in norm_b_map and norm_b_map[na]:
            matched_b_idx = norm_b_map[na].pop(0)
            used_b_indices.add(matched_b_idx)
            matched_pairs.append({
                "item_a": a,
                "item_b": unique_b[matched_b_idx],
                "similarity": 100.0,
                "match_type": "Tam Eşleşme"
            })
        else:
            candidates_a.append(a)

    # 2. Aşama: Eşik < 1.0 ise Fuzzy Benzerlik Eşlemesi
    assigned_a = set()
    if threshold < 1.0 and candidates_a:
        available_b_indices = [i for i in range(len(unique_b)) if i not in used_b_indices]
        pair_scores = []
        for a in candidates_a:
            for i_b in available_b_indices:
                score = calculate_text_similarity(a, unique_b[i_b])
                if score >= threshold:
                    pair_scores.append((score, a, i_b))

        # En yüksek benzerlikten en düşüğe açgözlü (greedy) eşleme
        pair_scores.sort(key=lambda x: x[0], reverse=True)
        for score, a, i_b in pair_scores:
            if a not in assigned_a and i_b not in used_b_indices:
                assigned_a.add(a)
                used_b_indices.add(i_b)
                pct = round(score * 100, 1)
                matched_pairs.append({
                    "item_a": a,
                    "item_b": unique_b[i_b],
                    "similarity": pct,
                    "match_type": f"Benzer (%{pct})"
                })

    unmatched_a = [a for a in candidates_a if a not in assigned_a]
    unmatched_b = [unique_b[i] for i in range(len(unique_b)) if i not in used_b_indices]

    total_unique = len(matched_pairs) + len(unmatched_a) + len(unmatched_b)
    jaccard = round(len(matched_pairs) / total_unique, 4) if total_unique > 0 else 1.0

    return {
        "intersection_count": len(matched_pairs),
        "only_a_count": len(unmatched_a),
        "only_b_count": len(unmatched_b),
        "total_unique": total_unique,
        "jaccard_similarity": jaccard,
        "matched_pairs": matched_pairs,
        "intersection_items": [f"{p['item_a']} ↔ {p['item_b']} (%{p['similarity']})" if p['item_a'] != p['item_b'] else p['item_a'] for p in matched_pairs],
        "only_a_items": sorted(unmatched_a),
        "only_b_items": sorted(unmatched_b),
        "threshold_used": threshold
    }


def execute_compare_two_excels(
    file_1: Any,
    file_2: Any,
    file_1_name: str = "1. Belge",
    file_2_name: str = "2. Belge",
    column_name: Optional[str] = None,
    col1_name: Optional[str] = None,
    col2_name: Optional[str] = None,
    sheet_name_1: Optional[str] = None,
    sheet_name_2: Optional[str] = None,
    threshold: float = 1.0
) -> Dict[str, Any]:
    """
    İki farklı Excel dosyasını okuyarak uygunsuzluklar arasındaki Venn küme matematiğini
    (yalnızca 1, yalnızca 2, ortak kesişim, Jaccard benzerliği) hesaplar.
    """
    try:
        df1 = read_excel_safe(file_1, sheet_name=sheet_name_1)
        df2 = read_excel_safe(file_2, sheet_name=sheet_name_2)

        col1 = col1_name
        col2 = col2_name

        if not col1 and column_name and str(column_name).strip() and column_name != "auto":
            norm_req = normalize_str(column_name)
            for c in df1.columns:
                if normalize_str(c) == norm_req or norm_req in normalize_str(c):
                    col1 = c
                    break

        if not col2 and column_name and str(column_name).strip() and column_name != "auto":
            norm_req = normalize_str(column_name)
            for c in df2.columns:
                if normalize_str(c) == norm_req or norm_req in normalize_str(c):
                    col2 = c
                    break

        if not col1:
            try:
                _, _, col1 = resolve_columns(df1)
            except Exception:
                col1 = df1.columns[0]

        if not col2:
            try:
                _, _, col2 = resolve_columns(df2)
            except Exception:
                col2 = df2.columns[0]

        items_1_raw = df1[col1].dropna().astype(str).tolist()
        items_2_raw = df2[col2].dropna().astype(str).tolist()

        venn_res = compute_fuzzy_venn(items_1_raw, items_2_raw, threshold=threshold)

        summary_note = (
            f"'{file_1_name}' ({len(set(items_1_raw))} benzersiz kayıt) ile '{file_2_name}' ({len(set(items_2_raw))} benzersiz kayıt) kıyaslandı. "
            f"Toplam {venn_res['total_unique']} farklı uygunsuzluk tespit edildi. "
            f"{venn_res['intersection_count']} uygunsuzluk her iki belgede de ortaktır (Jaccard Benzerlik Oranı: %{round(venn_res['jaccard_similarity'] * 100, 1)}). "
            f"{venn_res['only_a_count']} uygunsuzluk yalnızca 1. belgede, {venn_res['only_b_count']} uygunsuzluk ise yalnızca 2. belgede yer almaktadır."
        )

        return {
            "durum": "basarili",
            "dosya_1_adi": file_1_name,
            "dosya_2_adi": file_2_name,
            "karsilastirilan_sutun_1": col1,
            "karsilastirilan_sutun_2": col2,
            "ozet_yorum": summary_note,
            "venn_verileri": venn_res
        }
    except Exception as e:
        return {
            "durum": "hata",
            "hata_turu": type(e).__name__,
            "mesaj": str(e)
        }


def compare_two_excels_venn(
    file_path_a: str,
    file_path_b: str,
    column_name: Optional[str] = None,
    sheet_name_a: Optional[str] = None,
    sheet_name_b: Optional[str] = None
) -> str:
    """İki farklı Excel dosyasını Venn şeması küme matematiği ile kıyaslayan MCP aracı."""
    res = execute_compare_two_excels(
        file_1=file_path_a,
        file_2=file_path_b,
        file_1_name=os.path.basename(file_path_a),
        file_2_name=os.path.basename(file_path_b),
        column_name=column_name,
        sheet_name_1=sheet_name_a,
        sheet_name_2=sheet_name_b
    )
    return json.dumps(res, ensure_ascii=False, indent=2)


def extract_langflow_response(data: Any) -> str:
    """Langflow / LLM Suite Flow JSON çıktısından mesaj metnini ayıklar."""
    if not isinstance(data, dict):
        return str(data)

    try:
        outputs = data.get("outputs", [])
        if outputs and isinstance(outputs, list):
            first_out = outputs[0]
            sub_outputs = first_out.get("outputs", [])
            if sub_outputs and isinstance(sub_outputs, list):
                first_sub = sub_outputs[0]
                res_dict = first_sub.get("results", {})
                msg_obj = res_dict.get("message", {})
                if isinstance(msg_obj, dict) and "text" in msg_obj:
                    return str(msg_obj["text"])
                if isinstance(msg_obj, str) and msg_obj.strip():
                    return msg_obj

                messages = first_sub.get("messages", [])
                if messages and isinstance(messages, list):
                    last_msg = messages[-1]
                    if isinstance(last_msg, dict):
                        if "message" in last_msg and last_msg["message"]:
                            return str(last_msg["message"])
                        if "text" in last_msg and last_msg["text"]:
                            return str(last_msg["text"])

                artifacts = first_sub.get("artifacts", {})
                if isinstance(artifacts, dict) and "message" in artifacts:
                    return str(artifacts["message"])
    except Exception:
        pass

    for k in ["result", "output", "text", "message", "response"]:
        if k in data and isinstance(data[k], str) and data[k].strip():
            return data[k]

    return json.dumps(data, ensure_ascii=False, indent=2)


def call_aselsan_llm_suite_flow(
    base_url: str = "https://llmsuiteflow.aiplatform.aselsan.com.tr",
    flow_id: str = "c16ecc83-ff64-4fc7-bb52-0f594beaf70a",
    api_key: Optional[str] = None,
    tweaks_input: str = "ChatInput-jugL4",
    uploaded_files: Optional[List[Tuple[str, bytes]]] = None,
    prompt_text: str = "",
    session_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    ASELSAN AI Platform LLM Suite Flow API'sini çağırır.
    Dosyaları upload endpoint'ine yükler ve flow çalıştırma isteği yapar.
    """
    if not base_url:
        base_url = "https://llmsuiteflow.aiplatform.aselsan.com.tr"
    base_url = base_url.rstrip("/")

    if not flow_id:
        flow_id = "c16ecc83-ff64-4fc7-bb52-0f594beaf70a"

    headers = {}
    if api_key and str(api_key).strip():
        headers["x-api-key"] = str(api_key).strip()

    uploaded_file_paths = []

    if headers and uploaded_files:
        upload_url = f"{base_url}/api/v1/files/upload/{flow_id}"
        for fname, fbytes in uploaded_files:
            try:
                files_data = {
                    "file": (
                        fname,
                        io.BytesIO(fbytes),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                }
                up_resp = requests.post(upload_url, headers=headers, files=files_data, timeout=25, verify=False)
                if up_resp.status_code in [200, 201]:
                    up_json = up_resp.json()
                    fpath = up_json.get("file_path") or up_json.get("path") or up_json.get("name")
                    if fpath:
                        uploaded_file_paths.append(fpath)
            except Exception:
                pass

    if not session_id:
        session_id = str(uuid.uuid4())

    payload = {
        "output_type": "chat",
        "input_type": "chat",
        "input_value": prompt_text,
        "session_id": session_id
    }

    # Yalnızca dosya upload başarılı olduysa tweaks ekle, asla input_value'yu boş bırakma!
    if uploaded_file_paths and tweaks_input and str(tweaks_input).strip():
        payload["tweaks"] = {
            tweaks_input.strip(): {
                "files": uploaded_file_paths,
                "input_value": prompt_text
            }
        }

    run_url = f"{base_url}/api/v1/run/{flow_id}?stream=false"
    req_headers = {"Content-Type": "application/json", **headers}

    if not headers:
        current_user = KOKPIT_IDENTITY_STORE.get("current_user")
        is_admin = is_kokpit_user_admin(current_user) if "is_kokpit_user_admin" in globals() else False
        if is_admin:
            return {
                "durum": "bilgi",
                "session_id": session_id,
                "cevap": (
                    "### 📌 Matematiksel Venn Şeması ve Küme Analizi Hazır!\n\n"
                    "İki Excel dosyasındaki uygunsuzluklar başarıyla karşılaştırıldı ve görsel Venn şeması oluşturuldu.\n\n"
                    "**ASELSAN LLM Suite Flow Bağlantısı (Yönetici Bilgisi):**\n"
                    f"- **Hedef Akış:** `{base_url}/api/v1/run/{flow_id}`\n\n"
                    "Akıllı yapay zeka yorumu, kök neden analizi ve risk sınıflandırmasını almak için sağ üst köşedeki "
                    "**⚙️ Ayarlar** butonuna tıklayarak **API Key (x-api-key)** bilginizi kaydedebilirsiniz."
                )
            }
        return {
            "durum": "bilgi",
            "session_id": session_id,
            "cevap": (
                "### 📌 Matematiksel Venn Şeması ve Küme Analizi Hazır!\n\n"
                "İki Excel dosyasındaki uygunsuzluklar başarıyla karşılaştırıldı ve görsel Venn şeması oluşturuldu."
            )
        }

    try:
        resp = requests.post(run_url, headers=req_headers, json=payload, timeout=180, verify=False)
        
        # 400 hatası alındıysa ve tweaks varsa, tweaks olmadan (saf prompt ile) tekrar dene
        if resp.status_code == 400 and "tweaks" in payload:
            clean_payload = {
                "output_type": "chat",
                "input_type": "chat",
                "input_value": prompt_text,
                "session_id": session_id
            }
            resp = requests.post(run_url, headers=req_headers, json=clean_payload, timeout=180, verify=False)

        if resp.status_code >= 400:
            err_detail = ""
            try:
                err_json = resp.json()
                err_detail = err_json.get("detail") or err_json.get("message") or json.dumps(err_json, ensure_ascii=False)
            except Exception:
                err_detail = resp.text
            return {
                "durum": "hata",
                "mesaj": f"ASELSAN LLM Suite Flow hatası ({resp.status_code}): {err_detail}",
                "session_id": session_id
            }

        res_json = resp.json()
        extracted_text = extract_langflow_response(res_json)
        return {
            "durum": "basarili",
            "session_id": session_id,
            "cevap": extracted_text,
            "raw": res_json
        }
    except requests.exceptions.RequestException as req_err:
        err_msg = str(req_err)
        if hasattr(req_err, "response") and req_err.response is not None:
            try:
                err_msg += f" - Detay: {req_err.response.text}"
            except Exception:
                pass
        return {
            "durum": "hata",
            "mesaj": f"ASELSAN LLM Suite Flow bağlantı hatası: {err_msg}",
            "session_id": session_id
        }


def fetch_aselix_models(base_url: str = "https://api.aselixai.aselsan.com.tr/v1", api_key: Optional[str] = None) -> Dict[str, Any]:
    """
    ASELSAN AselixAI REST API üzerinden mevcut modelleri listeler (GET /models).
    Hem Milli Ağ (https://api.aselixai.aselsan.com.tr/v1) hem de Aynet Ağı 
    (https://agate.aselsan.com.tr/send/milli/api.aselixai/v1) ile uyumludur.
    """
    if not base_url:
        base_url = "https://api.aselixai.aselsan.com.tr/v1"
    base_url = base_url.rstrip("/")
    url = f"{base_url}/models"
    headers = {"Content-Type": "application/json"}
    if api_key and api_key.strip():
        headers["Authorization"] = f"Bearer {api_key.strip()}"
        headers["x-api-key"] = api_key.strip()

    fallback_models = [
        "aselix-llm-v1",
        "aselix-chat-default",
        "llama-3-8b-instruct",
        "mistral-7b-instruct",
        "qwen-2.5-7b-instruct"
    ]

    try:
        resp = requests.get(url, headers=headers, timeout=10, verify=False)
        if resp.status_code == 200:
            data = resp.json()
            model_ids = []
            if isinstance(data, dict):
                items = data.get("data") or data.get("models") or []
                if isinstance(items, list):
                    for it in items:
                        if isinstance(it, dict) and "id" in it:
                            model_ids.append(it["id"])
                        elif isinstance(it, str):
                            model_ids.append(it)
            elif isinstance(data, list):
                for it in data:
                    if isinstance(it, dict) and "id" in it:
                        model_ids.append(it["id"])
                    elif isinstance(it, str):
                        model_ids.append(it)

            return {
                "durum": "basarili",
                "kaynak": "canli_api",
                "base_url": base_url,
                "modeller": model_ids if model_ids else fallback_models,
                "mesaj": f"{len(model_ids)} model AselixAI API'den başarıyla listelendi."
            }
        else:
            return {
                "durum": "uyari",
                "hata_kodu": resp.status_code,
                "base_url": base_url,
                "mesaj": f"AselixAI API yanıtı ({resp.status_code}): {resp.text[:200]}",
                "modeller": fallback_models
            }
    except Exception as e:
        return {
            "durum": "uyari",
            "hata_turu": type(e).__name__,
            "base_url": base_url,
            "mesaj": f"API çağrısı uyarısı (İntranet veya anahtar kontrolü gerekebilir): {str(e)}",
            "modeller": fallback_models
        }


def call_aselix_chat(
    base_url: str = "https://api.aselixai.aselsan.com.tr/v1",
    api_key: Optional[str] = None,
    model: str = "aselix-llm-v1",
    messages: Optional[List[Dict[str, str]]] = None,
    temperature: float = 0.2,
    timeout: int = 50
) -> Dict[str, Any]:
    """
    ASELSAN AselixAI POST /chat/completions endpointine OpenAI uyumlu istek gönderir.
    1 dakikalık (varsayılan timeout=50sn) zaman aşımı sınırını kesinlikle aşmaz.
    """
    if not base_url:
        base_url = "https://api.aselixai.aselsan.com.tr/v1"
    base_url = base_url.rstrip("/")
    url = f"{base_url}/chat/completions"

    headers = {"Content-Type": "application/json"}
    if api_key and api_key.strip():
        headers["Authorization"] = f"Bearer {api_key.strip()}"
        headers["x-api-key"] = api_key.strip()

    payload = {
        "model": model or "aselix-llm-v1",
        "messages": messages or [{"role": "user", "content": "Merhaba"}],
        "temperature": temperature
    }

    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=timeout, verify=False)
        if resp.status_code == 200:
            data = resp.json()
            content = ""
            if isinstance(data, dict):
                choices = data.get("choices", [])
                if choices and isinstance(choices, list):
                    first_choice = choices[0]
                    if isinstance(first_choice, dict):
                        msg = first_choice.get("message", {})
                        if isinstance(msg, dict):
                            content = msg.get("content", "")
            if not content:
                content = json.dumps(data, ensure_ascii=False)
            return {
                "durum": "basarili",
                "model": model,
                "cevap": content,
                "raw": data
            }
        else:
            return {
                "durum": "hata",
                "hata_kodu": resp.status_code,
                "mesaj": f"AselixAI API Hatası ({resp.status_code}): {resp.text}"
            }
    except requests.exceptions.Timeout:
        return {
            "durum": "zaman_asimi",
            "hata_kodu": 408,
            "mesaj": f"AselixAI API isteği {timeout} saniye içinde yanıt vermedi (1 dakikalık zaman aşımı koruması devreye girdi)."
        }
    except Exception as e:
        return {
            "durum": "hata",
            "hata_turu": type(e).__name__,
            "mesaj": f"AselixAI API Çağrı Hatası: {str(e)}"
        }


def extract_json_array_from_llm(text: str) -> List[Dict[str, Any]]:
    """LLM metin çıktısından JSON dizisini ayıklar."""
    if not text:
        return []
    cleaned = text.strip()
    if "```json" in cleaned:
        cleaned = cleaned.split("```json")[1].split("```")[0].strip()
    elif "```" in cleaned:
        cleaned = cleaned.split("```")[1].split("```")[0].strip()

    start_idx = cleaned.find("[")
    end_idx = cleaned.rfind("]")
    if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
        cleaned = cleaned[start_idx:end_idx + 1]

    try:
        data = json.loads(cleaned)
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return []


def shortlist_candidates_for_batch(
    items_a_batch: List[str],
    candidates_b: List[str],
    top_k: int = 6,
    max_total: int = 60
) -> List[str]:
    """
    Her batch'e TÜM B listesini göndermek yerine (büyük listelerde prompt devasa olur ve
    50 sn içinde yanıt gelmez), her A maddesi için yerel benzerlikle en yakın top_k B adayını seçer.
    Birleşik aday listesi max_total ile sınırlandırılır -> prompt küçük, yanıt hızlı kalır.
    """
    unique_b = list(dict.fromkeys([clean_text_content(b) for b in candidates_b if clean_text_content(b)]))
    if len(unique_b) <= max_total:
        return unique_b

    scored_union: Dict[str, float] = {}
    for a in items_a_batch:
        scores = [(calculate_text_similarity(a, b), b) for b in unique_b]
        scores.sort(key=lambda x: x[0], reverse=True)
        for score, b in scores[:top_k]:
            if b not in scored_union or scored_union[b] < score:
                scored_union[b] = score

    ranked = sorted(scored_union.items(), key=lambda x: x[1], reverse=True)
    return [b for b, _ in ranked[:max_total]]


def compare_items_with_aselix_batch(
    base_url: str,
    api_key: str,
    model: str,
    items_a_batch: List[str],
    candidates_b: List[str],
    threshold: float = 0.8,
    timeout: int = 50
) -> Dict[str, Any]:
    """
    1 dakikalık (timeout=50sn) zaman aşımı limitini kesinlikle aşmayacak şekilde
    küçük bir batch (10-20 kayıt) için AselixAI API'si üzerinden anlamsal kıyaslama yapar.

    - B listesi her A maddesi için en yakın adaylara daraltılır (shortlist).
    - LLM'den metin yerine NUMARA (indeks) döndürmesi istenir -> eşleşmeler orijinal metne birebir bağlanır.
    - Her yanıtta tanılama (diagnostics) bilgisi döner: süre, kaynak, hata, aday sayısı vb.
    """
    t0 = time.time()
    clean_batch = [clean_text_content(x) for x in items_a_batch if clean_text_content(x)]
    clean_batch = list(dict.fromkeys(clean_batch))

    def _diag(extra: Dict[str, Any]) -> Dict[str, Any]:
        base = {
            "a_sayisi": len(clean_batch),
            "b_aday_sayisi": len(candidates_b),
            "sure_sn": round(time.time() - t0, 2),
        }
        base.update(extra)
        return base

    if not clean_batch or not candidates_b:
        return {
            "durum": "basarili",
            "matched_pairs": [],
            "unmatched_a": clean_batch,
            "kaynak": "bos_liste",
            "tanilama": _diag({})
        }

    # API Key yoksa yerel fuzzy matching kullan
    if not api_key or not api_key.strip():
        local_res = compute_fuzzy_venn(clean_batch, candidates_b, threshold=threshold)
        return {
            "durum": "basarili",
            "matched_pairs": local_res["matched_pairs"],
            "unmatched_a": local_res["only_a_items"],
            "kaynak": "yerel_ml",
            "tanilama": _diag({"not": "API anahtarı yok, yerel ML kullanıldı."})
        }

    shortlist = shortlist_candidates_for_batch(clean_batch, candidates_b)

    list_a_text = "\n".join([f"A{i+1}. {item}" for i, item in enumerate(clean_batch)])
    list_b_text = "\n".join([f"B{j+1}. {cand}" for j, cand in enumerate(shortlist)])

    system_prompt = (
        "Sen ASELSAN Sistem ve Kalite Güvence mühendisliği uzmanı bir analiz asistanısın. "
        "İki listedeki uygunsuzluk maddelerini karşılaştırıp anlamsal olarak AYNI teknik problemi ifade edenleri eşleştirirsin. "
        "Farklı kelimeler/eş anlamlılar kullanılsa da aynı arızaysa eşleştir. "
        "Yanıtını YALNIZCA geçerli bir JSON dizisi olarak üret, açıklama metni ekleme."
    )
    user_prompt = (
        f"Liste A'daki {len(clean_batch)} maddenin her biri için Liste B'de aynı teknik problemi ifade eden madde var mı?\n"
        "Her B maddesi en fazla bir A maddesiyle eşleşebilir.\n\n"
        "ÇIKTI FORMATI (yalnızca bu JSON dizisi, her A maddesi için bir eleman):\n"
        '[{"a": 1, "b": 3, "similarity": 92}, {"a": 2, "b": null, "similarity": 0}]\n'
        "a = A numarası, b = eşleşen B numarası (yoksa null), similarity = 0-100 arası anlamsal benzerlik.\n\n"
        f"--- LİSTE A ---\n{list_a_text}\n\n"
        f"--- LİSTE B ---\n{list_b_text}\n"
    )

    chat_res = call_aselix_chat(
        base_url=base_url,
        api_key=api_key,
        model=model or "aselix-llm-v1",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        temperature=0.1,
        timeout=timeout  # 1 dakikayı kesinlikle aşmayacak (50sn)
    )
    api_sure = round(time.time() - t0, 2)

    if chat_res.get("durum") != "basarili":
        # API zaman aşımı veya bağlantı hatası -> bu batch için yerel fuzzy distance ile kurtar
        local_res = compute_fuzzy_venn(clean_batch, candidates_b, threshold=threshold)
        return {
            "durum": "uyari",
            "api_hatasi": chat_res.get("mesaj"),
            "matched_pairs": local_res["matched_pairs"],
            "unmatched_a": local_res["only_a_items"],
            "kaynak": "yerel_fallback",
            "tanilama": _diag({
                "shortlist_sayisi": len(shortlist),
                "prompt_karakter": len(user_prompt),
                "api_sure_sn": api_sure,
                "api_durum": chat_res.get("durum"),
                "hata_kodu": chat_res.get("hata_kodu")
            })
        }

    raw_output = chat_res.get("cevap", "") or ""
    parsed = extract_json_array_from_llm(raw_output)
    # Geçerli ama boş JSON dizisi ("[]") = AI hiç eşleşme bulmadı; bu bir hata değildir
    compact = raw_output.replace(" ", "").replace("\n", "").replace("\r", "")
    parsed_ok = bool(parsed) or "[]" in compact

    matched_pairs: List[Dict[str, Any]] = []
    matched_a_idx: Set[int] = set()
    matched_b_idx: Set[int] = set()
    threshold_pct = threshold * 100

    def _to_index(val: Any, limit: int) -> Optional[int]:
        if val is None:
            return None
        s = str(val).strip().upper().lstrip("AB").rstrip(".")
        if not s.isdigit():
            return None
        idx = int(s) - 1
        return idx if 0 <= idx < limit else None

    for m in parsed:
        if not isinstance(m, dict):
            continue
        ia = _to_index(m.get("a"), len(clean_batch))
        ib = _to_index(m.get("b"), len(shortlist))
        try:
            sim = float(m.get("similarity") or 0.0)
        except (TypeError, ValueError):
            sim = 0.0
        if ia is None or ib is None or sim < threshold_pct:
            continue
        if ia in matched_a_idx or ib in matched_b_idx:
            continue
        matched_a_idx.add(ia)
        matched_b_idx.add(ib)
        matched_pairs.append({
            "item_a": clean_batch[ia],
            "item_b": shortlist[ib],
            "similarity": round(sim, 1),
            "match_type": f"AI Semantik (%{round(sim, 1)})"
        })

    # LLM'in eşleştiremediği / JSON'da eksik kalanlar için yerel fuzzy kontrolü
    unmatched_a = [a for i, a in enumerate(clean_batch) if i not in matched_a_idx]
    remaining_b = [b for j, b in enumerate(shortlist) if j not in matched_b_idx]
    fuzzy_added = 0
    if unmatched_a and remaining_b:
        local_fallback = compute_fuzzy_venn(unmatched_a, remaining_b, threshold=threshold)
        for pair in local_fallback["matched_pairs"]:
            matched_pairs.append(pair)
            fuzzy_added += 1
        unmatched_a = local_fallback["only_a_items"]

    return {
        "durum": "basarili" if parsed_ok else "uyari",
        "api_hatasi": None if parsed_ok else "AselixAI yanıtı JSON olarak ayrıştırılamadı; bu batch yerel ML ile tamamlandı.",
        "matched_pairs": matched_pairs,
        "unmatched_a": unmatched_a,
        "kaynak": "aselix_ai" if parsed_ok else "aselix_parse_hatasi",
        "tanilama": _diag({
            "shortlist_sayisi": len(shortlist),
            "prompt_karakter": len(user_prompt),
            "api_sure_sn": api_sure,
            "llm_eslesme": len(matched_pairs) - fuzzy_added,
            "fuzzy_ek_eslesme": fuzzy_added,
            "llm_yanit_onizleme": raw_output[:300]
        })
    }


# ==============================================================================
# KOKPİT KİMLİK AKTARIMI (SSO / IFRAME PUSH) & EMBEDDING ALTYAPISI
# ==============================================================================

import base64
import math
from collections import Counter
from datetime import datetime, timezone

ALLOWED_KOKPIT_ORIGINS = [
    "http://localhost:3004",
    "https://localhost:3004",
    "http://127.0.0.1:3004",
    "https://kokpit.sirketiniz.com",
    "https://kokpit.aselsan.com.tr",
    "http://kokpit.aselsan.com.tr"
]

# Bellekte aktif oturum ve dışarıdan gelen input deposu
KOKPIT_IDENTITY_STORE: Dict[str, Any] = {
    "current_user": None,
    "jwt_token": None,
    "token_claims": None,
    "source": None,
    "origin": None,
    "status": "waiting",
    "timestamp": None,
    "updated_at": None,
    "session_id": None,
    "external_inputs": {},
    "sessions_by_sicil": {},
    "history": []
}


def decode_jwt_payload_unverified(token: Optional[str]) -> Dict[str, Any]:
    """
    Harici bir Keycloak veya LDAP bağlantısı kurmadan (Rehber Bölüm 1 gereği)
    JWT Bearer Token'ın payload (orta) kısmını güvenli biçimde çözümler.
    """
    if not token or not isinstance(token, str):
        return {}
    clean_tok = token.strip()
    if clean_tok.lower().startswith("bearer "):
        clean_tok = clean_tok[7:].strip()
    parts = clean_tok.split(".")
    if len(parts) < 2:
        return {}
    try:
        payload_b64 = parts[1]
        padding = "=" * ((4 - len(payload_b64) % 4) % 4)
        decoded_bytes = base64.urlsafe_b64decode(payload_b64 + padding)
        claims = json.loads(decoded_bytes.decode("utf-8", errors="replace"))
        return claims if isinstance(claims, dict) else {}
    except Exception:
        return {}


def is_kokpit_user_admin(user: Optional[Dict[str, Any]]) -> bool:
    """
    Kullanıcının Admin yetkisine sahip olup olmadığını kontrol eder.
    - user.is_admin / user.isAdmin boolean alanı
    - user.role veya user.roles içerisinde Admin / Yönetici rolü
    - .env içindeki ADMIN_ROLES veya ADMIN_USERS (sicil / username) eşleşmesi
    """
    if not isinstance(user, dict) or not user:
        return False

    if user.get("is_admin") is True or user.get("isAdmin") is True:
        return True

    raw_admin_roles = os.environ.get(
        "ADMIN_ROLES",
        "admin,yönetici,yonetici,superuser,system_admin,sistem yöneticisi"
    )
    admin_roles = {r.strip().lower() for r in raw_admin_roles.split(",") if r.strip()}

    role_val = str(user.get("role") or user.get("rol") or "").strip().lower()
    if role_val and role_val in admin_roles:
        return True

    roles_list = user.get("roles")
    if isinstance(roles_list, list):
        for r in roles_list:
            if str(r or "").strip().lower() in admin_roles:
                return True

    raw_admin_users = os.environ.get("ADMIN_USERS", "")
    admin_users = {u.strip().lower() for u in raw_admin_users.split(",") if u.strip()}
    if admin_users:
        sicil = str(user.get("sicil") or "").strip().lower()
        username = str(user.get("username") or "").strip().lower()
        if (sicil and sicil in admin_users) or (username and username in admin_users):
            return True

    return False


def normalize_kokpit_user(raw_user: Optional[Dict[str, Any]], token: Optional[str] = None) -> Dict[str, Any]:
    """
    Rehber Bölüm 4'teki kullanıcı şemasını doğrular, eksik alanları varsa JWT içerisinden tamamlar
    ve dışarıdan gelen tüm ekstra kullanıcı özelliklerini korur.
    """
    claims = decode_jwt_payload_unverified(token)
    u = dict(raw_user) if isinstance(raw_user, dict) else {}

    sicil = str(
        u.get("sicil")
        or u.get("employeeNumber")
        or u.get("sicil_no")
        or claims.get("sicil")
        or claims.get("employee_number")
        or ""
    ).strip()

    username = str(
        u.get("username")
        or u.get("preferred_username")
        or claims.get("preferred_username")
        or claims.get("sub")
        or ""
    ).strip()

    isim = str(
        u.get("isim")
        or u.get("name")
        or u.get("fullName")
        or u.get("ad_soyad")
        or claims.get("name")
        or username
        or ""
    ).strip()

    mail = str(
        u.get("mail")
        or u.get("email")
        or claims.get("email")
        or ""
    ).strip()

    role = str(
        u.get("role")
        or u.get("rol")
        or (claims.get("realm_access", {}).get("roles", ["User"])[0] if isinstance(claims.get("realm_access"), dict) and claims.get("realm_access", {}).get("roles") else "User")
    ).strip()

    normalized = {
        **u,
        "id": u.get("id") if u.get("id") is not None else claims.get("id"),
        "sicil": sicil,
        "username": username,
        "isim": isim,
        "mail": mail,
        "picture": u.get("picture") or claims.get("picture") or (f"/uploads/photos/{sicil}.jpg" if sicil else ""),
        "role": role or "User",
        "mudurluk": u.get("mudurluk") or u.get("department") or claims.get("mudurluk") or "",
        "mudurluk_id": u.get("mudurluk_id") or claims.get("mudurluk_id"),
        "direktorluk": u.get("direktorluk") or claims.get("direktorluk") or "",
        "birim": u.get("birim") or claims.get("birim") or "",
        "unvan": u.get("unvan") or u.get("title") or claims.get("unvan") or "",
        "hierarchicLevel": u.get("hierarchicLevel") if u.get("hierarchicLevel") is not None else claims.get("hierarchicLevel")
    }
    normalized["is_admin"] = is_kokpit_user_admin(normalized)
    return normalized


def store_kokpit_identity_and_inputs(
    payload: Dict[str, Any],
    origin: Optional[str] = None
) -> Dict[str, Any]:
    """
    Hem Rehber Bölüm 4/5.C'deki `kokpit-identity` paketini hem de dışarıdan gelen
    ekstra input/config/embedding parametrelerini işler ve oturuma kaydeder.
    """
    raw_user = payload.get("user")
    token = payload.get("token") or payload.get("auth_token") or payload.get("jwt_token")
    source = payload.get("source") or "kokpit-iframe-bridge"
    msg_type = payload.get("type") or "kokpit-identity"
    status = payload.get("status") or "success"
    timestamp = payload.get("timestamp") or datetime.now(timezone.utc).isoformat()

    # Dışarıdan gelen ek input alanlarını topla
    ext_inputs: Dict[str, Any] = {}
    for key in ["inputs", "config", "data", "external_inputs", "embedding", "comparison"]:
        val = payload.get(key)
        if isinstance(val, dict):
            ext_inputs.update(val)

    # Üst seviyede doğrudan gönderilmiş olabilecek dış parametreleri de al
    direct_input_keys = [
        "api_key", "apiKey", "base_url", "baseUrl", "model", "embedding_model", "embeddingModel",
        "engine_mode", "engine", "threshold", "batch_size", "batchSize", "input_mode", "inputMode",
        "items_1", "items_2", "items_a", "items_b", "text_1", "text_2", "text1", "text2",
        "name_1", "name_2", "doc1_name", "doc2_name", "embeddings_1", "embeddings_2",
        "embeddings_a", "embeddings_b", "auto_run", "autoRun", "theme", "flow_id", "network"
    ]
    for k in direct_input_keys:
        if k in payload and payload[k] is not None:
            ext_inputs[k] = payload[k]

    user = normalize_kokpit_user(raw_user, token) if (raw_user or token) else None
    claims = decode_jwt_payload_unverified(token)
    now_iso = datetime.now(timezone.utc).isoformat()
    session_id = payload.get("session_id") or KOKPIT_IDENTITY_STORE.get("session_id") or str(uuid.uuid4())

    if user and (user.get("sicil") or user.get("username") or user.get("isim") or user.get("id") is not None):
        KOKPIT_IDENTITY_STORE["current_user"] = user
        if user.get("sicil"):
            KOKPIT_IDENTITY_STORE["sessions_by_sicil"][str(user["sicil"])] = {
                "user": user,
                "token": token,
                "updated_at": now_iso,
                "session_id": session_id
            }

    if token:
        KOKPIT_IDENTITY_STORE["jwt_token"] = token
        KOKPIT_IDENTITY_STORE["token_claims"] = claims

    if ext_inputs:
        merged_inputs = dict(KOKPIT_IDENTITY_STORE.get("external_inputs") or {})
        merged_inputs.update(ext_inputs)
        KOKPIT_IDENTITY_STORE["external_inputs"] = merged_inputs

    KOKPIT_IDENTITY_STORE["source"] = source
    KOKPIT_IDENTITY_STORE["type"] = msg_type
    KOKPIT_IDENTITY_STORE["origin"] = origin or payload.get("origin")
    KOKPIT_IDENTITY_STORE["status"] = status
    KOKPIT_IDENTITY_STORE["timestamp"] = timestamp
    KOKPIT_IDENTITY_STORE["updated_at"] = now_iso
    KOKPIT_IDENTITY_STORE["session_id"] = session_id

    history_entry = {
        "timestamp": now_iso,
        "type": msg_type,
        "source": source,
        "sicil": user.get("sicil") if user else None,
        "isim": user.get("isim") if user else None,
        "has_token": bool(token),
        "input_keys": list(ext_inputs.keys())
    }
    KOKPIT_IDENTITY_STORE["history"] = (KOKPIT_IDENTITY_STORE["history"] + [history_entry])[-25:]

    return {
        "session_id": session_id,
        "user": KOKPIT_IDENTITY_STORE["current_user"],
        "token_present": bool(KOKPIT_IDENTITY_STORE["jwt_token"]),
        "token_claims": KOKPIT_IDENTITY_STORE["token_claims"],
        "source": source,
        "timestamp": timestamp,
        "updated_at": now_iso,
        "external_inputs": KOKPIT_IDENTITY_STORE["external_inputs"]
    }


def compute_local_ngram_embedding(text: str) -> Dict[str, float]:
    """
    Çevrimdışı ortamda metin için karakter 3-gram ve kelime unigram tabanlı
    L2 normalize edilmiş seyrek vektör (embedding) üretir.
    """
    norm = normalize_str(clean_text_content(text))
    if not norm:
        return {}
    features: Counter = Counter()
    words = norm.split()
    for w in words:
        features[f"w:{w}"] += 2.0
    padded = f" {norm} "
    for i in range(len(padded) - 2):
        gram = padded[i:i + 3]
        features[f"c:{gram}"] += 1.0
    norm_sq = sum(v * v for v in features.values())
    if norm_sq <= 0:
        return {}
    inv_norm = 1.0 / math.sqrt(norm_sq)
    return {k: v * inv_norm for k, v in features.items()}


def cosine_similarity_vectors(vec_a: Any, vec_b: Any) -> float:
    """
    Hem yoğun (dense List[float]) hem de seyrek (sparse Dict[str, float])
    embedding vektörleri arasında kosinüs benzerliğini (0.0 - 1.0) hesaplar.
    """
    if isinstance(vec_a, dict) and isinstance(vec_b, dict):
        if not vec_a or not vec_b:
            return 0.0
        if len(vec_a) > len(vec_b):
            vec_a, vec_b = vec_b, vec_a
        dot = sum(val * vec_b.get(k, 0.0) for k, val in vec_a.items())
        return round(min(max(dot, 0.0), 1.0), 4)

    if isinstance(vec_a, (list, tuple)) and isinstance(vec_b, (list, tuple)):
        if not vec_a or not vec_b or len(vec_a) != len(vec_b):
            return 0.0
        dot = sum(float(a) * float(b) for a, b in zip(vec_a, vec_b))
        na = math.sqrt(sum(float(a) * float(a) for a in vec_a))
        nb = math.sqrt(sum(float(b) * float(b) for b in vec_b))
        if na <= 0 or nb <= 0:
            return 0.0
        return round(min(max(dot / (na * nb), 0.0), 1.0), 4)

    return 0.0


def call_aselix_embeddings(
     texts: List[str],
    base_url: str = "https://api.aselixai.aselsan.com.tr/v1",
    api_key: Optional[str] = None,
    model: str = "text-embedding-3-small",
    jwt_token: Optional[str] = None,
    timeout: int = 45
) -> Dict[str, Any]:
    """
    AselixAI / OpenAI uyumlu `POST /embeddings` servisini çağırarak
    metin listesi için vektör embedding'lerini döndürür.
    Eğer API anahtarı/token yoksa veya çevrimdışıysa yerel n-gram vektör embedding üretir.
    """
    clean_texts = [clean_text_content(t) for t in texts if clean_text_content(t)]
    if not clean_texts:
        return {"durum": "basarili", "kaynak": "bos", "embeddings": [], "model": model}

    auth_token = (api_key or "").strip() or (jwt_token or "").strip() or (KOKPIT_IDENTITY_STORE.get("jwt_token") or "").strip()
    if auth_token and base_url:
        url = f"{base_url.rstrip('/')}/embeddings"
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {auth_token}",
            "x-api-key": auth_token
        }
        payload = {
            "model": model or "text-embedding-3-small",
            "input": clean_texts
        }
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=timeout, verify=False)
            if resp.status_code == 200:
                data = resp.json()
                emb_list = []
                if isinstance(data, dict) and isinstance(data.get("data"), list):
                    sorted_items = sorted(data["data"], key=lambda x: x.get("index", 0) if isinstance(x, dict) else 0)
                    for item in sorted_items:
                        if isinstance(item, dict) and isinstance(item.get("embedding"), list):
                            emb_list.append(item["embedding"])
                if len(emb_list) == len(clean_texts):
                    return {
                        "durum": "basarili",
                        "kaynak": "aselix_embedding_api",
                        "model": model,
                        "boyut": len(emb_list[0]) if emb_list else 0,
                        "embeddings": emb_list
                    }
        except Exception as e:
            _safe_log(f"[EMBEDDING API FALLBACK] {str(e)}")

    # Çevrimdışı / Fallback: Yerel N-Gram Embedding Vektörleri
    local_embs = [compute_local_ngram_embedding(t) for t in clean_texts]
    return {
        "durum": "basarili",
        "kaynak": "yerel_ngram_embedding",
        "model": "local-char-word-ngram-v1",
        "embeddings": local_embs
    }


def compute_embedding_venn(
    items_a: List[str],
    items_b: List[str],
    threshold: float = 0.80,
    base_url: str = "https://api.aselixai.aselsan.com.tr/v1",
    api_key: Optional[str] = None,
    embedding_model: str = "text-embedding-3-small",
    external_embeddings_a: Optional[List[List[float]]] = None,
    external_embeddings_b: Optional[List[List[float]]] = None
) -> Dict[str, Any]:
    """
    İki metin listesini vektör embedding (dışarıdan gelen hazır vektörler, AselixAI Embedding API
    veya yerel N-gram vektör uzayı) kullanarak kosinüs benzerliği + hibrit metin benzerliği ile karşılaştırır.
    """
    clean_a = [clean_text_content(x) for x in items_a if clean_text_content(x)]
    clean_b = [clean_text_content(x) for x in items_b if clean_text_content(x)]
    unique_a = list(dict.fromkeys(clean_a))
    unique_b = list(dict.fromkeys(clean_b))

    if not unique_a or not unique_b:
        return compute_fuzzy_venn(unique_a, unique_b, threshold=threshold)

    # Dışarıdan hazır vektör embedding gelmiş mi kontrol et
    if (
        isinstance(external_embeddings_a, list)
        and isinstance(external_embeddings_b, list)
        and len(external_embeddings_a) == len(unique_a)
        and len(external_embeddings_b) == len(unique_b)
    ):
        vecs_a = external_embeddings_a
        vecs_b = external_embeddings_b
        emb_source = "dis_input_hazir_embedding"
    else:
        res_a = call_aselix_embeddings(unique_a, base_url=base_url, api_key=api_key, model=embedding_model)
        res_b = call_aselix_embeddings(unique_b, base_url=base_url, api_key=api_key, model=embedding_model)
        vecs_a = res_a.get("embeddings", [])
        vecs_b = res_b.get("embeddings", [])
        emb_source = res_a.get("kaynak", "yerel_ngram_embedding")

    matched_pairs: List[Dict[str, Any]] = []
    used_b_indices: Set[int] = set()
    assigned_a_indices: Set[int] = set()

    # 1. Tam eşleşmeler
    norm_b_map: Dict[str, List[int]] = {}
    for idx_b, b in enumerate(unique_b):
        nb = normalize_str(b)
        norm_b_map.setdefault(nb, []).append(idx_b)

    for idx_a, a in enumerate(unique_a):
        na = normalize_str(a)
        if na in norm_b_map and norm_b_map[na]:
            ib = norm_b_map[na].pop(0)
            used_b_indices.add(ib)
            assigned_a_indices.add(idx_a)
            matched_pairs.append({
                "item_a": a,
                "item_b": unique_b[ib],
                "similarity": 100.0,
                "match_type": "Tam Eşleşme (Embedding)"
            })

    # 2. Vektör Embedding Kosinüs Benzerliği
    pair_scores = []
    for idx_a, a in enumerate(unique_a):
        if idx_a in assigned_a_indices:
            continue
        va = vecs_a[idx_a] if idx_a < len(vecs_a) else compute_local_ngram_embedding(a)
        for idx_b, b in enumerate(unique_b):
            if idx_b in used_b_indices:
                continue
            vb = vecs_b[idx_b] if idx_b < len(vecs_b) else compute_local_ngram_embedding(b)
            cos_sim = cosine_similarity_vectors(va, vb)
            lex_sim = calculate_text_similarity(a, b)
            score = max(cos_sim, cos_sim * 0.65 + lex_sim * 0.35)
            if score >= threshold:
                pair_scores.append((score, idx_a, idx_b))

    pair_scores.sort(key=lambda x: x[0], reverse=True)
    for score, idx_a, idx_b in pair_scores:
        if idx_a not in assigned_a_indices and idx_b not in used_b_indices:
            assigned_a_indices.add(idx_a)
            used_b_indices.add(idx_b)
            pct = round(score * 100, 1)
            matched_pairs.append({
                "item_a": unique_a[idx_a],
                "item_b": unique_b[idx_b],
                "similarity": pct,
                "match_type": f"Embedding Vektör (%{pct})"
            })

    unmatched_a = [unique_a[i] for i in range(len(unique_a)) if i not in assigned_a_indices]
    unmatched_b = [unique_b[j] for j in range(len(unique_b)) if j not in used_b_indices]
    total_unique = len(matched_pairs) + len(unmatched_a) + len(unmatched_b)
    jaccard = round(len(matched_pairs) / total_unique, 4) if total_unique > 0 else 1.0

    return {
        "intersection_count": len(matched_pairs),
        "only_a_count": len(unmatched_a),
        "only_b_count": len(unmatched_b),
        "total_unique": total_unique,
        "jaccard_similarity": jaccard,
        "matched_pairs": matched_pairs,
        "intersection_items": [
            f"{p['item_a']} ↔ {p['item_b']} (%{p['similarity']})" if p["item_a"] != p["item_b"] else p["item_a"]
            for p in matched_pairs
        ],
        "only_a_items": sorted(unmatched_a),
        "only_b_items": sorted(unmatched_b),
        "threshold_used": threshold,
        "embedding_source": emb_source,
        "embedding_model": embedding_model
    }


# ==============================================================================
# UGESUDKOKPİT VERİTABANI (.env YAPILANDIRMASI & OTOMATİK TABLO/ŞEMA DENETİMİ)
# ==============================================================================

ENV_FILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")


def load_env_file_vars(env_path: Optional[str] = None, override: bool = False) -> Dict[str, str]:
    """
    Proje dizinindeki `.env` dosyasını okur ve hem sözlük olarak döndürür hem de
    `os.environ` içerisine aktarır. Harici paket (`python-dotenv`) olmasa bile çalışır.
    `override=False` olduğunda mevcut ortam değişkenleri ezilmez.
    """
    target_path = env_path or ENV_FILE_PATH
    loaded: Dict[str, str] = {}
    if not os.path.exists(target_path):
        return loaded
    try:
        with open(target_path, "r", encoding="utf-8") as f:
            for raw_line in f:
                line = raw_line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                key = k.strip()
                val = v.strip().strip('"').strip("'")
                if key:
                    loaded[key] = val
                    if override or key not in os.environ:
                        os.environ[key] = val
    except Exception as e:
        _safe_log(f"[ENV-WARN] .env okunurken uyarı: {e}")
    return loaded


# Modül yüklendiğinde .env değişkenlerini derhal os.environ'a aktar
load_env_file_vars()


def get_server_config() -> Dict[str, Any]:
    """
    `.env` dosyasından ve ortam değişkenlerinden HTTP sunucu host ve port bilgilerini
    dinamik olarak alır. (APP_PORT / PORT, APP_HOST / HOST)
    """
    load_env_file_vars()
    host = (os.environ.get("APP_HOST") or os.environ.get("HOST") or "0.0.0.0").strip()
    raw_port = (os.environ.get("APP_PORT") or os.environ.get("PORT") or "8000").strip()
    try:
        port = int(raw_port)
    except Exception:
        port = 8000
    return {
        "host": host,
        "port": port,
        "app_host": host,
        "app_port": port
    }


def get_db_config() -> Dict[str, Any]:
    """
    `.env` dosyasından ve ortam değişkenlerinden PostgreSQL bağlantı bilgilerini
    (host, port, dbname, user, password, schema, tablolar) dinamik olarak alır.
    """
    load_env_file_vars()
    host = os.environ.get("DB_HOST", "localhost").strip()
    port = os.environ.get("DB_PORT", "5432").strip()
    dbname = os.environ.get("DB_NAME", "ud_kokpit").strip()
    user = os.environ.get("DB_USER", "postgres").strip()
    password = os.environ.get("DB_PASSWORD", "563563").strip()
    timeout = int(os.environ.get("DB_CONNECT_TIMEOUT", "4") or 4)
    schema = os.environ.get("DB_SCHEMA", "teslimat_md").strip()
    bakiye_table = os.environ.get("DB_BAKIYE_TABLE", "bakiye_siparis").strip()
    kpi_table = os.environ.get("DB_KPI_TABLE", "dryrun_fkt_skt_kpi").strip()

    explicit_dsn = os.environ.get("UD_KOKPIT_DB_DSN", "").strip()
    dsn = explicit_dsn if explicit_dsn else (
        f"host={host} port={port} dbname={dbname} user={user} password={password}"
    )
    return {
        "host": host,
        "port": port,
        "dbname": dbname,
        "user": user,
        "password": password,
        "timeout": timeout,
        "schema": schema,
        "bakiye_table": bakiye_table,
        "kpi_table": kpi_table,
        "dsn": dsn,
        "env_file": ENV_FILE_PATH,
        "env_file_exists": os.path.exists(ENV_FILE_PATH)
    }


DEFAULT_UD_KOKPIT_DSN = get_db_config()["dsn"]


def fetch_bakiye_siparis_items(
    proje: Optional[str] = None,
    search: Optional[str] = None
) -> Dict[str, Any]:
    """
    DBeaver / UGESUDKOKPİT'in kullandığı PostgreSQL `ud_kokpit` veritabanındaki
    `teslimat_md.bakiye_siparis` tablosundan proje isimleri, SD (satis_belgesi),
    Kalem (kalem_numarasi) ve Ürün Kısa Metni (urun_kisa_metni) kayıtlarını çeker.
    Ayrıca `teslimat_md.dryrun_fkt_skt_kpi` tablosunun varlığını ve şema güncelliğini kontrol eder.
    Veritabanına erişilemezse yerel `Bakiye_Siparis_Raporu.xlsx` yedeğini kullanır.
    """
    db_cfg = get_db_config()
    records: List[Dict[str, Any]] = []
    kaynak = f"postgresql ({db_cfg['dbname']} -> {db_cfg['schema']}.{db_cfg['bakiye_table']})"
    db_err: Optional[str] = None
    kpi_schema_status: Optional[Dict[str, Any]] = None

    try:
        import psycopg2
        conn = psycopg2.connect(db_cfg["dsn"], connect_timeout=db_cfg["timeout"])
        try:
            # Site yüklendiğinde / veri çekildiğinde KPI tablosunun varlığını ve sütunlarını da denetle
            kpi_schema_status = ensure_dryrun_fkt_skt_kpi_table(conn)
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT DISTINCT ON (proje__faaliyet__ic_siparis, satis_belgesi, kalem_numarasi, urun_kisa_metni)
                        id,
                        COALESCE(proje__faaliyet__ic_siparis, '') AS proje_adi,
                        COALESCE(ana_proje, '') AS ana_proje,
                        COALESCE(uges_proje_tanim, '') AS uges_proje_tanim,
                        COALESCE(CAST(satis_belgesi AS TEXT), '') AS sd,
                        COALESCE(CAST(kalem_numarasi AS TEXT), '') AS kalem,
                        COALESCE(malzeme, '') AS malzeme,
                        COALESCE(urun_kisa_metni, tanim, '') AS urun_kisa_metni,
                        COALESCE(ilgili_bolum, '') AS ilgili_bolum,
                        COALESCE(direktorluk, '') AS direktorluk,
                        COALESCE(sorumlu, '') AS sorumlu,
                        COALESCE(musteri, '') AS musteri
                    FROM teslimat_md.bakiye_siparis
                    ORDER BY proje__faaliyet__ic_siparis, satis_belgesi, kalem_numarasi, urun_kisa_metni
                    """
                )
                cols = [d[0] for d in cur.description]
                for row in cur.fetchall():
                    r = dict(zip(cols, row))
                    p_adi = str(r.get("proje_adi") or "").strip()
                    sd_val = str(r.get("sd") or "").strip()
                    kalem_val = str(r.get("kalem") or "").strip()
                    urun_val = str(r.get("urun_kisa_metni") or "").strip()
                    malz_val = str(r.get("malzeme") or "").strip()
                    uges_tanim = str(r.get("uges_proje_tanim") or "").strip()

                    records.append({
                        "id": str(r.get("id") or f"{sd_val}_{kalem_val}"),
                        "proje_adi": p_adi,
                        "ana_proje": str(r.get("ana_proje") or "").strip(),
                        "uges_proje_tanim": uges_tanim,
                        "sd": sd_val,
                        "kalem": kalem_val,
                        "sd_kalem": f"{sd_val} / {kalem_val}" if sd_val or kalem_val else "-",
                        "malzeme": malz_val,
                        "urun_kisa_metni": urun_val,
                        "ilgili_bolum": str(r.get("ilgili_bolum") or "").strip(),
                        "direktorluk": str(r.get("direktorluk") or "").strip(),
                        "sorumlu": str(r.get("sorumlu") or "").strip(),
                        "musteri": str(r.get("musteri") or "").strip(),
                        "etiket": (
                            f"{p_adi}"
                            + (f" ({uges_tanim})" if uges_tanim else "")
                            + f" • SD: {sd_val} / Kalem: {kalem_val} • {urun_val}"
                            + (f" [{malz_val}]" if malz_val and malz_val != urun_val else "")
                        )
                    })
        finally:
            conn.close()
    except Exception as e:
        db_err = str(e)
        kaynak = "excel_fallback (Bakiye_Siparis_Raporu.xlsx)"
        excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Bakiye_Siparis_Raporu.xlsx")
        if os.path.exists(excel_path):
            try:
                bdf = pd.read_excel(excel_path)
                for idx, row in bdf.iterrows():
                    p_adi = clean_text_content(row.get("Proje  Faaliyet  Ic Siparis", ""))
                    sd_val = clean_text_content(row.get("Satis Belgesi", ""))
                    kalem_val = clean_text_content(row.get("Kalem Numarasi", ""))
                    urun_val = clean_text_content(row.get("Urun Kisa Metni", ""))
                    malz_val = clean_text_content(row.get("Malzeme", ""))
                    records.append({
                        "id": f"{sd_val}_{kalem_val}_{idx}",
                        "proje_adi": p_adi,
                        "ana_proje": "",
                        "uges_proje_tanim": "",
                        "sd": sd_val,
                        "kalem": kalem_val,
                        "sd_kalem": f"{sd_val} / {kalem_val}",
                        "malzeme": malz_val,
                        "urun_kisa_metni": urun_val,
                        "ilgili_bolum": clean_text_content(row.get("Ilgili Bolum", "")),
                        "direktorluk": clean_text_content(row.get("Direktorluk", "")),
                        "sorumlu": "",
                        "musteri": "",
                        "etiket": f"{p_adi} • SD: {sd_val} / Kalem: {kalem_val} • {urun_val}"
                    })
            except Exception:
                pass

    projeler = sorted(list({r["proje_adi"] for r in records if r.get("proje_adi")}))

    filtered = records
    if proje and proje.strip() and proje.strip().lower() != "all":
        p_req = proje.strip().lower()
        filtered = [r for r in filtered if r.get("proje_adi", "").lower() == p_req]

    if search and search.strip():
        q = normalize_str(search.strip())
        filtered = [
            r for r in filtered
            if q in normalize_str(r.get("etiket", ""))
            or q in normalize_str(r.get("sd", ""))
            or q in normalize_str(r.get("kalem", ""))
            or q in normalize_str(r.get("urun_kisa_metni", ""))
            or q in normalize_str(r.get("proje_adi", ""))
        ]

    return {
        "durum": "basarili",
        "kaynak": kaynak,
        "uyari": db_err,
        "kpi_tablosu_durumu": kpi_schema_status,
        "toplam_kayit": len(filtered),
        "projeler": projeler,
        "faz_secenekleri": ["Dry-Run", "FKT", "SKT"],
        "kayitlar": filtered
    }


def parse_bakiye_meta(raw_meta: Any) -> Dict[str, Any]:
    """Frontend veya API'den gelen bakiye_siparis proje/SD/kalem/ürün/faz seçim bilgisini ayrıştırır."""
    if not raw_meta:
        return {}
    if isinstance(raw_meta, dict):
        return raw_meta
    if isinstance(raw_meta, str):
        s = raw_meta.strip()
        if s.startswith("{"):
            try:
                parsed = json.loads(s)
                if isinstance(parsed, dict):
                    return parsed
            except Exception:
                pass
    return {}


def format_bakiye_meta_summary(meta: Dict[str, Any], default_doc_name: str = "") -> str:
    """Seçilen Proje, SD, Kalem, Ürün Kısa Metni ve Faz (Dry-Run/FKT/SKT) bilgisini okunaklı metne dönüştürür."""
    if not meta or not isinstance(meta, dict):
        return default_doc_name
    parts: List[str] = []
    proje = str(meta.get("proje_adi") or "").strip()
    uges = str(meta.get("uges_proje_tanim") or "").strip()
    sd = str(meta.get("sd") or "").strip()
    kalem = str(meta.get("kalem") or "").strip()
    urun = str(meta.get("urun_kisa_metni") or "").strip()
    faz = str(meta.get("faz") or "").strip()

    if proje:
        parts.append(f"Proje: {proje}" + (f" ({uges})" if uges else ""))
    if sd or kalem:
        parts.append(f"SD/Kalem: {sd or '-'} / {kalem or '-'}")
    if urun:
        parts.append(f"Ürün: {urun}")
    if faz:
        parts.append(f"Faz: {faz}")

    if not parts:
        return default_doc_name
    if default_doc_name:
        return f"{default_doc_name} [{' • '.join(parts)}]"
    return " • ".join(parts)


# Yerel yedek bellek (PostgreSQL geçici olarak erişilemezse test/çevrimdışı ortam için)
LOCAL_KPI_FALLBACK_STORE: List[Dict[str, Any]] = []


def ensure_dryrun_fkt_skt_kpi_table(conn) -> Dict[str, Any]:
    """
    `ud_kokpit` veritabanında `teslimat_md.dryrun_fkt_skt_kpi` tablosunun varlığını
    ve sütun değişikliklerini otomatik kontrol eder:
    1. Şema (`teslimat_md`) ve tablo (`dryrun_fkt_skt_kpi`) yoksa oluşturur.
    2. Tablo zaten varsa `information_schema.columns` üzerinden eksik sütun olup olmadığını
       denetler ve eksik sütunları `ALTER TABLE ... ADD COLUMN IF NOT EXISTS` ile ekler.
    Anahtar Sütunlar (PRIMARY KEY):
      proje_adi, sd, kalem, faaliyet_1_turu (daima 'Dry-Run'), faaliyet_2_turu, urun_kisa_metni
    """
    required_columns = [
        ("proje_adi", "TEXT NOT NULL DEFAULT ''"),
        ("sd", "TEXT NOT NULL DEFAULT ''"),
        ("kalem", "TEXT NOT NULL DEFAULT ''"),
        ("faaliyet_1_turu", "TEXT NOT NULL DEFAULT 'Dry-Run'"),
        ("faaliyet_2_turu", "TEXT NOT NULL DEFAULT 'FKT'"),
        ("urun_kisa_metni", "TEXT NOT NULL DEFAULT ''"),
        ("ilk_faaliyet_benzersiz_sayisi", "INTEGER NOT NULL DEFAULT 0"),
        ("ortak_uygunsuzluk_sayisi", "INTEGER NOT NULL DEFAULT 0"),
        ("ikinci_faaliyet_benzersiz_sayisi", "INTEGER NOT NULL DEFAULT 0"),
        ("toplam_benzersiz_uygunsuzluk_sayisi", "INTEGER NOT NULL DEFAULT 0"),
        ("jaccard_benzerlik_orani", "NUMERIC(6, 2) NOT NULL DEFAULT 0"),
        ("ilk_faaliyet_benzersiz_uygunsuzluklar", "JSONB NOT NULL DEFAULT '[]'::jsonb"),
        ("ortak_uygunsuzluklar", "JSONB NOT NULL DEFAULT '[]'::jsonb"),
        ("ikinci_faaliyet_benzersiz_uygunsuzluklar", "JSONB NOT NULL DEFAULT '[]'::jsonb"),
        ("reddedilen_eslesmeler", "JSONB NOT NULL DEFAULT '[]'::jsonb"),
        ("detay_json", "JSONB NOT NULL DEFAULT '{}'::jsonb"),
        ("uges_proje_tanim", "TEXT DEFAULT ''"),
        ("malzeme", "TEXT DEFAULT ''"),
        ("kaydeden_kullanici", "TEXT DEFAULT ''"),
        ("guncellenme_zamani", "TIMESTAMPTZ NOT NULL DEFAULT NOW()")
    ]

    added_columns: List[str] = []
    table_existed = False

    with conn.cursor() as cur:
        cur.execute("CREATE SCHEMA IF NOT EXISTS teslimat_md;")
        cur.execute(
            """
            SELECT EXISTS (
                SELECT 1
                FROM information_schema.tables
                WHERE table_schema = 'teslimat_md' AND table_name = 'dryrun_fkt_skt_kpi'
            );
            """
        )
        table_existed = bool(cur.fetchone()[0])

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS teslimat_md.dryrun_fkt_skt_kpi (
                proje_adi TEXT NOT NULL,
                sd TEXT NOT NULL,
                kalem TEXT NOT NULL,
                faaliyet_1_turu TEXT NOT NULL DEFAULT 'Dry-Run',
                faaliyet_2_turu TEXT NOT NULL,
                urun_kisa_metni TEXT NOT NULL,
                ilk_faaliyet_benzersiz_sayisi INTEGER NOT NULL DEFAULT 0,
                ortak_uygunsuzluk_sayisi INTEGER NOT NULL DEFAULT 0,
                ikinci_faaliyet_benzersiz_sayisi INTEGER NOT NULL DEFAULT 0,
                toplam_benzersiz_uygunsuzluk_sayisi INTEGER NOT NULL DEFAULT 0,
                jaccard_benzerlik_orani NUMERIC(6, 2) NOT NULL DEFAULT 0,
                ilk_faaliyet_benzersiz_uygunsuzluklar JSONB NOT NULL DEFAULT '[]'::jsonb,
                ortak_uygunsuzluklar JSONB NOT NULL DEFAULT '[]'::jsonb,
                ikinci_faaliyet_benzersiz_uygunsuzluklar JSONB NOT NULL DEFAULT '[]'::jsonb,
                reddedilen_eslesmeler JSONB NOT NULL DEFAULT '[]'::jsonb,
                detay_json JSONB NOT NULL DEFAULT '{}'::jsonb,
                uges_proje_tanim TEXT DEFAULT '',
                malzeme TEXT DEFAULT '',
                kaydeden_kullanici TEXT DEFAULT '',
                guncellenme_zamani TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                CONSTRAINT pk_dryrun_fkt_skt_kpi PRIMARY KEY (
                    proje_adi, sd, kalem, faaliyet_1_turu, faaliyet_2_turu, urun_kisa_metni
                )
            );
            """
        )

        # Mevcut tabloda eksik sütun varsa otomatik ekle (Şema değişikliği kontrolü)
        cur.execute(
            """
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = 'teslimat_md' AND table_name = 'dryrun_fkt_skt_kpi';
            """
        )
        existing_cols = {row[0] for row in cur.fetchall()}

        for col_name, col_def in required_columns:
            if col_name not in existing_cols:
                cur.execute(
                    f"ALTER TABLE teslimat_md.dryrun_fkt_skt_kpi ADD COLUMN IF NOT EXISTS {col_name} {col_def};"
                )
                added_columns.append(col_name)

    conn.commit()
    return {
        "tablo": "teslimat_md.dryrun_fkt_skt_kpi",
        "onceden_mevcut_muydu": table_existed,
        "olusturuldu_mu": not table_existed,
        "eklenen_yeni_sutunlar": added_columns,
        "toplam_sutun_sayisi": len(required_columns)
    }


def verify_and_init_database_on_startup() -> Dict[str, Any]:
    """
    Sunucu/site ayağa kalktığında otomatik çalışır:
    1. `.env` dosyasından PostgreSQL bağlantı ayarlarını yükler.
    2. `teslimat_md.bakiye_siparis` tablosunun varlığını ve güncel kayıt/proje sayısını kontrol eder.
    3. `teslimat_md.dryrun_fkt_skt_kpi` tablosunun varlığını ve sütun şemasını kontrol eder,
       tablo yoksa oluşturur, eksik sütun varsa otomatik günceller.
    """
    db_cfg = get_db_config()
    status_report: Dict[str, Any] = {
        "durum": "basarili",
        "env_dosyasi": db_cfg["env_file"],
        "env_dosyasi_mevcut": db_cfg["env_file_exists"],
        "sunucu": get_server_config(),
        "baglanti": {
            "host": db_cfg["host"],
            "port": db_cfg["port"],
            "dbname": db_cfg["dbname"],
            "user": db_cfg["user"],
            "schema": db_cfg["schema"]
        },
        "bakiye_siparis_tablosu": {
            "tablo": f"{db_cfg['schema']}.{db_cfg['bakiye_table']}",
            "mevcut": False,
            "kayit_sayisi": 0
        },
        "kpi_tablosu": {
            "tablo": f"{db_cfg['schema']}.{db_cfg['kpi_table']}",
            "mevcut": False,
            "kayit_sayisi": 0
        }
    }

    try:
        import psycopg2
        conn = psycopg2.connect(db_cfg["dsn"], connect_timeout=db_cfg["timeout"])
        try:
            kpi_info = ensure_dryrun_fkt_skt_kpi_table(conn)
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT EXISTS (
                        SELECT 1
                        FROM information_schema.tables
                        WHERE table_schema = %s AND table_name = %s
                    );
                    """,
                    (db_cfg["schema"], db_cfg["bakiye_table"])
                )
                bakiye_exists = bool(cur.fetchone()[0])
                bakiye_count = 0
                if bakiye_exists:
                    cur.execute(f"SELECT COUNT(*) FROM {db_cfg['schema']}.{db_cfg['bakiye_table']};")
                    bakiye_count = int(cur.fetchone()[0])

                cur.execute("SELECT COUNT(*) FROM teslimat_md.dryrun_fkt_skt_kpi;")
                kpi_count = int(cur.fetchone()[0])

            status_report["bakiye_siparis_tablosu"] = {
                "tablo": f"{db_cfg['schema']}.{db_cfg['bakiye_table']}",
                "mevcut": bakiye_exists,
                "kayit_sayisi": bakiye_count
            }
            status_report["kpi_tablosu"] = {
                **kpi_info,
                "mevcut": True,
                "kayit_sayisi": kpi_count
            }
            _safe_log(
                f"[DB-STARTUP] PostgreSQL ({db_cfg['host']}:{db_cfg['port']}/{db_cfg['dbname']}) doğrulandı -> "
                f"bakiye_siparis: {bakiye_count} satır, dryrun_fkt_skt_kpi: {kpi_count} satır "
                f"(Yeni oluşturuldu: {kpi_info['olusturuldu_mu']}, Eklenen sütunlar: {kpi_info['eklenen_yeni_sutunlar']})"
            )
        finally:
            conn.close()
    except Exception as e:
        status_report["durum"] = "uyari"
        status_report["hata"] = str(e)
        _safe_log(f"[DB-STARTUP-WARN] Veritabanı başlangıç kontrolü uyarısı: {e}")

    return status_report


def save_dryrun_fkt_skt_kpi_record(payload: Dict[str, Any]) -> Dict[str, Any]:
    """
    Mevcut (kullanıcının 'Benzer Değil / Decline' düzenlemeleri sonrasındaki) Venn sonuçlarını
    `teslimat_md.dryrun_fkt_skt_kpi` tablosuna UPSERT (INSERT ... ON CONFLICT DO UPDATE) olarak kaydeder.
    İlk faaliyet türü kural gereği daima 'Dry-Run' olarak sabitlenir.
    """
    proje_adi = str(payload.get("proje_adi") or "").strip() or "Belirtilmedi"
    sd = str(payload.get("sd") or "").strip() or "-"
    kalem = str(payload.get("kalem") or "").strip() or "-"
    # Kural: İlk faaliyet hep Dry-Run olmalı, 2. faaliyet ise FKT ya da SKT olmalı
    faaliyet_1_turu = "Dry-Run"
    raw_f2 = str(payload.get("faaliyet_2_turu") or "FKT").strip().upper()
    faaliyet_2_turu = "SKT" if raw_f2 == "SKT" else "FKT"
    urun_kisa_metni = str(payload.get("urun_kisa_metni") or "").strip() or "-"
    uges_proje_tanim = str(payload.get("uges_proje_tanim") or "").strip()
    malzeme = str(payload.get("malzeme") or "").strip()

    only_a_items: List[str] = [str(x) for x in (payload.get("ilk_faaliyet_benzersiz_uygunsuzluklar") or []) if str(x).strip()]
    only_b_items: List[str] = [str(x) for x in (payload.get("ikinci_faaliyet_benzersiz_uygunsuzluklar") or []) if str(x).strip()]
    common_items_raw = payload.get("ortak_uygunsuzluklar") or []
    declined_raw = payload.get("reddedilen_eslesmeler") or []

    only_a_count = len(only_a_items)
    common_count = len(common_items_raw)
    only_b_count = len(only_b_items)
    total_unique = only_a_count + common_count + only_b_count
    jaccard_pct = round((common_count / total_unique) * 100.0, 2) if total_unique > 0 else 0.0

    current_user = KOKPIT_IDENTITY_STORE.get("current_user") or {}
    kaydeden = (
        str(payload.get("kaydeden_kullanici") or "").strip()
        or current_user.get("isim")
        or current_user.get("username")
        or current_user.get("sicil")
        or "VennFlow Kullanıcı"
    )

    detay_json_obj = {
        "anahtar": {
            "proje_adi": proje_adi,
            "sd": sd,
            "kalem": kalem,
            "faaliyet_1_turu": faaliyet_1_turu,
            "faaliyet_2_turu": faaliyet_2_turu,
            "urun_kisa_metni": urun_kisa_metni,
            "uges_proje_tanim": uges_proje_tanim,
            "malzeme": malzeme
        },
        "sayilar": {
            "ilk_faaliyet_benzersiz_sayisi": only_a_count,
            "ortak_uygunsuzluk_sayisi": common_count,
            "ikinci_faaliyet_benzersiz_sayisi": only_b_count,
            "toplam_benzersiz_uygunsuzluk_sayisi": total_unique,
            "jaccard_benzerlik_orani": jaccard_pct
        },
        "uygunsuzluklar": {
            "ilk_faaliyet_dry_run_benzersiz": only_a_items,
            "ortak_benzer_uygunsuzluklar": common_items_raw,
            "ikinci_faaliyet_benzersiz": only_b_items,
            "kullanici_tarafindan_reddedilenler": declined_raw
        },
        "kaydeden_kullanici": kaydeden,
        "kayit_zamani": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    }

    record_summary = {
        "proje_adi": proje_adi,
        "sd": sd,
        "kalem": kalem,
        "faaliyet_1_turu": faaliyet_1_turu,
        "faaliyet_2_turu": faaliyet_2_turu,
        "urun_kisa_metni": urun_kisa_metni,
        "ilk_faaliyet_benzersiz_sayisi": only_a_count,
        "ortak_uygunsuzluk_sayisi": common_count,
        "ikinci_faaliyet_benzersiz_sayisi": only_b_count,
        "toplam_benzersiz_uygunsuzluk_sayisi": total_unique,
        "jaccard_benzerlik_orani": jaccard_pct,
        "ilk_faaliyet_benzersiz_uygunsuzluklar": only_a_items,
        "ortak_uygunsuzluklar": common_items_raw,
        "ikinci_faaliyet_benzersiz_uygunsuzluklar": only_b_items,
        "reddedilen_eslesmeler": declined_raw,
        "uges_proje_tanim": uges_proje_tanim,
        "malzeme": malzeme,
        "kaydeden_kullanici": kaydeden
    }

    try:
        import psycopg2
        from psycopg2.extras import Json
        db_cfg = get_db_config()
        conn = psycopg2.connect(db_cfg["dsn"], connect_timeout=db_cfg["timeout"])
        try:
            ensure_dryrun_fkt_skt_kpi_table(conn)
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO teslimat_md.dryrun_fkt_skt_kpi (
                        proje_adi, sd, kalem, faaliyet_1_turu, faaliyet_2_turu, urun_kisa_metni,
                        ilk_faaliyet_benzersiz_sayisi, ortak_uygunsuzluk_sayisi, ikinci_faaliyet_benzersiz_sayisi,
                        toplam_benzersiz_uygunsuzluk_sayisi, jaccard_benzerlik_orani,
                        ilk_faaliyet_benzersiz_uygunsuzluklar, ortak_uygunsuzluklar,
                        ikinci_faaliyet_benzersiz_uygunsuzluklar, reddedilen_eslesmeler, detay_json,
                        uges_proje_tanim, malzeme, kaydeden_kullanici, guncellenme_zamani
                    ) VALUES (
                        %s, %s, %s, %s, %s, %s,
                        %s, %s, %s,
                        %s, %s,
                        %s, %s,
                        %s, %s, %s,
                        %s, %s, %s, NOW()
                    )
                    ON CONFLICT (proje_adi, sd, kalem, faaliyet_1_turu, faaliyet_2_turu, urun_kisa_metni)
                    DO UPDATE SET
                        ilk_faaliyet_benzersiz_sayisi = EXCLUDED.ilk_faaliyet_benzersiz_sayisi,
                        ortak_uygunsuzluk_sayisi = EXCLUDED.ortak_uygunsuzluk_sayisi,
                        ikinci_faaliyet_benzersiz_sayisi = EXCLUDED.ikinci_faaliyet_benzersiz_sayisi,
                        toplam_benzersiz_uygunsuzluk_sayisi = EXCLUDED.toplam_benzersiz_uygunsuzluk_sayisi,
                        jaccard_benzerlik_orani = EXCLUDED.jaccard_benzerlik_orani,
                        ilk_faaliyet_benzersiz_uygunsuzluklar = EXCLUDED.ilk_faaliyet_benzersiz_uygunsuzluklar,
                        ortak_uygunsuzluklar = EXCLUDED.ortak_uygunsuzluklar,
                        ikinci_faaliyet_benzersiz_uygunsuzluklar = EXCLUDED.ikinci_faaliyet_benzersiz_uygunsuzluklar,
                        reddedilen_eslesmeler = EXCLUDED.reddedilen_eslesmeler,
                        detay_json = EXCLUDED.detay_json,
                        uges_proje_tanim = EXCLUDED.uges_proje_tanim,
                        malzeme = EXCLUDED.malzeme,
                        kaydeden_kullanici = EXCLUDED.kaydeden_kullanici,
                        guncellenme_zamani = NOW()
                    RETURNING guncellenme_zamani;
                    """,
                    (
                        proje_adi, sd, kalem, faaliyet_1_turu, faaliyet_2_turu, urun_kisa_metni,
                        only_a_count, common_count, only_b_count,
                        total_unique, jaccard_pct,
                        Json(only_a_items), Json(common_items_raw),
                        Json(only_b_items), Json(declined_raw), Json(detay_json_obj),
                        uges_proje_tanim, malzeme, kaydeden
                    )
                )
                ts_row = cur.fetchone()
                cur.execute("SELECT COUNT(*) FROM teslimat_md.dryrun_fkt_skt_kpi;")
                total_rows_in_db = cur.fetchone()[0]
            conn.commit()
            record_summary["guncellenme_zamani"] = str(ts_row[0]) if ts_row else ""
            return {
                "durum": "basarili",
                "tablo": "teslimat_md.dryrun_fkt_skt_kpi",
                "veritabani": "ud_kokpit (PostgreSQL)",
                "mesaj": (
                    f"Veriler 'teslimat_md.dryrun_fkt_skt_kpi' tablosuna başarıyla kaydedildi "
                    f"({proje_adi} • SD: {sd} / Kalem: {kalem} • {faaliyet_1_turu} vs {faaliyet_2_turu})."
                ),
                "toplam_tablo_kaydi": total_rows_in_db,
                "kayit": record_summary
            }
        finally:
            conn.close()
    except Exception as e:
        # Yedek belleğe kaydet
        LOCAL_KPI_FALLBACK_STORE.append(record_summary)
        return {
            "durum": "uyari",
            "tablo": "teslimat_md.dryrun_fkt_skt_kpi (yerel_yedek)",
            "mesaj": f"PostgreSQL bağlantısı kurulamadı ({str(e)}), kayıt yerel belleğe alındı.",
            "kayit": record_summary
        }


def fetch_dryrun_fkt_skt_kpi_records(
    limit: int = 100,
    proje: Optional[str] = None,
    sd: Optional[str] = None,
    kalem: Optional[str] = None,
    faaliyet_2_turu: Optional[str] = None,
    urun: Optional[str] = None,
    search: Optional[str] = None
) -> Dict[str, Any]:
    """
    `teslimat_md.dryrun_fkt_skt_kpi` tablosundaki mevcut KPI kayıtlarını
    isteğe bağlı filtrelerle (Proje Adı, SD, Kalem, 2. Faaliyet Türü, Ürün Kısa Metni, Genel Arama) listeler.
    """
    try:
        import psycopg2
        db_cfg = get_db_config()
        conn = psycopg2.connect(db_cfg["dsn"], connect_timeout=db_cfg["timeout"])
        try:
            ensure_dryrun_fkt_skt_kpi_table(conn)
            where_clauses: List[str] = []
            params: List[Any] = []

            if proje and proje.strip() and proje.strip().upper() != "ALL":
                where_clauses.append("proje_adi ILIKE %s")
                params.append(f"%{proje.strip()}%")

            if sd and sd.strip():
                where_clauses.append("sd ILIKE %s")
                params.append(f"%{sd.strip()}%")

            if kalem and kalem.strip():
                where_clauses.append("kalem ILIKE %s")
                params.append(f"%{kalem.strip()}%")

            if faaliyet_2_turu and faaliyet_2_turu.strip() and faaliyet_2_turu.strip().upper() != "ALL":
                where_clauses.append("UPPER(faaliyet_2_turu) = %s")
                params.append(faaliyet_2_turu.strip().upper())

            if urun and urun.strip():
                where_clauses.append("urun_kisa_metni ILIKE %s")
                params.append(f"%{urun.strip()}%")

            if search and search.strip():
                q = f"%{search.strip()}%"
                where_clauses.append(
                    "("
                    "proje_adi ILIKE %s OR "
                    "sd ILIKE %s OR "
                    "kalem ILIKE %s OR "
                    "urun_kisa_metni ILIKE %s OR "
                    "COALESCE(uges_proje_tanim, '') ILIKE %s OR "
                    "COALESCE(malzeme, '') ILIKE %s OR "
                    "COALESCE(kaydeden_kullanici, '') ILIKE %s OR "
                    "CAST(ilk_faaliyet_benzersiz_uygunsuzluklar AS TEXT) ILIKE %s OR "
                    "CAST(ortak_uygunsuzluklar AS TEXT) ILIKE %s OR "
                    "CAST(ikinci_faaliyet_benzersiz_uygunsuzluklar AS TEXT) ILIKE %s"
                    ")"
                )
                params.extend([q] * 10)

            where_sql = ("WHERE " + " AND ".join(where_clauses)) if where_clauses else ""

            with conn.cursor() as cur:
                cur.execute(
                    f"""
                    SELECT
                        proje_adi,
                        sd,
                        kalem,
                        faaliyet_1_turu,
                        faaliyet_2_turu,
                        urun_kisa_metni,
                        ilk_faaliyet_benzersiz_sayisi,
                        ortak_uygunsuzluk_sayisi,
                        ikinci_faaliyet_benzersiz_sayisi,
                        toplam_benzersiz_uygunsuzluk_sayisi,
                        CAST(jaccard_benzerlik_orani AS FLOAT) AS jaccard_benzerlik_orani,
                        ilk_faaliyet_benzersiz_uygunsuzluklar,
                        ortak_uygunsuzluklar,
                        ikinci_faaliyet_benzersiz_uygunsuzluklar,
                        reddedilen_eslesmeler,
                        uges_proje_tanim,
                        malzeme,
                        kaydeden_kullanici,
                        CAST(guncellenme_zamani AS TEXT) AS guncellenme_zamani
                    FROM teslimat_md.dryrun_fkt_skt_kpi
                    {where_sql}
                    ORDER BY guncellenme_zamani DESC
                    LIMIT %s;
                    """,
                    tuple(params + [limit])
                )
                cols = [d[0] for d in cur.description]
                rows = [dict(zip(cols, r)) for r in cur.fetchall()]

                cur.execute(
                    """
                    SELECT DISTINCT proje_adi
                    FROM teslimat_md.dryrun_fkt_skt_kpi
                    WHERE proje_adi IS NOT NULL AND TRIM(proje_adi) <> ''
                    ORDER BY proje_adi;
                    """
                )
                distinct_projects = [r[0] for r in cur.fetchall()]

            return {
                "durum": "basarili",
                "tablo": "teslimat_md.dryrun_fkt_skt_kpi",
                "toplam_kayit": len(rows),
                "projeler": distinct_projects,
                "kayitlar": rows
            }
        finally:
            conn.close()
    except Exception as e:
        fallback_rows = list(reversed(LOCAL_KPI_FALLBACK_STORE[-limit:]))
        if proje and proje.strip() and proje.strip().upper() != "ALL":
            p_low = proje.strip().lower()
            fallback_rows = [r for r in fallback_rows if p_low in str(r.get("proje_adi") or "").lower()]
        if sd and sd.strip():
            s_low = sd.strip().lower()
            fallback_rows = [r for r in fallback_rows if s_low in str(r.get("sd") or "").lower()]
        if kalem and kalem.strip():
            k_low = kalem.strip().lower()
            fallback_rows = [r for r in fallback_rows if k_low in str(r.get("kalem") or "").lower()]
        if faaliyet_2_turu and faaliyet_2_turu.strip() and faaliyet_2_turu.strip().upper() != "ALL":
            f2_up = faaliyet_2_turu.strip().upper()
            fallback_rows = [r for r in fallback_rows if str(r.get("faaliyet_2_turu") or "").upper() == f2_up]
        if urun and urun.strip():
            u_low = urun.strip().lower()
            fallback_rows = [r for r in fallback_rows if u_low in str(r.get("urun_kisa_metni") or "").lower()]
        if search and search.strip():
            q_low = search.strip().lower()
            fallback_rows = [
                r for r in fallback_rows
                if q_low in json.dumps(r, ensure_ascii=False).lower()
            ]
        distinct_projects = sorted({str(r.get("proje_adi") or "") for r in LOCAL_KPI_FALLBACK_STORE if r.get("proje_adi")})
        return {
            "durum": "uyari",
            "tablo": "teslimat_md.dryrun_fkt_skt_kpi",
            "uyari": str(e),
            "toplam_kayit": len(fallback_rows),
            "projeler": distinct_projects,
            "kayitlar": fallback_rows
        }


def delete_dryrun_fkt_skt_kpi_record(
    payload: Dict[str, Any],
    requester_user: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """
    `teslimat_md.dryrun_fkt_skt_kpi` tablosundan belirtilen bileşik anahtara sahip kaydı siler.
    Yalnızca Admin yetkisine sahip kullanıcılar kayıt silebilir.
    """
    user_to_check = requester_user if isinstance(requester_user, dict) and requester_user else KOKPIT_IDENTITY_STORE.get("current_user")
    if not is_kokpit_user_admin(user_to_check):
        return {
            "durum": "yetkisiz",
            "mesaj": "Kayıt silme işlemi yalnızca Admin (Yönetici) yetkisine sahip kullanıcılar tarafından yapılabilir."
        }

    proje_adi = str(payload.get("proje_adi") or "").strip()
    sd = str(payload.get("sd") or "").strip()
    kalem = str(payload.get("kalem") or "").strip()
    faaliyet_1_turu = str(payload.get("faaliyet_1_turu") or "Dry-Run").strip()
    faaliyet_2_turu = str(payload.get("faaliyet_2_turu") or "").strip()
    urun_kisa_metni = str(payload.get("urun_kisa_metni") or "").strip()

    if not proje_adi or not sd or not kalem or not faaliyet_2_turu:
        return {
            "durum": "hata",
            "mesaj": "Silinecek kaydın anahtar alanları (proje_adi, sd, kalem, faaliyet_2_turu, urun_kisa_metni) eksik."
        }

    try:
        import psycopg2
        db_cfg = get_db_config()
        conn = psycopg2.connect(db_cfg["dsn"], connect_timeout=db_cfg["timeout"])
        try:
            ensure_dryrun_fkt_skt_kpi_table(conn)
            with conn.cursor() as cur:
                cur.execute(
                    """
                    DELETE FROM teslimat_md.dryrun_fkt_skt_kpi
                    WHERE proje_adi = %s
                      AND sd = %s
                      AND kalem = %s
                      AND faaliyet_1_turu = %s
                      AND faaliyet_2_turu = %s
                      AND urun_kisa_metni = %s;
                    """,
                    (proje_adi, sd, kalem, faaliyet_1_turu, faaliyet_2_turu, urun_kisa_metni)
                )
                deleted_count = cur.rowcount
                cur.execute("SELECT COUNT(*) FROM teslimat_md.dryrun_fkt_skt_kpi;")
                remaining_count = cur.fetchone()[0]
            conn.commit()

            # Yerel yedek listeden de temizle
            LOCAL_KPI_FALLBACK_STORE[:] = [
                r for r in LOCAL_KPI_FALLBACK_STORE
                if not (
                    r.get("proje_adi") == proje_adi
                    and r.get("sd") == sd
                    and r.get("kalem") == kalem
                    and r.get("faaliyet_1_turu") == faaliyet_1_turu
                    and r.get("faaliyet_2_turu") == faaliyet_2_turu
                    and r.get("urun_kisa_metni") == urun_kisa_metni
                )
            ]

            if deleted_count == 0:
                return {
                    "durum": "bulunamadi",
                    "silinen_adet": 0,
                    "kalan_kayit_sayisi": remaining_count,
                    "mesaj": "Belirtilen anahtara ait kayıt bulunamadı veya daha önce silinmiş."
                }

            return {
                "durum": "basarili",
                "silinen_adet": deleted_count,
                "kalan_kayit_sayisi": remaining_count,
                "silen_kullanici": (user_to_check or {}).get("isim") or (user_to_check or {}).get("username") or "Admin",
                "mesaj": (
                    f"Kayıt başarıyla silindi: {proje_adi} • SD: {sd} / Kalem: {kalem} "
                    f"({faaliyet_1_turu} vs {faaliyet_2_turu} • {urun_kisa_metni})."
                )
            }
        finally:
            conn.close()
    except Exception as e:
        before_len = len(LOCAL_KPI_FALLBACK_STORE)
        LOCAL_KPI_FALLBACK_STORE[:] = [
            r for r in LOCAL_KPI_FALLBACK_STORE
            if not (
                r.get("proje_adi") == proje_adi
                and r.get("sd") == sd
                and r.get("kalem") == kalem
                and r.get("faaliyet_1_turu") == faaliyet_1_turu
                and r.get("faaliyet_2_turu") == faaliyet_2_turu
                and r.get("urun_kisa_metni") == urun_kisa_metni
            )
        ]
        deleted_local = before_len - len(LOCAL_KPI_FALLBACK_STORE)
        if deleted_local > 0:
            return {
                "durum": "basarili",
                "silinen_adet": deleted_local,
                "mesaj": f"Kayıt yerel yedek tablodan silindi ({proje_adi} • {sd}/{kalem})."
            }
        return {
            "durum": "hata",
            "mesaj": f"Kayıt silinirken veritabanı hatası oluştu: {str(e)}"
        }


def generate_bulk_template_excel_bytes(proje: Optional[str] = None) -> bytes:
    """
    Toplu veya tekli uygunsuzluk veri girişi için `teslimat_md.bakiye_siparis`
    verileriyle entegre örnek Excel (.xlsx) şablonu üretir.
    Kullanıcı Sütun A (SD) ve Sütun B (Kalem) değerlerini seçtiğinde veya yazdığında,
    diğer sütunlar (Proje Adı, Ürün Kısa Metni, Malzeme/UGES Tanım) Excel formülüyle otomatik dolar.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation

    bakiye_data = fetch_bakiye_siparis_items(proje=proje)
    items = bakiye_data.get("kayitlar") or []

    wb = openpyxl.Workbook()

    # 1. Sayfa: Uygunsuzluk_Veri_Girisi
    ws_entry = wb.active
    ws_entry.title = "Uygunsuzluk_Veri_Girisi"

    # 2. Sayfa: Bakiye_Siparis_Listesi (Referans, VLOOKUP Anahtarı ve Açılır Liste Kaynağı)
    ws_ref = wb.create_sheet(title="Bakiye_Siparis_Listesi")

    ref_headers = [
        "SD_Kalem_Anahtar",
        "SD",
        "Kalem",
        "Proje_Adi",
        "Urun_Kisa_Metni",
        "Malzeme",
        "UGES_Proje_Tanim",
        "Tekil_SD_Secenekleri",
        "Tekil_Kalem_Secenekleri"
    ]
    ws_ref.append(ref_headers)

    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    for col_idx in range(1, len(ref_headers) + 1):
        c = ws_ref.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")

    unique_sds: List[str] = []
    unique_kalems: List[str] = []

    for r_idx, it in enumerate(items, start=2):
        sd_val = str(it.get("sd") or "").strip()
        kalem_val = str(it.get("kalem") or "").strip()
        proj_val = str(it.get("proje_adi") or "").strip()
        urun_val = str(it.get("urun_kisa_metni") or "").strip()
        malz_val = str(it.get("malzeme") or "").strip()
        uges_val = str(it.get("uges_proje_tanim") or "").strip()
        lookup_key = f"{sd_val}|{kalem_val}"

        if sd_val and sd_val not in unique_sds:
            unique_sds.append(sd_val)
        if kalem_val and kalem_val not in unique_kalems:
            unique_kalems.append(kalem_val)

        ws_ref.cell(row=r_idx, column=1, value=lookup_key)
        ws_ref.cell(row=r_idx, column=2, value=sd_val)
        ws_ref.cell(row=r_idx, column=3, value=kalem_val)
        ws_ref.cell(row=r_idx, column=4, value=proj_val)
        ws_ref.cell(row=r_idx, column=5, value=urun_val)
        ws_ref.cell(row=r_idx, column=6, value=malz_val or uges_val)
        ws_ref.cell(row=r_idx, column=7, value=uges_val)

    # Tekil SD ve Kalem listelerini H ve I sütunlarına yaz (Açılır listeler için tekrarsız kaynak)
    try:
        unique_kalems.sort(key=lambda x: int(x) if x.isdigit() else 999999)
    except Exception:
        pass

    for idx_s, sd_v in enumerate(unique_sds, start=2):
        ws_ref.cell(row=idx_s, column=8, value=sd_v)
    for idx_k, kal_v in enumerate(unique_kalems, start=2):
        ws_ref.cell(row=idx_k, column=9, value=kal_v)

    ref_max_row = max(2, len(items) + 1)
    sd_max_row = max(2, len(unique_sds) + 1)
    kalem_max_row = max(2, len(unique_kalems) + 1)

    # 1. Sayfa Başlıkları: SD ve Kalem girildiğinde Proje Adı, Ürün Kısa Metni ve Malzeme otomatik dolar
    entry_headers = [
        "SD (Seçin veya Yazın)",
        "Kalem (Seçin veya Yazın)",
        "Faaliyet_Turu (Dry-Run / FKT / SKT)",
        "Uygunsuzluk_Aciklamasi",
        "Proje_Adi (Otomatik Dolar)",
        "Urun_Kisa_Metni (Otomatik Dolar)",
        "Malzeme / UGES Tanım (Otomatik Dolar)"
    ]
    ws_entry.append(entry_headers)

    for col_idx in range(1, len(entry_headers) + 1):
        c = ws_entry.cell(row=1, column=col_idx)
        # İlk 4 sütun kullanıcı girişi (Mavi), son 3 sütun otomatik dolan alanlar (Mor/İndigo)
        bg_hex = "2563EB" if col_idx <= 4 else "4F46E5"
        c.fill = PatternFill(start_color=bg_hex, end_color=bg_hex, fill_type="solid")
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws_entry.row_dimensions[1].height = 30

    # Örnek satırları doldur (İlk 2 bakiye kalemiyle gerçekçi Dry-Run vs FKT/SKT örnekleri)
    sample_item_1 = items[0] if len(items) > 0 else {
        "sd": "60000180", "kalem": "50", "proje_adi": "MİLGEM", "urun_kisa_metni": "Sİper Blok-1 Füze"
    }
    sample_item_2 = items[1] if len(items) > 1 else sample_item_1

    sd_1 = str(sample_item_1.get("sd") or "60000180").strip()
    kal_1 = str(sample_item_1.get("kalem") or "50").strip()
    sd_2 = str(sample_item_2.get("sd") or sd_1).strip()
    kal_2 = str(sample_item_2.get("kalem") or kal_1).strip()

    sample_rows = [
        (sd_1, kal_1, "Dry-Run", "Pano içi etiketleme eksik"),
        (sd_1, kal_1, "Dry-Run", "Topraklama barası cıvatası gevşek"),
        (sd_1, kal_1, "Dry-Run", "Kablo kesiti standart dışı"),
        (sd_1, kal_1, "FKT", "etiketleme eksik pano içi"),
        (sd_1, kal_1, "FKT", "Topraklama barası cıvatası gevşek"),
        (sd_1, kal_1, "FKT", "Optik sensör kalibrasyon sapması"),
        (sd_2, kal_2, "Dry-Run", "Konnektör sızdırmazlık contası deforme"),
        (sd_2, kal_2, "Dry-Run", "Cıvata tork değeri yetersiz"),
        (sd_2, kal_2, "SKT", "Konnektör sızdırmazlık contası deforme"),
        (sd_2, kal_2, "SKT", "Klemens bağlantısı gevşek"),
    ]

    for row_num, (s_val, k_val, faz, desc) in enumerate(sample_rows, start=2):
        ws_entry.cell(row=row_num, column=1, value=s_val)
        ws_entry.cell(row=row_num, column=2, value=k_val)
        ws_entry.cell(row=row_num, column=3, value=faz)
        ws_entry.cell(row=row_num, column=4, value=desc)

    # 2..300 arası tüm satırlara SD (A) ve Kalem (B) girildiğinde diğer sütunları otomatik dolduran VLOOKUP formüllerini ekle
    auto_fill_bg = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    auto_fill_font = Font(color="1E293B", bold=True, size=10)
    for r in range(2, 301):
        lookup_expr = f'TRIM(A{r})&"|"&TRIM(B{r})'
        c_proj = ws_entry.cell(
            row=r, column=5,
            value=f'=IF(OR(A{r}="",B{r}=""),"",IFERROR(VLOOKUP({lookup_expr},Bakiye_Siparis_Listesi!$A$2:$G${ref_max_row},4,FALSE),""))'
        )
        c_proj.fill = auto_fill_bg
        c_proj.font = auto_fill_font

        c_urun = ws_entry.cell(
            row=r, column=6,
            value=f'=IF(OR(A{r}="",B{r}=""),"",IFERROR(VLOOKUP({lookup_expr},Bakiye_Siparis_Listesi!$A$2:$G${ref_max_row},5,FALSE),""))'
        )
        c_urun.fill = auto_fill_bg
        c_urun.font = auto_fill_font

        c_malz = ws_entry.cell(
            row=r, column=7,
            value=f'=IF(OR(A{r}="",B{r}=""),"",IFERROR(VLOOKUP({lookup_expr},Bakiye_Siparis_Listesi!$A$2:$G${ref_max_row},6,FALSE),""))'
        )
        c_malz.fill = auto_fill_bg

        for c_i in range(1, 8):
            ws_entry.cell(row=r, column=c_i).border = thin_border

    # Açılır Liste (Data Validation): Sütun A (SD) -> Tekil SD Listesi (Serbest yazmaya da izin verir)
    dv_sd = DataValidation(
        type="list",
        formula1=f"=Bakiye_Siparis_Listesi!$H$2:$H${sd_max_row}",
        allow_blank=True
    )
    dv_sd.showErrorMessage = False
    dv_sd.promptTitle = "SD Numarası"
    dv_sd.prompt = "Açılır listeden SD seçin veya doğrudan yazın."
    ws_entry.add_data_validation(dv_sd)
    dv_sd.add("A2:A500")

    # Açılır Liste (Data Validation): Sütun B (Kalem) -> Tekil Kalem Listesi (Serbest yazmaya da izin verir)
    dv_kalem = DataValidation(
        type="list",
        formula1=f"=Bakiye_Siparis_Listesi!$I$2:$I${kalem_max_row}",
        allow_blank=True
    )
    dv_kalem.showErrorMessage = False
    dv_kalem.promptTitle = "Kalem Numarası"
    dv_kalem.prompt = "Açılır listeden Kalem seçin veya doğrudan yazın. SD ve Kalem girildiğinde Proje ve Ürün otomatik dolar."
    ws_entry.add_data_validation(dv_kalem)
    dv_kalem.add("B2:B500")

    # Açılır Liste (Data Validation): Sütun C -> Dry-Run, FKT, SKT
    dv_faz = DataValidation(
        type="list",
        formula1='"Dry-Run,FKT,SKT"',
        allow_blank=True
    )
    dv_faz.promptTitle = "Faaliyet Türü"
    dv_faz.prompt = "1. Faaliyet için Dry-Run, 2. Faaliyet için FKT veya SKT seçin."
    ws_entry.add_data_validation(dv_faz)
    dv_faz.add("C2:C500")

    # Sütun genişlikleri
    col_widths_entry = [22, 22, 26, 52, 28, 38, 28]
    for idx, w in enumerate(col_widths_entry, start=1):
        col_letter = openpyxl.utils.get_column_letter(idx)
        ws_entry.column_dimensions[col_letter].width = w

    col_widths_ref = [24, 18, 14, 26, 38, 24, 28, 20, 20]
    for idx, w in enumerate(col_widths_ref, start=1):
        col_letter = openpyxl.utils.get_column_letter(idx)
        ws_ref.column_dimensions[col_letter].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def normalize_col_name(col_text: Any) -> str:
    """Sütun adlarını küçük harfe çevirip boşlukları alt çizgiye dönüştürür."""
    import re
    base = normalize_str(col_text)
    return re.sub(r"[^a-z0-9]+", "_", base).strip("_")


def _parse_sd_kalem_from_row(row_dict: Dict[str, Any], default_sd: str = "", default_kalem: str = "") -> Tuple[str, str, str, str]:
    """
    Excel satırından (SD ve Kalem sütunları veya birleşik SD_Kalem_Secimi sütunu)
    (sd, kalem, fallback_proje, fallback_urun) bilgisini çözümler.
    """
    norm_map = {normalize_col_name(str(k)): v for k, v in row_dict.items() if k is not None}

    def _val(*candidates: str) -> str:
        # 1. Önce tam eşleşen veya candidate ile başlayan sütunlara bak (örn. 'sd_secin_veya_yazin')
        for c in candidates:
            nc = normalize_col_name(c)
            for k, v in norm_map.items():
                if (k == nc or k.startswith(nc + "_")) and not k.startswith("sd_kalem"):
                    if v is not None and not pd.isna(v):
                        s = str(v).strip()
                        if s and not s.startswith("="):
                            return s
        # 2. Alt dizge eşleşmesi
        for c in candidates:
            nc = normalize_col_name(c)
            for k, v in norm_map.items():
                if nc in k and v is not None and not pd.isna(v):
                    s = str(v).strip()
                    if s and not s.startswith("="):
                        return s
        return ""

    sd_val = _val("sd", "satis_belgesi", "siparis_no")
    kalem_val = _val("kalem", "kalem_numarasi", "kalem_no")
    combo_str = _val("sd_kalem_secimi", "sd_kalem", "siparis_kalem")
    proj_val = _val("proje_adi", "proje")
    urun_val = _val("urun_kisa_metni", "urun")

    if combo_str and "|" in combo_str:
        parts = [p.strip() for p in combo_str.split("|")]
        if len(parts) >= 1 and "/" in parts[0]:
            sd_p, kal_p = [x.strip() for x in parts[0].split("/", 1)]
            sd_val = sd_val or sd_p
            kalem_val = kalem_val or kal_p
        if len(parts) >= 2 and not proj_val:
            proj_val = parts[1]
        if len(parts) >= 3 and not urun_val:
            urun_val = parts[2]
    elif combo_str and "/" in combo_str:
        sd_p, kal_p = [x.strip() for x in combo_str.split("/", 1)]
        sd_val = sd_val or sd_p
        kalem_val = kalem_val or kal_p

    if sd_val.endswith(".0"):
        sd_val = sd_val[:-2]
    if kalem_val.endswith(".0"):
        kalem_val = kalem_val[:-2]

    sd_val = sd_val or default_sd or "-"
    kalem_val = kalem_val or default_kalem or "-"
    return sd_val, kalem_val, proj_val, urun_val


def analyze_bulk_uploaded_excel(
    file_bytes: bytes,
    filename: str = "toplu_veri.xlsx",
    engine_mode: str = "local",
    threshold: float = 0.80,
    default_sd: str = "",
    default_kalem: str = "",
    base_url: str = "https://api.aselixai.aselsan.com.tr/v1",
    api_key: str = "",
    embedding_model: str = "text-embedding-3-small"
) -> Dict[str, Any]:
    """
    Toplu veri girişi şablonundan yüklenen Excel dosyasını okur,
    her satırı arka planda `teslimat_md.bakiye_siparis` (SD / Kalem) ile eşleştirir,
    SD/Kalem bazında Dry-Run ile FKT/SKT uygunsuzluklarını kıyaslayıp Venn kümelerini döndürür.
    """
    df = read_excel_safe(file_bytes)
    if df.empty:
        return {"durum": "hata", "mesaj": "Yüklenen Excel dosyasında veri satırı bulunamadı."}

    bakiye_res = fetch_bakiye_siparis_items()
    bakiye_items = bakiye_res.get("kayitlar") or []
    bakiye_lookup: Dict[Tuple[str, str], Dict[str, Any]] = {}
    for bit in bakiye_items:
        k = (str(bit.get("sd") or "").strip(), str(bit.get("kalem") or "").strip())
        bakiye_lookup[k] = bit

    # Sütunları tespit et
    records_by_key: Dict[Tuple[str, str], Dict[str, Any]] = {}
    valid_row_count = 0

    for _, row in df.iterrows():
        r_dict = row.to_dict()
        norm_map = {normalize_col_name(str(k)): v for k, v in r_dict.items() if k is not None}

        # Uygunsuzluk açıklaması
        desc = ""
        for cand in ["uygunsuzluk_aciklamasi", "uygunsuzluk", "aciklama", "hata", "bulgu", "tanim"]:
            for k, v in norm_map.items():
                if cand in k and v is not None and not pd.isna(v):
                    s = clean_text_content(v)
                    if s and not s.startswith("="):
                        desc = s
                        break
            if desc:
                break
        if not desc:
            continue

        # Faaliyet Türü (Dry-Run, FKT, SKT)
        raw_faz = ""
        for cand in ["faaliyet_turu", "faaliyet", "faz", "test_fazi", "asama"]:
            for k, v in norm_map.items():
                if cand in k and v is not None and not pd.isna(v):
                    s = str(v).strip()
                    if s:
                        raw_faz = s
                        break
            if raw_faz:
                break

        faz_up = raw_faz.upper()
        if "SKT" in faz_up or "SAHA" in faz_up:
            norm_faz = "SKT"
        elif "FKT" in faz_up or "FABR" in faz_up:
            norm_faz = "FKT"
        else:
            norm_faz = "Dry-Run"

        sd_val, kalem_val, fallback_proj, fallback_urun = _parse_sd_kalem_from_row(
            r_dict, default_sd=default_sd, default_kalem=default_kalem
        )
        g_key = (sd_val, kalem_val)

        if g_key not in records_by_key:
            matched_bakiye = bakiye_lookup.get(g_key)
            records_by_key[g_key] = {
                "sd": sd_val,
                "kalem": kalem_val,
                "bakiye_eslesti": matched_bakiye is not None,
                "proje_adi": (matched_bakiye or {}).get("proje_adi") or fallback_proj or "Genel / Belirtilmemiş Proje",
                "urun_kisa_metni": (matched_bakiye or {}).get("urun_kisa_metni") or fallback_urun or "-",
                "malzeme": (matched_bakiye or {}).get("malzeme") or "",
                "uges_proje_tanim": (matched_bakiye or {}).get("uges_proje_tanim") or "",
                "phases": {"Dry-Run": [], "FKT": [], "SKT": []}
            }

        records_by_key[g_key]["phases"][norm_faz].append(desc)
        valid_row_count += 1

    if not records_by_key:
        return {
            "durum": "hata",
            "mesaj": "Excel içerisinde geçerli uygunsuzluk açıklaması satırı bulunamadı. Lütfen şablon formatını kontrol edin."
        }

    analyzed_groups: List[Dict[str, Any]] = []
    for (sd_val, kalem_val), g_info in records_by_key.items():
        dry_list = g_info["phases"]["Dry-Run"]
        fkt_list = g_info["phases"]["FKT"]
        skt_list = g_info["phases"]["SKT"]

        # Hangi 2. faaliyet türleri kıyaslanacak?
        target_second_phases: List[Tuple[str, List[str]]] = []
        if fkt_list and skt_list:
            target_second_phases.append(("FKT", fkt_list))
            target_second_phases.append(("SKT", skt_list))
        elif skt_list:
            target_second_phases.append(("SKT", skt_list))
        else:
            target_second_phases.append(("FKT", fkt_list))

        for faz_2_name, faz_2_items in target_second_phases:
            if engine_mode == "embedding":
                venn_res = compute_embedding_venn(
                    items_a=dry_list,
                    items_b=faz_2_items,
                    threshold=threshold,
                    base_url=base_url,
                    api_key=api_key,
                    embedding_model=embedding_model
                )
            else:
                venn_res = compute_fuzzy_venn(dry_list, faz_2_items, threshold=threshold)

            venn_res["declined_pairs"] = []
            group_id = f"{sd_val}_{kalem_val}_DryRun_vs_{faz_2_name}"
            analyzed_groups.append({
                "grup_id": group_id,
                "proje_adi": g_info["proje_adi"],
                "sd": sd_val,
                "kalem": kalem_val,
                "urun_kisa_metni": g_info["urun_kisa_metni"],
                "malzeme": g_info["malzeme"],
                "uges_proje_tanim": g_info["uges_proje_tanim"],
                "bakiye_eslesti": g_info["bakiye_eslesti"],
                "faaliyet_1_turu": "Dry-Run",
                "faaliyet_2_turu": faz_2_name,
                "dry_run_ham_sayisi": len(dry_list),
                "ikinci_faaliyet_ham_sayisi": len(faz_2_items),
                "esik_orani": threshold,
                "engine_mode": engine_mode,
                "venn_verileri": venn_res
            })

    return {
        "durum": "basarili",
        "dosya_adi": filename,
        "toplam_okunan_satir": valid_row_count,
        "toplam_grup_sayisi": len(analyzed_groups),
        "gruplar": analyzed_groups
    }


# ==============================================================================
# AĞDAKİ HERKESE AÇIK HTTP WEB SERVİSİ (FastAPI / Starlette)
# ==============================================================================

try:
    from fastapi import FastAPI, Request, Response, UploadFile, File, Form
    from fastapi.middleware.cors import CORSMiddleware
    from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
    from fastapi.staticfiles import StaticFiles
    from pydantic import BaseModel, Field

    web_app = FastAPI(
        title="VennFlow AI - Excel Kıyaslama ve Venn Analiz Sunucusu",
        description=(
            "Intranet ortamında Excel verilerini Venn Şeması küme matematiği ile analiz eden REST API & MCP Server.\n\n"
            "ASELSAN AI Platform LLM Suite Flow (Langflow), AselixAI Embedding ve Kokpit SSO (Iframe Push) entegrasyonuna sahiptir."
        ),
        version="2.1.0"
    )

    web_app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    @web_app.on_event("startup")
    def on_web_app_startup():
        """Sunucu başladığında .env dosyasını okur ve veritabanı tablolarının varlığını/şemasını otomatik denetler."""
        verify_and_init_database_on_startup()

    @web_app.middleware("http")
    async def kokpit_iframe_and_identity_middleware(request: Request, call_next):
        """
        Rehber Bölüm 3 (Iframe İzinleri) ve Bölüm 4/5 (Kimlik Başlıkları) gereksinimleri:
        1. X-Frame-Options kısıtlamasını kaldırır, dış sitenin Kokpit içinde <iframe> olarak
           sorunsuz gömülebilmesi (embedding) için Content-Security-Policy frame-ancestors başlığını ekler.
        2. İsteklerde gelen Authorization: Bearer <token> veya X-Kokpit-Token başlıklarını yakalar.
        """
        auth_header = request.headers.get("authorization") or request.headers.get("x-kokpit-token")
        if auth_header:
            token_val = auth_header[7:].strip() if auth_header.lower().startswith("bearer ") else auth_header.strip()
            if token_val and token_val != KOKPIT_IDENTITY_STORE.get("jwt_token"):
                KOKPIT_IDENTITY_STORE["jwt_token"] = token_val
                KOKPIT_IDENTITY_STORE["token_claims"] = decode_jwt_payload_unverified(token_val)
                if not KOKPIT_IDENTITY_STORE.get("current_user"):
                    inferred_user = normalize_kokpit_user({}, token_val)
                    if inferred_user.get("sicil") or inferred_user.get("username") or inferred_user.get("isim"):
                        KOKPIT_IDENTITY_STORE["current_user"] = inferred_user
                        KOKPIT_IDENTITY_STORE["status"] = "success"

        response = await call_next(request)

        # Iframe gömme (embedding) engellerini temizle
        if "x-frame-options" in response.headers:
            del response.headers["x-frame-options"]
        response.headers["Content-Security-Policy"] = (
            "frame-ancestors 'self' * http://localhost:3004 https://localhost:3004 "
            "http://127.0.0.1:3004 https://*.aselsan.com.tr http://*.aselsan.com.tr;"
        )
        return response

    static_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
    if os.path.exists(static_dir):
        web_app.mount("/static", StaticFiles(directory=static_dir), name="static")

    class KokpitIdentityRequest(BaseModel):
        type: Optional[str] = Field("kokpit-identity", description="Mesaj tipi ('kokpit-identity')")
        status: Optional[str] = Field("success", description="Aktarım durumu ('success')")
        source: Optional[str] = Field("kokpit-iframe-bridge", description="Kaynak ('kokpit-iframe-bridge')")
        timestamp: Optional[str] = None
        token: Optional[str] = None
        user: Optional[Dict[str, Any]] = None
        origin: Optional[str] = None
        inputs: Optional[Dict[str, Any]] = None
        config: Optional[Dict[str, Any]] = None
        data: Optional[Dict[str, Any]] = None

        class Config:
            extra = "allow"

    class ExternalInputRequest(BaseModel):
        source: Optional[str] = "kokpit-external-input"
        token: Optional[str] = None
        user: Optional[Dict[str, Any]] = None
        items_1: Optional[List[str]] = None
        items_2: Optional[List[str]] = None
        text_1: Optional[str] = None
        text_2: Optional[str] = None
        name_1: Optional[str] = "1. Dış Veri"
        name_2: Optional[str] = "2. Dış Veri"
        embeddings_1: Optional[List[List[float]]] = None
        embeddings_2: Optional[List[List[float]]] = None
        engine_mode: Optional[str] = "local"  # 'local', 'embedding', 'aselix'
        threshold: Optional[float] = 0.80
        base_url: Optional[str] = None
        api_key: Optional[str] = None
        model: Optional[str] = None
        embedding_model: Optional[str] = "text-embedding-3-small"
        auto_compare: Optional[bool] = False
        config: Optional[Dict[str, Any]] = None

        class Config:
            extra = "allow"

    class EmbeddingRequest(BaseModel):
        texts: List[str]
        base_url: Optional[str] = "https://api.aselixai.aselsan.com.tr/v1"
        api_key: Optional[str] = ""
        model: Optional[str] = "text-embedding-3-small"

    class CompareEmbeddingsRequest(BaseModel):
        items_a: List[str]
        items_b: List[str]
        doc1_name: Optional[str] = "1. Veri Seti"
        doc2_name: Optional[str] = "2. Veri Seti"
        threshold: Optional[float] = 0.80
        base_url: Optional[str] = "https://api.aselixai.aselsan.com.tr/v1"
        api_key: Optional[str] = ""
        embedding_model: Optional[str] = "text-embedding-3-small"
        embeddings_a: Optional[List[List[float]]] = None
        embeddings_b: Optional[List[List[float]]] = None

    class AnalyzeRequest(BaseModel):
        file_path: Optional[str] = Field(None, description="Excel dosyasının yolu")
        compare_by: str = Field("phases", description="'phases' veya 'activities'")
        activity_name: Optional[str] = Field(None, description="Opsiyonel faaliyet adı")
        sheet_name: Optional[str] = Field(None, description="Opsiyonel sayfa adı")
        activity_column: Optional[str] = Field(None)
        phase_column: Optional[str] = Field(None)
        issue_column: Optional[str] = Field(None)

    class FlowChatRequest(BaseModel):
        message: str
        session_id: Optional[str] = None
        base_url: Optional[str] = "https://llmsuiteflow.aiplatform.aselsan.com.tr"
        flow_id: Optional[str] = "c16ecc83-ff64-4fc7-bb52-0f594beaf70a"
        api_key: Optional[str] = ""
        tweaks_input: Optional[str] = ""
        context: Optional[Dict[str, Any]] = None

    class SampleCompareRequest(BaseModel):
        base_url: Optional[str] = "https://llmsuiteflow.aiplatform.aselsan.com.tr"
        flow_id: Optional[str] = "c16ecc83-ff64-4fc7-bb52-0f594beaf70a"
        api_key: Optional[str] = ""
        tweaks_input: Optional[str] = ""

    class TestConnectionRequest(BaseModel):
        base_url: Optional[str] = "https://llmsuiteflow.aiplatform.aselsan.com.tr"
        flow_id: Optional[str] = "c16ecc83-ff64-4fc7-bb52-0f594beaf70a"
        api_key: Optional[str] = ""

    class AselixModelsRequest(BaseModel):
        base_url: Optional[str] = "https://api.aselixai.aselsan.com.tr/v1"
        api_key: Optional[str] = ""

    class AselixChatRequest(BaseModel):
        base_url: Optional[str] = "https://api.aselixai.aselsan.com.tr/v1"
        api_key: Optional[str] = ""
        model: Optional[str] = "aselix-llm-v1"
        messages: List[Dict[str, str]]
        temperature: Optional[float] = 0.2
        context: Optional[Dict[str, Any]] = None

    class BatchChunkRequest(BaseModel):
        base_url: Optional[str] = "https://api.aselixai.aselsan.com.tr/v1"
        api_key: Optional[str] = ""
        model: Optional[str] = "aselix-llm-v1"
        batch_index: int = 1
        total_batches: int = 1
        items_a_batch: List[str]
        candidates_b: List[str]
        threshold: Optional[float] = 0.8
        timeout_sec: Optional[int] = 50

    class BatchFinalizeRequest(BaseModel):
        base_url: Optional[str] = "https://api.aselixai.aselsan.com.tr/v1"
        api_key: Optional[str] = ""
        model: Optional[str] = "aselix-llm-v1"
        doc1_name: Optional[str] = "1. Veri"
        doc2_name: Optional[str] = "2. Veri"
        col1_name: Optional[str] = None
        col2_name: Optional[str] = None
        threshold: float = 0.8
        all_items_a: List[str]
        all_items_b: List[str]
        matched_pairs: List[Dict[str, Any]]
        need_executive_summary: bool = True
        meta_1: Optional[Dict[str, Any]] = None
        meta_2: Optional[Dict[str, Any]] = None

    @web_app.get("/", response_class=HTMLResponse)
    async def index(request: Request):
        accept_header = request.headers.get("accept", "")
        if "application/json" in accept_header and "text/html" not in accept_header:
            srv_cfg = get_server_config()
            host_header = request.headers.get("host") or f"{srv_cfg['host']}:{srv_cfg['port']}"
            return JSONResponse({
                "servis": "VennFlow AI - Excel Kıyaslama ve Venn Analiz Sunucusu",
                "durum": "aktif",
                "ag_erisim_adresi": f"http://{host_header}",
                "mcp_http_endpoint": "/mcp",
                "interaktif_swagger_arayuzu": "/docs",
                "web_arayuzu": f"http://{host_header}/"
            })

        index_file = os.path.join(static_dir, "index.html")
        if os.path.exists(index_file):
            return FileResponse(index_file)
        return HTMLResponse("<h1>VennFlow AI Web Arayüzü Hazır!</h1>")

    @web_app.get("/mcp")
    def mcp_get_status():
        return {
            "status": "ok",
            "mcp_protocol": "2024-11-05",
            "message": "MCP HTTP endpoint hazir. JSON-RPC POST istekleri bu adrese gonderilmelidir."
        }

    # --------------------------------------------------------------------------
    # KOKPİT KİMLİK AKTARIMI (SSO / IFRAME PUSH) & DIŞ INPUT UÇ NOKTALARI
    # --------------------------------------------------------------------------
    @web_app.post("/api/identity")
    async def api_receive_identity(request: Request):
        """
        Rehber Bölüm 5.C (Tier-2 Senaryosu) uyumlu Backend Kimlik ve Dış Input Alıcı Uç Noktası.
        Ana sistem (Kokpit) iframe postMessage köprüsünden veya doğrudan dış servislerden gelen
        { user, token, source, type, status, inputs, config } paketini doğrular ve oturuma kaydeder.
        """
        try:
            body = await request.json()
        except Exception:
            return JSONResponse({"error": "Geçersiz JSON gövdesi.", "durum": "hata"}, status_code=400)

        if not isinstance(body, dict):
            return JSONResponse({"error": "JSON nesnesi bekleniyor.", "durum": "hata"}, status_code=400)

        raw_user = body.get("user")
        token = body.get("token")
        has_ext_inputs = any(k in body for k in ["inputs", "config", "data", "items_1", "items_2", "text_1", "text_2", "embeddings_1", "embeddings_2"])

        # Eğer ne geçerli bir kullanıcı (sicil/username/isim/id) ne token ne de dış input varsa 400 döndür (Rehber Bölüm 5.C)
        if not token and not has_ext_inputs:
            if not isinstance(raw_user, dict) or not (
                raw_user.get("sicil") or raw_user.get("username") or raw_user.get("isim") or raw_user.get("id") is not None
            ):
                return JSONResponse({"error": "Geçersiz kullanıcı paketi", "durum": "hata"}, status_code=400)

        origin = request.headers.get("origin") or body.get("origin")
        stored = store_kokpit_identity_and_inputs(body, origin=origin)
        user_obj = stored.get("user") or {}
        isim_display = user_obj.get("isim") or user_obj.get("username") or user_obj.get("sicil") or "Kullanıcı"

        _safe_log(
            f"[KOKPIT-IDENTITY] Kimlik/Input alındı: {isim_display} "
            f"(Sicil: {user_obj.get('sicil', '-')}, Kaynak: {stored.get('source')})"
        )

        return {
            "status": "success",
            "durum": "basarili",
            "message": f"{isim_display} oturumu açıldı.",
            "session_id": stored["session_id"],
            "user": stored["user"],
            "token_present": stored["token_present"],
            "token_claims": stored["token_claims"],
            "source": stored["source"],
            "timestamp": stored["timestamp"],
            "external_inputs": stored["external_inputs"]
        }

    @web_app.get("/api/identity")
    async def api_get_identity(request: Request, sicil: Optional[str] = None):
        """
        Mevcut aktif Kokpit oturumunu, kullanıcı profilini ve dışarıdan iletilmiş
        input/embedding yapılandırmalarını döndürür.
        """
        user = KOKPIT_IDENTITY_STORE.get("current_user")
        if sicil and str(sicil) in KOKPIT_IDENTITY_STORE.get("sessions_by_sicil", {}):
            sess = KOKPIT_IDENTITY_STORE["sessions_by_sicil"][str(sicil)]
            user = sess.get("user")

        return {
            "status": "success" if user else "waiting",
            "durum": "basarili",
            "authenticated": bool(user),
            "session_id": KOKPIT_IDENTITY_STORE.get("session_id"),
            "user": user,
            "token_present": bool(KOKPIT_IDENTITY_STORE.get("jwt_token")),
            "token_claims": KOKPIT_IDENTITY_STORE.get("token_claims"),
            "source": KOKPIT_IDENTITY_STORE.get("source"),
            "origin": KOKPIT_IDENTITY_STORE.get("origin"),
            "timestamp": KOKPIT_IDENTITY_STORE.get("timestamp"),
            "updated_at": KOKPIT_IDENTITY_STORE.get("updated_at"),
            "external_inputs": KOKPIT_IDENTITY_STORE.get("external_inputs") or {},
            "allowed_origins": ALLOWED_KOKPIT_ORIGINS,
            "history": KOKPIT_IDENTITY_STORE.get("history", [])[-10:]
        }

    @web_app.delete("/api/identity")
    @web_app.post("/api/identity/logout")
    async def api_clear_identity():
        """Aktif Kokpit kimlik oturumunu ve geçici dış inputları temizler."""
        KOKPIT_IDENTITY_STORE["current_user"] = None
        KOKPIT_IDENTITY_STORE["jwt_token"] = None
        KOKPIT_IDENTITY_STORE["token_claims"] = None
        KOKPIT_IDENTITY_STORE["status"] = "logged_out"
        KOKPIT_IDENTITY_STORE["external_inputs"] = {}
        return {
            "status": "success",
            "durum": "basarili",
            "message": "Oturum ve dış input verileri temizlendi."
        }

    @web_app.post("/api/external-input")
    async def api_receive_external_input(req: ExternalInputRequest, request: Request):
        """
        Dış sistemlerden veya Kokpit iframe üst penceresinden gelen her türlü veri, metin,
        hazır embedding vektörleri, model/API yapılandırması ve kimlik bilgilerini alır.
        `auto_compare=True` ise anında kıyaslamayı çalıştırıp Venn sonucunu da döndürür.
        """
        raw_dict = req.model_dump(exclude_none=True) if hasattr(req, "model_dump") else req.dict(exclude_none=True)
        origin = request.headers.get("origin")
        stored = store_kokpit_identity_and_inputs(raw_dict, origin=origin)

        comparison_result = None
        items_1 = req.items_1 or ([line.strip() for line in (req.text_1 or "").splitlines() if line.strip()] if req.text_1 else [])
        items_2 = req.items_2 or ([line.strip() for line in (req.text_2 or "").splitlines() if line.strip()] if req.text_2 else [])

        if req.auto_compare and items_1 and items_2:
            engine = (req.engine_mode or "local").lower()
            eff_key = req.api_key or KOKPIT_IDENTITY_STORE.get("jwt_token") or ""
            eff_url = req.base_url or "https://api.aselixai.aselsan.com.tr/v1"
            thr = req.threshold if req.threshold is not None else 0.80

            if engine == "embedding" or (req.embeddings_1 and req.embeddings_2):
                venn_res = compute_embedding_venn(
                    items_a=items_1,
                    items_b=items_2,
                    threshold=thr,
                    base_url=eff_url,
                    api_key=eff_key,
                    embedding_model=req.embedding_model or "text-embedding-3-small",
                    external_embeddings_a=req.embeddings_1,
                    external_embeddings_b=req.embeddings_2
                )
            else:
                venn_res = compute_fuzzy_venn(items_1, items_2, threshold=thr)

            jaccard_pct = round(venn_res["jaccard_similarity"] * 100, 1)
            comparison_result = {
                "durum": "basarili",
                "engine_mode": engine,
                "dosya_1_adi": req.name_1 or "1. Dış Veri",
                "dosya_2_adi": req.name_2 or "2. Dış Veri",
                "esik_orani": thr,
                "ozet_yorum": (
                    f"Dışarıdan alınan '{req.name_1}' ({len(items_1)} kayıt) ve '{req.name_2}' ({len(items_2)} kayıt) "
                    f"{engine} motoru ile kıyaslandı (Jaccard: %{jaccard_pct})."
                ),
                "venn_verileri": venn_res
            }

        return {
            "status": "success",
            "durum": "basarili",
            "session_id": stored["session_id"],
            "user": stored["user"],
            "external_inputs": stored["external_inputs"],
            "comparison_result": comparison_result
        }

    @web_app.get("/api/external-input")
    async def api_get_external_input():
        """Dışarıdan gönderilmiş son input ve yapılandırma paketini döndürür."""
        return {
            "status": "success",
            "durum": "basarili",
            "user": KOKPIT_IDENTITY_STORE.get("current_user"),
            "external_inputs": KOKPIT_IDENTITY_STORE.get("external_inputs") or {}
        }

    @web_app.get("/api/bakiye-siparis/items")
    def api_bakiye_siparis_items(
        proje: Optional[str] = None,
        search: Optional[str] = None
    ):
        """
        UGESUDKOKPİT veritabanındaki (`ud_kokpit` -> `teslimat_md.bakiye_siparis`)
        proje isimlerini, SD (satis_belgesi), Kalem (kalem_numarasi),
        Ürün Kısa Metni (urun_kisa_metni) ve Faz seçeneklerini (Dry-Run / FKT / SKT) döndürür.
        """
        return fetch_bakiye_siparis_items(proje=proje, search=search)

    @web_app.post("/api/bakiye-siparis/save-kpi")
    async def api_save_dryrun_fkt_skt_kpi(request: Request):
        """
        Kullanıcının inceleyip onayladığı (ve gerekirse 'Benzer Değil / Decline' ile ayırdığı)
        Venn analiz verilerini `teslimat_md.dryrun_fkt_skt_kpi` tablosuna kaydeder.
        """
        try:
            body = await request.json()
        except Exception:
            return JSONResponse({"durum": "hata", "mesaj": "Geçersiz JSON gövdesi."}, status_code=400)
        if not isinstance(body, dict):
            return JSONResponse({"durum": "hata", "mesaj": "JSON nesnesi bekleniyor."}, status_code=400)
        return save_dryrun_fkt_skt_kpi_record(body)

    @web_app.get("/api/bakiye-siparis/kpi-records")
    def api_get_dryrun_fkt_skt_kpi_records(
        limit: int = 100,
        proje: Optional[str] = None,
        sd: Optional[str] = None,
        kalem: Optional[str] = None,
        faaliyet_2_turu: Optional[str] = None,
        urun: Optional[str] = None,
        search: Optional[str] = None
    ):
        """
        `teslimat_md.dryrun_fkt_skt_kpi` tablosundaki kayıtlı KPI verilerini
        Proje Adı, SD, Kalem, 2. Faaliyet Türü (FKT/SKT), Ürün Kısa Metni ve Genel Arama filtreleriyle döndürür.
        """
        return fetch_dryrun_fkt_skt_kpi_records(
            limit=limit,
            proje=proje,
            sd=sd,
            kalem=kalem,
            faaliyet_2_turu=faaliyet_2_turu,
            urun=urun,
            search=search
        )

    @web_app.post("/api/bakiye-siparis/delete-kpi")
    @web_app.delete("/api/bakiye-siparis/kpi-records")
    async def api_delete_dryrun_fkt_skt_kpi(request: Request):
        """
        Yalnızca Admin yetkisine sahip kullanıcılar için `teslimat_md.dryrun_fkt_skt_kpi`
        tablosundan belirtilen bileşik anahtarlı kaydı siler.
        """
        try:
            body = await request.json()
        except Exception:
            return JSONResponse({"durum": "hata", "mesaj": "Geçersiz JSON gövdesi."}, status_code=400)
        if not isinstance(body, dict):
            return JSONResponse({"durum": "hata", "mesaj": "JSON nesnesi bekleniyor."}, status_code=400)

        requester_user = KOKPIT_IDENTITY_STORE.get("current_user")
        res = delete_dryrun_fkt_skt_kpi_record(body, requester_user=requester_user)
        if res.get("durum") == "yetkisiz":
            return JSONResponse(res, status_code=403)
        if res.get("durum") == "hata":
            return JSONResponse(res, status_code=400)
        return res

    @web_app.get("/api/bakiye-siparis/db-status")
    def api_get_database_status():
        """`.env` bağlantı ayarlarını, `bakiye_siparis` ve `dryrun_fkt_skt_kpi` tablolarının varlık/şema durumunu kontrol eder."""
        return verify_and_init_database_on_startup()

    @web_app.get("/api/bulk/template-excel")
    def api_download_bulk_template_excel(proje: Optional[str] = None):
        """
        `teslimat_md.bakiye_siparis` içindeki SD / Kalem seçeneklerini açılır liste (Data Validation)
        olarak barındıran örnek toplu veri girişi Excel (.xlsx) şablonunu indirir.
        """
        xlsx_bytes = generate_bulk_template_excel_bytes(proje=proje)
        safe_proj = f"_{proje.strip().replace(' ', '_')}" if proje and proje.strip() and proje.strip().upper() != "ALL" else ""
        fname = f"VennFlow_Uygunsuzluk_Sablonu{safe_proj}.xlsx"
        return Response(
            content=xlsx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f'attachment; filename="{fname}"'
            }
        )

    @web_app.post("/api/bulk/analyze-excel")
    async def api_bulk_analyze_excel(
        file: UploadFile = File(...),
        engine_mode: str = Form("local"),
        threshold: float = Form(0.80),
        default_sd: Optional[str] = Form(""),
        default_kalem: Optional[str] = Form(""),
        base_url: Optional[str] = Form("https://api.aselixai.aselsan.com.tr/v1"),
        api_key: Optional[str] = Form(""),
        embedding_model: Optional[str] = Form("text-embedding-3-small")
    ):
        """
        Doldurulan şablon Excel dosyasını okur, satırlardaki SD/Kalem bilgisini arka planda
        `teslimat_md.bakiye_siparis` ile eşleştirir ve her SD/Kalem grubu için Dry-Run vs FKT/SKT
        Venn kıyaslamasını üretir.
        """
        try:
            f_bytes = await file.read()
            eff_key = (api_key or "").strip() or (KOKPIT_IDENTITY_STORE.get("jwt_token") or "").strip()
            return analyze_bulk_uploaded_excel(
                file_bytes=f_bytes,
                filename=file.filename or "toplu_veri.xlsx",
                engine_mode=engine_mode,
                threshold=threshold,
                default_sd=default_sd or "",
                default_kalem=default_kalem or "",
                base_url=base_url or "https://api.aselixai.aselsan.com.tr/v1",
                api_key=eff_key,
                embedding_model=embedding_model or "text-embedding-3-small"
            )
        except Exception as e:
            return {"durum": "hata", "mesaj": f"Toplu Excel analizi sırasında hata oluştu: {str(e)}"}

    @web_app.post("/api/bulk/save-all-kpi")
    async def api_bulk_save_all_kpi(request: Request):
        """
        Toplu veri girişi sekmesinde incelenen ve onaylanan tüm SD/Kalem gruplarını
        tek seferde `teslimat_md.dryrun_fkt_skt_kpi` tablosuna kaydeder.
        """
        try:
            body = await request.json()
        except Exception:
            return JSONResponse({"durum": "hata", "mesaj": "Geçersiz JSON gövdesi."}, status_code=400)

        gruplar = body.get("gruplar") if isinstance(body, dict) else []
        if not isinstance(gruplar, list) or not gruplar:
            return JSONResponse({"durum": "hata", "mesaj": "Kaydedilecek grup listesi bulunamadı."}, status_code=400)

        saved_results = []
        errors = []
        for g in gruplar:
            if not isinstance(g, dict):
                continue
            res = save_dryrun_fkt_skt_kpi_record(g)
            if res.get("durum") in ("basarili", "uyari"):
                saved_results.append(res.get("kayit"))
            else:
                errors.append(res.get("mesaj"))

        return {
            "durum": "basarili" if saved_results else "hata",
            "kaydedilen_grup_sayisi": len(saved_results),
            "hata_sayisi": len(errors),
            "hatalar": errors,
            "mesaj": f"Toplam {len(saved_results)} adet SD/Kalem faaliyet kıyaslama grubu 'teslimat_md.dryrun_fkt_skt_kpi' tablosuna başarıyla aktarıldı.",
            "kayitlar": saved_results
        }

    @web_app.post("/api/aselix/embeddings")
    def api_aselix_embeddings(req: EmbeddingRequest):
        """Metin listesi için AselixAI / OpenAI uyumlu veya yerel N-gram vektör embedding üretir."""
        return call_aselix_embeddings(
            texts=req.texts,
            base_url=req.base_url or "https://api.aselixai.aselsan.com.tr/v1",
            api_key=req.api_key,
            model=req.model or "text-embedding-3-small"
        )

    @web_app.post("/api/compare/embeddings")
    def api_compare_embeddings(req: CompareEmbeddingsRequest):
        """İki metin listesini veya dışarıdan gelen hazır embedding vektörlerini kosinüs benzerliği ile kıyaslar."""
        thr = req.threshold if req.threshold is not None else 0.80
        venn_res = compute_embedding_venn(
            items_a=req.items_a,
            items_b=req.items_b,
            threshold=thr,
            base_url=req.base_url or "https://api.aselixai.aselsan.com.tr/v1",
            api_key=req.api_key,
            embedding_model=req.embedding_model or "text-embedding-3-small",
            external_embeddings_a=req.embeddings_a,
            external_embeddings_b=req.embeddings_b
        )
        jaccard_pct = round(venn_res["jaccard_similarity"] * 100, 1)
        summary_note = (
            f"'{req.doc1_name}' ({len(set(req.items_a))} kayıt) ile '{req.doc2_name}' ({len(set(req.items_b))} kayıt) "
            f"Embedding Vektör Uzayı ({venn_res.get('embedding_source')}) ile kıyaslandı. "
            f"Benzerlik eşiği %{round(thr * 100)} olarak uygulandı. "
            f"Toplam {venn_res['total_unique']} farklı uygunsuzluk tespit edildi. "
            f"{venn_res['intersection_count']} uygunsuzluk her iki grupta da ortaktır (Jaccard: %{jaccard_pct})."
        )
        return {
            "durum": "basarili",
            "engine_mode": "embedding",
            "dosya_1_adi": req.doc1_name,
            "dosya_2_adi": req.doc2_name,
            "esik_orani": thr,
            "ozet_yorum": summary_note,
            "venn_verileri": venn_res,
            "model_kullanilan": req.embedding_model
        }

    @web_app.post("/api/analyze")
    def api_analyze(req: AnalyzeRequest):
        return execute_analyze_venn(
            file_path=req.file_path or "",
            compare_by=req.compare_by,
            activity_name=req.activity_name,
            sheet_name=req.sheet_name,
            activity_column=req.activity_column,
            phase_column=req.phase_column,
            issue_column=req.issue_column
        )

    @web_app.get("/api/activities")
    def api_activities(file_path: Optional[str] = None, sheet_name: Optional[str] = None):
        return execute_list_activities_and_phases(file_path=file_path or "", sheet_name=sheet_name)

    @web_app.get("/api/preview")
    def api_preview(file_path: Optional[str] = None, max_rows: int = 5, sheet_name: Optional[str] = None):
        return execute_preview_excel_data(file_path=file_path or "", max_rows=max_rows, sheet_name=sheet_name)

    @web_app.post("/api/flow/test-connection")
    async def api_flow_test_connection(req: TestConnectionRequest):
        """Kullanıcının girdiği API Key ve Flow ID'yi test etmek için basit bir merhaba çağrısı yapar."""
        return call_aselsan_llm_suite_flow(
            base_url=req.base_url or "https://llmsuiteflow.aiplatform.aselsan.com.tr",
            flow_id=req.flow_id or "c16ecc83-ff64-4fc7-bb52-0f594beaf70a",
            api_key=req.api_key,
            tweaks_input="",
            prompt_text="Merhaba, bağlantı testi yapıyorum. Lütfen 1 cümle ile bağlandığını teyit et."
        )

    @web_app.post("/api/flow/compare")
    async def api_flow_compare(
        file_1: UploadFile = File(...),
        file_2: UploadFile = File(...),
        base_url: str = Form("https://llmsuiteflow.aiplatform.aselsan.com.tr"),
        flow_id: str = Form("c16ecc83-ff64-4fc7-bb52-0f594beaf70a"),
        api_key: Optional[str] = Form(""),
        tweaks_input: str = Form(""),
        column_name: Optional[str] = Form(None)
    ):
        """İki Excel dosyasını alır, yerel Venn matematiğini hesaplar ve ASELSAN Flow'u çağırarak analiz üretir."""
        file_1_bytes = await file_1.read()
        file_2_bytes = await file_2.read()

        comp_res = execute_compare_two_excels(
            file_1=file_1_bytes,
            file_2=file_2_bytes,
            file_1_name=file_1.filename or "1. Belge",
            file_2_name=file_2.filename or "2. Belge",
            column_name=column_name
        )

        if comp_res.get("durum") != "basarili":
            return comp_res

        v = comp_res["venn_verileri"]

        sample_common = "\n- ".join(v["intersection_items"][:10])
        sample_only_a = "\n- ".join(v["only_a_items"][:8])
        sample_only_b = "\n- ".join(v["only_b_items"][:8])

        prompt = (
            f"Ekteki iki Excel dosyasındaki uygunsuzluk verileri başarıyla karşılaştırıldı:\n\n"
            f"• 1. Belge ({file_1.filename}): Toplam {v['only_a_count'] + v['intersection_count']} uygunsuzluk.\n"
            f"• 2. Belge ({file_2.filename}): Toplam {v['only_b_count'] + v['intersection_count']} uygunsuzluk.\n"
            f"• Ortak Kesişim (Her iki belgede tekrarlayan): {v['intersection_count']} uygunsuzluk.\n"
            f"• Yalnızca 1. Belgede olan (2. belgede çözülmüş): {v['only_a_count']} uygunsuzluk.\n"
            f"• Yalnızca 2. Belgede olan (Yeni ortaya çıkan): {v['only_b_count']} uygunsuzluk.\n"
            f"• Jaccard Benzerlik İndeksi: %{round(v['jaccard_similarity'] * 100, 1)}\n\n"
            f"Örnek Ortak Uygunsuzluklar:\n- {sample_common or 'Yok'}\n\n"
            f"Örnek Yeni Ortaya Çıkanlar (Yalnızca 2. Belge):\n- {sample_only_b or 'Yok'}\n\n"
            f"Örnek Çözülenler (Yalnızca 1. Belge):\n- {sample_only_a or 'Yok'}\n\n"
            f"Lütfen bir Sistem Mühendisi ve Kalite Güvence Uzmanı olarak:\n"
            f"1. Kronik tekrarlayan ortak uygunsuzlukların kök neden risk analizini yap.\n"
            f"2. Yeni ortaya çıkan beklenmedik sapmaları değerlendir.\n"
            f"3. Kalite ve test ekiplerinin derhal alması gereken 3 somut önleyici aksiyon önerisi sun.\n"
            f"Yanıtını profesyonel, maddeler halinde ve yönetici özeti şeklinde yapılandır."
        )

        flow_res = call_aselsan_llm_suite_flow(
            base_url=base_url,
            flow_id=flow_id,
            api_key=api_key,
            tweaks_input=tweaks_input,
            uploaded_files=[
                (file_1.filename or "file_1.xlsx", file_1_bytes),
                (file_2.filename or "file_2.xlsx", file_2_bytes)
            ],
            prompt_text=prompt
        )

        comp_res["flow_analiz_ozeti"] = flow_res.get("cevap") or flow_res.get("mesaj") or ""
        comp_res["session_id"] = flow_res.get("session_id")
        comp_res["flow_durum"] = flow_res.get("durum")
        return comp_res

    @web_app.post("/api/flow/sample-compare")
    async def api_flow_sample_compare(req: SampleCompareRequest):
        """Hazır ornek_uygunsuzluklar.xlsx üzerinden örnek iki test fazını (Dry-run vs SKT) kıyaslar."""
        sample_path = os.path.abspath("ornek_uygunsuzluklar.xlsx")
        if not os.path.exists(sample_path):
            sample_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ornek_uygunsuzluklar.xlsx")

        if not os.path.exists(sample_path):
            return {"durum": "hata", "mesaj": "Örnek Excel dosyası bulunamadı."}

        df = pd.read_excel(sample_path)
        dry_df = df[df["Faz"].astype(str).str.contains("Dry", case=False, na=False)]
        skt_df = df[df["Faz"].astype(str).str.contains("SKT|Saha", case=False, na=False)]

        comp_res = execute_compare_two_excels(
            file_1=dry_df,
            file_2=skt_df,
            file_1_name="Ornek_Faz1_DryRun.xlsx",
            file_2_name="Ornek_Faz2_SKT.xlsx",
            column_name="Uygunsuzluk"
        )

        v = comp_res["venn_verileri"]
        prompt = (
            f"Ekteki iki test fazındaki uygunsuzluklar karşılaştırıldı:\n"
            f"Dry-run ({v['only_a_count'] + v['intersection_count']} uygunsuzluk) vs SKT ({v['only_b_count'] + v['intersection_count']} uygunsuzluk).\n"
            f"Ortak Kesişim: {v['intersection_count']} adet.\n"
            f"Yalnızca Dry-run'da olup SKT'de çözülen: {v['only_a_count']} adet.\n"
            f"SKT'de ilk kez ortaya çıkan: {v['only_b_count']} adet.\n"
            f"Lütfen kalite mühendisliği açısından bir değerlendirme özeti sun."
        )

        flow_res = call_aselsan_llm_suite_flow(
            base_url=req.base_url or "https://llmsuiteflow.aiplatform.aselsan.com.tr",
            flow_id=req.flow_id or "c16ecc83-ff64-4fc7-bb52-0f594beaf70a",
            api_key=req.api_key,
            tweaks_input=req.tweaks_input or "ChatInput-jugL4",
            prompt_text=prompt
        )

        comp_res["flow_analiz_ozeti"] = flow_res.get("cevap") or flow_res.get("mesaj") or ""
        comp_res["session_id"] = flow_res.get("session_id")
        comp_res["flow_durum"] = flow_res.get("durum")
        return comp_res

    @web_app.post("/api/flow/chat")
    async def api_flow_chat(req: FlowChatRequest):
        """Kullanıcının kıyaslanan veriler hakkında Flow Chatbot ile sohbet etmesini sağlar."""
        context_str = ""
        if req.context:
            context_str = (
                f"[Önceki Kıyaslama Bağlamı: {req.context.get('dosya_1')} ve {req.context.get('dosya_2')}. "
                f"Ortak Hatalar: {req.context.get('intersection_count')}, Yalnızca 1: {req.context.get('only_a_count')}, "
                f"Yalnızca 2: {req.context.get('only_b_count')}]\n\n"
            )

        full_prompt = f"{context_str}Kullanıcı Sorusu: {req.message}"

        flow_res = call_aselsan_llm_suite_flow(
            base_url=req.base_url or "https://llmsuiteflow.aiplatform.aselsan.com.tr",
            flow_id=req.flow_id or "c16ecc83-ff64-4fc7-bb52-0f594beaf70a",
            api_key=req.api_key,
            tweaks_input=req.tweaks_input or "ChatInput-jugL4",
            prompt_text=full_prompt,
            session_id=req.session_id
        )

        return flow_res

    @web_app.post("/api/compare/excel-preview")
    async def api_compare_excel_preview(file: UploadFile = File(...)):
        """
        Excel dosyasının sütunlarını ve ilk 5 satırını önizleme tablosu için döndürür.
        Ayrıca Excel içinde SD ve Kalem sütunları varsa teslimat_md.bakiye_siparis tablosundan
        Proje Adı, Ürün Kısa Metni ve Malzeme bilgilerini otomatik tespit edip döndürür.
        """
        try:
            content = await file.read()
            df = read_excel_safe(content)
            cols = [str(c) for c in df.columns]
            suggested = None
            try:
                _, _, suggested = resolve_columns(df)
            except Exception:
                pass
            if not suggested and cols:
                suggested = cols[0]

            # Excel içinde SD ve Kalem sütunları varsa bakiye_siparis ile otomatik eşleştir
            detected_bakiye = None
            detected_faz = None
            try:
                bakiye_res = fetch_bakiye_siparis_items()
                bakiye_items = bakiye_res.get("kayitlar") or []
                bakiye_lookup = {
                    (str(it.get("sd") or "").strip(), str(it.get("kalem") or "").strip()): it
                    for it in bakiye_items
                }
                for _, r in df.head(15).iterrows():
                    r_dict = r.to_dict()
                    sd_v, kal_v, fb_proj, fb_urun = _parse_sd_kalem_from_row(r_dict)
                    if sd_v and sd_v != "-" and kal_v and kal_v != "-":
                        matched = bakiye_lookup.get((sd_v, kal_v))
                        if matched:
                            detected_bakiye = matched
                        else:
                            detected_bakiye = {
                                "id": f"{sd_v}_{kal_v}",
                                "sd": sd_v,
                                "kalem": kal_v,
                                "sd_kalem": f"{sd_v} / {kal_v}",
                                "proje_adi": fb_proj or "Belirtilmemiş Proje",
                                "urun_kisa_metni": fb_urun or "-",
                                "malzeme": "",
                                "uges_proje_tanim": ""
                            }
                        # Faz sütunu varsa tespit et
                        norm_r = {normalize_col_name(str(k)): v for k, v in r_dict.items() if k is not None}
                        for cand in ["faaliyet_turu", "faaliyet", "faz", "test_fazi", "asama"]:
                            for k, v in norm_r.items():
                                if cand in k and v is not None and not pd.isna(v):
                                    fv = str(v).strip().upper()
                                    if "SKT" in fv:
                                        detected_faz = "SKT"
                                    elif "FKT" in fv:
                                        detected_faz = "FKT"
                                    elif "DRY" in fv:
                                        detected_faz = "Dry-Run"
                                    break
                            if detected_faz:
                                break
                        break
            except Exception:
                pass

            preview_df = df.head(5).astype(object).where(pd.notnull(df.head(5)), "")
            raw_preview_rows = preview_df.to_dict(orient="records")
            preview_rows = []
            for row_obj in raw_preview_rows:
                clean_row = {}
                for k, v in row_obj.items():
                    if v is None or (isinstance(v, float) and pd.isna(v)) or pd.isna(v):
                        clean_row[str(k)] = ""
                    else:
                        clean_row[str(k)] = v
                # Satırda SD ve Kalem varsa ve Proje/Ürün/Malzeme hücreleri boşsa otomatik doldur
                try:
                    r_sd, r_kal, _, _ = _parse_sd_kalem_from_row(clean_row)
                    if r_sd and r_sd != "-" and r_kal and r_kal != "-":
                        b_match = bakiye_lookup.get((r_sd, r_kal))
                        if b_match:
                            for col_k in list(clean_row.keys()):
                                nk = normalize_col_name(col_k)
                                if not str(clean_row[col_k]).strip():
                                    if "proje" in nk:
                                        clean_row[col_k] = b_match.get("proje_adi") or ""
                                    elif "urun" in nk or "kisa_metin" in nk:
                                        clean_row[col_k] = b_match.get("urun_kisa_metni") or ""
                                    elif "malzeme" in nk or "uges" in nk:
                                        clean_row[col_k] = b_match.get("malzeme") or b_match.get("uges_proje_tanim") or ""
                except Exception:
                    pass
                preview_rows.append(clean_row)

            return {
                "durum": "basarili",
                "dosya_adi": file.filename or "excel_belgesi.xlsx",
                "toplam_satir": len(df),
                "sutunlar": cols,
                "onerilen_sutun": suggested,
                "onizleme_satirlari": preview_rows,
                "tespit_edilen_bakiye": detected_bakiye,
                "tespit_edilen_faz": detected_faz
            }
        except Exception as e:
            return {
                "durum": "hata",
                "mesaj": f"Excel dosyası okunamadı: {str(e)}"
            }

    @web_app.post("/api/aselix/models")
    async def api_aselix_models(req: AselixModelsRequest):
        """Aselsan AselixAI API'sinden (GET /models) mevcut modelleri çeker."""
        return fetch_aselix_models(
            base_url=req.base_url or "https://api.aselixai.aselsan.com.tr/v1",
            api_key=req.api_key
        )

    @web_app.post("/api/aselix/chat")
    async def api_aselix_chat(req: AselixChatRequest):
        """Aselsan AselixAI (POST /chat/completions) üzerinden sohbet ve analiz yanıtı üretir."""
        msgs = []
        if req.context:
            ctx_info = (
                f"[Kıyaslanan Veriler: '{req.context.get('dosya_1', '1. Veri')}' vs '{req.context.get('dosya_2', '2. Veri')}'. "
                f"Ortak Kesişim: {req.context.get('intersection_count', 0)}, "
                f"Yalnızca 1: {req.context.get('only_a_count', 0)}, "
                f"Yalnızca 2: {req.context.get('only_b_count', 0)}, "
                f"Jaccard Benzerlik: %{round(float(req.context.get('jaccard_similarity', 0)) * 100, 1)}, "
                f"Uygulanan Benzerlik Eşiği: %{round(float(req.context.get('esik_orani', 0.8)) * 100)}]\n"
            )
            msgs.append({
                "role": "system",
                "content": (
                    "Sen ASELSAN sistem ve kalite güvence mühendisliği konusunda uzman kıdemli bir yapay zeka analiz asistanısın. "
                    f"Kullanıcının yaptığı son kıyaslama bağlamı:\n{ctx_info}\n"
                    "Soruları profesyonel, yapıcı, teknik ve somut aksiyon adımları içeren bir üslupla yanıtla."
                )
            })

        for m in req.messages:
            msgs.append(m)

        return call_aselix_chat(
            base_url=req.base_url or "https://api.aselixai.aselsan.com.tr/v1",
            api_key=req.api_key,
            model=req.model or "aselix-llm-v1",
            messages=msgs,
            temperature=req.temperature or 0.2
        )

    @web_app.post("/api/compare/run")
    async def api_compare_run(
        input_mode: str = Form("excel"),
        engine_mode: str = Form("local"),
        threshold: float = Form(0.80),
        file_1: Optional[UploadFile] = File(None),
        file_2: Optional[UploadFile] = File(None),
        col_1: Optional[str] = Form(None),
        col_2: Optional[str] = Form(None),
        text_1: Optional[str] = Form(None),
        text_2: Optional[str] = Form(None),
        name_1: Optional[str] = Form("1. Veri Seti"),
        name_2: Optional[str] = Form("2. Veri Seti"),
        base_url: Optional[str] = Form("https://api.aselixai.aselsan.com.tr/v1"),
        api_key: Optional[str] = Form(""),
        model: Optional[str] = Form("aselix-llm-v1"),
        meta_1: Optional[str] = Form(None),
        meta_2: Optional[str] = Form(None)
    ):
        """
        Hem Excel dosyası yükleme hem de Metin Kopyala-Yapıştır (Copy-Paste) için evrensel kıyaslama çalıştırır:
        1. Yerel String Distance / Fuzzy Matching (%50 - %100 threshold) ile çevrimdışı çalışır.
        2. İsteğe bağlı olarak Aselsan AselixAI API ile yönetici özeti, risk analizi ve aksiyon planı üretir.
        """
        items_1: List[str] = []
        items_2: List[str] = []

        doc1_name = name_1 or "1. Veri"
        doc2_name = name_2 or "2. Veri"
        chosen_col1 = col_1
        chosen_col2 = col_2
        parsed_meta_1 = parse_bakiye_meta(meta_1)
        if parsed_meta_1:
            parsed_meta_1["faz"] = "Dry-Run"
        parsed_meta_2 = parse_bakiye_meta(meta_2)
        if parsed_meta_2:
            parsed_meta_2["faz"] = "SKT" if str(parsed_meta_2.get("faz") or "").strip().upper() == "SKT" else "FKT"

        if input_mode == "excel":
            if not file_1 or not file_2:
                return {"durum": "hata", "mesaj": "Lütfen kıyaslanacak iki adet Excel dosyası yükleyin."}

            try:
                f1_bytes = await file_1.read()
                f2_bytes = await file_2.read()
                doc1_name = file_1.filename or "1. Belge"
                doc2_name = file_2.filename or "2. Belge"

                df1 = read_excel_safe(f1_bytes)
                df2 = read_excel_safe(f2_bytes)

                if not chosen_col1 or chosen_col1 == "auto" or chosen_col1 not in df1.columns:
                    try:
                        _, _, chosen_col1 = resolve_columns(df1)
                    except Exception:
                        chosen_col1 = df1.columns[0]

                if not chosen_col2 or chosen_col2 == "auto" or chosen_col2 not in df2.columns:
                    try:
                        _, _, chosen_col2 = resolve_columns(df2)
                    except Exception:
                        chosen_col2 = df2.columns[0]

                items_1 = df1[chosen_col1].dropna().astype(str).tolist()
                items_2 = df2[chosen_col2].dropna().astype(str).tolist()
            except Exception as e:
                return {"durum": "hata", "mesaj": f"Excel dosyaları işlenirken hata oluştu: {str(e)}"}

        else: # Metin Kopyala / Yapıştır
            if not text_1 or not text_1.strip() or not text_2 or not text_2.strip():
                return {"durum": "hata", "mesaj": "Lütfen her iki metin alanına da kıyaslanacak satırları girin."}

            items_1 = [line.strip() for line in (text_1 or "").splitlines() if line.strip()]
            items_2 = [line.strip() for line in (text_2 or "").splitlines() if line.strip()]
            chosen_col1 = "Metin Satırları"
            chosen_col2 = "Metin Satırları"

        if not items_1 or not items_2:
            return {"durum": "hata", "mesaj": "Kıyaslama yapılacak geçerli satır veya veri bulunamadı."}

        ai_summary = ""
        ai_durum = "kapali"
        effective_key = (api_key or "").strip() or (KOKPIT_IDENTITY_STORE.get("jwt_token") or "").strip()

        if engine_mode == "aselix" and effective_key:
            # 1 Dakikalık zaman aşımını aşmamak için 15'lik batch'ler halinde AselixAI'ye gönder
            batch_size = 15
            all_matched_pairs = []
            used_b = set()

            for i in range(0, len(items_1), batch_size):
                chunk_a = items_1[i:i + batch_size]
                chunk_res = compare_items_with_aselix_batch(
                    base_url=base_url or "https://api.aselixai.aselsan.com.tr/v1",
                    api_key=effective_key,
                    model=model or "aselix-llm-v1",
                    items_a_batch=chunk_a,
                    candidates_b=items_2,
                    threshold=threshold,
                    timeout=50  # 1 dakikayı kesinlikle aşmayacak koruma
                )
                for p in chunk_res.get("matched_pairs", []):
                    if p["item_b"] not in used_b:
                        used_b.add(p["item_b"])
                        all_matched_pairs.append(p)

            # Kalan eşleşmemişler için küme oluştur
            matched_a = {p["item_a"] for p in all_matched_pairs}
            clean_1 = [clean_text_content(x) for x in items_1 if clean_text_content(x)]
            clean_2 = [clean_text_content(x) for x in items_2 if clean_text_content(x)]
            unique_1 = list(dict.fromkeys(clean_1))
            unique_2 = list(dict.fromkeys(clean_2))

            only_a = [a for a in unique_1 if a not in matched_a]
            only_b = [b for b in unique_2 if b not in used_b]

            total_unique = len(all_matched_pairs) + len(only_a) + len(only_b)
            jaccard = round(len(all_matched_pairs) / total_unique, 4) if total_unique > 0 else 1.0

            venn_res = {
                "intersection_count": len(all_matched_pairs),
                "only_a_count": len(only_a),
                "only_b_count": len(only_b),
                "total_unique": total_unique,
                "jaccard_similarity": jaccard,
                "matched_pairs": all_matched_pairs,
                "intersection_items": [f"{p['item_a']} ↔ {p['item_b']} (%{p.get('similarity', 100)})" if p['item_a'] != p['item_b'] else p['item_a'] for p in all_matched_pairs],
                "only_a_items": sorted(only_a),
                "only_b_items": sorted(only_b),
                "threshold_used": threshold
            }

            jaccard_pct = round(jaccard * 100, 1)
            summary_note = (
                f"'{doc1_name}' ({len(unique_1)} kayıt) ile '{doc2_name}' ({len(unique_2)} kayıt) AselixAI API ile batchler halinde kıyaslandı. "
                f"Benzerlik eşiği %{round(threshold * 100)} uygulandı. Toplam {total_unique} farklı uygunsuzluk tespit edildi. "
                f"{len(all_matched_pairs)} uygunsuzluk her iki grupta da ortaktır (Jaccard: %{jaccard_pct}). "
                f"{len(only_a)} uygunsuzluk yalnızca 1. grupta, {len(only_b)} uygunsuzluk ise yalnızca 2. grupta yer almaktadır."
            )

            # Yönetici Özeti (timeout=50sn < 1dk)
            sample_common = "\n- ".join(venn_res["intersection_items"][:10]) or "Yok"
            sample_only_b = "\n- ".join(venn_res["only_b_items"][:8]) or "Yok"
            sample_only_a = "\n- ".join(venn_res["only_a_items"][:8]) or "Yok"

            system_msg = (
                "Sen ASELSAN sistem ve kalite güvence mühendisliği konusunda uzman kıdemli bir yapay zeka analiz asistanısın. "
                "İki test fazı veya veri seti arasındaki uygunsuzlukları kıyaslar, kök neden analizi yapar ve yönetici özeti sunarsın."
            )
            user_msg = (
                f"Kıyaslanan Belgeler: '{doc1_name}' vs '{doc2_name}'\n"
                f"Benzerlik Eşiği: %{round(threshold * 100)}\n\n"
                f"İstatistiksel Özet:\n"
                f"• Toplam Farklı Kayıt: {venn_res['total_unique']}\n"
                f"• Ortak Kesişim (Her iki belgede tekrarlayan): {venn_res['intersection_count']}\n"
                f"• Yalnızca 1. Belgede (Çözülmüş / giderilmiş): {venn_res['only_a_count']}\n"
                f"• Yalnızca 2. Belgede (Yeni ortaya çıkan): {venn_res['only_b_count']}\n"
                f"• Jaccard Benzerlik Oranı: %{jaccard_pct}\n\n"
                f"Örnek Ortak Uygunsuzluklar:\n- {sample_common}\n\n"
                f"Örnek Yeni Ortaya Çıkanlar (Yalnızca 2. Belge):\n- {sample_only_b}\n\n"
                f"Örnek Çözülenler (Yalnızca 1. Belge):\n- {sample_only_a}\n\n"
                f"Lütfen bir Sistem Mühendisi ve Kalite Güvence Uzmanı olarak:\n"
                f"1. Kronik tekrarlayan ortak uygunsuzlukların kök neden risk analizini yap.\n"
                f"2. Yeni ortaya çıkan beklenmedik sapmaları değerlendir.\n"
                f"3. Kalite ve test ekiplerinin derhal alması gereken 3 somut önleyici aksiyon önerisi sun.\n"
                f"Yanıtını profesyonel, maddeler halinde ve yönetici özeti şeklinde yapılandır."
            )
            chat_res = call_aselix_chat(
                base_url=base_url or "https://api.aselixai.aselsan.com.tr/v1",
                api_key=effective_key,
                model=model or "aselix-llm-v1",
                messages=[
                    {"role": "system", "content": system_msg},
                    {"role": "user", "content": user_msg}
                ],
                temperature=0.2,
                timeout=50  # 1 dakikayı kesinlikle aşmayacak koruma
            )
            if chat_res.get("durum") == "basarili":
                ai_summary = chat_res.get("cevap", "")
                ai_durum = "basarili"
            else:
                ai_summary = f"AselixAI API Uyarısı: {chat_res.get('mesaj', '')}"
                ai_durum = "uyari"

        elif engine_mode == "embedding":
            ext_in = KOKPIT_IDENTITY_STORE.get("external_inputs") or {}
            ext_emb_a = ext_in.get("embeddings_1") or ext_in.get("embeddings_a")
            ext_emb_b = ext_in.get("embeddings_2") or ext_in.get("embeddings_b")
            emb_model_name = ext_in.get("embedding_model") or (model if model and "embed" in model.lower() else "text-embedding-3-small")

            venn_res = compute_embedding_venn(
                items_a=items_1,
                items_b=items_2,
                threshold=threshold,
                base_url=base_url or "https://api.aselixai.aselsan.com.tr/v1",
                api_key=effective_key,
                embedding_model=emb_model_name,
                external_embeddings_a=ext_emb_a,
                external_embeddings_b=ext_emb_b
            )
            jaccard_pct = round(venn_res["jaccard_similarity"] * 100, 1)
            summary_note = (
                f"'{doc1_name}' ({len(set(items_1))} kayıt) ile '{doc2_name}' ({len(set(items_2))} kayıt) "
                f"Vektör Embedding ({venn_res.get('embedding_source')}) ile kıyaslandı. "
                f"Benzerlik eşiği %{round(threshold * 100)} olarak uygulandı. "
                f"Toplam {venn_res['total_unique']} farklı uygunsuzluk tespit edildi. "
                f"{venn_res['intersection_count']} uygunsuzluk her iki belgede de ortaktır (Jaccard Benzerlik Oranı: %{jaccard_pct}). "
                f"{venn_res['only_a_count']} uygunsuzluk yalnızca 1. grupta, {venn_res['only_b_count']} uygunsuzluk ise yalnızca 2. grupta yer almaktadır."
            )

        else: # engine_mode == "local"
            venn_res = compute_fuzzy_venn(items_1, items_2, threshold=threshold)
            jaccard_pct = round(venn_res["jaccard_similarity"] * 100, 1)
            summary_note = (
                f"'{doc1_name}' ({len(set(items_1))} kayıt) ile '{doc2_name}' ({len(set(items_2))} kayıt) yerel ML ile kıyaslandı. "
                f"Benzerlik eşiği %{round(threshold * 100)} olarak uygulandı. "
                f"Toplam {venn_res['total_unique']} farklı uygunsuzluk tespit edildi. "
                f"{venn_res['intersection_count']} uygunsuzluk her iki belgede de ortaktır (Jaccard Benzerlik Oranı: %{jaccard_pct}). "
                f"{venn_res['only_a_count']} uygunsuzluk yalnızca 1. grupta, {lenn_b if 'lenn_b' in locals() else venn_res['only_b_count']} uygunsuzluk ise yalnızca 2. grupta yer almaktadır."
            )

        meta_1_str = format_bakiye_meta_summary(parsed_meta_1)
        meta_2_str = format_bakiye_meta_summary(parsed_meta_2)
        if meta_1_str or meta_2_str:
            meta_prefix = f"[1. Belge: {meta_1_str or doc1_name}  ↔  2. Belge: {meta_2_str or doc2_name}] "
            summary_note = meta_prefix + summary_note

        return {
            "durum": "basarili",
            "input_mode": input_mode,
            "engine_mode": engine_mode,
            "dosya_1_adi": doc1_name,
            "dosya_2_adi": doc2_name,
            "belge_1_meta": parsed_meta_1,
            "belge_2_meta": parsed_meta_2,
            "karsilastirilan_sutun_1": chosen_col1,
            "karsilastirilan_sutun_2": chosen_col2,
            "esik_orani": threshold,
            "ozet_yorum": summary_note,
            "venn_verileri": venn_res,
            "flow_analiz_ozeti": ai_summary,
            "ai_durum": ai_durum,
            "model_kullanilan": model,
            "kokpit_user": KOKPIT_IDENTITY_STORE.get("current_user")
        }

    @web_app.post("/api/compare/extract-items")
    async def api_compare_extract_items(
        file_1: UploadFile = File(...),
        file_2: UploadFile = File(...),
        col_1: Optional[str] = Form("auto"),
        col_2: Optional[str] = Form("auto")
    ):
        """
        İki Excel dosyasından seçilen sütunların satırlarını liste olarak döndürür.
        Tarayıcı bu listeleri batch'lere bölüp her birini ayrı (<1 dk) istekle gönderir.
        """
        try:
            f1_bytes = await file_1.read()
            f2_bytes = await file_2.read()
            df1 = read_excel_safe(f1_bytes)
            df2 = read_excel_safe(f2_bytes)

            def _pick(df: pd.DataFrame, col: Optional[str]) -> str:
                if col and col != "auto" and col in df.columns:
                    return col
                try:
                    _, _, c = resolve_columns(df)
                    return c
                except Exception:
                    return df.columns[0]

            c1 = _pick(df1, col_1)
            c2 = _pick(df2, col_2)
            items_1 = [clean_text_content(x) for x in df1[c1].dropna().astype(str).tolist()]
            items_2 = [clean_text_content(x) for x in df2[c2].dropna().astype(str).tolist()]
            items_1 = [x for x in items_1 if x]
            items_2 = [x for x in items_2 if x]
            return {
                "durum": "basarili",
                "dosya_1_adi": file_1.filename or "1. Belge",
                "dosya_2_adi": file_2.filename or "2. Belge",
                "sutun_1": str(c1),
                "sutun_2": str(c2),
                "items_1": items_1,
                "items_2": items_2
            }
        except Exception as e:
            return {"durum": "hata", "mesaj": f"Excel dosyaları okunamadı: {str(e)}"}

    @web_app.post("/api/compare/batch-chunk")
    def api_compare_batch_chunk(req: BatchChunkRequest):
        """
        1 dakikalık zaman aşımını aşmayacak şekilde tek bir batch (10-20 satır) için
        AselixAI anlamsal kıyaslama isteği yürütür.
        Senkron 'def' -> FastAPI threadpool'da çalışır, bloklayan HTTP çağrısı sunucuyu dondurmaz.
        """
        t_start = time.time()
        _safe_log(f"[BATCH {req.batch_index}/{req.total_batches}] Başladı: {len(req.items_a_batch)} A satırı, {len(req.candidates_b)} B adayı")
        res = compare_items_with_aselix_batch(
            base_url=req.base_url or "https://api.aselixai.aselsan.com.tr/v1",
            api_key=req.api_key,
            model=req.model or "aselix-llm-v1",
            items_a_batch=req.items_a_batch,
            candidates_b=req.candidates_b,
            threshold=req.threshold or 0.8,
            timeout=req.timeout_sec or 50
        )
        elapsed = round(time.time() - t_start, 2)
        res["batch_index"] = req.batch_index
        res["total_batches"] = req.total_batches
        res["gecen_sure_sn"] = elapsed
        _safe_log(
            f"[BATCH {req.batch_index}/{req.total_batches}] Bitti: {elapsed}sn, kaynak={res.get('kaynak')}, "
            f"eşleşme={len(res.get('matched_pairs', []))}"
            + (f", HATA={res.get('api_hatasi')}" if res.get('api_hatasi') else "")
        )
        return res

    @web_app.post("/api/compare/batch-finalize")
    def api_compare_batch_finalize(req: BatchFinalizeRequest):
        """
        Tüm batch'lerden toplanan eşleşmeleri birleştirir, Venn şeması küme istatistiklerini
        hesaplar ve AselixAI ile hızlı bir yönetici özeti (< 1 dakika) üretir.
        """
        clean_a = [clean_text_content(x) for x in req.all_items_a if clean_text_content(x)]
        clean_b = [clean_text_content(x) for x in req.all_items_b if clean_text_content(x)]
        unique_a = list(dict.fromkeys(clean_a))
        unique_b = list(dict.fromkeys(clean_b))

        matched_pairs = req.matched_pairs or []
        used_a = set()
        used_b = set()
        dedup_matched = []

        for p in matched_pairs:
            ia = clean_text_content(p.get("item_a"))
            ib = clean_text_content(p.get("item_b"))
            if ia and ib and ia not in used_a and ib not in used_b:
                used_a.add(ia)
                used_b.add(ib)
                dedup_matched.append(p)

        only_a = [a for a in unique_a if a not in used_a]
        only_b = [b for b in unique_b if b not in used_b]

        total_unique = len(dedup_matched) + len(only_a) + len(only_b)
        jaccard = round(len(dedup_matched) / total_unique, 4) if total_unique > 0 else 1.0

        venn_res = {
            "intersection_count": len(dedup_matched),
            "only_a_count": len(only_a),
            "only_b_count": len(only_b),
            "total_unique": total_unique,
            "jaccard_similarity": jaccard,
            "matched_pairs": dedup_matched,
            "intersection_items": [f"{p['item_a']} ↔ {p['item_b']} (%{p.get('similarity', 100)})" if p['item_a'] != p['item_b'] else p['item_a'] for p in dedup_matched],
            "only_a_items": sorted(only_a),
            "only_b_items": sorted(only_b),
            "threshold_used": req.threshold
        }

        jaccard_pct = round(jaccard * 100, 1)
        summary_note = (
            f"'{req.doc1_name}' ({len(unique_a)} kayıt) ile '{req.doc2_name}' ({len(unique_b)} kayıt) kıyaslandı. "
            f"Benzerlik eşiği %{round(req.threshold * 100)} olarak uygulandı. "
            f"Toplam {total_unique} farklı uygunsuzluk tespit edildi. "
            f"{len(dedup_matched)} uygunsuzluk her iki grupta da ortaktır (Jaccard: %{jaccard_pct}). "
            f"{len(only_a)} uygunsuzluk yalnızca 1. grupta, {len(only_b)} uygunsuzluk ise yalnızca 2. grupta yer almaktadır."
        )

        ai_summary = ""
        ai_durum = "kapali"

        if req.need_executive_summary and req.api_key and req.api_key.strip():
            sample_common = "\n- ".join(venn_res["intersection_items"][:10]) or "Yok"
            sample_only_b = "\n- ".join(venn_res["only_b_items"][:8]) or "Yok"
            sample_only_a = "\n- ".join(venn_res["only_a_items"][:8]) or "Yok"

            system_msg = (
                "Sen ASELSAN sistem ve kalite güvence mühendisliği konusunda uzman kıdemli bir yapay zeka analiz asistanısın. "
                "Kıyaslanan iki test fazı için kök neden risk analizi ve somut aksiyon adımları sunarsın."
            )
            user_msg = (
                f"Kıyaslanan Belgeler: '{req.doc1_name}' vs '{req.doc2_name}'\n"
                f"İstatistiksel Özet:\n"
                f"• Toplam Farklı Kayıt: {total_unique}\n"
                f"• Ortak Kesişim: {len(dedup_matched)}\n"
                f"• Yalnızca 1. Belgede (Çözülen): {len(only_a)}\n"
                f"• Yalnızca 2. Belgede (Yeni Ortaya Çıkan): {len(only_b)}\n"
                f"• Jaccard Benzerlik Oranı: %{jaccard_pct}\n\n"
                f"Örnek Ortak Uygunsuzluklar:\n- {sample_common}\n\n"
                f"Örnek Yeni Ortaya Çıkanlar (Yalnızca 2. Belge):\n- {sample_only_b}\n\n"
                f"Örnek Çözülenler (Yalnızca 1. Belge):\n- {sample_only_a}\n\n"
                "Lütfen Kalite Güvence Uzmanı olarak:\n"
                "1. Kronik tekrarlayan ortak uygunsuzlukların kök neden risk analizini yap.\n"
                "2. Yeni ortaya çıkan beklenmedik sapmaları değerlendir.\n"
                "3. Kalite ekiplerinin derhal alması gereken 3 somut önleyici aksiyon önerisi sun.\n"
                "Kısa, net ve yönetici özeti şeklinde yapılandır."
            )
            chat_res = call_aselix_chat(
                base_url=req.base_url or "https://api.aselixai.aselsan.com.tr/v1",
                api_key=req.api_key,
                model=req.model or "aselix-llm-v1",
                messages=[
                    {"role": "system", "content": system_msg},
                    {"role": "user", "content": user_msg}
                ],
                temperature=0.2,
                timeout=50  # 1 dakikayı kesinlikle aşmayacak koruma
            )
            if chat_res.get("durum") == "basarili":
                ai_summary = chat_res.get("cevap", "")
                ai_durum = "basarili"
            else:
                ai_summary = f"AselixAI API Uyarısı: {chat_res.get('mesaj', '')}"
                ai_durum = "uyari"

        parsed_meta_1 = parse_bakiye_meta(req.meta_1)
        parsed_meta_2 = parse_bakiye_meta(req.meta_2)
        meta_1_str = format_bakiye_meta_summary(parsed_meta_1)
        meta_2_str = format_bakiye_meta_summary(parsed_meta_2)
        if meta_1_str or meta_2_str:
            summary_note = f"[1. Belge: {meta_1_str or req.doc1_name}  ↔  2. Belge: {meta_2_str or req.doc2_name}] " + summary_note

        return {
            "durum": "basarili",
            "engine_mode": "aselix",
            "dosya_1_adi": req.doc1_name,
            "dosya_2_adi": req.doc2_name,
            "belge_1_meta": parsed_meta_1,
            "belge_2_meta": parsed_meta_2,
            "karsilastirilan_sutun_1": req.col1_name,
            "karsilastirilan_sutun_2": req.col2_name,
            "esik_orani": req.threshold,
            "ozet_yorum": summary_note,
            "venn_verileri": venn_res,
            "flow_analiz_ozeti": ai_summary,
            "ai_durum": ai_durum,
            "model_kullanilan": req.model
        }

    # --------------------------------------------------------------------------
    # MODEL CONTEXT PROTOCOL (MCP) JSON-RPC 2.0 HTTP İŞLEYİCİSİ
    # --------------------------------------------------------------------------
    async def process_single_mcp_jsonrpc(payload: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Tek bir MCP JSON-RPC 2.0 isteğini işler."""
        if not isinstance(payload, dict):
            return {
                "jsonrpc": "2.0",
                "id": None,
                "error": {"code": -32600, "message": "Invalid Request: Payload must be a JSON object"}
            }

        method = payload.get("method")
        params = payload.get("params", {}) or {}
        has_id = "id" in payload
        req_id = payload.get("id")

        # 1. MCP Başlatma (Handshake)
        if method == "initialize":
            res = {
                "protocolVersion": "2024-11-05",
                "capabilities": {
                    "tools": {"listChanged": False},
                    "resources": {"subscribe": False, "listChanged": False},
                    "prompts": {"listChanged": False}
                },
                "serverInfo": {
                    "name": "venn-nonconformity-server",
                    "version": "1.2.0"
                }
            }
            return {"jsonrpc": "2.0", "id": req_id, "result": res}

        # 2. İstemci Bildirimi (Notification - yanıt verilmez)
        elif method in ["notifications/initialized", "initialized"]:
            return None if not has_id else {"jsonrpc": "2.0", "id": req_id, "result": {}}

        # 3. Ping / Logging
        elif method in ["ping", "logging/setLevel"]:
            return {"jsonrpc": "2.0", "id": req_id, "result": {}}

        # 4. Araç Listesi Keşfi (tools/list)
        elif method == "tools/list":
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "tools": [
                        {
                            "name": "analyze_venn_nonconformities",
                            "description": (
                                "Excel dosyasındaki test aşamalarını (Dry-run, SKT, FKT) veya faaliyetleri/illeri "
                                "Venn Şeması küme matematiği ile analiz eder. Hangi uygunsuzlukların sadece Dry-run'da "
                                "olduğunu, hangilerinin SKT ve FKT'de ortak veya kronik tekrarlayan olduğunu hesaplar."
                            ),
                            "inputSchema": {
                                "type": "object",
                                "properties": {
                                    "compare_by": {
                                        "type": "string",
                                        "enum": ["phases", "activities"],
                                        "default": "phases",
                                        "description": "'phases' (Dry-run, SKT, FKT arası) veya 'activities' (şehirler/faaliyetler arası)"
                                    },
                                    "activity_name": {
                                        "type": "string",
                                        "description": "Filtrelenecek faaliyet veya il adı (örn: 'kayseri', 'izmir'). Belirtilmezse tüm dosya analiz edilir."
                                    },
                                    "file_path": {
                                        "type": "string",
                                        "description": "Excel dosya yolu (boş bırakılırsa sunucudaki Uygunsuzluk.xlsx otomatik kullanılır)"
                                    }
                                },
                                "required": []
                            }
                        },
                        {
                            "name": "list_activities_and_phases",
                            "description": "Excel dosyasında mevcut olan faaliyet adlarını, illeri, test fazlarını ve kayıt sayılarını listeler.",
                            "inputSchema": {
                                "type": "object",
                                "properties": {
                                    "file_path": {
                                        "type": "string",
                                        "description": "Excel dosya yolu (boş bırakılırsa Uygunsuzluk.xlsx kullanılır)"
                                    }
                                },
                                "required": []
                            }
                        },
                        {
                            "name": "preview_excel_data",
                            "description": "Excel dosyasının sütunlarını ve ilk birkaç örnek satırını incelemek için kullanılır.",
                            "inputSchema": {
                                "type": "object",
                                "properties": {
                                    "file_path": {
                                        "type": "string",
                                        "description": "Excel dosya yolu (boş bırakılırsa Uygunsuzluk.xlsx kullanılır)"
                                    },
                                    "max_rows": {
                                        "type": "integer",
                                        "default": 5,
                                        "description": "Gösterilecek örnek satır sayısı"
                                    }
                                },
                                "required": []
                            }
                        },
                        {
                            "name": "compare_two_excels_venn",
                            "description": "İki farklı Excel dosyasını (örneğin iki ayrı test aşaması veya versiyonu) karşılaştırarak Venn şeması küme analizi (yalnızca A, yalnızca B, ortak kesişim ve Jaccard benzerliği) yapar.",
                            "inputSchema": {
                                "type": "object",
                                "properties": {
                                    "file_path_a": {
                                        "type": "string",
                                        "description": "1. Excel dosyasının mutlak yolu"
                                    },
                                    "file_path_b": {
                                        "type": "string",
                                        "description": "2. Excel dosyasının mutlak yolu"
                                    },
                                    "column_name": {
                                        "type": "string",
                                        "description": "Kıyaslanacak uygunsuzluk sütun adı (belirtilmezse otomatik tespit edilir)"
                                    }
                                },
                                "required": ["file_path_a", "file_path_b"]
                            }
                        }
                    ]
                }
            }

        # 5. Araç Çağrısı (tools/call)
        elif method == "tools/call":
            tool_name = params.get("name")
            args = params.get("arguments", {}) or {}

            # Varsayılan dosya yolu güvencesi (LLM dosya yolunu belirtmezse Uygunsuzluk.xlsx kullan)
            if not args.get("file_path"):
                args["file_path"] = ""

            try:
                if tool_name in ["analyze_venn_nonconformities", "analyze_venn"]:
                    res = execute_analyze_venn(**args)
                elif tool_name in ["list_activities_and_phases", "list_activities"]:
                    res = execute_list_activities_and_phases(**args)
                elif tool_name in ["preview_excel_data", "preview_data"]:
                    res = execute_preview_excel_data(**args)
                elif tool_name in ["compare_two_excels_venn", "compare_two_excels", "compare_excels"]:
                    f_a = args.get("file_path_a") or args.get("file_1") or ""
                    f_b = args.get("file_path_b") or args.get("file_2") or ""
                    c_name = args.get("column_name")
                    res = execute_compare_two_excels(file_1=f_a, file_2=f_b, column_name=c_name)
                else:
                    return {
                        "jsonrpc": "2.0",
                        "id": req_id,
                        "error": {"code": -32601, "message": f"Bilinmeyen MCP aracı: '{tool_name}'"}
                    }


                return {
                    "jsonrpc": "2.0",
                    "id": req_id,
                    "result": {
                        "content": [
                            {
                                "type": "text",
                                "text": json.dumps(res, ensure_ascii=False, indent=2)
                            }
                        ],
                        "isError": res.get("durum") == "hata"
                    }
                }
            except Exception as call_err:
                return {
                    "jsonrpc": "2.0",
                    "id": req_id,
                    "result": {
                        "content": [
                            {
                                "type": "text",
                                "text": f"Araç çalıştırma hatası: {str(call_err)}"
                            }
                        ],
                        "isError": True
                    }
                }

        # 6. Kaynaklar ve İstemler (Resources / Prompts) Boş Dönüşü
        elif method == "resources/list":
            return {"jsonrpc": "2.0", "id": req_id, "result": {"resources": []}}

        elif method == "resources/templates/list":
            return {"jsonrpc": "2.0", "id": req_id, "result": {"resourceTemplates": []}}

        elif method == "prompts/list":
            return {"jsonrpc": "2.0", "id": req_id, "result": {"prompts": []}}

        # Bilinmeyen metot
        else:
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "error": {"code": -32601, "message": f"Method '{method}' not found"}
            }

    async def handle_mcp_http_request(request: Request):
        """MCP HTTP isteklerini (tekil ve toplu) ele alır."""
        try:
            body = await request.json()
        except Exception:
            body_bytes = await request.body()
            try:
                body = json.loads(body_bytes.decode("utf-8"))
            except Exception as parse_err:
                return Response(
                    content=json.dumps({
                        "jsonrpc": "2.0",
                        "id": None,
                        "error": {"code": -32700, "message": f"JSON parse error: {str(parse_err)}"}
                    }),
                    status_code=400,
                    media_type="application/json"
                )

        if isinstance(body, list):
            responses = []
            for item in body:
                resp = await process_single_mcp_jsonrpc(item)
                if resp is not None:
                    responses.append(resp)
            return Response(
                content=json.dumps(responses, ensure_ascii=False),
                media_type="application/json"
            )
        else:
            resp = await process_single_mcp_jsonrpc(body)
            if resp is None:
                return Response(status_code=204)
            return Response(
                content=json.dumps(resp, ensure_ascii=False),
                media_type="application/json"
            )

    @web_app.post("/mcp")
    @web_app.post("/mcp/")
    @web_app.post("/v1/mcp")
    @web_app.post("/")
    async def mcp_all_post_endpoints(request: Request):
        return await handle_mcp_http_request(request)

except Exception as import_err:
    web_app = None


# ==============================================================================
# CLI ve Sunucu Başlatma
# ==============================================================================

def main():
    load_env_file_vars()
    srv_cfg = get_server_config()
    default_host = srv_cfg["host"]
    default_port = srv_cfg["port"]

    parser = argparse.ArgumentParser(
        description="Venn Diagram Nonconformity RAG & HTTP Server (Intranet Compatible)"
    )
    parser.add_argument(
        "--transport",
        choices=["http", "stdio", "sse"],
        default="http",
        help="Protokol: 'http' (Ağdaki herkese açık REST API - Önerilen), 'stdio' veya 'sse'."
    )
    parser.add_argument(
        "--host",
        default=None,
        help=f"Ağ adresi (varsayılan: .env içindeki APP_HOST/HOST veya '{default_host}')."
    )
    parser.add_argument(
        "--port",
        type=int,
        default=None,
        help=f"Port numarası (varsayılan: .env içindeki APP_PORT/PORT veya {default_port})."
    )

    args = parser.parse_args()

    effective_host = args.host if args.host is not None else default_host
    effective_port = args.port if args.port is not None else default_port

    # 1. Ağdaki Herkese Açık Standart HTTP Modu (Önerilen)
    if args.transport in ["http", "sse"]:
        import uvicorn
        sys.stderr.write(
            f"\n"
            f"===============================================================================\n"
            f"[HTTP SUNUCUSU] Agdaki Herkese Acik Olarak Baslatiliyor!\n"
            f"Adres: http://{effective_host}:{effective_port}\n"
            f"Tarayicidan erisim / Dokumantasyon: http://<SANAL_PC_IP>:{effective_port}/docs\n"
            f"Yapilandirma: .env (APP_PORT={default_port}, APP_HOST={default_host})\n"
            f"Durdurmak icin CTRL+C tuslarina basin.\n"
            f"===============================================================================\n\n"
        )
        uvicorn.run(web_app or "app:web_app", host=effective_host, port=effective_port, reload=False)

    # 2. Yerel stdio Modu (Aynı PC'de LLM istemcisi için)
    elif args.transport == "stdio":
        try:
            from mcp.server.mcpserver import MCPServer
            mcp_srv = MCPServer("VennRAGServer")
            @mcp_srv.tool()
            def analyze_venn(file_path: str, compare_by: str = "phases", activity_name: Optional[str] = None) -> str:
                return analyze_venn_nonconformities(file_path, compare_by, activity_name)
            mcp_srv.run(transport="stdio")
        except Exception as e:
            sys.stderr.write(f"[HATA] stdio modu baslatilamadi: {e}\n")
            sys.exit(1)


if __name__ == "__main__":
    main()
