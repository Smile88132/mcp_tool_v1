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
        return {
            "durum": "bilgi",
            "session_id": session_id,
            "cevap": (
                "### 📌 Matematiksel Venn Şeması ve Küme Analizi Hazır!\n\n"
                "İki Excel dosyasındaki uygunsuzluklar başarıyla karşılaştırıldı ve görsel Venn şeması oluşturuldu.\n\n"
                "**ASELSAN LLM Suite Flow Bağlantısı:**\n"
                f"- **Hedef Akış:** `{base_url}/api/v1/run/{flow_id}`\n\n"
                "Akıllı yapay zeka yorumu, kök neden analizi ve risk sınıflandırmasını almak için sağ üst köşedeki "
                "**⚙️ Ayarlar** butonuna tıklayarak **API Key (x-api-key)** bilginizi kaydedebilirsiniz."
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
    küçük bir batch (10-15 kayıt) için AselixAI API'si üzerinden anlamsal kıyaslama yapar.
    """
    if not items_a_batch or not candidates_b:
        return {
            "durum": "basarili",
            "matched_pairs": [],
            "unmatched_a": items_a_batch,
            "kaynak": "bos_liste"
        }

    # Eğer API Key yoksa yerel fuzzy matching kullan
    if not api_key or not api_key.strip():
        local_res = compute_fuzzy_venn(items_a_batch, candidates_b, threshold=threshold)
        return {
            "durum": "basarili",
            "matched_pairs": local_res["matched_pairs"],
            "unmatched_a": local_res["only_a_items"],
            "kaynak": "yerel_ml"
        }

    # Liste formatını hazırla
    list_a_text = "\n".join([f"{i+1}. {clean_text_content(item)}" for i, item in enumerate(items_a_batch)])
    list_b_text = "\n".join([f"{j+1}. {clean_text_content(cand)}" for j, cand in enumerate(candidates_b)])

    system_prompt = (
        "Sen ASELSAN Sistem ve Kalite Güvence mühendisliği uzmanı bir yapay zeka analiz asistanısın. "
        "Görevin, iki test fazından alınan uygunsuzluk maddelerini karşılaştırmak ve anlamsal olarak aynı "
        "veya eşdeğer arıza/sapmayı ifade edenleri eşleştirmektir. "
        "Farklı kelimeler veya eş anlamlı terimler kullanılmış olsa bile aynı teknik problemi ifade ediyorsa eşleştir. "
        "Yanıtını YALNIZCA geçerli bir JSON dizisi olarak üret, başka hiçbir açıklama metni ekleme."
    )

    user_prompt = (
        f"Aşağıdaki Liste A'da yer alan {len(items_a_batch)} adet uygunsuzluğu incele. "
        "Her bir Liste A maddesi için Liste B'deki adaylar arasında anlamsal olarak aynı teknik arızayı "
        "ifade eden bir karşılık var mı değerlendir.\n\n"
        "ÇIKTI FORMATI (Yalnızca aşağıdaki JSON şemasına uygun JSON dizisi döndür):\n"
        "[\n"
        "  {\n"
        '    "item_a": "1. listedeki tam metin",\n'
        '    "item_b": "2. listedeki eşleşen tam metin (eşleşme yoksa null)",\n'
        '    "is_match": true veya false,\n'
        '    "similarity": 50 ile 100 arasında benzerlik yüzdesi (eşleşme yoksa 0),\n'
        '    "reason": "Kısa eşleşme gerekçesi"\n'
        "  }\n"
        "]\n\n"
        f"--- LİSTE A ({len(items_a_batch)} Kayıt) ---\n{list_a_text}\n\n"
        f"--- LİSTE B ({len(candidates_b)} Aday Kayıt) ---\n{list_b_text}\n"
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

    if chat_res.get("durum") != "basarili":
        # API zaman aşımı veya bağlantı hatası durumunda yerel fuzzy distance ile kurtar
        local_res = compute_fuzzy_venn(items_a_batch, candidates_b, threshold=threshold)
        return {
            "durum": "uyari",
            "api_hatasi": chat_res.get("mesaj"),
            "matched_pairs": local_res["matched_pairs"],
            "unmatched_a": local_res["only_a_items"],
            "kaynak": "yerel_fallback"
        }

    raw_output = chat_res.get("cevap", "")
    parsed_matches = extract_json_array_from_llm(raw_output)

    matched_pairs = []
    matched_a = set()
    matched_b = set()
    threshold_pct = threshold * 100

    for m in parsed_matches:
        if not isinstance(m, dict):
            continue
        item_a = str(m.get("item_a") or "").strip()
        item_b = str(m.get("item_b") or "").strip()
        is_match = bool(m.get("is_match"))
        sim = float(m.get("similarity") or 0.0)

        if is_match and item_b and item_b.lower() not in ["null", "none", "yok"] and sim >= threshold_pct:
            # item_b'yi candidates_b içindeki tam metin ile eşle
            matched_cand = None
            for cand in candidates_b:
                if normalize_str(cand) == normalize_str(item_b) or cand in item_b or item_b in cand:
                    matched_cand = cand
                    break
            if not matched_cand:
                matched_cand = item_b

            matched_pairs.append({
                "item_a": item_a,
                "item_b": matched_cand,
                "similarity": round(sim, 1),
                "match_type": f"AI Semantik (%{round(sim, 1)})",
                "ai_reason": str(m.get("reason") or "AselixAI anlamsal eşleşmesi")
            })
            matched_a.add(item_a)
            matched_b.add(matched_cand)

    # LLM'in eşleştiremediği veya JSON'da eksik kalanlar için fuzzy mesafe kontrolü
    unmatched_a = [a for a in items_a_batch if a not in matched_a]
    remaining_b = [b for b in candidates_b if b not in matched_b]

    if unmatched_a and remaining_b:
        local_fallback = compute_fuzzy_venn(unmatched_a, remaining_b, threshold=threshold)
        for pair in local_fallback["matched_pairs"]:
            matched_pairs.append(pair)
            matched_a.add(pair["item_a"])
        unmatched_a = local_fallback["only_a_items"]

    return {
        "durum": "basarili",
        "matched_pairs": matched_pairs,
        "unmatched_a": unmatched_a,
        "kaynak": "aselix_ai"
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
            "ASELSAN AI Platform LLM Suite Flow (Langflow) entegrasyonu ile akıllı yapay zeka yorumları üretir."
        ),
        version="2.0.0"
    )

    web_app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    static_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
    if os.path.exists(static_dir):
        web_app.mount("/static", StaticFiles(directory=static_dir), name="static")

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

    @web_app.get("/", response_class=HTMLResponse)
    async def index(request: Request):
        accept_header = request.headers.get("accept", "")
        if "application/json" in accept_header and "text/html" not in accept_header:
            return JSONResponse({
                "servis": "VennFlow AI - Excel Kıyaslama ve Venn Analiz Sunucusu",
                "durum": "aktif",
                "ag_erisim_adresi": "http://0.0.0.0:8000",
                "mcp_http_endpoint": "/mcp",
                "interaktif_swagger_arayuzu": "/docs",
                "web_arayuzu": "http://0.0.0.0:8000/"
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
        """Excel dosyasının sütunlarını ve ilk 5 satırını önizleme tablosu için döndürür."""
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

            preview_df = df.head(5).copy()
            preview_rows = preview_df.where(pd.notnull(preview_df), None).to_dict(orient="records")

            return {
                "durum": "basarili",
                "dosya_adi": file.filename or "excel_belgesi.xlsx",
                "toplam_satir": len(df),
                "sutunlar": cols,
                "onerilen_sutun": suggested,
                "onizleme_satirlari": preview_rows
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
        model: Optional[str] = Form("aselix-llm-v1")
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

        if engine_mode == "aselix" and api_key and api_key.strip():
            # 1 Dakikalık zaman aşımını aşmamak için 15'lik batch'ler halinde AselixAI'ye gönder
            batch_size = 15
            all_matched_pairs = []
            used_b = set()

            for i in range(0, len(items_1), batch_size):
                chunk_a = items_1[i:i + batch_size]
                chunk_res = compare_items_with_aselix_batch(
                    base_url=base_url or "https://api.aselixai.aselsan.com.tr/v1",
                    api_key=api_key,
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
                api_key=api_key,
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

        else: # engine_mode == "local"
            venn_res = compute_fuzzy_venn(items_1, items_2, threshold=threshold)
            jaccard_pct = round(venn_res["jaccard_similarity"] * 100, 1)
            summary_note = (
                f"'{doc1_name}' ({len(set(items_1))} kayıt) ile '{doc2_name}' ({len(set(items_2))} kayıt) yerel ML ile kıyaslandı. "
                f"Benzerlik eşiği %{round(threshold * 100)} olarak uygulandı. "
                f"Toplam {venn_res['total_unique']} farklı uygunsuzluk tespit edildi. "
                f"{venn_res['intersection_count']} uygunsuzluk her iki belgede de ortaktır (Jaccard Benzerlik Oranı: %{jaccard_pct}). "
                f"{venn_res['only_a_count']} uygunsuzluk yalnızca 1. grupta, {venn_res['only_b_count']} uygunsuzluk ise yalnızca 2. grupta yer almaktadır."
            )

        return {
            "durum": "basarili",
            "input_mode": input_mode,
            "engine_mode": engine_mode,
            "dosya_1_adi": doc1_name,
            "dosya_2_adi": doc2_name,
            "karsilastirilan_sutun_1": chosen_col1,
            "karsilastirilan_sutun_2": chosen_col2,
            "esik_orani": threshold,
            "ozet_yorum": summary_note,
            "venn_verileri": venn_res,
            "flow_analiz_ozeti": ai_summary,
            "ai_durum": ai_durum,
            "model_kullanilan": model
        }

    @web_app.post("/api/compare/batch-chunk")
    async def api_compare_batch_chunk(req: BatchChunkRequest):
        """
        1 dakikalık zaman aşımını aşmayacak şekilde tek bir batch (10-15 satır) için
        AselixAI anlamsal kıyaslama isteği yürütür.
        """
        t_start = time.time()
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
        return res

    @web_app.post("/api/compare/batch-finalize")
    async def api_compare_batch_finalize(req: BatchFinalizeRequest):
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
            ia = p.get("item_a")
            ib = p.get("item_b")
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

        return {
            "durum": "basarili",
            "engine_mode": "aselix",
            "dosya_1_adi": req.doc1_name,
            "dosya_2_adi": req.doc2_name,
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
        default="0.0.0.0",
        help="Ağ adresi (varsayılan: 0.0.0.0, intranet üzerindeki herkese açık)."
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port numarası (varsayılan: 8000)."
    )

    args = parser.parse_args()

    # 1. Ağdaki Herkese Açık Standart HTTP Modu (Önerilen)
    if args.transport in ["http", "sse"]:
        import uvicorn
        sys.stderr.write(
            f"\n"
            f"===============================================================================\n"
            f"[HTTP SUNUCUSU] Agdaki Herkese Acik Olarak Baslatiliyor!\n"
            f"Adres: http://{args.host}:{args.port}\n"
            f"Tarayicidan erisim / Dokumantasyon: http://<SANAL_PC_IP>:{args.port}/docs\n"
            f"Durdurmak icin CTRL+C tuslarina basin.\n"
            f"===============================================================================\n\n"
        )
        uvicorn.run(web_app or "app:web_app", host=args.host, port=args.port, reload=False)

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
