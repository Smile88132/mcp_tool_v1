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
import json
import argparse
import unicodedata
from typing import Any, Dict, List, Optional, Set, Tuple
import pandas as pd

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


def read_excel_safe(file_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """Excel dosyasını güvenli şekilde okur."""
    if not os.path.isabs(file_path):
        file_path = os.path.abspath(file_path)

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Belirtilen dosya bulunamadı: '{file_path}'")

    if not os.path.isfile(file_path):
        raise ValueError(f"Belirtilen yol bir dosya değil: '{file_path}'")

    valid_extensions = (".xlsx", ".xls", ".xlsm")
    if not file_path.lower().endswith(valid_extensions):
        raise ValueError(f"Geçersiz dosya uzantısı. Desteklenen uzantılar: {valid_extensions}. Verilen: '{file_path}'")

    try:
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)
    except Exception as e:
        raise RuntimeError(f"Excel dosyası okunurken hata oluştu: {str(e)}")

    if df.empty:
        raise ValueError(f"Excel dosyası ('{file_path}') boş veri içeriyor.")

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


# ==============================================================================
# AĞDAKİ HERKESE AÇIK HTTP WEB SERVİSİ (FastAPI / Starlette)
# ==============================================================================

try:
    from fastapi import FastAPI
    from pydantic import BaseModel, Field

    web_app = FastAPI(
        title="Venn Diagram Nonconformity Analysis API",
        description=(
            "Intranet ortamında Excel verilerini Venn Şeması küme matematiği ile analiz eden REST API.\n\n"
            "Ağdaki tüm LLM'ler, RAG sistemleri ve web tarayıcıları doğrudan bağlanabilir."
        ),
        version="1.2.0"
    )

    class AnalyzeRequest(BaseModel):
        file_path: str = Field(..., description="Excel dosyasının tam dosya yolu (örn: C:\\...\\Uygunsuzluk.xlsx)")
        compare_by: str = Field("phases", description="'phases' (fazlar arası) veya 'activities' (faaliyetler/iller arası)")
        activity_name: Optional[str] = Field(None, description="Yalnızca belirli bir ili/faaliyeti analiz etmek için (örn: 'kayseri')")
        sheet_name: Optional[str] = Field(None, description="Opsiyonel çalışma sayfası adı")
        activity_column: Optional[str] = Field(None, description="Opsiyonel özel faaliyet/il sütunu")
        phase_column: Optional[str] = Field(None, description="Opsiyonel özel faz sütunu")
        issue_column: Optional[str] = Field(None, description="Opsiyonel özel uygunsuzluk sütunu")

    @web_app.get("/")
    def index():
        return {
            "servis": "Venn Şeması Uygunsuzluk Analiz Servisi (Intranet)",
            "durum": "aktif",
            "ag_erisim_adresi": "http://0.0.0.0:8000",
            "interaktif_swagger_arayuzu": "/docs",
            "endpoints": {
                "POST /api/analyze": "Venn Şeması küme analizi",
                "GET /api/activities": "Mevcut faaliyetler ve fazlar",
                "GET /api/preview": "Excel önizleme"
            }
        }

    @web_app.post("/api/analyze")
    def api_analyze(req: AnalyzeRequest):
        return execute_analyze_venn(
            file_path=req.file_path,
            compare_by=req.compare_by,
            activity_name=req.activity_name,
            sheet_name=req.sheet_name,
            activity_column=req.activity_column,
            phase_column=req.phase_column,
            issue_column=req.issue_column
        )

    @web_app.get("/api/activities")
    def api_activities(file_path: str, sheet_name: Optional[str] = None):
        return execute_list_activities_and_phases(file_path=file_path, sheet_name=sheet_name)

    @web_app.get("/api/preview")
    def api_preview(file_path: str, max_rows: int = 5, sheet_name: Optional[str] = None):
        return execute_preview_excel_data(file_path=file_path, max_rows=max_rows, sheet_name=sheet_name)

    # RAG ve MCP için standart JSON-RPC HTTP Endpoint
    @web_app.post("/mcp")
    async def mcp_http_endpoint(payload: Dict[str, Any]):
        method = payload.get("method")
        params = payload.get("params", {})
        req_id = payload.get("id", 1)

        if method == "tools/list":
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "tools": [
                        {
                            "name": "analyze_venn_nonconformities",
                            "description": "Excel'deki uygunsuzlukları Venn şeması ile analiz eder.",
                            "inputSchema": {
                                "type": "object",
                                "properties": {
                                    "file_path": {"type": "string"},
                                    "compare_by": {"type": "string", "default": "phases"},
                                    "activity_name": {"type": "string"}
                                },
                                "required": ["file_path"]
                            }
                        },
                        {
                            "name": "list_activities_and_phases",
                            "description": "Faaliyet ve faz listesini keşfeder.",
                            "inputSchema": {
                                "type": "object",
                                "properties": {"file_path": {"type": "string"}},
                                "required": ["file_path"]
                            }
                        }
                    ]
                }
            }
        elif method == "tools/call":
            tool_name = params.get("name")
            args = params.get("arguments", {})
            if tool_name == "analyze_venn_nonconformities":
                res = execute_analyze_venn(**args)
            elif tool_name == "list_activities_and_phases":
                res = execute_list_activities_and_phases(**args)
            elif tool_name == "preview_excel_data":
                res = execute_preview_excel_data(**args)
            else:
                return {"jsonrpc": "2.0", "id": req_id, "error": {"code": -32601, "message": "Bilinmeyen tool"}}
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {"content": [{"type": "text", "text": json.dumps(res, ensure_ascii=False, indent=2)}]}
            }
        else:
            return {"jsonrpc": "2.0", "id": req_id, "result": {}}

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
