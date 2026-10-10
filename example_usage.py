"""
Hızlı Doğrulama ve Örnek Kullanım Betiği
Bu betik, Uygunsuzluk.xlsx gerçek verisi üzerinde MCP analiz motorunu çalıştırarak
hem Fazlar (Dry-run, SKT, FKT) hem de İller/Kapsamlar arası Venn küme analizini gösterir.
"""

import os
import json
from app import analyze_venn_nonconformities, list_activities_and_phases, preview_excel_data

def run_demo():
    # Gerçek veri dosyası varsa onu, yoksa örnek dosyayı kullan
    if os.path.exists("Uygunsuzluk.xlsx"):
        excel_path = os.path.abspath("Uygunsuzluk.xlsx")
    else:
        excel_path = os.path.abspath("ornek_uygunsuzluklar.xlsx")

    print(f"[*] İncelenen Excel Dosyası: {excel_path}\n")

    # 1. Önizleme ve Otomatik Rol Tespiti
    print("=" * 65)
    print("1. ÖNİZLEME VE OTOMATİK SÜTUN ROLÜ TESPİTİ")
    print("=" * 65)
    prev = json.loads(preview_excel_data(excel_path, max_rows=3))
    print(f"Toplam Satır: {prev.get('toplam_satir')}")
    print("Tespit Edilen Roller:", json.dumps(prev.get("tespit_edilen_roller"), ensure_ascii=False))

    # 2. Dosyadaki Faaliyet ve Fazları Listeleme
    print("\n" + "=" * 65)
    print("2. FAALİYET/İL VE FAZ LİSTESİ KEŞFİ")
    print("=" * 65)
    summary = json.loads(list_activities_and_phases(excel_path))
    print("Mevcut Test Fazları:", summary.get("mevcut_fazlar"))
    print("\nİlk 5 Faaliyet/İl Dağılımı:")
    for item in summary.get("faaliyetler_veya_kapsamlar", [])[:5]:
        print(f" - {item['faaliyet_veya_kapsam']}: {item['mevcut_fazlar']} -> {item['benzersiz_uygunsuzluk_sayisi']} benzersiz uygunsuzluk")

    # 3. Fazlar Arası Venn Şeması Analizi (Kayseri veya ilk grup)
    print("\n" + "=" * 65)
    print("3. FAZLAR ARASI VENN ŞEMASI (Örn: 'kayseri' ili)")
    print("=" * 65)
    venn_res = json.loads(analyze_venn_nonconformities(
        file_path=excel_path,
        compare_by="phases",
        activity_name="kayseri"
    ))
    if venn_res.get("durum") == "basarili":
        print(f"RAG Özet Yorumu:\n-> {venn_res.get('rag_ozet_yorumu')}\n")
        print("Venn Bölgeleri Dağılımı:")
        bolgeler = venn_res.get("venn_analizi", {}).get("venn_bolgeleri", {})
        for bolge_adi, data in bolgeler.items():
            print(f"  - {bolge_adi}: {data['adet']} adet")
    else:
        print("Hata:", venn_res)

    # 4. İller / Faaliyetler Arası Küme Karşılaştırması
    print("\n" + "=" * 65)
    print("4. İLLER / FAALİYETLER ARASI KÜME KARŞILAŞTIRMASI")
    print("=" * 65)
    act_venn_res = json.loads(analyze_venn_nonconformities(
        file_path=excel_path,
        compare_by="activities"
    ))
    if act_venn_res.get("durum") == "basarili":
        print(f"Karşılaştırılan Sütun: {act_venn_res.get('karsilastirilan_sutun')}")
        print("Örnek İkili Kesişimler:")
        for pair, info in list(act_venn_res.get("venn_analizi", {}).get("ikili_kesisimler", {}).items())[:5]:
            print(f"  * {pair}: {info['kesisen_adet']} ortak uygunsuzluk (Jaccard: {info['jaccard_benzerlik']})")
            if info['kesisen_ogeler']:
                print(f"    Örnek Ortak: '{info['kesisen_ogeler'][0]}'")

if __name__ == "__main__":
    run_demo()
