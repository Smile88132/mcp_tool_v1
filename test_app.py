"""
Test script for app.py tools in local/offline environment.
"""

import json
import os
import unittest
from app import analyze_venn_nonconformities, list_activities_and_phases, preview_excel_data

class TestVennMCP(unittest.TestCase):
    def setUp(self):
        self.excel_path = os.path.abspath("ornek_uygunsuzluklar.xlsx")

    def test_preview_excel_data(self):
        res = preview_excel_data(self.excel_path)
        data = json.loads(res)
        self.assertEqual(data["durum"], "basarili")
        self.assertIn("ornek_veriler", data)
        self.assertGreater(data["toplam_satir"], 0)

    def test_list_activities_and_phases(self):
        res = list_activities_and_phases(self.excel_path)
        data = json.loads(res)
        self.assertEqual(data["durum"], "basarili")
        self.assertIn("faaliyetler", data)
        self.assertIn("mevcut_fazlar", data)
        act_names = [a["faaliyet_adi"] for a in data["faaliyetler"]]
        self.assertIn("Elektrik Montaj", act_names)
        self.assertIn("Mekanik Montaj", act_names)

    def test_analyze_venn_phases_single_activity(self):
        res = analyze_venn_nonconformities(
            file_path=self.excel_path,
            compare_by="phases",
            activity_name="Elektrik Montaj"
        )
        data = json.loads(res)
        self.assertEqual(data["durum"], "basarili")
        self.assertIn("venn_analizi", data)
        venn = data["venn_analizi"]
        # 'Pano içi etiketleme eksik' was in all 3 phases (Dry-run, SKT, FKT)
        all_three = venn["venn_bolgeleri"]["ucunde_ortak_kesisim"]
        self.assertEqual(all_three["adet"], 1)
        self.assertIn("Pano içi etiketleme eksik", all_three["ogeler"])
        
        # Sadece Dry-run
        sadece_dry = venn["venn_bolgeleri"]["sadece_Dry-run"]
        self.assertIn("Kablo kesiti standart dışı", sadece_dry["ogeler"])
        self.assertIn("Topraklama direnci yüksek", sadece_dry["ogeler"])

    def test_analyze_venn_phases_all_activities(self):
        res = analyze_venn_nonconformities(
            file_path=self.excel_path,
            compare_by="phases"
        )
        data = json.loads(res)
        self.assertEqual(data["durum"], "basarili")
        self.assertIn("rag_ozet_yorumu", data)

    def test_analyze_venn_activities(self):
        res = analyze_venn_nonconformities(
            file_path=self.excel_path,
            compare_by="activities"
        )
        data = json.loads(res)
        self.assertEqual(data["durum"], "basarili")
        self.assertEqual(data["analiz_tipi"], "faaliyetler_arasi_kume_analizi")
        self.assertIn("venn_analizi", data)

    def test_invalid_file_path(self):
        res = analyze_venn_nonconformities(
            file_path="olmayan_dosya.xlsx"
        )
        data = json.loads(res)
        self.assertEqual(data["durum"], "hata")
        self.assertEqual(data["hata_turu"], "FileNotFoundError")

    def test_nonexistent_activity(self):
        res = analyze_venn_nonconformities(
            file_path=self.excel_path,
            compare_by="phases",
            activity_name="Olmayan Faaliyet"
        )
        data = json.loads(res)
        self.assertEqual(data["durum"], "hata")
        self.assertIn("mevcut_faaliyetler", data)

    def test_real_uygunsuzluk_xlsx(self):
        real_file = os.path.abspath("Uygunsuzluk.xlsx")
        if os.path.exists(real_file):
            # 1. Preview
            prev = json.loads(preview_excel_data(real_file))
            self.assertEqual(prev["durum"], "basarili")
            self.assertEqual(prev["tespit_edilen_roller"]["faz"], "Faaliyet")
            self.assertEqual(prev["tespit_edilen_roller"]["uygunsuzluk"], "Açıklama")

            # 2. Phases (Kayseri)
            kayseri_venn = json.loads(analyze_venn_nonconformities(
                file_path=real_file,
                compare_by="phases",
                activity_name="kayseri"
            ))
            self.assertEqual(kayseri_venn["durum"], "basarili")
            self.assertEqual(kayseri_venn["venn_analizi"]["ozet_istatistikler"]["toplam_benzersiz_Dry-run"], 799)
            self.assertEqual(kayseri_venn["venn_analizi"]["ozet_istatistikler"]["toplam_benzersiz_SKT"], 79)

            # 3. Cities (activities)
            cities_venn = json.loads(analyze_venn_nonconformities(
                file_path=real_file,
                compare_by="activities"
            ))
            self.assertEqual(cities_venn["durum"], "basarili")
            self.assertIn("izmir <-> adana", cities_venn["venn_analizi"]["ikili_kesisimler"])

if __name__ == "__main__":
    unittest.main()

