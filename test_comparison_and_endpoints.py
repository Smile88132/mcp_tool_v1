"""
Unit tests for new two-Excel comparison and flow logic.
"""

import json
import os
import unittest
import pandas as pd
from app import execute_compare_two_excels, compare_two_excels_venn, extract_langflow_response

class TestTwoExcelComparison(unittest.TestCase):
    def setUp(self):
        # Create two simple test DataFrames
        self.df1 = pd.DataFrame({
            "Uygunsuzluk": ["Hata 1", "Hata 2", "Hata 3", "Ortak Hata A", "Ortak Hata B"]
        })
        self.df2 = pd.DataFrame({
            "Uygunsuzluk": ["Hata 4", "Hata 5", "Ortak Hata A", "Ortak Hata B"]
        })

    def test_compare_dataframes_venn(self):
        res = execute_compare_two_excels(
            file_1=self.df1,
            file_2=self.df2,
            file_1_name="Versiyon_1.xlsx",
            file_2_name="Versiyon_2.xlsx"
        )
        self.assertEqual(res["durum"], "basarili")
        v = res["venn_verileri"]
        
        # Intersection
        self.assertEqual(v["intersection_count"], 2)
        self.assertIn("Ortak Hata A", v["intersection_items"])
        self.assertIn("Ortak Hata B", v["intersection_items"])
        
        # Only in 1
        self.assertEqual(v["only_a_count"], 3)
        self.assertIn("Hata 1", v["only_a_items"])
        
        # Only in 2
        self.assertEqual(v["only_b_count"], 2)
        self.assertIn("Hata 4", v["only_b_items"])
        
        # Total unique
        self.assertEqual(v["total_unique"], 7)
        # Jaccard = 2 / 7 = 0.2857
        self.assertAlmostEqual(v["jaccard_similarity"], 2/7, places=3)

    def test_compare_with_sample_excel(self):
        sample_path = os.path.abspath("ornek_uygunsuzluklar.xlsx")
        res_str = compare_two_excels_venn(sample_path, sample_path)
        data = json.loads(res_str)
        self.assertEqual(data["durum"], "basarili")
        # Comparing file with itself should have 100% intersection
        v = data["venn_verileri"]
        self.assertEqual(v["only_a_count"], 0)
        self.assertEqual(v["only_b_count"], 0)
        self.assertEqual(v["jaccard_similarity"], 1.0)

    def test_extract_langflow_response(self):
        # Test standard nested Langflow response
        sample_langflow_json = {
            "session_id": "test-session",
            "outputs": [
                {
                    "outputs": [
                        {
                            "results": {
                                "message": {
                                    "text": "ASELSAN LLM Suite Analiz Çıktısı: Her iki fazda da kritik pano hataları gözlenmiştir."
                                }
                            }
                        }
                    ]
                }
            ]
        }
        extracted = extract_langflow_response(sample_langflow_json)
        self.assertIn("ASELSAN LLM Suite Analiz Çıktısı", extracted)

if __name__ == "__main__":
    unittest.main()
