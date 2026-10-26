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

    def test_calculate_text_similarity(self):
        from app import calculate_text_similarity
        # Identical strings
        self.assertEqual(calculate_text_similarity("Pano montaj hatası", "Pano montaj hatası"), 1.0)
        # Word order difference
        sim_order = calculate_text_similarity("Pano montaj hatası", "montaj hatası pano")
        self.assertGreaterEqual(sim_order, 0.85)
        # Completely different
        sim_diff = calculate_text_similarity("Kablo kesiti hatalı", "Yazılım derleme hatası")
        self.assertLess(sim_diff, 0.40)

    def test_compute_fuzzy_venn(self):
        from app import compute_fuzzy_venn
        list_a = ["Pano içi etiketleme eksik", "Kablo gevşek", "Hidrolik sızdırmazlık arızası"]
        list_b = ["etiketleme eksik pano içi", "Kablo gevşek", "Yazılım derleme hatası"]
        
        # Exact match (threshold 1.0)
        res_exact = compute_fuzzy_venn(list_a, list_b, threshold=1.0)
        self.assertEqual(res_exact["intersection_count"], 1) # Only "Kablo gevşek"
        
        # Fuzzy match (threshold 0.80)
        res_fuzzy = compute_fuzzy_venn(list_a, list_b, threshold=0.80)
        self.assertEqual(res_fuzzy["intersection_count"], 2) # Both "Kablo gevşek" and reordered "etiketleme..."
        self.assertEqual(res_fuzzy["only_a_count"], 1)
        self.assertEqual(res_fuzzy["only_b_count"], 1)

    def test_fetch_aselix_models_fallback(self):
        from app import fetch_aselix_models
        # Even with dummy/unreachable URL, it should return models list gracefully
        res = fetch_aselix_models(base_url="https://api.aselixai.aselsan.com.tr/v1", api_key="dummy_key")
        self.assertIn("modeller", res)
        self.assertGreater(len(res["modeller"]), 0)

    def test_batch_comparison_offline_fallback(self):
        from app import compare_items_with_aselix_batch
        # When no API key is provided, it must safely execute via local fuzzy distance
        batch_a = ["Pano etiketleme eksik", "Kablo kopuk"]
        cands_b = ["etiketleme eksik pano", "Hidrolik sızıntı"]
        res = compare_items_with_aselix_batch(
            base_url="https://api.aselixai.aselsan.com.tr/v1",
            api_key="",
            model="aselix-llm-v1",
            items_a_batch=batch_a,
            candidates_b=cands_b,
            threshold=0.8,
            timeout=50
        )
        self.assertEqual(res["durum"], "basarili")
        self.assertEqual(len(res["matched_pairs"]), 1)
        self.assertEqual(res["matched_pairs"][0]["item_a"], "Pano etiketleme eksik")

    def test_batch_endpoints_with_testclient(self):
        from fastapi.testclient import TestClient
        from app import web_app
        client = TestClient(web_app)
        
        # 1. Batch Chunk endpoint
        chunk_res = client.post("/api/compare/batch-chunk", json={
            "batch_index": 1,
            "total_batches": 1,
            "items_a_batch": ["Kablo gevşek"],
            "candidates_b": ["Kablo gevşek", "Optik hata"],
            "threshold": 0.8,
            "api_key": ""
        })
        self.assertEqual(chunk_res.status_code, 200)
        c_data = chunk_res.json()
        self.assertEqual(c_data["durum"], "basarili")
        self.assertEqual(len(c_data["matched_pairs"]), 1)

        # 2. Batch Finalize endpoint
        finalize_res = client.post("/api/compare/batch-finalize", json={
            "doc1_name": "Test1",
            "doc2_name": "Test2",
            "threshold": 0.8,
            "all_items_a": ["Kablo gevşek", "Hata A"],
            "all_items_b": ["Kablo gevşek", "Hata B"],
            "matched_pairs": c_data["matched_pairs"],
            "need_executive_summary": False
        })
        self.assertEqual(finalize_res.status_code, 200)
        f_data = finalize_res.json()
        self.assertEqual(f_data["durum"], "basarili")
        self.assertEqual(f_data["venn_verileri"]["intersection_count"], 1)
        self.assertEqual(f_data["venn_verileri"]["only_a_count"], 1)
        self.assertEqual(f_data["venn_verileri"]["only_b_count"], 1)


if __name__ == "__main__":
    unittest.main()
