"""
Unit tests for new two-Excel comparison and flow logic.
"""

import json
import os
import unittest
from unittest import mock
import pandas as pd
from app import execute_compare_two_excels, compare_two_excels_venn, extract_langflow_response

class TestTwoExcelComparison(unittest.TestCase):
    def setUp(self):
        import app
        app.ACTIVE_SESSIONS.clear()
        app.KOKPIT_IDENTITY_STORE["current_user"] = None
        app.KOKPIT_IDENTITY_STORE["jwt_token"] = None
        app.KOKPIT_IDENTITY_STORE["token_claims"] = None
        app.KOKPIT_IDENTITY_STORE["external_inputs"] = {}
        app.KOKPIT_IDENTITY_STORE["sessions_by_sicil"] = {}
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


class TestAselixBatchPipeline(unittest.TestCase):
    """AselixAI API'si mock'lanarak batch kıyaslamanın uçtan uca doğrulanması."""

    def _mock_response(self, content):
        from unittest.mock import MagicMock
        r = MagicMock()
        r.status_code = 200
        r.json.return_value = {"choices": [{"message": {"content": content}}]}
        return r

    def test_ai_index_matching(self):
        from unittest.mock import patch
        import app
        a = ["Acil durdurma butonu takılıyor", "Röle soketi kırık"]
        b = ["Optik sensör sapması", "Acil stop butonu mekanik takılma"]
        llm = '```json\n[{"a": 1, "b": 2, "similarity": 93}, {"a": 2, "b": null, "similarity": 0}]\n```'
        with patch.object(app.requests, "post", return_value=self._mock_response(llm)):
            res = app.compare_items_with_aselix_batch("http://x/v1", "key", "m", a, b, threshold=0.8)
        self.assertEqual(res["kaynak"], "aselix_ai")
        self.assertEqual(len(res["matched_pairs"]), 1)
        self.assertEqual(res["matched_pairs"][0]["item_a"], a[0])
        self.assertEqual(res["matched_pairs"][0]["item_b"], b[1])
        self.assertIn("tanilama", res)

    def test_timeout_falls_back_to_local(self):
        from unittest.mock import patch
        import app
        with patch.object(app.requests, "post", side_effect=app.requests.exceptions.Timeout()):
            res = app.compare_items_with_aselix_batch(
                "http://x/v1", "key", "m", ["Kablo gevşek"], ["Kablo gevşek", "Hata"], threshold=0.8, timeout=1)
        self.assertEqual(res["kaynak"], "yerel_fallback")
        self.assertTrue(res["api_hatasi"])
        self.assertEqual(len(res["matched_pairs"]), 1)

    def test_shortlist_limits_prompt_size(self):
        from app import shortlist_candidates_for_batch
        b = [f"Uygunsuzluk kaydı numara {i} farklı açıklama" for i in range(500)] + ["Pano etiketi eksik"]
        sl = shortlist_candidates_for_batch(["Pano etiketleme eksik"], b, top_k=6, max_total=60)
        self.assertLessEqual(len(sl), 60)
        self.assertIn("Pano etiketi eksik", sl)

    def test_full_batch_flow_via_endpoints(self):
        """15'erli batch'lerle 40 satır: chunk -> finalize, AI mock'lu."""
        from unittest.mock import patch
        from fastapi.testclient import TestClient
        import app
        client = TestClient(app.web_app)
        a = [f"Hata kodu {i} test sapması" for i in range(40)]
        b = a[:20] + [f"Yeni hata {i}" for i in range(10)]
        llm = "[]"  # AI eşleşme bulmasa da fuzzy tam eşleşmeleri yakalamalı
        pairs = []
        with patch.object(app.requests, "post", return_value=self._mock_response(llm)):
            for i in range(0, 40, 15):
                r = client.post("/api/compare/batch-chunk", json={
                    "api_key": "k", "items_a_batch": a[i:i + 15], "candidates_b": b,
                    "threshold": 0.95, "batch_index": i // 15 + 1, "total_batches": 3})
                self.assertEqual(r.status_code, 200)
                pairs += r.json()["matched_pairs"]
        f = client.post("/api/compare/batch-finalize", json={
            "threshold": 0.95, "all_items_a": a, "all_items_b": b,
            "matched_pairs": pairs, "need_executive_summary": False}).json()
        v = f["venn_verileri"]
        self.assertEqual(v["intersection_count"], 20)
        self.assertEqual(v["only_a_count"], 20)
        self.assertEqual(v["only_b_count"], 10)

    def test_extract_items_endpoint(self):
        from fastapi.testclient import TestClient
        from app import web_app
        client = TestClient(web_app)
        path = os.path.abspath("ornek_uygunsuzluklar.xlsx")
        with open(path, "rb") as f1, open(path, "rb") as f2:
            r = client.post("/api/compare/extract-items",
                            files={"file_1": ("a.xlsx", f1.read()), "file_2": ("b.xlsx", f2.read())},
                            data={"col_1": "auto", "col_2": "auto"})
        d = r.json()
        self.assertEqual(d["durum"], "basarili")
        self.assertGreater(len(d["items_1"]), 0)


class TestKokpitIdentityAndEmbedding(unittest.TestCase):
    """KIMLIK_AKTARIMI_ENTEGRASYON_REHBERI.md uyumlu SSO, Iframe Push, Dış Input ve Embedding testleri."""

    def setUp(self):
        from fastapi.testclient import TestClient
        from app import web_app
        self.client = TestClient(web_app)

    def test_iframe_headers_allow_embedding(self):
        resp = self.client.get("/")
        self.assertEqual(resp.status_code, 200)
        self.assertNotIn("x-frame-options", [k.lower() for k in resp.headers.keys()])
        csp = resp.headers.get("Content-Security-Policy", "")
        self.assertIn("frame-ancestors", csp)
        self.assertIn("*", csp)
        self.assertEqual(resp.headers.get("Access-Control-Allow-Origin"), "*")

    def test_kokpit_identity_push_and_get(self):
        payload = {
            "type": "kokpit-identity",
            "status": "success",
            "source": "kokpit-iframe-bridge",
            "timestamp": "2026-10-29T10:00:00.000Z",
            "token": "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NSIsInByZWZlcnJlZF91c2VybmFtZSI6ImFobWV0LnlpbG1heiIsIm5hbWUiOiJBaG1ldCBZxLFsbWF6In0.sig",
            "user": {
                "id": 145,
                "sicil": "12345",
                "username": "ahmet.yilmaz",
                "isim": "Ahmet Yılmaz",
                "mail": "ahmet.yilmaz@sirket.com",
                "picture": "/uploads/photos/12345.jpg",
                "role": "Admin",
                "mudurluk": "UGES Teslimat Yönetim Müdürlüğü",
                "mudurluk_id": 191148,
                "direktorluk": "Program Direktörlüğü",
                "birim": "Planlama Birimi",
                "unvan": "Uzman",
                "hierarchicLevel": 2
            },
            "inputs": {
                "engine_mode": "embedding",
                "threshold": 0.80,
                "items_1": ["Pano etiketleme eksik", "Kablo gevşek"],
                "items_2": ["etiketleme eksik pano", "Sensör hatası"]
            }
        }
        post_res = self.client.post("/api/identity", json=payload)
        self.assertEqual(post_res.status_code, 200)
        data = post_res.json()
        self.assertEqual(data["status"], "success")
        self.assertEqual(data["user"]["sicil"], "12345")
        self.assertEqual(data["user"]["isim"], "Ahmet Yılmaz")
        self.assertTrue(data["token_present"])
        self.assertEqual(data["external_inputs"]["engine_mode"], "embedding")

        get_res = self.client.get("/api/identity")
        self.assertEqual(get_res.status_code, 200)
        gdata = get_res.json()
        self.assertTrue(gdata["authenticated"])
        self.assertEqual(gdata["user"]["mudurluk_id"], 191148)

    def test_all_incoming_identity_requests_accepted(self):
        # Dışarıdan gelen tüm talepler onay gerekmeksizin doğrudan kabul edilir
        res = self.client.post("/api/identity", json={"user": {}})
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.json()["status"], "success")

        res_bare = self.client.post("/api/identity", json={"sicil": "23365", "isim": "İsmail Yılmaz"})
        self.assertEqual(res_bare.status_code, 200)
        self.assertEqual(res_bare.json()["user"]["sicil"], "23365")

    def test_external_input_with_embeddings_and_auto_compare(self):
        ext_payload = {
            "source": "kokpit-iframe-bridge",
            "items_1": ["Pano içi etiketleme eksik", "Kablo kopuk"],
            "items_2": ["etiketleme eksik pano içi", "Yazılım hatası"],
            "embeddings_1": [[1.0, 0.0, 0.0], [0.0, 1.0, 0.0]],
            "embeddings_2": [[0.98, 0.02, 0.0], [0.0, 0.0, 1.0]],
            "engine_mode": "embedding",
            "threshold": 0.80,
            "auto_compare": True
        }
        res = self.client.post("/api/external-input", json=ext_payload)
        self.assertEqual(res.status_code, 200)
        d = res.json()
        self.assertEqual(d["status"], "success")
        self.assertIsNotNone(d["comparison_result"])
        v = d["comparison_result"]["venn_verileri"]
        self.assertEqual(v["intersection_count"], 1)
        self.assertEqual(v["embedding_source"], "dis_input_hazir_embedding")

    def test_compare_run_with_embedding_engine(self):
        res = self.client.post("/api/compare/run", data={
            "input_mode": "paste",
            "engine_mode": "embedding",
            "threshold": 0.75,
            "text_1": "Pano içi etiketleme eksik\nTopraklama cıvatası gevşek",
            "text_2": "etiketleme eksik pano içi\nOptik sensör hatası"
        })
        self.assertEqual(res.status_code, 200)
        d = res.json()
        self.assertEqual(d["durum"], "basarili")
        self.assertEqual(d["engine_mode"], "embedding")
        self.assertEqual(d["venn_verileri"]["intersection_count"], 1)

    def test_bakiye_siparis_items_endpoint(self):
        res = self.client.get("/api/bakiye-siparis/items")
        self.assertEqual(res.status_code, 200)
        d = res.json()
        self.assertEqual(d["durum"], "basarili")
        self.assertEqual(d["faz_secenekleri"], ["Dry-Run", "FKT", "SKT"])
        self.assertTrue(len(d["kayitlar"]) > 0)
        first = d["kayitlar"][0]
        self.assertIn("proje_adi", first)
        self.assertIn("sd", first)
        self.assertIn("kalem", first)
        self.assertIn("urun_kisa_metni", first)

    def test_compare_run_with_bakiye_meta_and_phases(self):
        meta_1 = json.dumps({
            "proje_adi": "Deneme Proje 1",
            "uges_proje_tanim": "UDeneme1",
            "sd": "987654",
            "kalem": "2",
            "urun_kisa_metni": "Malzeme2",
            "faz": "Dry-Run"
        })
        meta_2 = json.dumps({
            "proje_adi": "Deneme Proje 1",
            "uges_proje_tanim": "UDeneme1",
            "sd": "987654",
            "kalem": "2",
            "urun_kisa_metni": "Malzeme2",
            "faz": "FKT"
        })
        res = self.client.post("/api/compare/run", data={
            "input_mode": "paste",
            "engine_mode": "local",
            "threshold": 0.80,
            "text_1": "Pano etiketleme eksik\nKablo gevşek",
            "text_2": "Pano etiketleme eksik\nSensör arızası",
            "meta_1": meta_1,
            "meta_2": meta_2
        })
        self.assertEqual(res.status_code, 200)
        d = res.json()
        self.assertEqual(d["durum"], "basarili")
        self.assertEqual(d["belge_1_meta"]["faz"], "Dry-Run")
        self.assertEqual(d["belge_2_meta"]["faz"], "FKT")
        self.assertEqual(d["belge_1_meta"]["sd"], "987654")
        self.assertIn("Dry-Run", d["ozet_yorum"])
        self.assertIn("FKT", d["ozet_yorum"])

    def test_save_and_get_dryrun_fkt_skt_kpi(self):
        payload = {
            "proje_adi": "Deneme Proje 1",
            "sd": "987654",
            "kalem": "2",
            "faaliyet_1_turu": "FKT",  # Gönderilse bile backend kural gereği Dry-Run yapmalı
            "faaliyet_2_turu": "SKT",
            "urun_kisa_metni": "Malzeme2",
            "uges_proje_tanim": "UDeneme1",
            "malzeme": "5865-7567-0110",
            "ilk_faaliyet_benzersiz_uygunsuzluklar": [
                "Kablo kesiti standart dışı",
                "Topraklama direnci yüksek",
                "Cıvata tork değeri yetersiz"
            ],
            "ortak_uygunsuzluklar": [
                {
                    "ilk_faaliyet_uygunsuzluk": "Pano içi etiketleme eksik",
                    "ikinci_faaliyet_uygunsuzluk": "Pano içi etiketleme eksik",
                    "benzerlik_yuzdesi": 100.0
                }
            ],
            "ikinci_faaliyet_benzersiz_uygunsuzluklar": [
                "Flanş sızdırmazlığı arızalı",
                "Klemens bağlantısı gevşek",
                "Cıvata tork değeri yetersiz"
            ],
            "reddedilen_eslesmeler": [
                {
                    "ilk_faaliyet_uygunsuzluk": "Cıvata tork değeri yetersiz",
                    "ikinci_faaliyet_uygunsuzluk": "Cıvata tork değeri yetersiz",
                    "eski_benzerlik_yuzdesi": 100.0
                }
            ]
        }
        save_res = self.client.post("/api/bakiye-siparis/save-kpi", json=payload)
        self.assertEqual(save_res.status_code, 200)
        sdata = save_res.json()
        self.assertIn(sdata["durum"], ["basarili", "uyari"])
        self.assertEqual(sdata["kayit"]["faaliyet_1_turu"], "Dry-Run")
        self.assertEqual(sdata["kayit"]["faaliyet_2_turu"], "SKT")
        self.assertEqual(sdata["kayit"]["ilk_faaliyet_benzersiz_sayisi"], 3)
        self.assertEqual(sdata["kayit"]["ortak_uygunsuzluk_sayisi"], 1)
        self.assertEqual(sdata["kayit"]["ikinci_faaliyet_benzersiz_sayisi"], 3)

        list_res = self.client.get("/api/bakiye-siparis/kpi-records")
        self.assertEqual(list_res.status_code, 200)
        ldata = list_res.json()
        self.assertTrue(len(ldata["kayitlar"]) >= 1)

    def test_env_and_db_startup_verification(self):
        status_res = self.client.get("/api/bakiye-siparis/db-status")
        self.assertEqual(status_res.status_code, 200)
        s = status_res.json()
        self.assertTrue(s["env_dosyasi_mevcut"])
        self.assertEqual(s["baglanti"]["dbname"], "ud_kokpit")
        self.assertEqual(s["baglanti"]["schema"], "teslimat_md")
        self.assertIn("kpi_tablosu", s)
        self.assertIn("sunucu", s)
        self.assertIn("app_port", s["sunucu"])
        self.assertEqual(s["sunucu"]["app_port"], 9000)

        # .env üzerinden veya ortam değişkeninden port değiştirilebilirlik testi
        from unittest import mock
        from app import get_server_config
        srv = get_server_config()
        self.assertEqual(srv["port"], 9000)
        with mock.patch.dict(os.environ, {"APP_PORT": "9090"}):
            self.assertEqual(get_server_config()["port"], 9090)
        with mock.patch.dict(os.environ, {"PORT": "7070", "APP_PORT": ""}):
            self.assertEqual(get_server_config()["port"], 7070)

    def test_admin_only_llm_suite_and_role_visibility(self):
        # 1) Standart kullanıcı (User) oturumu açıldığında LLM Suite Flow URL ve API Key uyarısı dönmemeli
        self.client.post(
            "/api/identity",
            json={
                "type": "kokpit-identity",
                "status": "success",
                "source": "kokpit-iframe-bridge",
                "user": {
                    "sicil": "67890",
                    "username": "mehmet.kaya",
                    "isim": "Mehmet Kaya",
                    "role": "User"
                }
            }
        )
        id_res_user = self.client.get("/api/identity").json()
        self.assertFalse(id_res_user["user"]["is_admin"])

        sample_user_res = self.client.post(
            "/api/flow/sample-compare",
            json={
                "base_url": "https://api.aselixai.aselsan.com.tr/v1",
                "flow_id": "c16ecc83-ff64-4fc7-bb52-0f594beaf70a",
                "api_key": ""
            }
        ).json()
        self.assertNotIn("ASELSAN LLM Suite Flow Bağlantısı", sample_user_res.get("flow_analiz_ozeti", ""))
        self.assertNotIn("Hedef Akış:", sample_user_res.get("flow_analiz_ozeti", ""))
        self.assertNotIn("API Key", sample_user_res.get("flow_analiz_ozeti", ""))

        # 2) Admin kullanıcısı oturumu açıldığında LLM Suite Flow URL ve Ayarlar bilgisi görünmeli
        self.client.post(
            "/api/identity",
            json={
                "type": "kokpit-identity",
                "status": "success",
                "source": "kokpit-iframe-bridge",
                "user": {
                    "sicil": "12345",
                    "username": "ahmet.yilmaz",
                    "isim": "Ahmet Yılmaz",
                    "role": "Admin"
                }
            }
        )
        id_res_admin = self.client.get("/api/identity").json()
        self.assertTrue(id_res_admin["user"]["is_admin"])

        sample_admin_res = self.client.post(
            "/api/flow/sample-compare",
            json={
                "base_url": "https://api.aselixai.aselsan.com.tr/v1",
                "flow_id": "c16ecc83-ff64-4fc7-bb52-0f594beaf70a",
                "api_key": ""
            }
        ).json()
        self.assertIn("ASELSAN LLM Suite Flow Bağlantısı", sample_admin_res.get("flow_analiz_ozeti", ""))
        self.assertIn("Hedef Akış:", sample_admin_res.get("flow_analiz_ozeti", ""))
        self.assertIn("Ayarlar", sample_admin_res.get("flow_analiz_ozeti", ""))

        # Temizle
        self.client.delete("/api/identity")

    def test_kpi_records_filtering_and_admin_delete(self):
        # 1) Geçici test KPI kaydı ekle
        test_payload = {
            "proje_adi": "TEST-FILTRE-SIL-PROJESI",
            "sd": "99999999",
            "kalem": "999",
            "faaliyet_1_turu": "Dry-Run",
            "faaliyet_2_turu": "SKT",
            "urun_kisa_metni": "TEST RADAR MODULU",
            "ilk_faaliyet_benzersiz_uygunsuzluklar": ["Test hata A1"],
            "ortak_uygunsuzluklar": [{"ilk_faaliyet_uygunsuzluk": "Ortak X", "ikinci_faaliyet_uygunsuzluk": "Ortak X", "benzerlik_yuzdesi": 100}],
            "ikinci_faaliyet_benzersiz_uygunsuzluklar": ["Test hata B1"]
        }
        save_res = self.client.post("/api/bakiye-siparis/save-kpi", json=test_payload)
        self.assertEqual(save_res.status_code, 200)

        # 2) Tablo filtreleme sorgularını test et
        filter_res = self.client.get(
            "/api/bakiye-siparis/kpi-records",
            params={
                "proje": "TEST-FILTRE-SIL-PROJESI",
                "sd": "99999999",
                "faaliyet_2_turu": "SKT",
                "search": "RADAR"
            }
        )
        self.assertEqual(filter_res.status_code, 200)
        fdata = filter_res.json()
        self.assertEqual(len(fdata["kayitlar"]), 1)
        self.assertEqual(fdata["kayitlar"][0]["urun_kisa_metni"], "TEST RADAR MODULU")

        # 3) Standart kullanıcı (User) silmeye çalışırsa 403 Yetkisiz dönmeli
        self.client.post(
            "/api/identity",
            json={
                "type": "kokpit-identity",
                "status": "success",
                "source": "kokpit-iframe-bridge",
                "user": {"sicil": "67890", "username": "mehmet.kaya", "isim": "Mehmet Kaya", "role": "User"}
            }
        )
        del_unauth = self.client.post("/api/bakiye-siparis/delete-kpi", json=test_payload)
        self.assertEqual(del_unauth.status_code, 403)
        self.assertEqual(del_unauth.json()["durum"], "yetkisiz")

        # 4) Admin kullanıcısı silmeye çalışırsa başarıyla silinmeli
        self.client.post(
            "/api/identity",
            json={
                "type": "kokpit-identity",
                "status": "success",
                "source": "kokpit-iframe-bridge",
                "user": {"sicil": "12345", "username": "ahmet.yilmaz", "isim": "Ahmet Yılmaz", "role": "Admin"}
            }
        )
        del_admin = self.client.post("/api/bakiye-siparis/delete-kpi", json=test_payload)
        self.assertEqual(del_admin.status_code, 200)
        self.assertEqual(del_admin.json()["durum"], "basarili")
        self.assertEqual(del_admin.json()["silinen_adet"], 1)

        # Silindiğini doğrula
        after_res = self.client.get(
            "/api/bakiye-siparis/kpi-records",
            params={"proje": "TEST-FILTRE-SIL-PROJESI", "sd": "99999999"}
        ).json()
        self.assertEqual(len(after_res["kayitlar"]), 0)

        self.client.delete("/api/identity")

    def test_bulk_template_and_analyze_and_tabs_ui(self):
        # 1) Ana sayfa HTML içinde Home, VennFlow, Toplu Veri Girişi sekmeleri ve KPI Pop-up modalı var mı?
        html_res = self.client.get("/")
        self.assertEqual(html_res.status_code, 200)
        html_text = html_res.text
        self.assertIn('id="pageTabBtnHome"', html_text)
        self.assertIn('id="pageTabBtnVennflow"', html_text)
        self.assertIn('id="pageTabBtnBulk"', html_text)
        self.assertIn('id="kpiRecordsModal"', html_text)

        # 2) Örnek toplu Excel şablonunu indir
        tpl_res = self.client.get("/api/bulk/template-excel")
        self.assertEqual(tpl_res.status_code, 200)
        self.assertIn(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            tpl_res.headers.get("content-type", "")
        )
        excel_bytes = tpl_res.content
        self.assertTrue(len(excel_bytes) > 1000)

        # 3) İndirilen örnek Excel'i geri yükleyip arka planda SD/Kalem eşleştirmesini ve Venn analizini doğrula
        analyze_res = self.client.post(
            "/api/bulk/analyze-excel",
            files={"file": ("VennFlow_Toplu_Veri_Girisi_Sablonu.xlsx", excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            data={"engine_mode": "local", "threshold": "0.8"}
        )
        self.assertEqual(analyze_res.status_code, 200)
        adata = analyze_res.json()
        self.assertEqual(adata["durum"], "basarili")
        self.assertGreaterEqual(adata["toplam_grup_sayisi"], 1)

        first_grp = adata["gruplar"][0]
        self.assertEqual(first_grp["faaliyet_1_turu"], "Dry-Run")
        self.assertIn(first_grp["faaliyet_2_turu"], ("FKT", "SKT"))
        self.assertTrue(first_grp["sd"])
        self.assertTrue(first_grp["kalem"])
        self.assertTrue(first_grp["proje_adi"])
        self.assertTrue(first_grp["urun_kisa_metni"])
        self.assertIn("venn_verileri", first_grp)
        # Örnek şablonda 2 ortak, 2 yalnızca Dry-Run, 1 yalnızca FKT var
        self.assertEqual(first_grp["venn_verileri"]["intersection_count"], 2)

        # 4) Excel şablonunun A=SD, B=Kalem sütun yapısını ve E-G otomatik VLOOKUP formüllerini doğrula
        import io
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=False)
        ws = wb["Uygunsuzluk_Veri_Girisi"]
        self.assertIn("SD", str(ws["A1"].value))
        self.assertIn("Kalem", str(ws["B1"].value))
        self.assertIn("VLOOKUP(TRIM(A2)&\"|\"&TRIM(B2)", str(ws["E2"].value))
        self.assertIn("VLOOKUP(TRIM(A2)&\"|\"&TRIM(B2)", str(ws["F2"].value))
        self.assertIn("VLOOKUP(TRIM(A2)&\"|\"&TRIM(B2)", str(ws["G2"].value))

        # 5) VennFlow tekli önizlemede (/api/compare/excel-preview) SD ve Kalem sütunlarından bakiye bilgisi otomatik doluyor mu?
        preview_res = self.client.post(
            "/api/compare/excel-preview",
            files={"file": ("VennFlow_Toplu_Veri_Girisi_Sablonu.xlsx", excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        )
        self.assertEqual(preview_res.status_code, 200)
        pdata = preview_res.json()
        self.assertIsNotNone(pdata.get("tespit_edilen_bakiye"))
        self.assertEqual(pdata["tespit_edilen_bakiye"]["sd"], first_grp["sd"])
        self.assertEqual(pdata["tespit_edilen_bakiye"]["kalem"], first_grp["kalem"])
        self.assertEqual(pdata["tespit_edilen_bakiye"]["proje_adi"], first_grp["proje_adi"])


class TestAselixAIEnhancementsAndResilience(unittest.TestCase):
    def setUp(self):
        from fastapi.testclient import TestClient
        import app
        self.client = TestClient(app.web_app)

    def tearDown(self):
        import app
        app.ACTIVE_SESSIONS.clear()
        app.KOKPIT_IDENTITY_STORE["current_user"] = None
        app.KOKPIT_IDENTITY_STORE["jwt_token"] = None
        app.KOKPIT_IDENTITY_STORE["token_claims"] = None
        app.KOKPIT_IDENTITY_STORE["external_inputs"] = {}
        app.KOKPIT_IDENTITY_STORE["sessions_by_sicil"] = {}

    def test_get_effective_aselix_config(self):
        import app
        # 1. Varsayılanlar
        with mock.patch.dict(os.environ, {"ASELIX_API_KEY": "", "ASELIX_BASE_URL": "", "ASELIX_TIMEOUT": "", "ASELIX_MODEL": ""}, clear=True):
            eff_url, eff_key, eff_model, eff_timeout = app.get_effective_aselix_config()
            self.assertEqual(eff_url, "https://api.aselixai.aselsan.com.tr/v1")
            self.assertEqual(eff_model, "openai/gpt-oss-120b")
            self.assertEqual(eff_timeout, 90)

        # 2. .env önceliği
        with mock.patch.dict(os.environ, {
            "ASELIX_API_KEY": "env-secret-key",
            "ASELIX_BASE_URL": "https://agate.aselsan.com.tr/send/milli/api.aselixai/v1",
            "ASELIX_TIMEOUT": "120"
        }):
            eff_url, eff_key, eff_model, eff_timeout = app.get_effective_aselix_config()
            self.assertEqual(eff_key, "env-secret-key")
            self.assertIn("agate.aselsan.com.tr", eff_url)
            self.assertEqual(eff_timeout, 120)

        # 3. Parametre önceliği (ortam değişkenini ezer)
        with mock.patch.dict(os.environ, {"ASELIX_API_KEY": "env-secret-key"}):
            _, eff_key, _, _ = app.get_effective_aselix_config(api_key="direct-param-key")
            self.assertEqual(eff_key, "direct-param-key")

    def test_extract_json_array_resilience(self):
        import app
        # A) Standard JSON
        txt1 = '[{"a": 1, "b": 2, "similarity": 95}]'
        self.assertEqual(len(app.extract_json_array_from_llm(txt1)), 1)

        # B) LLM with <think> tag and markdown block
        txt2 = """<think>
        Düşünce: 1. madde 3. madde ile eşleşiyor.
        </think>
        İşte eşleştirme sonucu:
        ```json
        [
          {"a": 1, "b": 3, "similarity": 90},
          {"a": 2, "b": null, "similarity": 0}
        ]
        ```
        """
        arr2 = app.extract_json_array_from_llm(txt2)
        self.assertEqual(len(arr2), 2)
        self.assertEqual(arr2[0]["b"], 3)

        # C) Python literal eval single quotes
        txt3 = "[{'a': 1, 'b': 2, 'similarity': 88}, {'a': 2, 'b': None, 'similarity': 0}]"
        arr3 = app.extract_json_array_from_llm(txt3)
        self.assertEqual(len(arr3), 2)

        # D) Regex object recovery when brackets are broken
        txt4 = 'Analiz tamamlandı: {"a": 1, "b": 2, "similarity": 85} ve ayrıca {"a": 2, "b": 3, "similarity": 90}'
        arr4 = app.extract_json_array_from_llm(txt4)
        self.assertEqual(len(arr4), 2)

        # E) List input (TypeError: expected string or bytes-like object, got 'list')
        list_input_1 = [{"a": 1, "b": 2, "similarity": 95}]
        arr5 = app.extract_json_array_from_llm(list_input_1)
        self.assertEqual(len(arr5), 1)
        self.assertEqual(arr5[0]["b"], 2)

        # F) Multi-block content list (OpenAI style [{"type": "text", "text": "..."}])
        list_input_2 = [{"type": "text", "text": '[{"a": 2, "b": 4, "similarity": 89}]'}]
        arr6 = app.extract_json_array_from_llm(list_input_2)
        self.assertEqual(len(arr6), 1)
        self.assertEqual(arr6[0]["b"], 4)

        # G) String list input
        list_input_3 = ['[{"a": 3, "b": null, "similarity": 0}]']
        arr7 = app.extract_json_array_from_llm(list_input_3)
        self.assertEqual(len(arr7), 1)
        self.assertEqual(arr7[0]["a"], 3)

    def test_call_aselix_chat_dual_network_failover(self):
        import app
        call_urls = []

        def mock_post(url, *args, **kwargs):
            call_urls.append(url)
            # Birincil (Milli Ağ) zaman aşımı versin
            if "api.aselixai.aselsan.com.tr" in url:
                import requests
                raise requests.exceptions.Timeout("Milli Ağ Gateway Timeout")
            # İkincil (Aynet) başarılı dönsün
            resp = mock.MagicMock()
            resp.status_code = 200
            resp.json.return_value = {"choices": [{"message": {"content": "Aynet üzerinden selamlar"}}]}
            return resp

        with mock.patch("requests.post", side_effect=mock_post):
            res = app.call_aselix_chat(
                base_url="https://api.aselixai.aselsan.com.tr/v1",
                api_key="test-key",
                messages=[{"role": "user", "content": "test"}]
            )
            self.assertEqual(res["durum"], "basarili")
            self.assertEqual(res["aktif_ag"], "aynet")
            self.assertEqual(res["cevap"], "Aynet üzerinden selamlar")
            # Hem Milli Ağ hem Aynet çağrılmış olmalı
            self.assertTrue(any("api.aselixai" in u for u in call_urls))
            self.assertTrue(any("agate.aselsan" in u for u in call_urls))

    def test_aselix_config_and_test_connection_endpoints(self):
        # 1. /api/aselix/config
        with mock.patch.dict(os.environ, {"ASELIX_API_KEY": "dummy-key", "ASELIX_TIMEOUT": "90"}):
            res = self.client.get("/api/aselix/config")
            self.assertEqual(res.status_code, 200)
            data = res.json()
            self.assertEqual(data["durum"], "basarili")
            self.assertTrue(data["has_server_api_key"])
            self.assertEqual(data["default_timeout"], 90)
            self.assertEqual(data["default_batch_size"], 5)

        # 2. /api/aselix/test-connection
        with mock.patch("app.fetch_aselix_models", return_value={"durum": "basarili", "base_url": "https://api.aselixai.aselsan.com.tr/v1", "aktif_ag": "milli", "modeller": ["openai/gpt-oss-120b", "aselix-llm-v1"]}):
            with mock.patch("app.call_aselix_chat", return_value={"durum": "basarili", "cevap": "Canlıyım", "aktif_ag": "milli"}):
                conn_res = self.client.post("/api/aselix/test-connection", json={"api_key": "k"})
                self.assertEqual(conn_res.status_code, 200)
                cdata = conn_res.json()
                self.assertEqual(cdata["durum"], "basarili")
                self.assertEqual(cdata["aktif_ag"], "milli")
                self.assertIsNotNone(cdata["chat_ping"])

        # 3. GET /api/aselix/models (Sayfa yenilemede otomatik modelleri çekme)
        with mock.patch("app.fetch_aselix_models", return_value={"durum": "basarili", "modeller": ["openai/gpt-oss-120b", "aselix-llm-v1"]}):
            models_res = self.client.get("/api/aselix/models")
            self.assertEqual(models_res.status_code, 200)
            m_data = models_res.json()
            self.assertEqual(m_data["durum"], "basarili")
            self.assertIn("openai/gpt-oss-120b", m_data["modeller"])

    def test_strip_bakiye_metadata_from_text(self):
        from app import strip_bakiye_metadata_from_text
        # Pipe-delimited bakiye format
        raw_1 = "987654 / 2 | Deneme Proje 1 | Malzeme2 | Pano içi etiketleme eksik"
        cleaned_1 = strip_bakiye_metadata_from_text(raw_1)
        self.assertEqual(cleaned_1, "Pano içi etiketleme eksik")

        # Prefix format
        raw_2 = "SD: 987654 Kalem: 2 Proje: Radar | Kablo kesiti standart dışı"
        cleaned_2 = strip_bakiye_metadata_from_text(raw_2)
        self.assertEqual(cleaned_2, "Kablo kesiti standart dışı")

        # Clean text without bakiye
        raw_3 = "Topraklama barası cıvatası gevşek"
        cleaned_3 = strip_bakiye_metadata_from_text(raw_3)
        self.assertEqual(cleaned_3, "Topraklama barası cıvatası gevşek")

    def test_create_similarity_aware_batches(self):
        from app import create_similarity_aware_batches
        items_a = [
            "Pano içi etiketleme eksik",
            "Topraklama barası cıvatası gevşek",
            "Hidrolik yağ sızıntısı mevcut",
            "Pano kapağı kilit arızası",
            "Kablo kanalı kapağı açık",
            "Farklı başka bir problem A",
            "Farklı başka bir problem B"
        ]
        items_b = [
            "etiketleme eksik pano içi",
            "cıvata gevşek topraklama barası",
            "yağ kaçağı hidrolik sistem",
            "yazılım derleme hatası"
        ]
        batches = create_similarity_aware_batches(items_a, items_b, batch_size=5)
        self.assertGreaterEqual(len(batches), 2)
        self.assertLessEqual(len(batches[0]["items_a"]), 5)
        # En yüksek potansiyelli maddeler ilk batch'te toplanmalı
        first_batch_texts = batches[0]["items_a"]
        self.assertTrue(
            "Pano içi etiketleme eksik" in first_batch_texts or
            "Topraklama barası cıvatası gevşek" in first_batch_texts
        )
        self.assertIn("cluster_label", batches[0])

    def test_api_compare_plan_batches_endpoint(self):
        payload = {
            "items_a": ["Hata 1", "Hata 2", "Hata 3", "Hata 4", "Hata 5", "Hata 6"],
            "items_b": ["Hata 1", "Hata 2"],
            "batch_size": 5
        }
        res = self.client.post("/api/compare/plan-batches", json=payload)
        self.assertEqual(res.status_code, 200)
        data = res.json()
        self.assertEqual(data["durum"], "basarili")
        self.assertEqual(data["batch_size"], 5)
        self.assertEqual(len(data["batches"]), 2)

    def test_api_aselix_chat_multi_turn(self):
        with mock.patch("app.call_aselix_chat", return_value={"durum": "basarili", "cevap": "Sohbet cevabı", "aktif_ag": "milli"}):
            chat_res = self.client.post("/api/aselix/chat", json={
                "messages": [
                    {"role": "user", "content": "İlk soru"},
                    {"role": "assistant", "content": "İlk cevap"},
                    {"role": "user", "content": "İkinci soru"}
                ],
                "timeout_sec": 90
            })
            self.assertEqual(chat_res.status_code, 200)
            self.assertEqual(chat_res.json()["cevap"], "Sohbet cevabı")

    def test_sample_compare_delegates_to_aselix_without_401_error(self):
        # AselixAI yapılandırmasıyla /api/flow/sample-compare çağrıldığında
        # 401 hatası vermek yerine call_aselix_chat ile başarılı analiz döndürmeli
        with mock.patch("app.call_aselix_chat", return_value={"durum": "basarili", "cevap": "AselixAI Kök Neden Analizi Başarılı"}):
            res = self.client.post("/api/flow/sample-compare", json={
                "base_url": "https://api.aselixai.aselsan.com.tr/v1",
                "api_key": "test-aselix-token"
            })
            self.assertEqual(res.status_code, 200)
            data = res.json()
            self.assertEqual(data["durum"], "basarili")
            self.assertEqual(data["engine_mode"], "aselix")
            self.assertEqual(data["flow_analiz_ozeti"], "AselixAI Kök Neden Analizi Başarılı")


class TestMultiUserSessionIsolation(unittest.TestCase):
    def setUp(self):
        from fastapi.testclient import TestClient
        import app
        self.client = TestClient(app.web_app)

    def tearDown(self):
        import app
        app.ACTIVE_SESSIONS.clear()
        app.KOKPIT_IDENTITY_STORE["current_user"] = None
        app.KOKPIT_IDENTITY_STORE["jwt_token"] = None
        app.KOKPIT_IDENTITY_STORE["token_claims"] = None
        app.KOKPIT_IDENTITY_STORE["external_inputs"] = {}
        app.KOKPIT_IDENTITY_STORE["sessions_by_sicil"] = {}

    def test_multi_user_sessions_isolated(self):
        # 1. Kullanıcı A oturum açar (Admin, Ahmet Yılmaz)
        payload_a = {
            "type": "kokpit-identity",
            "source": "kokpit-iframe-bridge",
            "user": {
                "sicil": "12345",
                "isim": "Ahmet Yılmaz",
                "username": "ahmet.yilmaz",
                "role": "Admin"
            },
            "token": "token-user-a",
            "inputs": {"engine_mode": "aselix", "threshold": 85}
        }
        res_a = self.client.post("/api/identity", json=payload_a, headers={"X-Session-ID": "sess_user_a"})
        self.assertEqual(res_a.status_code, 200)
        self.assertEqual(res_a.json()["session_id"], "sess_user_a")

        # 2. Kullanıcı B oturum açar (Standart Kullanıcı, Mehmet Kaya)
        payload_b = {
            "type": "kokpit-identity",
            "source": "kokpit-iframe-bridge",
            "user": {
                "sicil": "67890",
                "isim": "Mehmet Kaya",
                "username": "mehmet.kaya",
                "role": "User"
            },
            "token": "token-user-b",
            "inputs": {"engine_mode": "local", "threshold": 75}
        }
        res_b = self.client.post("/api/identity", json=payload_b, headers={"X-Session-ID": "sess_user_b"})
        self.assertEqual(res_b.status_code, 200)
        self.assertEqual(res_b.json()["session_id"], "sess_user_b")

        # 3. Kullanıcı C henüz kimlik göndermemiş yeni bir sekmedir
        res_c = self.client.get("/api/identity", headers={"X-Session-ID": "sess_user_c"})
        self.assertEqual(res_c.status_code, 200)
        data_c = res_c.json()
        self.assertFalse(data_c["authenticated"])
        self.assertIsNone(data_c["user"])
        self.assertEqual(data_c["status"], "waiting")
        self.assertEqual(data_c["session_id"], "sess_user_c")

        # 4. Kullanıcı A kendi oturumunu sorguladığında Ahmet Yılmaz döner
        get_a = self.client.get("/api/identity", headers={"X-Session-ID": "sess_user_a"})
        self.assertEqual(get_a.status_code, 200)
        data_a = get_a.json()
        self.assertTrue(data_a["authenticated"])
        self.assertEqual(data_a["user"]["isim"], "Ahmet Yılmaz")
        self.assertEqual(data_a["user"]["role"], "Admin")
        self.assertEqual(data_a["external_inputs"]["threshold"], 85)

        # 5. Kullanıcı B kendi oturumunu sorguladığında Mehmet Kaya döner
        get_b = self.client.get("/api/identity", headers={"X-Session-ID": "sess_user_b"})
        self.assertEqual(get_b.status_code, 200)
        data_b = get_b.json()
        self.assertTrue(data_b["authenticated"])
        self.assertEqual(data_b["user"]["isim"], "Mehmet Kaya")
        self.assertEqual(data_b["user"]["role"], "User")
        self.assertEqual(data_b["external_inputs"]["threshold"], 75)

    def test_refresh_simulation_preserves_session(self):
        # Sayfa yenileme simülasyonu: Cookie veya başlık üzerinden aynı oturum korunmalıdır
        res_first = self.client.get("/api/identity", headers={"X-Session-ID": "sess_persisted_tab"})
        self.assertEqual(res_first.json()["session_id"], "sess_persisted_tab")

        # Kimlik ekle
        self.client.post("/api/identity", json={
            "user": {"sicil": "11111", "isim": "Ayşe Demir", "role": "User"}
        }, headers={"X-Session-ID": "sess_persisted_tab"})

        # F5 / Yenileme: Aynı session_id ile tekrar çağrıldığında Ayşe Demir oturumu korunur
        res_f5 = self.client.get("/api/identity", headers={"X-Session-ID": "sess_persisted_tab"})
        self.assertTrue(res_f5.json()["authenticated"])
        self.assertEqual(res_f5.json()["user"]["isim"], "Ayşe Demir")

    def test_logout_isolation(self):
        # A ve B oturumları açılır
        self.client.post("/api/identity", json={"user": {"sicil": "A1", "isim": "User A"}}, headers={"X-Session-ID": "sess_logout_a"})
        self.client.post("/api/identity", json={"user": {"sicil": "B1", "isim": "User B"}}, headers={"X-Session-ID": "sess_logout_b"})

        # Yalnızca A oturumu kapatılır
        del_a = self.client.delete("/api/identity", headers={"X-Session-ID": "sess_logout_a"})
        self.assertEqual(del_a.status_code, 200)

        # A oturumu kapalı (unauthenticated) olmalı
        check_a = self.client.get("/api/identity", headers={"X-Session-ID": "sess_logout_a"})
        self.assertFalse(check_a.json()["authenticated"])

        # B oturumu ise aktif kalmaya devam etmeli
        check_b = self.client.get("/api/identity", headers={"X-Session-ID": "sess_logout_b"})
        self.assertTrue(check_b.json()["authenticated"])
        self.assertEqual(check_b.json()["user"]["isim"], "User B")

    def test_external_inputs_isolation(self):
        # Kullanıcı X özel embedding girdileri yollar
        self.client.post("/api/external-input", json={
            "session_id": "sess_input_x",
            "items_1": ["Örnek X1"],
            "items_2": ["Örnek X2"],
            "model": "ozel-model-x"
        }, headers={"X-Session-ID": "sess_input_x"})

        # Kullanıcı Y external inputlarını sorguladığında Kullanıcı X'in verisini görmemeli
        get_y = self.client.get("/api/external-input", headers={"X-Session-ID": "sess_input_y"})
        self.assertEqual(get_y.status_code, 200)
        self.assertEqual(get_y.json().get("external_inputs", {}), {})

        # Kullanıcı X ise kendi girdisini görmeli
        get_x = self.client.get("/api/external-input", headers={"X-Session-ID": "sess_input_x"})
        self.assertEqual(get_x.status_code, 200)
        self.assertIn("items_1", get_x.json().get("external_inputs", {}))

    def test_is_chat_llm_model_and_filtering(self):
        from app import is_chat_llm_model, get_effective_aselix_config, fetch_aselix_models
        # Non-chat models must be rejected
        self.assertFalse(is_chat_llm_model("FLUX.2-dev"))
        self.assertFalse(is_chat_llm_model("stabilityai/stable-diffusion-xl-base-1.0"))
        self.assertFalse(is_chat_llm_model("openai/whisper-large-v3"))
        self.assertFalse(is_chat_llm_model("BAAI/bge-m3"))
        self.assertFalse(is_chat_llm_model("intfloat/multilingual-e5-large"))

        # Valid chat LLMs must be accepted
        self.assertTrue(is_chat_llm_model("openai/gpt-oss-120b"))
        self.assertTrue(is_chat_llm_model("meta-llama/Meta-Llama-3-8B-Instruct"))
        self.assertTrue(is_chat_llm_model("mistralai/Mistral-7B-Instruct-v0.3"))

        # Non-chat model in get_effective_aselix_config should fallback to openai/gpt-oss-120b
        _, _, eff_model, _ = get_effective_aselix_config(model="FLUX.2-dev")
        self.assertEqual(eff_model, "openai/gpt-oss-120b")

        # fetch_aselix_models endpoint should filter out FLUX/whisper and prioritize openai/gpt-oss-120b
        res = self.client.get("/api/aselix/models")
        self.assertEqual(res.status_code, 200)
        models = res.json().get("modeller", [])
        self.assertIn("openai/gpt-oss-120b", models)
        self.assertEqual(models[0], "openai/gpt-oss-120b")
        for m in models:
            self.assertTrue(is_chat_llm_model(m), f"Non-chat model leaked into models list: {m}")


if __name__ == "__main__":
    unittest.main()
