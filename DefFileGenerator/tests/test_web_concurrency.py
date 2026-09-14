import io
import logging
import threading
import unittest

from fastapi.testclient import TestClient

from web.app import app

client = TestClient(app)


class TestWebConcurrency(unittest.TestCase):
    def setUp(self):
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_concurrent_uploads(self):
        # Spin up 10 threads simulating concurrent users uploading a file
        results = []

        def upload_worker(idx):
            csv_content = f"Address,Name,Type\n100,Var{idx},U16\n"
            response = client.post(
                "/api/convert",
                data={
                    "manufacturer": f"Mfg{idx}",
                    "model": f"Model{idx}",
                    "protocol": "modbusRTU",
                    "category": "Inverter",
                    "address_offset": 0,
                    "forced_write": "",
                },
                files={
                    "file": (f"test_{idx}.csv", io.BytesIO(csv_content.encode("utf-8")), "text/csv")
                },
            )
            results.append(
                (
                    idx,
                    response.status_code,
                    response.json() if response.status_code == 200 else None,
                )
            )

        threads = []
        for i in range(10):
            t = threading.Thread(target=upload_worker, args=(i,))
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        for idx, status, data in results:
            self.assertEqual(status, 200)
            self.assertTrue(data["success"])
            self.assertIn(f"mfg{idx}", data["filename"])
            self.assertIn(f"model{idx}", data["filename"])
            # Ensure the output CSV belongs to THIS request
            self.assertIn(f"Var{idx}", data["csv_content"])


if __name__ == "__main__":
    unittest.main()
