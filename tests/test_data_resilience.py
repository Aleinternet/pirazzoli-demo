"""Offline regression tests: python -m unittest discover -s tests -v."""

import io
import unittest
from datetime import date, timedelta
from unittest.mock import Mock, patch

import requests
from openpyxl import Workbook

from api import index as api


def response(data=None, status=200):
    result = Mock(status_code=status)
    result.json.return_value = data
    if status >= 400:
        result.raise_for_status.side_effect = requests.HTTPError("private URL?apikey=secret")
    return result


def workbook(company="MINERVA", rows=3):
    book = Workbook()
    sheet = book.active
    sheet.title = "IMPO MAR"
    sheet.append(["REFERENCIA CLIENTE", "NUMERO DE DESPACHO", "FECHA DE LIBERACION", "T/C LIBERACION CARGA"])
    for index in range(rows):
        sheet.append([f"{company}-{index}", str(index), date(2026, 4, 1) + timedelta(days=index), None])
    data = io.BytesIO()
    book.save(data)
    return {"content": data.getvalue(), "file_name": f"Comex_{company}_2026_01.xlsx", "last_modified": "2026-10-14T10:00:00Z"}


class DataResilienceTests(unittest.TestCase):
    def setUp(self):
        for cache in (api.CMF_RATE_CACHE, api.PDF_DATE_CACHE, api.CMF_RETRY_AFTER, api.OPTIONAL_BLOCKED_UNTIL):
            cache.clear()
        api.API_CACHE.update(payload=None, last_build_ts=0, last_modified=None)
        self.addCleanup(patch.stopall)
        patch.object(api, "CMF_API_KEY", "test-key-not-a-credential").start()
        # All tests must stay offline, even if a new unmocked HTTP call is added.
        patch.object(requests.sessions.Session, "request", side_effect=AssertionError("Unexpected network request")).start()
        self.client = api.app.test_client()

    def test_cold_start_cmf_timeout_keeps_every_company_and_row(self):
        files = [workbook("MINERVA", 80), workbook("OTRA", 80)]
        with patch.object(api, "get_token", return_value="test-token"), \
             patch.object(api, "fetch_excel_files", return_value=files), \
             patch.object(api.requests, "get", side_effect=requests.ConnectTimeout("apikey=secret")) as get:
            result = self.client.get("/api")
        self.assertEqual(result.status_code, 200)
        payload = result.get_json()
        self.assertTrue(payload["ok"])
        self.assertTrue(payload["ratesPending"])
        self.assertEqual(len(payload["files"]), 2)
        self.assertEqual([len(f["sheets"][0]["rows"]) for f in payload["files"]], [80, 80])
        self.assertEqual(get.call_count, 1, "Do not repeat timeout for each row/company")
        self.assertLessEqual(sum(get.call_args.kwargs["timeout"]), 5)
        self.assertNotIn("secret", result.get_data(as_text=True))
        self.assertEqual(payload["files"][0]["sheets"][0]["rows"][0]["values"]["T/C LIBERACION CARGA"], "")

    def test_http_errors_and_malformed_responses_do_not_escape(self):
        cases = [
            requests.ReadTimeout("secret"), requests.ConnectionError("secret"),
            response(status=401), response(status=429), response(status=500),
            response(data=[]), response(data={"Dolares": [{"Fecha": "bad-date", "Valor": "900,00"}]}),
            response(data={"Dolares": [{"Fecha": "2026-04-01", "Valor": "NaN"}]}),
            response(data={"Dolares": [{"Fecha": "2026-04-01", "Valor": "0"}]}),
            response(data={"Dolares": "invalid"}),
        ]
        malformed_json = response()
        malformed_json.json.side_effect = ValueError("private upstream body")
        cases.append(malformed_json)
        for case in cases:
            with self.subTest(case=type(case).__name__):
                api.OPTIONAL_BLOCKED_UNTIL.clear()
                budget = api.EnrichmentBudget()
                kwargs = {"side_effect": case} if isinstance(case, Exception) else {"return_value": case}
                with patch.object(api.requests, "get", **kwargs) as get:
                    self.assertIsNone(api.fetch_cmf_dollar_for_date(date(2026, 4, 1), {}, budget))
                    self.assertIsNone(api.fetch_cmf_dollar_for_date(date(2026, 4, 2), {}, budget))
                self.assertEqual(get.call_count, 1)
                self.assertTrue(budget.pending)

    def test_circuit_recovers_and_successful_values_are_cached(self):
        good = response({"Dolares": [{"Fecha": "2026-04-01", "Valor": "950,12"}]})
        with patch.object(api.time, "monotonic", return_value=100) as clock, \
             patch.object(api.requests, "get", side_effect=[requests.ConnectTimeout(), good]) as get:
            self.assertIsNone(api.fetch_cmf_dollar_for_date(date(2026, 4, 1), api.CMF_RATE_CACHE))
            self.assertNotIn("2026-04-01", api.CMF_RATE_CACHE)
            clock.return_value = 161
            self.assertEqual(api.fetch_cmf_dollar_for_date(date(2026, 4, 1), api.CMF_RATE_CACHE), "950,12")
            self.assertEqual(api.fetch_cmf_dollar_for_date(date(2026, 4, 1), api.CMF_RATE_CACHE), "950,12")
            self.assertEqual(get.call_count, 2)

    def test_weekend_uses_closest_previous_rate_never_a_future_rate(self):
        history = {"Dolares": [
            {"Fecha": "2026-04-02", "Valor": "900,00"},
            {"Fecha": "2026-04-06", "Valor": "999,00"},
            {"Fecha": "2026-04-03", "Valor": "950,12"},
        ]}
        with patch.object(api.requests, "get", side_effect=[response(status=404), response(history)]) as get:
            value = api.fetch_cmf_dollar_for_date(date(2026, 4, 5), {})
        self.assertEqual(value, "950,12")
        self.assertEqual(get.call_count, 2)

    def test_missing_date_is_retried_after_cooldown(self):
        good = response({"Dolares": [{"Fecha": "2026-04-01", "Valor": "950,12"}]})
        with patch.object(api.time, "monotonic", return_value=100) as clock, \
             patch.object(api.requests, "get", side_effect=[response(status=404), response(status=404), good]) as get:
            self.assertIsNone(api.fetch_cmf_dollar_for_date(date(2026, 4, 1), api.CMF_RATE_CACHE))
            self.assertIsNone(api.fetch_cmf_dollar_for_date(date(2026, 4, 1), api.CMF_RATE_CACHE))
            self.assertEqual(get.call_count, 2)
            clock.return_value = 161
            self.assertEqual(api.fetch_cmf_dollar_for_date(date(2026, 4, 1), api.CMF_RATE_CACHE), "950,12")

    def test_shared_budget_stops_new_requests_but_cached_values_still_work(self):
        budget = api.EnrichmentBudget()
        with patch.object(api.time, "monotonic", return_value=100) as clock, \
             patch.object(api.requests, "get", return_value=response({"Dolares": [{"Fecha": "2026-04-01", "Valor": "950,12"}]})) as get:
            self.assertEqual(api.fetch_cmf_dollar_for_date(date(2026, 4, 1), api.CMF_RATE_CACHE, budget), "950,12")
            clock.return_value = 109
            for day in range(2, 29):
                self.assertIsNone(api.fetch_cmf_dollar_for_date(date(2026, 4, day), api.CMF_RATE_CACHE, budget))
            self.assertEqual(api.fetch_cmf_dollar_for_date(date(2026, 4, 1), api.CMF_RATE_CACHE, budget), "950,12")
            self.assertEqual(get.call_count, 1)
        self.assertTrue(budget.pending)

    def test_existing_excel_rates_are_preserved(self):
        headers = ["FECHA DE LIBERACION", "T/C LIBERACION CARGA"]
        rows = [{"values": {headers[0]: "01/04/2026", headers[1]: "945,50"}}]
        with patch.object(api.requests, "get") as get:
            api.enrich_exchange_rate_columns(rows, headers, "test")
        get.assert_not_called()
        self.assertEqual(rows[0]["values"][headers[1]], "945,50")

    def test_pdf_outage_is_optional_and_does_not_poison_cache(self):
        headers = ["COMPROBANTE PAGO TGR", "T/C DOLAR LEGALIZACION"]
        rows = [{"values": {headers[0]: {"text": "PDF", "url": f"https://example.test/{i}.pdf"}, headers[1]: ""}} for i in range(50)]
        budget = api.EnrichmentBudget()
        with patch.object(api.requests, "get", side_effect=requests.ReadTimeout()) as get:
            self.assertIs(api.enrich_exchange_rate_columns(rows, headers, "test", budget), rows)
        self.assertEqual(get.call_count, 1)
        self.assertTrue(budget.pending)
        self.assertEqual(api.PDF_DATE_CACHE, {})

    def test_degraded_payload_rebuilds_soon_and_recovers(self):
        good = response({"Dolares": [{"Fecha": "2026-04-01", "Valor": "950,12"}]})
        with patch.object(api, "get_token", return_value="test"), \
             patch.object(api, "fetch_excel_files", return_value=[workbook(rows=1)]) as files, \
             patch.object(api.time, "time", return_value=1000) as wall, \
             patch.object(api.time, "monotonic", return_value=100) as clock, \
             patch.object(api.requests, "get", side_effect=[requests.ConnectTimeout(), good]):
            self.assertTrue(api.build_payload()["ratesPending"])
            self.assertTrue(api.build_payload()["ratesPending"])
            self.assertEqual(files.call_count, 1)
            wall.return_value = 1061
            clock.return_value = 161
            recovered = api.build_payload()
            self.assertFalse(recovered["ratesPending"])
            self.assertEqual(files.call_count, 2)

    def test_essential_failure_returns_safe_error_or_marked_cached_data(self):
        with patch.object(api, "build_payload", side_effect=RuntimeError("https://private/?apikey=secret")):
            unavailable = self.client.get("/api")
            self.assertEqual(unavailable.status_code, 503)
            self.assertFalse(unavailable.get_json()["ok"])
            self.assertNotIn("secret", unavailable.get_data(as_text=True))
            api.API_CACHE["payload"] = {"ok": True, "files": [], "lastModified": "2026-10-14"}
            stale = self.client.get("/api")
            self.assertEqual(stale.status_code, 200)
            self.assertTrue(stale.get_json()["stale"])
            self.assertNotIn("secret", stale.get_data(as_text=True))
            self.assertNotIn("stale", api.API_CACHE["payload"])

    def test_history_does_not_expose_upstream_errors(self):
        with patch.object(api.requests, "get", side_effect=requests.ConnectTimeout("apikey=secret")):
            result = self.client.get("/api/row-history?file_key=test&sheet_name=test&row_identity=test")
        self.assertEqual(result.status_code, 503)
        self.assertNotIn("secret", result.get_data(as_text=True))


if __name__ == "__main__":
    unittest.main()
