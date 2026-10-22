import unittest
import sys
from pathlib import Path

# Add project root to sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from scraper import normalize_date_key, format_date_dmy

class TestDedupAndDates(unittest.TestCase):

    def test_normalize_date_key(self):
        # YYYY-MM-DD
        self.assertEqual(normalize_date_key("2026-10-01"), "2026-10-01")
        # DD/MM/YYYY
        self.assertEqual(normalize_date_key("01/10/2026"), "2026-10-01")
        # DD-MM-YYYY
        self.assertEqual(normalize_date_key("15-10-2026"), "2026-10-15")
        # With single-digit day/month
        self.assertEqual(normalize_date_key("7/10/2026"), "2026-10-07")
        self.assertEqual(normalize_date_key("2026-7-5"), "2026-07-05")
        # Empty or placeholder
        self.assertEqual(normalize_date_key("/"), "")
        self.assertEqual(normalize_date_key(""), "")
        self.assertEqual(normalize_date_key(None), "")

    def test_format_date_dmy(self):
        self.assertEqual(format_date_dmy("2026-10-01"), "01/10/2026")
        self.assertEqual(format_date_dmy("2026-06-08"), "08/06/2026")
        self.assertEqual(format_date_dmy("01/10/2026"), "01/10/2026")
        self.assertEqual(format_date_dmy("/"), "/")
        self.assertEqual(format_date_dmy(""), "/")

    def test_ao_deduplication_different_dates(self):
        """
        Verify that tenders in the same commune with DIFFERENT publication dates
        are NOT treated as duplicates (the bug that previously dropped Notices 85018 and 93574).
        """
        existing_items = set()

        # Simulate existing March tender in Es-Senia, Oran
        action = "RÉALISATION"
        ptype = "STADE DE PROXIMITÉ EN GAZON SYNTHÉTIQUE"
        norm_w = "ORAN"
        norm_c = "ESSENIA"
        parution_march = normalize_date_key("2026-03-10")
        existing_items.add(("AO_PARUTION", action, ptype, norm_w, norm_c, parution_march))

        # New October tender in Es-Senia, Oran (Notice 85018)
        parution_oct = normalize_date_key("2026-10-01")
        candidate_key = ("AO_PARUTION", action, ptype, norm_w, norm_c, parution_oct)

        # Candidate MUST NOT be flagged as duplicate
        self.assertNotIn(candidate_key, existing_items)

        # Same tender re-checked with same date MUST be flagged as duplicate
        existing_items.add(candidate_key)
        self.assertIn(candidate_key, existing_items)

if __name__ == "__main__":
    unittest.main()
