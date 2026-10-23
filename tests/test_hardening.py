"""
Regression tests for the October 2026 audit: missed notices, merged lots/budgets, swapped dates,
shifted Sheet rows, dead OCR models, silent crawl stops, and the TAPIS/MOQUETTE scope.
"""
import json
import logging
import sys
import unittest
from pathlib import Path
from unittest import mock

import requests

# Add project root to sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import scraper
from gspread.utils import a1_to_rowcol
from scraper import (
    AlgerieMarchesScraper,
    ListingFetchError,
    ai_confirmed_scope,
    append_to_gsheet,
    classify_type_projet,
    clean_annonceur,
    clean_date_parution_ao,
    detect_scope,
    extract_ad_ids,
    merge_ai_data,
    normalize_ai_data,
    normalize_delai,
    parse_budget,
    parse_commune,
    prefilter_decision,
    sheet_text,
    to_local_date,
)

# Keep test runs out of scraper.log
scraper.log.setLevel(logging.CRITICAL)


class TestScope(unittest.TestCase):

    def test_gazon_plurals_and_inflections(self):
        for text in [
            "Réalisation de deux terrains en gazons synthétiques",
            "Pose de pelouses naturelles",
            "Stade gazonné de la commune",
            "Travaux d'engazonnement du stade",
        ]:
            with self.subTest(text=text):
                self.assertEqual(detect_scope(text), "GAZON")

    def test_gazon_arabic_with_prefixes(self):
        for text in [
            "تكسية الملعب بالعشب الاصطناعي",
            "انجاز ملعب جواري بالعشب",
            "أشغال التعشيب الاصطناعي",
            "تكسية اصطناعية لملعب كرة القدم",
        ]:
            with self.subTest(text=text):
                self.assertEqual(detect_scope(text), "GAZON")

    def test_tapis_and_moquette_for_mosques_and_buildings(self):
        for text in [
            "Fourniture et pose de tapis pour les mosquées (lof unique)",
            "Fourniture et pose de moquette pour le siège de la wilaya",
            "Moquettage des salles de conférence",
            "اقتناء زرابي للمساجد",
            "تزويد المسجد بالسجاد",
            "Fourniture de tapis �  la mosquée El Feth",
        ]:
            with self.subTest(text=text):
                self.assertEqual(detect_scope(text), "TAPIS")

    def test_tapis_that_is_not_a_carpet(self):
        for text in [
            "Acquisition de tapis roulants pour la salle de sport",
            "Renforcement de la chaussée par un tapis en béton bitumineux",
            "Mise en oeuvre d'un tapis d'enrobé sur le CW 12",
            "Acquisition de tatamis et tapis de judo",
            "Fourniture de convoyeurs à tapis",
        ]:
            with self.subTest(text=text):
                self.assertIsNone(detect_scope(text))

    def test_gazon_equipment_is_not_a_turf_project(self):
        self.assertIsNone(detect_scope("Acquisition de Matériel Agricole (tondeuses �  gazon)"))
        self.assertIsNone(detect_scope("Tonte du gazon des espaces verts"))

    def test_gazon_carpet_counts_as_gazon(self):
        self.assertEqual(detect_scope("Fourniture et pose de tapis en gazon synthétique"), "GAZON")

    def test_prefilter_decisions(self):
        cases = [
            ("Réalisation aires de jeux en gazon synthétique", "explicit"),
            ("Fourniture et pose de tapis pour les mosquées", "explicit"),
            ("Réalisation de deux stades de proximité", "ocr"),
            ("Réalisation d'un stade avec drainage et éclairage public", "ocr"),
            ("Equipement de la mosquée El Feth", "ocr"),
            ("Aménagement de la route menant vers le stade", "skip"),
            ("Achèvement de la salle de sport", "skip"),
            ("Réalisation d'une salle de sports au lycée brahmi ismail", "skip"),
            ("fourniture de divers articles (entretien, literie, tissus, ménagers, sports et quincaillerie)", "skip"),
            ("Acquisition de Matériel Agricole (tondeuses à gazon)", "skip"),
            ("Réalisation d'un lycée type 800/200", "skip"),
        ]
        for text, expected in cases:
            with self.subTest(text=text):
                self.assertEqual(prefilter_decision(text)[0], expected)

    def test_type_projet_for_tapis(self):
        self.assertEqual(classify_type_projet("Fourniture et pose de tapis pour les mosquées"), "TAPIS MOSQUÉE")
        self.assertEqual(classify_type_projet("Fourniture de moquette pour les bureaux"), "MOQUETTE")
        self.assertEqual(classify_type_projet("Fourniture de tapis pour la maison de jeunes"), "TAPIS")
        self.assertEqual(classify_type_projet("Fourniture et pose de tapis en gazon synthétique"), "GAZON SYNTHÉTIQUE")


class TestAmounts(unittest.TestCase):

    def test_amounts_never_merge_with_dates(self):
        self.assertEqual(parse_budget("du 12/08/2026 : 19 736 328,50 DA"), 19736328.5)
        self.assertEqual(parse_budget("19.736.328,50 DA du 12/08/2026"), 19736328.5)
        self.assertEqual(parse_budget("attribué le 15.06.2026 1 250 000,00 DA"), 1250000)
        self.assertEqual(parse_budget("15.06.2026"), "/")

    def test_ht_and_ttc_pick_ttc(self):
        self.assertEqual(parse_budget("12 500 000,00 DA HT soit 14 875 000,00 DA TTC"), 14875000)
        self.assertEqual(parse_budget("HT : 12 500 000 DA ; TTC : 14 875 000 DA"), 14875000)

    def test_formats(self):
        self.assertEqual(parse_budget("770 525.00 دج"), 770525)
        self.assertEqual(parse_budget("12 500 000,00 DA"), 12500000)
        self.assertEqual(parse_budget("263.139.211,58 DA"), 263139211.58)
        self.assertEqual(parse_budget("25,892,020.00"), 25892020)
        self.assertEqual(parse_budget("Lot 3 1 250 000 DA"), 1250000)
        self.assertEqual(parse_budget(12500000), 12500000)
        self.assertEqual(parse_budget(None), "/")
        self.assertEqual(parse_budget("montant non communiqué"), "/")


class TestTextFields(unittest.TestCase):

    def test_parse_commune(self):
        self.assertEqual(parse_commune("Réalisation d'un stade", "COMMUNE DE OULED DERRADJ"), "OULED DERRADJ")
        self.assertEqual(parse_commune("Réalisation", "COMMUNE DE MEDEA"), "MEDEA")
        self.assertEqual(parse_commune("Réalisation", "Commune d'Es Sénia"), "ES SENIA")
        self.assertEqual(parse_commune("Aménagement du stade EL REMAL commune de TEBESSA", "DJS"), "TEBESSA")
        self.assertEqual(parse_commune("Réalisation d'un stade", "DEP de la wilaya"), "/")

    def test_clean_annonceur_word_boundaries(self):
        self.assertEqual(clean_annonceur("Direction des equipements publics de la wilaya de Chlef"), "DEP DE LA WILAYA")
        self.assertEqual(clean_annonceur("SARL INDEPENDANT TRAVAUX"), "SARL INDEPENDANT TRAVAUX")
        self.assertEqual(clean_annonceur("Assemblée populaire communale de Baraki"), "COMMUNE")
        self.assertEqual(clean_annonceur(""), "/")

    def test_delai_ignores_contestation_window(self):
        text = "dans un délai de 10 jours à compter de la première publication. Délai d'exécution : 45 jours"
        self.assertEqual(normalize_delai(text), "45 JOURS")
        self.assertEqual(normalize_delai("مدة الإنجاز 03 أشهر"), "03 MOIS")
        self.assertEqual(normalize_delai("quarante-cinq (45) jours"), "45 JOURS")
        self.assertEqual(normalize_delai("15 Jours"), "15 JOURS")
        self.assertEqual(normalize_delai(None), "/")


class TestDates(unittest.TestCase):

    def test_to_local_date_uses_algerian_time(self):
        self.assertEqual(to_local_date("2026-10-21T23:00:00.000Z"), "2026-10-22")
        self.assertEqual(to_local_date("2026-10-22T00:00:00.000Z"), "2026-10-22")
        self.assertEqual(to_local_date("2026-10-22"), "2026-10-22")
        self.assertEqual(to_local_date(None), "")

    def test_date_parution_ao_validation(self):
        self.assertEqual(clean_date_parution_ao("2026/06/08", "2026-10-21"), "08/06/2026")
        self.assertEqual(clean_date_parution_ao("08-06-2026", "2026-10-21"), "08/06/2026")
        self.assertEqual(clean_date_parution_ao("21/11/2026", "2026-10-21"), "/")  # after the award
        self.assertEqual(clean_date_parution_ao("08/06/2015", "2026-10-21"), "/")  # decree year
        self.assertEqual(clean_date_parution_ao("31/02/2026", "2026-10-21"), "/")  # not a date
        self.assertEqual(clean_date_parution_ao("/", "2026-10-21"), "/")

    def test_sheet_text_forces_literal_dates(self):
        self.assertEqual(sheet_text("02/11/2026"), "'02/11/2026")
        self.assertEqual(sheet_text("/"), "/")
        self.assertEqual(sheet_text(""), "/")

    def test_extract_ad_ids(self):
        self.assertEqual(extract_ad_ids('=HYPERLINK("https://github.com/x/tree/main/data/scans/2026-10-19/95953", "S")'), {"95953"})
        self.assertEqual(extract_ad_ids('=HYPERLINK("https://algeriemarches.com/annonces/396720-tapis#ad-97497", "A")'), {"97497"})


class TestAINormalization(unittest.TestCase):

    def test_odd_value_types_do_not_crash(self):
        data = normalize_ai_data({"budget": 12500000, "wilaya": None, "commune": ["ORAN", None],
                                  "is_gazon": "true", "gazon_mention": "gazon synthétique"})
        self.assertEqual(data["budget"], "12500000")
        self.assertEqual(data["wilaya"], "/")
        self.assertEqual(data["commune"], "ORAN")
        self.assertEqual(data["scope"], "GAZON")
        self.assertEqual(data["scope_evidence"], "gazon synthétique")
        self.assertEqual(normalize_ai_data({}), {})
        self.assertEqual(normalize_ai_data("not json"), {})

    def test_lots_are_normalized(self):
        data = normalize_ai_data({"scope": "tapis", "lots": [
            {"numero": 1, "objet": "Tapis mosquée A", "in_scope": "true", "budget": "2 562 004,79 DA"},
            "garbage",
        ]})
        self.assertEqual(data["scope"], "TAPIS")
        self.assertEqual(len(data["lots"]), 1)
        self.assertTrue(data["lots"][0]["in_scope"])
        self.assertEqual(data["lots"][0]["numero"], "1")
        self.assertEqual(data["lots"][0]["delai"], "/")

    def test_second_page_can_confirm_scope(self):
        page1 = normalize_ai_data({"scope": "NONE", "scope_evidence": "NONE", "commune": "MILA"})
        page2 = normalize_ai_data({"scope": "GAZON", "scope_evidence": "revêtement en gazon synthétique", "budget": "1 000 000 DA"})
        merged = merge_ai_data(page1, page2)
        self.assertEqual(merged["scope"], "GAZON")
        self.assertEqual(merged["commune"], "MILA")
        self.assertEqual(merged["budget"], "1 000 000 DA")
        self.assertEqual(ai_confirmed_scope(merged), "GAZON")

    def test_scope_without_real_evidence_is_not_trusted(self):
        hallucinated = normalize_ai_data({"scope": "GAZON", "scope_evidence": "aménagement du stade communal"})
        self.assertIsNone(ai_confirmed_scope(hallucinated))


AO_HEADER = ['', '', 'N°', 'DATE', "TITRE D'APPEL D'OFFRE", 'Nombre de projet', 'TYPE DE PROJET', 'DATE DE PARUTION',
             "DATE D'ECHEANCE", 'ANNONCEUR ', 'WILAYA ', 'SCANS', 'CATEGORIE', 'SCANS']
AA_HEADER = ['N°', 'DATE', "TITRE D'AVIS D'ATTRIBUTION", ' Nombre de projet', 'TYPE DE PROJET', 'DATE DE PARUTION   ',
             "DATE D'ECHEANCE", 'ANNONCEUR ', 'WILAYA ', 'COMMUNE', "DATE D'ATRRIBUTION", 'BUDGET', 'ATTRIBUTION PAR',
             'DELAI', 'WILAYA', 'SCANS']


class FakeWorksheet:
    """Minimal gspread Worksheet stand-in that records writes."""

    def __init__(self, title, rows, row_count=None):
        self.title = title
        self.values = [list(r) for r in rows]
        self.row_count = row_count if row_count is not None else len(rows) + 10
        self.updates = []

    def get_all_values(self, value_render_option=None):
        return [list(r) for r in self.values]

    def add_rows(self, n):
        self.row_count += n

    def update(self, values=None, range_name=None, value_input_option=None):
        self.updates.append((range_name, values, value_input_option))
        row, col = a1_to_rowcol(range_name)
        for i, vals in enumerate(values):
            while len(self.values) < row + i:
                self.values.append([])
            target = self.values[row + i - 1]
            while len(target) < col - 1 + len(vals):
                target.append("")
            target[col - 1:col - 1 + len(vals)] = [str(v) for v in vals]


class TestSheetWriter(unittest.TestCase):

    def ao_sheet(self, header=AO_HEADER, row_count=None):
        existing = ['', '', '276', '19/10/2026', 'RÉALISATION', '1', 'STADE COMMUNAL', '19/10/2026', '02/11/2026', 'COMMUNE',
                    'TIARET', 'FRENDA', 'TRAVAUX PUBLICS',
                    '=HYPERLINK("https://github.com/x/tree/main/data/scans/2026-10-19/95953", "📁 Voir 2 Scans")']
        return FakeWorksheet("APPEL D'OFFRE 2026", [['', '', '', '', 'Mise à jour'], [], header, existing], row_count)

    def ao_item(self, **overrides):
        item = {"id": "97781", "slug": "123-terrains", "titre": "Réalisation de 06 terrains sportifs en gazon synthétique",
                "wilaya": "Sétif", "commune": "BOUTALEB", "annonceur": "DJS de la wilaya de Sétif",
                "date_parution": "2026-10-21", "date_echeance": "2026-11-04", "scans": ["/x/AM_97781_a.jpg"]}
        item.update(overrides)
        return item

    def test_ao_row_starts_at_n_column_without_padding(self):
        ws = self.ao_sheet()
        self.assertEqual(append_to_gsheet([self.ao_item()], "appels-doffres", worksheet=ws), 1)
        range_name, values, option = ws.updates[0]
        self.assertEqual(range_name, "C5")
        self.assertEqual(option, "USER_ENTERED")
        row = values[0]
        self.assertEqual(len(row), 12)
        self.assertEqual(row[0], 277)
        self.assertEqual(row[1], "'21/10/2026")   # DATE
        self.assertEqual(row[5], "'21/10/2026")   # DATE DE PARUTION
        self.assertEqual(row[6], "'04/11/2026")   # DATE D'ECHEANCE: stays 4 Nov, not April 11
        self.assertIn("/97781/", row[11])

    def test_ao_duplicate_id_is_skipped(self):
        ws = self.ao_sheet()
        self.assertEqual(append_to_gsheet([self.ao_item(id="95953")], "appels-doffres", worksheet=ws), 0)
        self.assertEqual(ws.updates, [])

    def test_ao_notice_without_scan_keeps_its_id(self):
        ws = self.ao_sheet()
        append_to_gsheet([self.ao_item(scans=[])], "appels-doffres", worksheet=ws)
        self.assertIn("#ad-97781", ws.updates[0][1][0][11])
        self.assertEqual(append_to_gsheet([self.ao_item(scans=[], date_echeance="2026-11-20")], "appels-doffres", worksheet=ws), 0)

    def test_grows_sheet_when_full(self):
        ws = self.ao_sheet(row_count=4)
        append_to_gsheet([self.ao_item()], "appels-doffres", worksheet=ws)
        self.assertGreaterEqual(ws.row_count, 5)

    def test_refuses_to_write_when_layout_changed(self):
        ws = self.ao_sheet(header=AO_HEADER[:-1] + ["NOTES"])
        self.assertEqual(append_to_gsheet([self.ao_item()], "appels-doffres", worksheet=ws), -1)
        self.assertEqual(ws.updates, [])

    def aa_item(self):
        lots = [
            {"numero": "01", "objet": "Lot 01 aire de jeux en gazon synthétique cité A", "in_scope": True, "commune": "/",
             "budget": "6 959 358,00 DA", "entreprise_attributaire": "AOUACHERIA ALAA", "delai": "45 JOURS"},
            {"numero": "02", "objet": "Lot 02 aire de jeux en gazon synthétique cité B", "in_scope": True, "commune": "/",
             "budget": "5 663 686,00 DA", "entreprise_attributaire": "AOUACHERIA ALAA", "delai": "45 JOURS"},
            {"numero": "03", "objet": "Lot 03 aire de jeux en gazon synthétique cité C", "in_scope": True, "commune": "/",
             "budget": "/", "entreprise_attributaire": "AOUACHERIA ALAA", "delai": "55 JOURS"},
        ]
        return {"id": "97497", "slug": "1-aires", "titre": "Réalisation aires de jeux en gazon synthétique",
                "wilaya": "Constantine", "commune": "CONSTANTINE", "annonceur": "Commune de Constantine",
                "date_parution": "2026-10-21", "date_echeance": "2026-12-31", "date_parution_ao": "21/09/2026",
                "entreprise_concernee": "AOUACHERIA ALAA", "montant": "24 389 912,75 DA", "delai": "55 JOURS",
                "lots": lots, "scans": ["/x/a.jpg", "/x/b.jpg"]}

    def test_aa_one_row_per_lot_never_the_total(self):
        ws = FakeWorksheet("AVIS D'ATTRIBUTIONS 2026", [[], [], AA_HEADER])
        self.assertEqual(append_to_gsheet([self.aa_item()], "avis-attribution", worksheet=ws), 3)
        range_name, rows, _ = ws.updates[0]
        self.assertEqual(range_name, "A4")
        self.assertEqual([r[0] for r in rows], [1, 2, 3])
        self.assertEqual([r[11] for r in rows], [6959358, 5663686, "/"])
        self.assertEqual([r[13] for r in rows], ["45 JOURS", "45 JOURS", "55 JOURS"])
        self.assertEqual(rows[0][1], "'2026-10-21")
        self.assertEqual(rows[0][5], "'21/09/2026")
        self.assertEqual(rows[0][10], "'2026-10-21")
        self.assertEqual(rows[0][4], "AIRE DE JEUX")
        # Re-running the same award writes nothing
        self.assertEqual(append_to_gsheet([self.aa_item()], "avis-attribution", worksheet=ws), 0)

    def test_aa_same_award_under_another_id_is_skipped(self):
        ws = FakeWorksheet("AVIS D'ATTRIBUTIONS 2026", [[], [], AA_HEADER])
        append_to_gsheet([self.aa_item()], "avis-attribution", worksheet=ws)
        other = self.aa_item()
        other["id"] = "97498"  # Arabic version of the same notice
        self.assertEqual(append_to_gsheet([other], "avis-attribution", worksheet=ws), 0)


def fake_response(status, payload=None, text=""):
    r = mock.Mock()
    r.status_code = status
    r.json.return_value = payload or {}
    r.text = text or json.dumps(payload or {})
    return r


def gemini_ok(data):
    return fake_response(200, {"candidates": [{"content": {"parts": [{"text": json.dumps(data)}]}}]})


class TestOcrChain(unittest.TestCase):

    def setUp(self):
        scraper._nvidia_state.update(failures=0, disabled=False)
        patches = [
            mock.patch.object(scraper, "NVIDIA_API_KEY", "nv-test"),
            mock.patch.object(scraper, "GEMINI_API_KEY", "gm-test"),
            mock.patch.object(scraper.time, "sleep"),
        ]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        self.addCleanup(scraper._nvidia_state.update, failures=0, disabled=False)

    def test_nvidia_is_skipped_after_repeated_timeouts(self):
        calls = []

        def post(url, **kwargs):
            calls.append(url)
            if "nvidia" in url:
                raise requests.exceptions.Timeout("read timeout")
            return gemini_ok({"scope": "GAZON", "scope_evidence": "gazon"})

        with mock.patch.object(scraper.requests, "post", side_effect=post):
            for _ in range(5):
                self.assertEqual(scraper.analyze_scan(b"\xff\xd8 scan")["scope"], "GAZON")
        self.assertEqual(sum("nvidia" in u for u in calls), scraper.NVIDIA_MAX_CONSECUTIVE_FAILURES)
        self.assertTrue(scraper._nvidia_state["disabled"])

    def test_gemini_retries_rate_limits_then_falls_back(self):
        responses = {"gemini-2.5-flash": [fake_response(429), fake_response(429), fake_response(429)],
                     "gemini-flash-latest": [gemini_ok({"scope": "TAPIS", "scope_evidence": "tapis"})]}

        def post(url, **kwargs):
            if "nvidia" in url:
                return fake_response(503)
            model = url.split("/models/")[1].split(":")[0]
            return responses[model].pop(0)

        with mock.patch.object(scraper.requests, "post", side_effect=post):
            data = scraper.analyze_scan(b"\xff\xd8 scan")
        self.assertEqual(data["scope"], "TAPIS")
        self.assertEqual(responses["gemini-2.5-flash"], [])  # all three attempts used

    def test_total_failure_returns_empty(self):
        with mock.patch.object(scraper.requests, "post", return_value=fake_response(404)):
            self.assertEqual(scraper.analyze_scan(b"\xff\xd8 scan"), {})


def bare_scraper():
    s = AlgerieMarchesScraper.__new__(AlgerieMarchesScraper)
    s.failures, s.unverified, s.ad_errors, s._session_refreshed = [], [], 0, False
    s.drive_mgr = None
    return s


def listing(ann_id, date):
    return {"id_annonce": ann_id, "titre": f"Annonce {ann_id}", "slug": f"{ann_id}-x", "date_parution": f"{date}T00:00:00.000Z"}


class TestCrawl(unittest.TestCase):

    def setUp(self):
        p = mock.patch.object(scraper.time, "sleep")
        p.start()
        self.addCleanup(p.stop)

    def test_stops_only_when_a_whole_page_is_older(self):
        pages = {
            1: [listing(1, "2026-10-22"), listing(2, "2026-10-10")],  # a late upload on page 1 must not stop the crawl
            2: [listing(3, "2026-10-20"), listing(1, "2026-10-22")],  # shifted duplicate
            3: [listing(4, "2026-10-14"), listing(5, "2026-10-13")],
            4: [listing(6, "2026-10-12")],
        }
        s = bare_scraper()
        with mock.patch.object(s, "fetch_listings", side_effect=lambda page, notice_type: pages[page]) as fetch:
            got = s.collect_listings("appels-doffres", since_date="2026-10-15")
        self.assertEqual([a["id_annonce"] for a in got], [1, 3])
        self.assertEqual(fetch.call_count, 3)
        self.assertEqual(s.failures, [])

    def test_listing_failure_is_reported_not_silent(self):
        s = bare_scraper()

        def fetch(page, notice_type):
            if page == 2:
                raise ListingFetchError("page 2: HTTP 502")
            return [listing(page, "2026-10-22")]

        with mock.patch.object(s, "fetch_listings", side_effect=fetch):
            got = s.collect_listings("avis-attribution", since_date="2026-10-15")
        self.assertEqual(len(got), 1)
        self.assertEqual(len(s.failures), 1)

    def test_page_cap_is_reported(self):
        s = bare_scraper()
        with mock.patch.object(scraper, "MAX_PAGES_CATCHUP", 2), \
             mock.patch.object(s, "fetch_listings", side_effect=lambda page, notice_type: [listing(page, "2026-10-22")]):
            s.collect_listings("appels-doffres", since_date="2026-10-15")
        self.assertIn("page cap", s.failures[0])

    def test_fetch_listings_retries_then_raises(self):
        s = bare_scraper()
        s.s = mock.Mock()
        s.s.get.return_value = fake_response(502)
        with self.assertRaises(ListingFetchError):
            s.fetch_listings(1, "appels-doffres")
        self.assertEqual(s.s.get.call_count, scraper.LISTING_RETRIES * 2)

    def test_one_bad_notice_does_not_stop_the_cycle(self):
        s = bare_scraper()
        ads = [listing(10, "2026-10-22") | {"titre": "Pose de gazon synthétique A"},
               listing(11, "2026-10-22") | {"titre": "Pose de gazon synthétique B"},
               listing(12, "2026-10-22") | {"titre": "Achèvement d'un lycée"}]
        verdicts = {"13": {"verdict": "reject", "v": scraper.VERDICTS_VERSION, "checked": "2026-10-22"},
                    "14": {"verdict": "match", "v": scraper.VERDICTS_VERSION, "checked": "2026-10-22", "row": {"id": "14"}}}
        ads += [listing(13, "2026-10-22") | {"titre": "Stade 13"}, listing(14, "2026-10-22") | {"titre": "Gazon 14"}]

        def process(a, *args):
            if a["id_annonce"] == 10:
                raise AttributeError("'NoneType' object has no attribute 'strip'")
            return {"id": str(a["id_annonce"])}

        with mock.patch.object(s, "collect_listings", return_value=ads), \
             mock.patch.object(s, "process_ad", side_effect=process) as proc:
            results = s.scrape_notice_type("appels-doffres", "run", since_date="2026-10-15", verdicts=verdicts)
        self.assertEqual(sorted(r["id"] for r in results), ["11", "14"])
        self.assertEqual(s.ad_errors, 1)
        self.assertEqual(proc.call_count, 2)  # 12 filtered out, 13 cached reject, 14 cached match


class TestScopeGate(unittest.TestCase):

    def run_ad(self, titre, ai, scans=(Path("p1.jpg"),)):
        s = bare_scraper()
        verdicts = {}
        ad = listing(99999901, "2026-10-22") | {"titre": titre}
        decision, info = prefilter_decision(titre)
        with mock.patch.object(s, "fetch_detail_with_retry", return_value={"wilaya": "Oran", "annonceur": "Commune d'Oran"}), \
             mock.patch.object(s, "download_scans", return_value=list(scans)), \
             mock.patch.object(s, "read_scans", return_value=ai):
            row = s.process_ad(ad, "avis-attribution", "run", decision, info if decision == "explicit" else None, verdicts)
        return s, row, verdicts

    def test_explicit_text_wins_over_ai_no(self):
        _, row, verdicts = self.run_ad("Réalisation d'un terrain en gazon synthétique",
                                       normalize_ai_data({"scope": "NONE", "scope_evidence": "NONE"}))
        self.assertEqual(row["scope"], "GAZON")
        self.assertEqual(verdicts["99999901"]["verdict"], "match")

    def test_unreadable_scan_is_retried_not_rejected(self):
        s, row, verdicts = self.run_ad("Réalisation de deux stades de proximité", {})
        self.assertIsNone(row)
        self.assertEqual(s.unverified, ["99999901"])
        self.assertEqual(verdicts, {})

    def test_ai_confirmed_tapis_with_lots(self):
        ai = normalize_ai_data({"scope": "TAPIS", "scope_evidence": "Fourniture et pose de tapis pour les mosquées",
                                "objet_marche": "Fourniture et pose de tapis pour les mosquées", "lots": [
                                    {"objet": "Lot 1 tapis mosquée El Feth", "in_scope": True, "budget": "2 562 004,79 DA"},
                                    {"objet": "Lot 2 peinture façade", "in_scope": False, "budget": "900 000 DA"}]})
        _, row, _ = self.run_ad("Equipement de la mosquée El Feth", ai)
        self.assertEqual(row["scope"], "TAPIS")
        self.assertEqual(row["ptype"], "TAPIS MOSQUÉE")
        self.assertEqual([lot["objet"] for lot in row["lots"]], ["Lot 1 tapis mosquée El Feth"])

    def test_ai_no_rejects_and_caches(self):
        _, row, verdicts = self.run_ad("Réalisation de deux stades de proximité",
                                       normalize_ai_data({"scope": "NONE", "scope_evidence": "NONE"}))
        self.assertIsNone(row)
        self.assertEqual(verdicts["99999901"]["verdict"], "reject")


if __name__ == "__main__":
    unittest.main()
