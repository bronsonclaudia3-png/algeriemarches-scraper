import unittest
import sys
from pathlib import Path

# Add project root to sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from scraper import (
    strip_accents,
    classify_action,
    classify_type_projet,
    parse_commune,
    clean_annonceur,
    parse_budget,
    parse_delai,
    extract_json_from_text,
    GAZON_EXPLICIT_KEYWORDS,
    NON_GAZON_EXCLUSIONS,
)

class TestClassifiers(unittest.TestCase):

    def test_strip_accents(self):
        self.assertEqual(strip_accents("RÉALISATION"), "REALISATION")
        self.assertEqual(strip_accents("aménagement"), "amenagement")
        self.assertEqual(strip_accents("Tébessa"), "Tebessa")

    def test_classify_action(self):
        cases = [
            ("Fourniture et pose de gazon synthétique", "FOURNITURE ET POSE"),
            ("Réalisation d'un terrain de proximité", "RÉALISATION"),
            ("Aménagement et revêtement d'un stade", "AMÉNAGEMENT ET REVÊTEMENT"),
            ("Réalisation et revêtement d'un terrain de sport", "RÉALISATION ET REVÊTEMENT"),
            ("Réhabilitation du stade communal", "RÉHABILITATION"),
            ("Travaux d'engazonnement", "TRAVAUX D'ENGAZONNEMENT"),
            ("Réfection de la pelouse", "RÉFECTION"),
            ("Travaux divers", "TRAVAUX"),
        ]
        for titre, expected in cases:
            with self.subTest(titre=titre):
                self.assertEqual(classify_action(titre), expected)

    def test_classify_type_projet(self):
        cases = [
            ("Stade de proximité en gazon synthétique", "STADE DE PROXIMITÉ"),
            ("Stade de football communal", "STADE DE FOOTBALL"),
            ("Stade communal d'Oran", "STADE COMMUNAL"),
            ("Stade municipal de Blida", "STADE MUNICIPAL"),
            ("Terrain de football en gazon", "TERRAIN DE FOOTBALL"),
            ("Terrain de sport combiné", "TERRAIN DE SPORT"),
            ("Aire de jeux pour enfants", "AIRE DE JEUX"),
            ("Revêtement de la cour d'école primaire", "COUR ECOLE PRIMAIRE"),
            ("Gazon synthétique pour complexe", "GAZON SYNTHÉTIQUE"),
        ]
        for titre, expected in cases:
            with self.subTest(titre=titre):
                self.assertEqual(classify_type_projet(titre), expected)

    def test_clean_annonceur(self):
        self.assertEqual(clean_annonceur("Commune de Morset"), "COMMUNE")
        self.assertEqual(clean_annonceur("APC D'ORAN"), "COMMUNE")
        self.assertEqual(clean_annonceur("DJS DE LA WILAYA DE SETIF"), "DJS DE LA WILAYA")
        self.assertEqual(clean_annonceur("Direction de la jeunesse et des sports"), "DJS DE LA WILAYA")
        self.assertEqual(clean_annonceur("Direction des equipements publics"), "DEP DE LA WILAYA")
        self.assertEqual(clean_annonceur("Direction de l'administration locale"), "DIRECTION DE L'ADMINISTRATION LOCALE DE LA WILAYA")

    def test_parse_budget(self):
        self.assertEqual(parse_budget("19 736 328,50 DA"), 19736328.5)
        self.assertEqual(parse_budget("31.558.205,00 DA"), 31558205.0)
        self.assertEqual(parse_budget("25892020 DA"), 25892020)
        self.assertEqual(parse_budget("/"), "/")

    def test_parse_delai(self):
        self.assertEqual(parse_delai(45, ""), "45 JOURS")
        self.assertEqual(parse_delai(0, "Délai d'exécution fixé à 03 MOIS"), "03 MOIS")
        self.assertEqual(parse_delai(0, "durée de réalisation: 60 JOURS"), "60 JOURS")
        self.assertEqual(parse_delai(0, "aucune info"), "/")

    def test_extract_json_from_text(self):
        # Plain JSON
        j1 = '{"is_gazon": true, "budget": "1000 DA"}'
        self.assertEqual(extract_json_from_text(j1).get("is_gazon"), True)

        # Markdown code fence JSON
        j2 = 'Here is the result:\n```json\n{"is_gazon": false, "commune": "ORAN"}\n```'
        self.assertEqual(extract_json_from_text(j2).get("commune"), "ORAN")

        # Freeform text containing JSON object
        j3 = 'Explanation before {"is_gazon": true} explanation after'
        self.assertEqual(extract_json_from_text(j3).get("is_gazon"), True)

    def test_gazon_scope_regexes(self):
        # Must match explicit gazon
        self.assertTrue(GAZON_EXPLICIT_KEYWORDS.search("Revêtement en gazon synthétique du terrain"))
        self.assertTrue(GAZON_EXPLICIT_KEYWORDS.search("Pose de pelouse synthétique"))
        self.assertTrue(GAZON_EXPLICIT_KEYWORDS.search("أشغال تكسية بالعشب الاصطناعي للملعب"))
        self.assertTrue(GAZON_EXPLICIT_KEYWORDS.search("Travaux d'engazonnement"))

        # Non-gazon exclusions must catch civil/irrelevant works
        self.assertTrue(NON_GAZON_EXCLUSIONS.search("Glissement de terrain sur route communale"))
        self.assertTrue(NON_GAZON_EXCLUSIONS.search("Éclairage public et voirie"))
        self.assertTrue(NON_GAZON_EXCLUSIONS.search("Fourniture de bureau et matériel informatique"))
        self.assertTrue(NON_GAZON_EXCLUSIONS.search("Revêtement en enrobé bitumineux"))

if __name__ == "__main__":
    unittest.main()
