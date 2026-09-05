import re, unicodedata

def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def classify_action(titre: str) -> str:
    """Classify the action for Col 3."""
    t_norm = strip_accents(titre).upper()
    
    # Priority ordered patterns
    patterns = [
        ("AMENAGEMENT ET REVETEMENT", r"\bAMENAGEMENT\b.*\bREVETEMENT\b|\bREVETEMENT\b.*\bAMENAGEMENT\b"),
        ("REALISATION ET REVETEMENT", r"\bREALISATION\b.*\bREVETEMENT\b|\bREVETEMENT\b.*\bREALISATION\b"),
        ("REALISATION ET REHABILITATION", r"\bREALISATION\b.*\bREHABILITATION\b"),
        ("REALISATION ET AMENAGEMENT", r"\bREALISATION\b.*\bAMENAGEMENT\b"),
        ("AMENAGEMENT ET REHABILITATION", r"\bAMENAGEMENT\b.*\bREHABILITATION\b|\bREHABILITATION\b.*\bAMENAGEMENT\b"),
        ("TRAVAUX D'ENGAZONNEMENT", r"\bTRAVAUX D'?ENGAZONNEMENT\b|\bENGAZONNEMENT\b"),
        ("FOURNITURE ET POSE", r"\bFOURNITURE ET POSE\b|\bFOURNITURE\b.*\bPOSE\b"),
        ("FOURNITURE ET INSTALLATION", r"\bFOURNITURE ET INSTALLATION\b"),
        ("RÉALISATION", r"\bREALISATION\b"),
        ("AMÉNAGEMENT", r"\bAMENAGEMENT\b"),
        ("REVÊTEMENT", r"\bREVETEMENT\b"),
        ("RÉHABILITATION", r"\bREHABILITATION\b|\bRH[EÉ]ABILITATION\b"),
        ("FOURNITURE", r"\bFOURNITURE\b"),
        ("RÉNOVATION", r"\bRENOVATION\b"),
        ("ACHEVEMENT", r"\bACHEVEMENT\b"),
        ("CONSTRUCTION", r"\bCONSTRUCTION\b"),
        ("REFECTION", r"\bREFECTION\b"),
        ("COUVERTURE", r"\bCOUVERTURE\b"),
        ("ETUDE ET SUIVI", r"\bETUDE ET SUIVI\b|\bSUIVI\b"),
        ("TRAVAUX", r"\bTRAVAUX\b"),
    ]
    for action_label, pattern in patterns:
        if re.search(pattern, t_norm):
            return action_label
    return "TRAVAUX"

def classify_type_projet(titre: str) -> str:
    """Classify the facility / project type for Col 5."""
    t_norm = strip_accents(titre).upper()

    # Priority ordered target types
    facility_patterns = [
        ("STADE DE PROXIMITÉ", r"\bSTADES?\s+DE\s+PROXIMIT[EÉ]\b"),
        ("STADE DE FOOTBALL", r"\bSTADES?\s+DE\s+FOOTBALL\b|\bSTADES?\s+DE\s+FOOT\b"),
        ("STADE COMMUNAL", r"\bSTADES?\s+COMMUNAU?X?\b|\bSTADES?\s+COMMUNALES?\b"),
        ("STADE MUNICIPAL", r"\bSTADES?\s+MUNICIPAU?X?\b|\bSTADES?\s+MUNICIPALES?\b"),
        ("STADE MATICO", r"\bSTADES?\s+MATICO\b"),
        ("STADE", r"\bSTADES?\b"),
        ("TERRAIN DE FOOTBALL", r"\bTERRAINS?\s+DE\s+FOOTBALL\b|\bTERRAINS?\s+DE\s+FOOT\b"),
        ("TERRAIN DE PROXIMITÉ", r"\bTERRAINS?\s+DE\s+PROXIMIT[EÉ]\b"),
        ("TERRAIN DE SPORT", r"\bTERRAINS?\s+DE\s+SPORTS?\b|\bTERRAINS?\s+SPORTIFS?\b"),
        ("TERRAIN DE JEU", r"\bTERRAINS?\s+DE\s+JEUX?\b"),
        ("TERRAIN GAZONNÉ", r"\bTERRAINS?\s+GAZONN[EÉ]S?\b"),
        ("TERRAIN", r"\bTERRAINS?\b"),
        ("AIRE DE JEUX", r"\bAIRES?\s+DE\s+JEUX?\b|\bESPACES?\s+DE\s+JEUX?\b"),
        ("MATICO", r"\bMATICO\b|\bMATIQUO\b"),
        ("COUR ECOLE PRIMAIRE", r"\bCOURS?\s+D'?ECOLES?\s+PRIMAIRES?\b|\bCOURS?\s+ECOLES?\b"),
        ("COUR", r"\bCOURS?\b"),
        ("PISTE D'ATHLETISME", r"\bPISTES?\s+D'?ATHLETISME\b|\bATHLETISME\b"),
        ("SALLE DE SPORT", r"\bSALLES?\s+MULTI[\s\-]SPORTS?\b|\bSALLES?\s+DE\s+SPORTS?\b|\bCOMPLEXES?\s+SPORTIFS?\b"),
        ("GAZON SYNTHÉTIQUE", r"\bGAZONS?\s+SYNTH[EÉ]TIQUES?\b|\bPELOUSES?\s+SYNTH[EÉ]TIQUES?\b|\bGAZONS?\b"),
    ]
    for facility_label, pattern in facility_patterns:
        if re.search(pattern, t_norm):
            return facility_label
    return "INFRASTRUCTURE SPORTIVE"

# Test cases from Algerian tenders
test_cases = [
    "Fourniture et pose de gazon synthétique pour stade communal",
    "Réalisation d'un terrain de proximité en gazon synthétique",
    "Aménagement et revêtement d'un stade de football",
    "Travaux d'engazonnement du terrain de sport municipal",
    "Réhabilitation et couverture d'une salle multi-sports",
    "Aménagement d'une aire de jeux pour enfants",
    "Revêtement en gazon synthétique de la cour de l'école primaire",
    "Réalisation d'un stade matico",
    "Achèvement des travaux de la piste d'athlétisme",
]

for tc in test_cases:
    print(f"Title: {tc}")
    print(f"  -> Action (Col 3): {classify_action(tc)}")
    print(f"  -> Type   (Col 5): {classify_type_projet(tc)}\n")
