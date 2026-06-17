from __future__ import annotations

import json
import math
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from typing import Any, Dict, List, Tuple

import pandas as pd
import streamlit as st

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt, RGBColor
except Exception:  # pragma: no cover - l'export DOCX reste optionnel si la dépendance manque
    Document = None
    OxmlElement = None
    qn = None

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except Exception:  # pragma: no cover - l'export Excel reste optionnel si la dépendance manque
    Workbook = None

st.set_page_config(
    page_title="Système expert CGP – Holding familiale",
    page_icon="🏛️",
    layout="wide",
)

UNKNOWN = "Non renseigné"
NA = "Non applicable"
YES = "Oui"
NO = "Non"
UNCERTAIN = "Incertain"

# =============================================================================
# 1. Données de référence : risques, conséquences, outils, actions
# =============================================================================

@dataclass
class RiskDefinition:
    code: str
    objectif: str
    libelle: str
    consequences: List[str]
    outils: List[str]
    actions_preventives: List[str]
    professionnels: List[str]
    gravite_base: int


RISK_DEFINITIONS: Dict[str, RiskDefinition] = {
    "dilution": RiskDefinition(
        "dilution", "Pérennité / contrôle", "Dilution progressive du capital",
        ["Dispersion des titres entre plusieurs branches familiales", "Affaiblissement du pouvoir de décision familial", "Risque de perte de cohérence stratégique"],
        ["Holding familiale", "Clause d’agrément", "Clause de préemption", "Actions de préférence", "Pacte d’associés"],
        ["Cartographier les futurs associés", "Encadrer les entrées et sorties du capital", "Anticiper les transmissions successives"],
        ["CGP", "Avocat", "Notaire"], 4,
    ),
    "tiers": RiskDefinition(
        "tiers", "Pérennité / contrôle", "Entrée de tiers au capital",
        ["Perte d’indépendance familiale", "Conflit stratégique avec un associé extérieur", "Remise en cause du projet familial"],
        ["Clause d’agrément", "Clause de préemption", "Pacte d’associés", "Statuts adaptés"],
        ["Formaliser les règles de cession", "Prévoir un droit de rachat prioritaire", "Vérifier la cohérence statuts / pacte"],
        ["CGP", "Avocat"], 4,
    ),
    "blocage_gouvernance": RiskDefinition(
        "blocage_gouvernance", "Pérennité / contrôle", "Blocage de gouvernance",
        ["Paralysie décisionnelle", "Conflits entre associés actifs et passifs", "Perte de valeur ou ralentissement de l’entreprise"],
        ["Pacte d’associés", "Charte familiale", "Conseil de famille", "Comité stratégique", "Statuts de SAS"],
        ["Clarifier les rôles", "Définir les règles de majorité", "Organiser un dialogue familial"],
        ["CGP", "Avocat", "Conseil en gouvernance familiale"], 4,
    ),
    "successeur": RiskDefinition(
        "successeur", "Pérennité / contrôle", "Absence ou mauvaise préparation du successeur",
        ["Rupture de continuité managériale", "Perte de confiance des salariés et partenaires", "Cession subie ou direction externe non anticipée"],
        ["Gouvernance transitoire", "Direction externe temporaire", "Actions de préférence", "Pacte d’associés", "Calendrier de transmission"],
        ["Identifier les intentions des héritiers", "Évaluer les compétences du successeur", "Prévoir un calendrier progressif"],
        ["CGP", "Expert-comptable", "Conseil en gouvernance", "Avocat"], 5,
    ),
    "conflit_heritiers": RiskDefinition(
        "conflit_heritiers", "Équité entre héritiers", "Conflit repreneur / non repreneurs",
        ["Sentiment d’injustice", "Blocage de la holding", "Tensions durables entre branches familiales"],
        ["Donation-partage", "Soulte", "Pacte d’associés", "Mécanisme de sortie", "Droits financiers différenciés"],
        ["Distinguer égalité et équité", "Organiser une valorisation indépendante", "Anticiper la liquidité des non repreneurs", "Expliquer les choix avant les actes"],
        ["CGP", "Notaire", "Avocat", "Expert-comptable"], 4,
    ),
    "contestation": RiskDefinition(
        "contestation", "Équité entre héritiers", "Contestation successorale",
        ["Remise en cause des opérations de transmission", "Contentieux familial", "Instabilité juridique de la holding"],
        ["Donation-partage", "Audit civil", "Valorisation indépendante", "Démembrement", "Testament"],
        ["Vérifier la réserve héréditaire", "Documenter la valorisation", "Faire intervenir le notaire dès le diagnostic"],
        ["CGP", "Notaire", "Avocat", "Expert-comptable"], 4,
    ),
    "liquidite": RiskDefinition(
        "liquidite", "Équité entre héritiers", "Soulte trop lourde ou manque de liquidité",
        ["Difficulté de paiement des soultes ou droits", "Pression excessive sur les dividendes", "Cession forcée d’actifs"],
        ["Soulte échelonnée", "Family Buy-Out", "Paiement différé/fractionné", "Assurance-vie", "Politique de distribution encadrée"],
        ["Simuler les besoins de liquidité", "Mesurer la capacité de financement", "Prévoir les conditions de sortie des héritiers non actifs"],
        ["CGP", "Expert-comptable", "Notaire", "Avocat"], 4,
    ),
    "fiscalite": RiskDefinition(
        "fiscalite", "Optimisation fiscale", "Charge fiscale excessive",
        ["Droits de mutation élevés", "Besoin de liquidité immédiate", "Risque de cession partielle pour financer la fiscalité"],
        ["Pacte Dutreil", "Donation anticipée", "Démembrement", "Paiement différé et fractionné"],
        ["Réaliser une simulation fiscale", "Anticiper la transmission", "Évaluer la fiscalité résiduelle"],
        ["CGP", "Notaire", "Avocat fiscaliste", "Expert-comptable"], 4,
    ),
    "dutreil": RiskDefinition(
        "dutreil", "Optimisation fiscale", "Remise en cause du Pacte Dutreil",
        ["Perte de l’exonération partielle", "Imposition complémentaire imprévue", "Fragilisation financière des héritiers"],
        ["Audit Dutreil", "Documentation holding animatrice", "Suivi des engagements", "Revue juridique annuelle"],
        ["Vérifier l’éligibilité", "Documenter la qualité de holding animatrice", "Suivre les engagements de conservation"],
        ["CGP", "Avocat fiscaliste", "Notaire", "Expert-comptable"], 5,
    ),
    "conjoint": RiskDefinition(
        "conjoint", "Protection familiale", "Fragilisation du conjoint et des proches",
        ["Insuffisance de revenus du conjoint survivant", "Dépendance aux décisions des enfants", "Tensions conjoint / héritiers"],
        ["Régime matrimonial", "Donation entre époux", "Assurance-vie", "Démembrement", "Prévoyance"],
        ["Analyser les droits du conjoint", "Simuler un décès ou une incapacité", "Organiser des revenus ou capitaux disponibles"],
        ["CGP", "Notaire", "Assureur", "Avocat"], 4,
    ),
    "dependance": RiskDefinition(
        "dependance", "Protection familiale", "Dépendance excessive au patrimoine professionnel",
        ["Exposition aux aléas économiques de l’entreprise", "Manque de liquidité familiale", "Difficulté à protéger conjoint et héritiers"],
        ["Diversification patrimoniale", "Holding patrimoniale", "Assurance-vie", "Prévoyance", "Investissements complémentaires"],
        ["Mesurer le poids de l’entreprise", "Constituer des actifs liquides", "Prévoir des scénarios de crise"],
        ["CGP", "Expert-comptable", "Notaire"], 4,
    ),
    "suivi": RiskDefinition(
        "suivi", "Suivi", "Inadaptation progressive de la stratégie",
        ["Outils devenus inadaptés", "Non-respect d’engagements fiscaux ou juridiques", "Perte de cohérence dans le temps"],
        ["Rendez-vous de suivi", "Revue des statuts", "Actualisation du pacte", "Revue patrimoniale annuelle"],
        ["Planifier une revue annuelle", "Réévaluer les objectifs familiaux", "Contrôler les engagements"],
        ["CGP", "Notaire", "Avocat", "Expert-comptable"], 3,
    ),
}


TOOL_JUSTIFICATIONS_BY_RISK: Dict[str, Dict[str, str]] = {
    "dilution": {
        "Holding familiale": "Elle centralise la détention des titres et limite la dispersion directe du capital entre plusieurs branches familiales.",
        "Clause d’agrément": "Elle permet de contrôler l’entrée de nouveaux associés et d’éviter qu’un tiers non souhaité intègre le capital.",
        "Clause de préemption": "Elle donne une priorité de rachat aux associés familiaux avant toute cession à un tiers.",
        "Actions de préférence": "Elles peuvent renforcer certains droits politiques afin de préserver le contrôle familial malgré l’évolution de l’actionnariat.",
        "Pacte d’associés": "Il formalise les règles d’entrée, de sortie et de conservation des titres entre membres de la famille.",
    },
    "tiers": {
        "Clause d’agrément": "Elle soumet l’entrée d’un nouvel associé à l’accord préalable des associés ou de l’organe compétent.",
        "Clause de préemption": "Elle permet aux associés familiaux de se porter acquéreurs en priorité avant une cession extérieure.",
        "Pacte d’associés": "Il encadre contractuellement les cessions et réduit le risque d’entrée non anticipée d’un tiers.",
        "Statuts adaptés": "Ils permettent d’intégrer directement les règles de contrôle des cessions dans l’organisation juridique de la société.",
    },
    "blocage_gouvernance": {
        "Pacte d’associés": "Il permet d’organiser les règles de décision, d’information, de sortie et de résolution des désaccords.",
        "Charte familiale": "Elle clarifie les valeurs, les rôles et les principes de fonctionnement de la famille autour de l’entreprise.",
        "Conseil de famille": "Il crée un espace de dialogue distinct de la gestion opérationnelle et limite la transformation des désaccords en blocages.",
        "Comité stratégique": "Il permet d’associer certaines parties prenantes aux grandes orientations sans confondre gouvernance familiale et direction opérationnelle.",
        "Statuts de SAS": "La souplesse statutaire de la SAS permet d’adapter les organes de décision aux besoins de la famille.",
    },
    "successeur": {
        "Gouvernance transitoire": "Elle accompagne le passage progressif du pouvoir entre le dirigeant et le repreneur.",
        "Direction externe temporaire": "Elle constitue une solution de continuité lorsque le successeur familial n’est pas encore prêt ou identifié.",
        "Actions de préférence": "Elles peuvent attribuer des droits adaptés au repreneur ou organiser une montée en puissance progressive.",
        "Pacte d’associés": "Il encadre la période de transition, les droits du repreneur et les équilibres avec les autres héritiers.",
        "Calendrier de transmission": "Il rend la transmission opérationnelle en fixant les étapes de transfert du pouvoir et des responsabilités.",
    },
    "conflit_heritiers": {
        "Donation-partage": "Elle organise la répartition du patrimoine du vivant du dirigeant et réduit le risque de contestation entre héritiers.",
        "Soulte": "Elle permet de compenser financièrement les héritiers non repreneurs lorsque l’entreprise est attribuée principalement au repreneur.",
        "Pacte d’associés": "Il fixe les droits et obligations des héritiers associés, notamment les règles de sortie, d’information et de distribution.",
        "Mécanisme de sortie": "Il évite que les héritiers non repreneurs se sentent enfermés dans une structure peu liquide.",
        "Droits financiers différenciés": "Ils peuvent concilier pouvoir renforcé du repreneur et protection économique des héritiers non actifs.",
    },
    "contestation": {
        "Donation-partage": "Elle permet d’anticiper le partage et de stabiliser les équilibres successoraux avant l’ouverture de la succession.",
        "Audit civil": "Il vérifie la cohérence de la stratégie avec la réserve héréditaire, les droits du conjoint et la situation familiale.",
        "Valorisation indépendante": "Elle réduit le risque de contestation sur la valeur des titres transmis.",
        "Démembrement": "Il permet d’organiser une transmission progressive tout en maintenant certains droits au profit du dirigeant ou du conjoint.",
        "Testament": "Il complète l’organisation successorale lorsque certaines volontés doivent être formalisées hors donation.",
    },
    "liquidite": {
        "Soulte échelonnée": "Elle évite une charge immédiate excessive pour le repreneur ou la holding.",
        "Family Buy-Out": "Il structure le financement de la reprise familiale et peut faciliter la compensation des héritiers non repreneurs.",
        "Paiement différé/fractionné": "Il étale le paiement de certaines charges fiscales et réduit la pression immédiate sur la trésorerie.",
        "Assurance-vie": "Elle peut fournir des capitaux disponibles pour compenser ou sécuriser certains héritiers.",
        "Politique de distribution encadrée": "Elle donne de la visibilité aux associés familiaux sans imposer une distribution désorganisée.",
    },
    "fiscalite": {
        "Pacte Dutreil": "Il peut réduire significativement l’assiette taxable de la transmission si les conditions sont respectées.",
        "Donation anticipée": "Elle permet d’organiser la transmission dans le temps et de mieux maîtriser la charge fiscale.",
        "Démembrement": "Il peut réduire la base taxable de la nue-propriété tout en maintenant certains droits économiques.",
        "Paiement différé et fractionné": "Il n’allège pas la fiscalité, mais en facilite le financement dans le temps.",
    },
    "dutreil": {
        "Audit Dutreil": "Il permet de vérifier l’éligibilité au régime avant la mise en œuvre et de réduire le risque de remise en cause.",
        "Documentation holding animatrice": "Elle sécurise la qualification de la holding lorsque le bénéfice du régime dépend de son rôle d’animation.",
        "Suivi des engagements": "Il garantit que les conditions de conservation et les obligations associées restent respectées dans le temps.",
        "Revue juridique annuelle": "Elle permet d’identifier les opérations ou évolutions susceptibles d’affecter la sécurité du dispositif.",
    },
    "conjoint": {
        "Régime matrimonial": "Son analyse permet d’identifier les droits du conjoint et d’adapter la protection du couple avant la transmission.",
        "Donation entre époux": "Elle peut augmenter les droits du conjoint survivant et offrir davantage de souplesse successorale.",
        "Assurance-vie": "Elle peut transmettre des capitaux disponibles au conjoint en dehors du règlement successoral ordinaire.",
        "Démembrement": "Il peut préserver des revenus ou une jouissance au profit du conjoint tout en préparant la transmission aux enfants.",
        "Prévoyance": "Elle apporte une liquidité rapide en cas de décès ou d’incapacité du dirigeant.",
    },
    "dependance": {
        "Diversification patrimoniale": "Elle réduit la dépendance de la famille à la seule valeur de l’entreprise.",
        "Holding patrimoniale": "Elle peut structurer progressivement des actifs distincts de l’entreprise opérationnelle.",
        "Assurance-vie": "Elle constitue un support de liquidité et de transmission complémentaire au patrimoine professionnel.",
        "Prévoyance": "Elle protège la famille contre un événement brutal affectant le dirigeant ou l’entreprise.",
        "Investissements complémentaires": "Ils permettent de constituer des sources de revenus et de liquidité en dehors de l’entreprise familiale.",
    },
    "suivi": {
        "Rendez-vous de suivi": "Ils permettent de vérifier régulièrement que la stratégie reste adaptée à la famille et à l’entreprise.",
        "Revue des statuts": "Elle évite que les règles de gouvernance deviennent inadaptées à l’évolution de l’actionnariat familial.",
        "Actualisation du pacte": "Elle permet d’ajuster les règles entre associés lorsque la situation familiale ou patrimoniale évolue.",
        "Revue patrimoniale annuelle": "Elle contrôle la cohérence globale entre objectifs, risques, outils et besoins de liquidité.",
    },
}

DEFAULT_ANSWERS: Dict[str, Any] = {
    "client_name": "",
    "client_age": 0,
    "company_name": "",
    "company_activity": "",
    "company_form": UNKNOWN,
    "cgp_name": "",
    "entretien_date": "",
    "rapport_orientation": "Équilibré",
    "niveau_detail": "Détaillé",
    "objectif_libre": "",
    "attentes_client": "",
    "contraintes_client": "",
    "personnes_a_associer": "",
    "observations": "",
    "maturite_projet": UNKNOWN,
    "delai_transmission": UNKNOWN,
    "urgence_evenement": UNKNOWN,
    "qualite_information": UNKNOWN,
    "objectifs": [],
    "objective_weights": {},
    "nb_enfants": 0,
    "conjoint_present": UNKNOWN,
    "famille_recomposee": UNKNOWN,
    "dialogue_familial": UNKNOWN,
    "accord_conjoint": UNKNOWN,
    "heritier_repreneur": UNKNOWN,
    "autres_heritiers_actifs": UNKNOWN,
    "volonte_non_repreneurs": UNKNOWN,
    "soulte_envisagee": UNKNOWN,
    "capacite_financement_soulte": UNKNOWN,
    "valorisation_independante": UNKNOWN,
    "audit_civil": UNKNOWN,
    "conjoint_dependant": UNKNOWN,
    "protection_conjoint_prevue": UNKNOWN,
    "regime_matrimonial_adapte": UNKNOWN,
    "valeur_entreprise": 0,
    "poids_entreprise": 0,
    "actifs_liquides": UNKNOWN,
    "endettement_familial": UNKNOWN,
    "besoin_revenus_famille": UNKNOWN,
    "diversification": UNKNOWN,
    "prevoyance": UNKNOWN,
    "gouvernance_formalisee": UNKNOWN,
    "associes_actifs_passifs": UNKNOWN,
    "clauses_entree_sortie": UNKNOWN,
    "politique_dividendes_definie": UNKNOWN,
    "entreprise_dependante_dirigeant": UNKNOWN,
    "calendrier_transmission": UNKNOWN,
    "successeur_prepare": UNKNOWN,
    "pacte_dutreil": UNKNOWN,
    "holding_animatrice": UNKNOWN,
    "simulation_fiscale": UNKNOWN,
    "audit_dutreil": UNKNOWN,
    "suivi_engagements": UNKNOWN,
    "suivi_annuel": UNKNOWN,
    "formalisation_rapport": UNKNOWN,
}

STEP_KEYS: Dict[int, List[str]] = {
    1: ["client_name", "client_age", "company_name", "company_activity", "company_form", "cgp_name", "entretien_date", "rapport_orientation", "niveau_detail", "objectif_libre", "attentes_client", "contraintes_client", "personnes_a_associer", "observations", "maturite_projet", "delai_transmission", "urgence_evenement", "qualite_information", "objectifs", "objective_weights"],
    2: ["nb_enfants", "conjoint_present", "famille_recomposee", "dialogue_familial", "accord_conjoint", "heritier_repreneur", "autres_heritiers_actifs", "volonte_non_repreneurs", "soulte_envisagee", "capacite_financement_soulte", "conjoint_dependant", "protection_conjoint_prevue", "regime_matrimonial_adapte", "valorisation_independante", "audit_civil"],
    3: ["valeur_entreprise", "poids_entreprise", "actifs_liquides", "endettement_familial", "besoin_revenus_famille", "diversification", "prevoyance"],
    4: ["clauses_entree_sortie", "gouvernance_formalisee", "associes_actifs_passifs", "politique_dividendes_definie", "entreprise_dependante_dirigeant", "successeur_prepare", "calendrier_transmission"],
    5: ["simulation_fiscale", "pacte_dutreil", "holding_animatrice", "audit_dutreil", "suivi_engagements", "suivi_annuel", "formalisation_rapport"],
    6: [],
}

PRIORITY_ORDER = {"Critique": 4, "Élevé": 3, "Moyen": 2, "Faible": 1, "Inexistant": 0}
PRIORITY_COLORS = {"Critique": "#7f1d1d", "Élevé": "#b45309", "Moyen": "#1d4ed8", "Faible": "#047857", "Inexistant": "#6b7280"}

OBJECTIVE_WEIGHT_OPTIONS = ["Important", "Très important", "Prioritaire"]
OBJECTIVE_WEIGHT_FACTORS = {"Important": 1.0, "Très important": 1.2, "Prioritaire": 1.4}
OBJECTIVE_DISPLAY_ORDER = [
    "Conserver le contrôle familial",
    "Transmettre l’entreprise",
    "Optimiser la fiscalité",
    "Protéger le conjoint et les proches",
    "Préserver l’équité entre les héritiers",
    "Diversifier le patrimoine",
]
RISK_TO_OBJECTIVE = {
    "dilution": "Conserver le contrôle familial",
    "tiers": "Conserver le contrôle familial",
    "blocage_gouvernance": "Conserver le contrôle familial",
    "successeur": "Transmettre l’entreprise",
    "conflit_heritiers": "Préserver l’équité entre les héritiers",
    "contestation": "Préserver l’équité entre les héritiers",
    "liquidite": "Préserver l’équité entre les héritiers",
    "fiscalite": "Optimiser la fiscalité",
    "dutreil": "Optimiser la fiscalité",
    "conjoint": "Protéger le conjoint et les proches",
    "dependance": "Diversifier le patrimoine",
    "suivi": "Suivi global",
}

# =============================================================================
# 2. Initialisation, widgets, validation explicite
# =============================================================================

def safe_key(text: str) -> str:
    return "".join(ch.lower() if ch.isalnum() else "_" for ch in text).strip("_")

def get_objective_weight(answers: Dict[str, Any], objective: str) -> str:
    weights = answers.get("objective_weights") or {}
    return weights.get(objective, "Important")

def get_objective_factor(answers: Dict[str, Any], objective: str) -> float:
    return OBJECTIVE_WEIGHT_FACTORS.get(get_objective_weight(answers, objective), 1.0)

def weight_for_risk(code: str, answers: Dict[str, Any] | None) -> Tuple[str, float]:
    answers = answers or {}
    objectives = list(answers.get("objectifs") or [])
    if code == "suivi":
        if not objectives:
            return "Important", 1.0
        labels = [get_objective_weight(answers, obj) for obj in objectives]
        label = max(labels, key=lambda lab: OBJECTIVE_WEIGHT_FACTORS.get(lab, 1.0))
        return f"Suivi global ({label.lower()})", OBJECTIVE_WEIGHT_FACTORS.get(label, 1.0)
    objective = RISK_TO_OBJECTIVE.get(code, "")
    if code == "liquidite" and "Diversifier le patrimoine" in objectives:
        objective = "Diversifier le patrimoine"
    if code == "dependance" and "Diversifier le patrimoine" not in objectives and "Protéger le conjoint et les proches" in objectives:
        objective = "Protéger le conjoint et les proches"
    label = get_objective_weight(answers, objective) if objective else "Important"
    return label, OBJECTIVE_WEIGHT_FACTORS.get(label, 1.0)

def init_app() -> None:
    if "answers" not in st.session_state:
        st.session_state.answers = dict(DEFAULT_ANSWERS)
    if "draft_answers" not in st.session_state:
        # Les réponses en cours de saisie sont conservées séparément des réponses validées.
        # Cela permet de revenir en arrière dans le questionnaire sans perdre les champs déjà remplis.
        st.session_state.draft_answers = dict(st.session_state.answers)
    if "validated_steps" not in st.session_state:
        st.session_state.validated_steps = set()
    if "last_validation" not in st.session_state:
        st.session_state.last_validation = {}
    st.session_state.setdefault("current_step", 1)
    st.session_state.setdefault("app_page", "Questionnaire adaptatif")
    for key, value in DEFAULT_ANSWERS.items():
        st.session_state.setdefault(f"w_{key}", st.session_state.draft_answers.get(key, st.session_state.answers.get(key, value)))


def reset_app() -> None:
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    init_app()


def sync_draft_key(key: str) -> None:
    """Copie immédiatement la valeur d'un widget dans le brouillon."""
    widget_key = f"w_{key}"
    if widget_key in st.session_state:
        st.session_state.draft_answers[key] = st.session_state[widget_key]


def sync_widget_from_draft(key: str) -> None:
    """Initialise le widget avec la dernière valeur saisie, ou avec la valeur validée si aucun brouillon n'existe."""
    value = st.session_state.draft_answers.get(key, st.session_state.answers.get(key, DEFAULT_ANSWERS.get(key)))
    st.session_state.setdefault(f"w_{key}", value)


def set_answer_and_draft(key: str, value: Any) -> None:
    """Met à jour simultanément la réponse validée, le brouillon et le widget si possible."""
    st.session_state.answers[key] = value
    st.session_state.draft_answers[key] = value
    widget_key = f"w_{key}"
    if widget_key in st.session_state:
        st.session_state[widget_key] = value


def get_draft(key: str) -> Any:
    if f"w_{key}" in st.session_state:
        st.session_state.draft_answers[key] = st.session_state[f"w_{key}"]
    return st.session_state.draft_answers.get(key, st.session_state.answers.get(key, DEFAULT_ANSWERS.get(key)))


def selected_objectives_from_draft() -> set:
    return set(get_draft("objectifs") or [])


def save_step(step: int) -> Tuple[bool, str]:
    """Copie les réponses visibles/en cours vers answers. Les résultats utilisent uniquement answers."""
    keys = STEP_KEYS.get(step, [])
    if step == 1 and not get_draft("objectifs"):
        return False, "Sélectionne au moins un objectif avant de valider l’étape 1."

    for key in keys:
        if f"w_{key}" in st.session_state:
            st.session_state.draft_answers[key] = st.session_state[f"w_{key}"]
        st.session_state.answers[key] = st.session_state.draft_answers.get(key, st.session_state.answers.get(key, DEFAULT_ANSWERS.get(key)))

    if step == 1:
        selected = list(st.session_state.answers.get("objectifs") or [])
        weights = {}
        draft_weights = st.session_state.draft_answers.get("objective_weights", {}) or {}
        for objective in selected:
            widget_key = f"w_objective_weight_{safe_key(objective)}"
            weights[objective] = st.session_state.get(widget_key, draft_weights.get(objective, "Important"))
        st.session_state.answers["objective_weights"] = weights
        st.session_state.draft_answers["objective_weights"] = weights
        st.session_state["w_objective_weights"] = weights

    # Nettoyage léger des branches non applicables pour éviter les anciennes réponses parasites.
    a = st.session_state.answers
    objectifs = set(a.get("objectifs") or [])

    if "Transmettre l’entreprise" not in objectifs or int(a.get("nb_enfants") or 0) == 0:
        for key in ["heritier_repreneur", "autres_heritiers_actifs", "soulte_envisagee", "capacite_financement_soulte", "successeur_prepare", "calendrier_transmission"]:
            set_answer_and_draft(key, UNKNOWN)

    if a.get("heritier_repreneur") != YES:
        for key in ["autres_heritiers_actifs", "volonte_non_repreneurs", "soulte_envisagee", "capacite_financement_soulte"]:
            set_answer_and_draft(key, UNKNOWN)

    if a.get("conjoint_present") != YES:
        set_answer_and_draft("accord_conjoint", UNKNOWN)

    if a.get("soulte_envisagee") != YES:
        set_answer_and_draft("capacite_financement_soulte", UNKNOWN)

    if "Protéger le conjoint et les proches" not in objectifs or a.get("conjoint_present") != YES:
        for key in ["conjoint_dependant", "protection_conjoint_prevue", "regime_matrimonial_adapte"]:
            set_answer_and_draft(key, UNKNOWN)

    if a.get("pacte_dutreil") != YES:
        for key in ["holding_animatrice", "audit_dutreil", "suivi_engagements"]:
            set_answer_and_draft(key, UNKNOWN)

    st.session_state.validated_steps.add(step)
    st.session_state.last_validation[step] = datetime.now().strftime("%H:%M:%S")
    return True, f"Étape {step} validée. Les réponses sont enregistrées dans le diagnostic."


def yes_no_unknown(label: str, key: str, options: List[str] | None = None, horizontal: bool = True) -> Any:
    sync_widget_from_draft(key)
    opts = options or [UNKNOWN, YES, NO]
    return st.radio(label, opts, key=f"w_{key}", horizontal=horizontal, on_change=sync_draft_key, args=(key,))


def number_input(label: str, key: str, **kwargs: Any) -> Any:
    sync_widget_from_draft(key)
    return st.number_input(label, key=f"w_{key}", on_change=sync_draft_key, args=(key,), **kwargs)


def slider(label: str, key: str, min_value: int, max_value: int) -> Any:
    sync_widget_from_draft(key)
    return st.slider(label, min_value, max_value, key=f"w_{key}", on_change=sync_draft_key, args=(key,))


def objective_weight_buttons(objective: str) -> None:
    """Affiche la pondération de l'objectif sous forme de boutons, sans curseur."""
    key = f"w_objective_weight_{safe_key(objective)}"
    weights = st.session_state.draft_answers.setdefault("objective_weights", {})
    if key not in st.session_state:
        st.session_state[key] = weights.get(objective, (st.session_state.answers.get("objective_weights", {}) or {}).get(objective, "Important"))

    current = st.session_state.get(key, weights.get(objective, "Important"))
    weights[objective] = current
    st.markdown(f"**{objective}**")
    cols = st.columns(len(OBJECTIVE_WEIGHT_OPTIONS))
    for col, option in zip(cols, OBJECTIVE_WEIGHT_OPTIONS):
        with col:
            if st.button(
                option,
                key=f"btn_{key}_{safe_key(option)}",
                type="primary" if current == option else "secondary",
                use_container_width=True,
                help="Sélectionne le niveau d’importance de cet objectif.",
            ):
                st.session_state[key] = option
                st.session_state.draft_answers.setdefault("objective_weights", {})[objective] = option
                st.rerun()
    st.caption(f"Niveau retenu : {st.session_state.draft_answers.get('objective_weights', {}).get(objective, 'Important')}")


def selectbox(label: str, key: str, options: List[str]) -> Any:
    sync_widget_from_draft(key)
    current = st.session_state.get(f"w_{key}", st.session_state.draft_answers.get(key, DEFAULT_ANSWERS.get(key)))
    if current not in options:
        st.session_state[f"w_{key}"] = options[0] if options else None
        st.session_state.draft_answers[key] = st.session_state[f"w_{key}"]
    return st.selectbox(label, options, key=f"w_{key}", on_change=sync_draft_key, args=(key,))


def text_input_field(label: str, key: str, placeholder: str = "") -> Any:
    sync_widget_from_draft(key)
    return st.text_input(label, key=f"w_{key}", placeholder=placeholder, on_change=sync_draft_key, args=(key,))


def text_area_field(label: str, key: str, placeholder: str = "", height: int = 90) -> Any:
    sync_widget_from_draft(key)
    return st.text_area(label, key=f"w_{key}", placeholder=placeholder, height=height, on_change=sync_draft_key, args=(key,))


def validate_buttons(step: int) -> None:
    """Bouton unique de validation : enregistre l'étape et passe à la suivante."""
    c1, c2, _ = st.columns([1.1, 2.2, 4.7])
    with c1:
        if st.button("← Étape précédente", disabled=step == 1):
            st.session_state.current_step = max(1, step - 1)
            st.rerun()
    with c2:
        label = "Valider et afficher la synthèse" if step == 5 else "Valider et passer à l’étape suivante"
        if st.button(label, type="primary"):
            ok, msg = save_step(step)
            if ok:
                st.session_state.current_step = min(6, step + 1)
                st.rerun()
            else:
                st.error(msg)

# =============================================================================
# 3. Scoring
# =============================================================================

def is_yes(a: Dict, key: str) -> bool:
    return a.get(key) == YES


def is_no(a: Dict, key: str) -> bool:
    return a.get(key) == NO


def is_known(a: Dict, key: str) -> bool:
    return a.get(key) not in [UNKNOWN, NA, None, ""]


def add_points(points: Dict[str, int], evidence: Dict[str, List[str]], risk: str, value: int, reason: str) -> None:
    if value <= 0:
        return
    points[risk] += value
    evidence.setdefault(risk, []).append(f"+{value} — {reason}")


def clamp(value: int, low: int = 0, high: int = 5) -> int:
    return max(low, min(high, value))


def niveau(score: int) -> str:
    if score == 0:
        return "Inexistant"
    if score >= 16:
        return "Critique"
    if score >= 10:
        return "Élevé"
    if score >= 5:
        return "Moyen"
    return "Faible"


def score_to_label(value: int) -> str:
    labels = {0: "Aucun signal", 1: "Faible", 2: "Moyen", 3: "Marqué", 4: "Fort", 5: "Très fort"}
    return labels.get(value, str(value))


def calculate_risks(a: Dict) -> Tuple[pd.DataFrame, Dict[str, List[str]]]:
    points = {code: 0 for code in RISK_DEFINITIONS}
    evidence: Dict[str, List[str]] = {code: [] for code in RISK_DEFINITIONS}

    objectifs = set(a.get("objectifs") or [])
    nb_enfants = int(a.get("nb_enfants") or 0)
    poids = int(a.get("poids_entreprise") or 0)
    valeur = int(a.get("valeur_entreprise") or 0)

    if not objectifs:
        return build_df(points, evidence, a)

    objectif_controle = "Conserver le contrôle familial" in objectifs
    objectif_transmission = "Transmettre l’entreprise" in objectifs
    objectif_equite = "Préserver l’équité entre les héritiers" in objectifs
    objectif_fiscal = "Optimiser la fiscalité" in objectifs
    objectif_protection = "Protéger le conjoint et les proches" in objectifs
    objectif_diversification = "Diversifier le patrimoine" in objectifs

    # Signaux transversaux liés à la maturité et à la qualité du dossier.
    if a.get("maturite_projet") == "Transmission urgente":
        add_points(points, evidence, "successeur", 2, "Projet présenté comme urgent")
        add_points(points, evidence, "fiscalite", 1, "Urgence pouvant réduire les marges de planification fiscale")
        add_points(points, evidence, "contestation", 1, "Urgence pouvant limiter la pédagogie familiale")
        add_points(points, evidence, "suivi", 1, "Projet urgent nécessitant un suivi renforcé")
    elif a.get("maturite_projet") == "Transmission prévue":
        add_points(points, evidence, "suivi", 1, "Transmission déjà prévue : nécessité de coordination et de suivi")

    if a.get("delai_transmission") == "Moins de 12 mois":
        add_points(points, evidence, "successeur", 2, "Horizon de transmission inférieur à 12 mois")
        add_points(points, evidence, "fiscalite", 1, "Horizon court pour organiser l’optimisation fiscale")
        add_points(points, evidence, "liquidite", 1, "Horizon court pour organiser les besoins de liquidité")
    elif a.get("delai_transmission") == "1 à 3 ans":
        add_points(points, evidence, "suivi", 1, "Transmission envisagée à moyen terme : suivi de trajectoire nécessaire")

    if is_yes(a, "urgence_evenement"):
        add_points(points, evidence, "successeur", 2, "Événement d’urgence ou de fragilité signalé")
        add_points(points, evidence, "conjoint", 2, "Événement de fragilité pouvant affecter les proches")
        add_points(points, evidence, "liquidite", 1, "Événement de fragilité pouvant créer un besoin de liquidité rapide")
        add_points(points, evidence, "suivi", 1, "Événement de fragilité imposant une coordination rapide")

    if a.get("qualite_information") == "À vérifier":
        add_points(points, evidence, "suivi", 2, "Qualité de l’information à vérifier avant recommandation")
    elif a.get("qualite_information") == "Partiellement confirmées":
        add_points(points, evidence, "suivi", 1, "Informations partiellement confirmées")

    if objectif_controle:
        add_points(points, evidence, "dilution", 1, "Objectif déclaré de conservation du contrôle familial")
        add_points(points, evidence, "tiers", 1, "Objectif déclaré de protection du capital familial")
        if nb_enfants >= 2:
            add_points(points, evidence, "dilution", 2, "Présence de plusieurs héritiers")
        if nb_enfants >= 3:
            add_points(points, evidence, "dilution", 1, "Nombre élevé d’héritiers")
        if is_no(a, "clauses_entree_sortie"):
            add_points(points, evidence, "dilution", 2, "Absence de clauses d’entrée/sortie alors que le contrôle est recherché")
            add_points(points, evidence, "tiers", 3, "Absence de clause d’agrément ou de préemption identifiée")

    if is_yes(a, "associes_actifs_passifs"):
        add_points(points, evidence, "blocage_gouvernance", 3, "Coexistence d’associés actifs et passifs")
        if is_no(a, "politique_dividendes_definie"):
            add_points(points, evidence, "blocage_gouvernance", 2, "Politique de dividendes non définie malgré la présence d’associés actifs et passifs")
            add_points(points, evidence, "liquidite", 1, "Absence de politique de distribution pour les associés non actifs")
    if is_no(a, "gouvernance_formalisee") and (objectif_controle or nb_enfants >= 2):
        add_points(points, evidence, "blocage_gouvernance", 3, "Gouvernance insuffisamment formalisée")
    if is_no(a, "dialogue_familial") and nb_enfants >= 2:
        add_points(points, evidence, "blocage_gouvernance", 1, "Dialogue familial non organisé")
        add_points(points, evidence, "conflit_heritiers", 1, "Dialogue familial non organisé")

    if is_yes(a, "entreprise_dependante_dirigeant"):
        add_points(points, evidence, "successeur", 3, "Entreprise fortement dépendante du dirigeant actuel")
        add_points(points, evidence, "dependance", 1, "Dépendance opérationnelle au dirigeant")

    if objectif_transmission:
        if a.get("heritier_repreneur") == NO:
            add_points(points, evidence, "successeur", 4, "Aucun héritier repreneur identifié malgré l’objectif de transmission")
        elif a.get("heritier_repreneur") == UNCERTAIN:
            add_points(points, evidence, "successeur", 3, "Successeur familial incertain")
        elif a.get("heritier_repreneur") == YES:
            if is_no(a, "successeur_prepare"):
                add_points(points, evidence, "successeur", 3, "Successeur identifié mais préparation insuffisante")
            if is_no(a, "calendrier_transmission"):
                add_points(points, evidence, "successeur", 2, "Absence de calendrier de transmission du pouvoir")

    if objectif_equite or objectif_transmission:
        if a.get("heritier_repreneur") == YES and nb_enfants >= 2:
            add_points(points, evidence, "conflit_heritiers", 3, "Un enfant reprend alors que plusieurs héritiers existent")
            add_points(points, evidence, "liquidite", 1, "Compensation éventuelle des autres héritiers à organiser")
            if is_no(a, "autres_heritiers_actifs"):
                add_points(points, evidence, "conflit_heritiers", 3, "Les autres héritiers ne sont pas impliqués dans l’entreprise")
            if a.get("volonte_non_repreneurs") in ["Sortir du capital", "Recevoir principalement une compensation"]:
                add_points(points, evidence, "liquidite", 2, "Les héritiers non repreneurs souhaitent une sortie ou une compensation")
                add_points(points, evidence, "conflit_heritiers", 1, "Attente de sortie ou de compensation des non repreneurs")
            elif a.get("volonte_non_repreneurs") == "Incertain / non abordé":
                add_points(points, evidence, "conflit_heritiers", 2, "Volonté des héritiers non repreneurs non clarifiée")
                add_points(points, evidence, "liquidite", 1, "Sortie ou compensation des héritiers non repreneurs non clarifiée")
        if is_yes(a, "soulte_envisagee"):
            add_points(points, evidence, "liquidite", 1, "Soulte envisagée : besoin de financement à vérifier")
            if is_no(a, "capacite_financement_soulte"):
                add_points(points, evidence, "liquidite", 4, "Soulte envisagée mais capacité de financement non validée")
                add_points(points, evidence, "conflit_heritiers", 1, "Soulte potentiellement difficile à financer")
        if is_no(a, "valorisation_independante") and (valeur >= 500_000 or nb_enfants >= 2):
            add_points(points, evidence, "contestation", 3, "Absence de valorisation indépendante")
            add_points(points, evidence, "conflit_heritiers", 1, "Valorisation susceptible d’être discutée")
        if is_yes(a, "famille_recomposee"):
            add_points(points, evidence, "contestation", 2, "Famille recomposée : droits et attentes potentiellement divergents")
            add_points(points, evidence, "conjoint", 2, "Famille recomposée : vigilance renforcée sur le conjoint")
        if is_yes(a, "conjoint_present") and a.get("accord_conjoint") in [NO, UNCERTAIN]:
            add_points(points, evidence, "conjoint", 2, "Accord du conjoint absent ou incertain")
            add_points(points, evidence, "contestation", 1, "Adhésion du conjoint non sécurisée")
        if is_no(a, "audit_civil") and (nb_enfants >= 2 or is_yes(a, "famille_recomposee") or is_yes(a, "conjoint_present")):
            add_points(points, evidence, "contestation", 2, "Aucun audit civil ou successoral réalisé")

    if objectif_fiscal:
        add_points(points, evidence, "fiscalite", 2, "Objectif déclaré d’optimisation fiscale")
        if valeur >= 3_000_000:
            add_points(points, evidence, "fiscalite", 2, "Valeur d’entreprise élevée")
        elif valeur >= 1_000_000:
            add_points(points, evidence, "fiscalite", 1, "Valeur d’entreprise significative")
        if is_no(a, "simulation_fiscale"):
            add_points(points, evidence, "fiscalite", 2, "Aucune simulation fiscale préalable")
    elif objectif_transmission and valeur >= 3_000_000:
        add_points(points, evidence, "fiscalite", 1, "Valeur élevée dans un contexte de transmission")

    if a.get("pacte_dutreil") == YES:
        add_points(points, evidence, "dutreil", 2, "Pacte Dutreil envisagé")
        if a.get("holding_animatrice") == UNCERTAIN:
            add_points(points, evidence, "dutreil", 4, "Qualification de holding animatrice incertaine")
        elif a.get("holding_animatrice") == NO:
            add_points(points, evidence, "dutreil", 3, "Holding déclarée non animatrice")
        if is_no(a, "audit_dutreil"):
            add_points(points, evidence, "dutreil", 2, "Audit Dutreil non réalisé")
        if is_no(a, "suivi_engagements"):
            add_points(points, evidence, "dutreil", 3, "Suivi des engagements Dutreil non prévu")

    # Protection familiale : les réponses sont prises en compte dès qu’un conjoint est présent.
    # La pondération est renforcée si l’objectif « protéger le conjoint et les proches » est explicitement retenu.
    if is_yes(a, "conjoint_present"):
        bonus_protection = 1 if objectif_protection else 0
        if is_yes(a, "conjoint_dependant"):
            add_points(points, evidence, "conjoint", 2 + bonus_protection, "Conjoint financièrement dépendant")
        if is_no(a, "protection_conjoint_prevue"):
            add_points(points, evidence, "conjoint", 2 + bonus_protection, "Protection du conjoint non prévue")
        if is_no(a, "regime_matrimonial_adapte"):
            add_points(points, evidence, "conjoint", 1 + bonus_protection, "Régime matrimonial non analysé ou non adapté")

    if poids >= 60:
        add_points(points, evidence, "dependance", 4, "Entreprise représentant au moins 60 % du patrimoine")
        add_points(points, evidence, "liquidite", 3, "Patrimoine fortement concentré dans l’entreprise")
    elif poids >= 40:
        add_points(points, evidence, "dependance", 2, "Entreprise représentant au moins 40 % du patrimoine")
        add_points(points, evidence, "liquidite", 1, "Liquidité potentielle à vérifier")

    if a.get("actifs_liquides") == "Faible":
        add_points(points, evidence, "liquidite", 3, "Actifs liquides faibles hors entreprise")
        add_points(points, evidence, "dependance", 2, "Patrimoine liquide ou diversifié insuffisant")
        if is_yes(a, "conjoint_present"):
            add_points(points, evidence, "conjoint", 1, "Faibles liquidités disponibles pour sécuriser les proches")
    elif a.get("actifs_liquides") == "Moyen":
        add_points(points, evidence, "liquidite", 1, "Actifs liquides moyens : simulation à confirmer")

    if a.get("endettement_familial") == "Élevé":
        add_points(points, evidence, "liquidite", 2, "Endettement personnel ou familial élevé")
        add_points(points, evidence, "conjoint", 1, "Endettement pouvant fragiliser les proches")
    elif a.get("endettement_familial") == "Moyen":
        add_points(points, evidence, "liquidite", 1, "Endettement familial à surveiller")

    # Besoin de revenus : signal transversal important.
    # Il influence la liquidité, la protection familiale, la gouvernance et la dépendance au patrimoine professionnel.
    if a.get("besoin_revenus_famille") == "Élevé":
        add_points(points, evidence, "liquidite", 3, "Besoin de revenus réguliers élevé : pression potentielle sur la trésorerie et les distributions")
        add_points(points, evidence, "blocage_gouvernance", 1, "Besoin de revenus pouvant générer des divergences sur les dividendes")
        if is_yes(a, "conjoint_present") or objectif_protection:
            add_points(points, evidence, "conjoint", 2, "Besoin de revenus élevé pour sécuriser le conjoint ou les proches")
        else:
            add_points(points, evidence, "conjoint", 1, "Besoin de revenus familial à analyser même hors objectif explicite de protection")
        if poids >= 40 or a.get("actifs_liquides") in ["Faible", "Moyen"]:
            add_points(points, evidence, "dependance", 1, "Besoin de revenus renforçant la dépendance au patrimoine professionnel")
        if a.get("actifs_liquides") == "Faible":
            add_points(points, evidence, "liquidite", 1, "Besoin de revenus élevé combiné à de faibles liquidités")
        if is_no(a, "politique_dividendes_definie"):
            add_points(points, evidence, "blocage_gouvernance", 2, "Besoin de revenus élevé sans politique de distribution définie")
            add_points(points, evidence, "liquidite", 1, "Politique de distribution à définir pour répondre aux besoins de revenus")
        elif a.get("politique_dividendes_definie") == UNKNOWN:
            add_points(points, evidence, "blocage_gouvernance", 1, "Politique de distribution non renseignée malgré un besoin de revenus élevé")
    elif a.get("besoin_revenus_famille") == "Moyen":
        add_points(points, evidence, "liquidite", 1, "Besoin de revenus familiaux à intégrer")
        if is_yes(a, "conjoint_present") or objectif_protection:
            add_points(points, evidence, "conjoint", 1, "Besoin de revenus moyen à intégrer dans la protection familiale")
        if is_no(a, "politique_dividendes_definie"):
            add_points(points, evidence, "blocage_gouvernance", 1, "Besoin de revenus moyen sans politique de distribution définie")

    if is_no(a, "diversification") and (objectif_diversification or poids >= 40):
        add_points(points, evidence, "dependance", 2, "Diversification patrimoniale insuffisante")
    if is_no(a, "prevoyance") and is_yes(a, "conjoint_present"):
        add_points(points, evidence, "conjoint", 1, "Prévoyance non identifiée")
        add_points(points, evidence, "dependance", 1, "Absence de mécanisme assurantiel de protection")

    if objectifs:
        if is_no(a, "suivi_annuel"):
            add_points(points, evidence, "suivi", 3, "Aucun rendez-vous de suivi annuel prévu")
        if a.get("pacte_dutreil") == YES and is_no(a, "suivi_engagements"):
            add_points(points, evidence, "suivi", 2, "Engagements fiscaux sans processus de suivi")
        if is_no(a, "formalisation_rapport"):
            add_points(points, evidence, "suivi", 2, "Absence de rapport écrit prévu")

    return build_df(points, evidence, a)


def build_df(points: Dict[str, int], evidence: Dict[str, List[str]], answers: Dict[str, Any] | None = None) -> Tuple[pd.DataFrame, Dict[str, List[str]]]:
    rows = []
    for code, definition in RISK_DEFINITIONS.items():
        raw_prob = points.get(code, 0)
        weight_label, weight_factor = weight_for_risk(code, answers)
        prob = clamp(math.ceil(raw_prob * weight_factor)) if raw_prob > 0 else 0
        grav = definition.gravite_base if prob > 0 else 0
        score = prob * grav if prob > 0 else 0
        niv = niveau(score)
        rows.append({
            "Code": code,
            "Objectif": definition.objectif,
            "Risque": definition.libelle,
            "Probabilité": prob,
            "Gravité": grav,
            "Score": score,
            "Niveau": niv,
            "Pondération objectif": weight_label if raw_prob > 0 else "Non applicable",
            "Points bruts": raw_prob,
            "Facteur de pondération": weight_factor if raw_prob > 0 else 0,
            "Signaux retenus": " | ".join(evidence.get(code, [])) if evidence.get(code) else "Aucun signal détecté : risque inexistant au vu des réponses validées.",
            "Outils à étudier": ", ".join(definition.outils) if prob > 0 else "Aucun outil spécifique à ce stade",
            "Justification des outils": tool_justifications_text(code) if prob > 0 else "Aucune justification : risque inexistant au vu des réponses validées.",
            "Actions préventives": ", ".join(definition.actions_preventives) if prob > 0 else "Aucune action prioritaire à ce stade",
            "Professionnels": ", ".join(definition.professionnels) if prob > 0 else "Non prioritaire",
            "Rang": PRIORITY_ORDER[niv],
        })
    return pd.DataFrame(rows).sort_values(["Rang", "Score", "Gravité", "Risque"], ascending=[False, False, False, True]).reset_index(drop=True), evidence


def missing_data(a: Dict) -> List[str]:
    missing: List[str] = []
    objectifs = set(a.get("objectifs") or [])
    if not objectifs:
        missing.append("Aucun objectif n’a encore été validé.")
    if not is_known(a, "maturite_projet"):
        missing.append("Préciser le niveau de maturité du projet.")
    if not is_known(a, "delai_transmission"):
        missing.append("Préciser l’horizon de transmission envisagé.")
    if not is_known(a, "qualite_information"):
        missing.append("Qualifier le niveau de fiabilité des informations recueillies.")
    if "Transmettre l’entreprise" in objectifs and int(a.get("nb_enfants") or 0) > 0 and not is_known(a, "heritier_repreneur"):
        missing.append("Indiquer si un héritier repreneur est identifié.")
    if a.get("heritier_repreneur") == YES and int(a.get("nb_enfants") or 0) >= 2 and not is_known(a, "volonte_non_repreneurs"):
        missing.append("Préciser la volonté probable des héritiers non repreneurs.")
    if "Protéger le conjoint et les proches" in objectifs and is_yes(a, "conjoint_present"):
        for k, label in [("conjoint_dependant", "dépendance financière du conjoint"), ("protection_conjoint_prevue", "protection déjà prévue"), ("regime_matrimonial_adapte", "analyse du régime matrimonial"), ("accord_conjoint", "adhésion du conjoint au projet")]:
            if not is_known(a, k):
                missing.append(f"Préciser la {label}.")
    if "Optimiser la fiscalité" in objectifs and not is_known(a, "simulation_fiscale"):
        missing.append("Préciser si une simulation fiscale a été réalisée.")
    if not is_known(a, "endettement_familial"):
        missing.append("Préciser le niveau d’endettement personnel ou familial.")
    if not is_known(a, "besoin_revenus_famille"):
        missing.append("Préciser le besoin de revenus réguliers pour la famille.")
    if a.get("besoin_revenus_famille") in ["Moyen", "Élevé"] and not is_known(a, "politique_dividendes_definie"):
        missing.append("Préciser si une politique de distribution de dividendes est définie, car le besoin de revenus est significatif.")
    if a.get("besoin_revenus_famille") in ["Moyen", "Élevé"] and not is_known(a, "prevoyance"):
        missing.append("Préciser l’existence d’une prévoyance ou assurance décès, car le besoin de revenus est significatif.")
    if not is_known(a, "entreprise_dependante_dirigeant"):
        missing.append("Préciser si l’entreprise dépend fortement du dirigeant actuel.")
    return missing

# =============================================================================
# 4. Affichage
# =============================================================================

def render_list(items: List[str]) -> str:
    if not items:
        return "<p class='small-note'>Aucun élément.</p>"
    return "<ul>" + "".join(f"<li>{item}</li>" for item in items) + "</ul>"


def get_tool_justifications(risk_code: str) -> Dict[str, str]:
    """Renvoie les justifications pédagogiques des outils associés à un risque."""
    definition = RISK_DEFINITIONS[risk_code]
    specific = TOOL_JUSTIFICATIONS_BY_RISK.get(risk_code, {})
    return {
        tool: specific.get(
            tool,
            f"Cet outil est proposé car il constitue une piste d’analyse pertinente pour réduire le risque « {definition.libelle} »."
        )
        for tool in definition.outils
    }


def render_tool_justifications(risk_code: str) -> str:
    justifications = get_tool_justifications(risk_code)
    if not justifications:
        return "<p class='small-note'>Aucune justification disponible.</p>"
    rows = "".join(
        f"<tr><td><strong>{tool}</strong></td><td>{why}</td></tr>"
        for tool, why in justifications.items()
    )
    return f"""
    <table class="justification-table">
        <thead><tr><th>Outil proposé</th><th>Pourquoi cet outil est proposé</th></tr></thead>
        <tbody>{rows}</tbody>
    </table>
    """


def tool_justifications_text(risk_code: str) -> str:
    return " | ".join(f"{tool} : {why}" for tool, why in get_tool_justifications(risk_code).items())


def priority_badge(priority: str) -> str:
    color = PRIORITY_COLORS.get(priority, "#374151")
    return f"<span class='badge' style='background:{color}'>{priority}</span>"


def card(title: str, body: str, accent: str = "#0f4c81") -> None:
    st.markdown(f"""
    <div class="card" style="border-left: 6px solid {accent};">
        <h3>{title}</h3>
        <div>{body}</div>
    </div>
    """, unsafe_allow_html=True)


def build_action_plan(df: pd.DataFrame) -> pd.DataFrame:
    """Construit un plan d'action en trois temps à partir des risques détectés."""
    rows: List[Dict[str, str]] = []
    detected = df[df["Score"] > 0].copy()
    for _, row in detected.iterrows():
        definition = RISK_DEFINITIONS[row["Code"]]
        risk_name = definition.libelle
        level = row["Niveau"]
        if level in ["Critique", "Élevé"]:
            rows.append({
                "Horizon": "Actions immédiates",
                "Priorité": level,
                "Action": f"Traiter en priorité le risque « {risk_name} » : {definition.actions_preventives[0]}.",
                "Risques concernés": risk_name,
                "Professionnels": ", ".join(definition.professionnels),
            })
            if row["Code"] in ["dutreil", "fiscalite", "contestation"]:
                rows.append({
                    "Horizon": "Actions immédiates",
                    "Priorité": level,
                    "Action": "Organiser un audit technique avec les professionnels compétents avant toute mise en œuvre.",
                    "Risques concernés": risk_name,
                    "Professionnels": ", ".join(definition.professionnels),
                })
        if level in ["Critique", "Élevé", "Moyen"]:
            rows.append({
                "Horizon": "Actions à moyen terme",
                "Priorité": level,
                "Action": f"Étudier et arbitrer les outils suivants : {', '.join(definition.outils[:4])}.",
                "Risques concernés": risk_name,
                "Professionnels": ", ".join(definition.professionnels),
            })
            if len(definition.actions_preventives) > 1:
                rows.append({
                    "Horizon": "Actions à moyen terme",
                    "Priorité": level,
                    "Action": definition.actions_preventives[1],
                    "Risques concernés": risk_name,
                    "Professionnels": ", ".join(definition.professionnels),
                })
        if level in ["Critique", "Élevé", "Moyen", "Faible"]:
            rows.append({
                "Horizon": "Suivi annuel",
                "Priorité": level,
                "Action": "Réévaluer le risque, vérifier l’adéquation des outils et actualiser la stratégie si la situation familiale, patrimoniale ou fiscale évolue.",
                "Risques concernés": risk_name,
                "Professionnels": "CGP, puis professionnels spécialisés selon le risque",
            })

    if rows:
        order = {"Actions immédiates": 0, "Actions à moyen terme": 1, "Suivi annuel": 2}
        priority_order = {"Critique": 0, "Élevé": 1, "Moyen": 2, "Faible": 3, "Inexistant": 4}
        out = pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)
        out["_h"] = out["Horizon"].map(order)
        out["_p"] = out["Priorité"].map(priority_order)
        return out.sort_values(["_h", "_p", "Risques concernés"]).drop(columns=["_h", "_p"]).reset_index(drop=True)
    return pd.DataFrame(columns=["Horizon", "Priorité", "Action", "Risques concernés", "Professionnels"])


def markdown_report(client_name: str, answers: Dict, df: pd.DataFrame) -> str:
    detected = df[df["Score"] > 0].copy()
    lines = ["# Rapport de diagnostic préventif – Holding familiale"]
    if client_name:
        lines.append(f"\nClient / dossier : **{client_name}**")
    lines.append("\n## 1. Synthèse du diagnostic")
    lines.append(f"- Objectifs validés : {', '.join(answers.get('objectifs') or []) if answers.get('objectifs') else 'Non renseigné'}")
    lines.append(f"- Nombre d’enfants : {answers.get('nb_enfants')}")
    lines.append(f"- Héritier repreneur identifié : {answers.get('heritier_repreneur')}")
    lines.append(f"- Poids estimé de l’entreprise dans le patrimoine : {answers.get('poids_entreprise')} %")
    lines.append(f"- Actifs liquides hors entreprise : {answers.get('actifs_liquides')}")
    lines.append(f"- Pacte Dutreil envisagé : {answers.get('pacte_dutreil')}")
    lines.append("\n## 2. Risques détectés et solutions à étudier")
    if detected.empty:
        lines.append("Aucun risque significatif n’a été détecté à partir des réponses validées.")
    else:
        for _, row in detected.iterrows():
            lines.append(f"\n### {row['Risque']} — {row['Niveau']} / score {row['Score']}")
            lines.append(f"Objectif concerné : {row['Objectif']}")
            lines.append(f"Signaux retenus : {row['Signaux retenus']}")
            lines.append(f"Outils à étudier : {row['Outils à étudier']}")
            lines.append(f"Pourquoi ces outils sont proposés : {row['Justification des outils']}")
            lines.append(f"Actions préventives : {row['Actions préventives']}")
            lines.append(f"Professionnels à mobiliser : {row['Professionnels']}")
    lines.append("\n## 3. Risques inexistants")
    zero = df[df["Score"] == 0]["Risque"].tolist()
    lines.append(", ".join(zero) if zero else "Aucun.")
    gaps = missing_data(answers)
    if gaps:
        lines.append("\n## 4. Informations manquantes à compléter")
        for g in gaps:
            lines.append(f"- {g}")
    lines.append("\n## Limites")
    lines.append("Ce rapport est un support d’aide à la décision. Il ne constitue pas une consultation juridique, fiscale ou notariale. Les pistes proposées doivent être validées avec les professionnels compétents.")
    return "\n".join(lines)


def create_docx_report(client_name: str, answers: Dict, df: pd.DataFrame, evidence: Dict[str, List[str]]) -> bytes:
    """Génère un rapport Word dont le contenu varie réellement selon l’orientation choisie."""
    if Document is None:
        raise RuntimeError("La dépendance python-docx n'est pas installée.")

    orientation = str(answers.get("rapport_orientation") or "Équilibré")
    orientation_lower = orientation.lower()
    is_synthetic = "synth" in orientation_lower
    is_pedagogical = "pédagog" in orientation_lower or "pedagog" in orientation_lower
    is_technical = "technique" in orientation_lower or "approfondi" in orientation_lower

    detected = df[df["Score"] > 0].copy()
    top3 = detected.head(3).copy()

    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    styles = doc.styles
    try:
        styles["Normal"].font.name = "Aptos"
        styles["Normal"].font.size = Pt(10)
        styles["Heading 1"].font.color.rgb = RGBColor(15, 76, 129)
        styles["Heading 2"].font.color.rgb = RGBColor(11, 114, 133)
    except Exception:
        pass

    def shade_cell(cell, fill: str) -> None:
        if OxmlElement is None or qn is None:
            return
        tc_pr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), fill)
        tc_pr.append(shd)

    def set_cell_text(cell, text: str, bold: bool = False, color: str | None = None, size: int | None = None) -> None:
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(str(text))
        run.bold = bold
        if color:
            run.font.color.rgb = RGBColor.from_string(color)
        if size:
            run.font.size = Pt(size)

    def add_bullets(items: List[str], max_items: int | None = None) -> None:
        shown = list(items or [])
        if max_items is not None:
            shown = shown[:max_items]
        if not shown:
            doc.add_paragraph("Aucun élément renseigné à ce stade.")
            return
        for item in shown:
            doc.add_paragraph(str(item), style="List Bullet")

    def add_small_note(text: str, color: str = "475569") -> None:
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.italic = True
        run.font.color.rgb = RGBColor.from_string(color)
        run.font.size = Pt(9)

    def add_label_value_table(title: str, rows: List[Tuple[str, str]], compact: bool = False) -> None:
        doc.add_heading(title, level=2)
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"
        set_cell_text(table.rows[0].cells[0], "Information", bold=True, color="FFFFFF", size=9)
        set_cell_text(table.rows[0].cells[1], "Réponse validée", bold=True, color="FFFFFF", size=9)
        shade_cell(table.rows[0].cells[0], "0F4C81")
        shade_cell(table.rows[0].cells[1], "0F4C81")
        for label, value in rows:
            cells = table.add_row().cells
            cells[0].text = str(label)
            cells[1].text = str(value)
        if compact:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(9)

    def add_tool_justification_table(risk_code: str, max_tools: int | None = None) -> None:
        justifications = get_tool_justifications(risk_code)
        if max_tools is not None:
            justifications = dict(list(justifications.items())[:max_tools])
        if not justifications:
            doc.add_paragraph("Aucune justification disponible à ce stade.")
            return
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"
        set_cell_text(table.rows[0].cells[0], "Outil proposé", bold=True, color="FFFFFF", size=9)
        set_cell_text(table.rows[0].cells[1], "Pourquoi cet outil est proposé", bold=True, color="FFFFFF", size=9)
        shade_cell(table.rows[0].cells[0], "0F4C81")
        shade_cell(table.rows[0].cells[1], "0F4C81")
        for tool, why in justifications.items():
            cells = table.add_row().cells
            cells[0].text = str(tool)
            cells[1].text = str(why)

    def add_cover_page() -> None:
        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = title.add_run("Rapport de diagnostic préventif\nHolding familiale")
        r.bold = True
        r.font.size = Pt(24)
        r.font.color.rgb = RGBColor(15, 76, 129)
        doc.add_paragraph()
        subtitle = doc.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rs = subtitle.add_run(f"Format retenu : {orientation}")
        rs.bold = True
        rs.font.size = Pt(13)
        rs.font.color.rgb = RGBColor(11, 114, 133)
        doc.add_paragraph()
        meta_rows = [
            ("Client / dossier", client_name or answers.get("client_name") or "Non renseigné"),
            ("Entreprise", answers.get("company_name") or "Non renseigné"),
            ("CGP / cabinet", answers.get("cgp_name") or "Non renseigné"),
            ("Date de l’entretien", answers.get("entretien_date") or datetime.now().strftime("%d/%m/%Y")),
        ]
        add_label_value_table("Informations générales", [(k, str(v)) for k, v in meta_rows], compact=True)
        doc.add_paragraph()
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run("Objectif → Risque → Outil → Prévention → Suivi")
        run.bold = True
        run.font.color.rgb = RGBColor(15, 76, 129)
        doc.add_page_break()

    def add_static_toc() -> None:
        doc.add_heading("Sommaire", level=1)
        if is_synthetic:
            toc_items = [
                "1. Synthèse exécutive",
                "2. Informations clés du dossier",
                "3. Risques prioritaires",
                "4. Solutions et actions prioritaires",
                "5. Points à valider",
                "6. Clause de prudence",
            ]
        elif is_pedagogical:
            toc_items = [
                "1. Synthèse exécutive",
                "2. Informations recueillies",
                "3. Lecture pédagogique des risques",
                "4. Pourquoi ces outils sont proposés",
                "5. Plan d’action expliqué",
                "6. Points à valider",
                "7. Clause de prudence",
            ]
        elif is_technical:
            toc_items = [
                "1. Synthèse exécutive",
                "2. Informations complètes recueillies",
                "3. Méthodologie et scoring",
                "4. Pondération des objectifs",
                "5. Matrice probabilité / gravité",
                "6. Analyse technique par risque",
                "7. Plan d’action hiérarchisé",
                "8. Points à valider et informations manquantes",
                "9. Clause de prudence",
            ]
        else:
            toc_items = [
                "1. Synthèse exécutive",
                "2. Informations recueillies",
                "3. Pondération des objectifs",
                "4. Top 3 des risques prioritaires",
                "5. Matrice probabilité / gravité",
                "6. Analyse et solutions par risque",
                "7. Plan d’action en trois temps",
                "8. Points à valider",
                "9. Clause de prudence",
            ]
        for item in toc_items:
            doc.add_paragraph(item, style="List Number")
        doc.add_page_break()

    def add_executive_summary() -> None:
        doc.add_heading("1. Synthèse exécutive", level=1)
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"
        cell = table.rows[0].cells[0]
        shade_cell(cell, "EAF4F8")
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run("Lecture du diagnostic\n")
        r.bold = True
        r.font.size = Pt(13)
        r.font.color.rgb = RGBColor(15, 76, 129)
        if detected.empty:
            text = "Aucun risque significatif n’a été détecté à partir des réponses validées. Le diagnostic doit être complété si de nouvelles informations apparaissent."
        else:
            top = detected.iloc[0]
            crit = int((detected["Niveau"] == "Critique").sum())
            high = int((detected["Niveau"] == "Élevé").sum())
            base = f"{len(detected)} risque(s) détecté(s), dont {crit} critique(s) et {high} élevé(s). Le risque principal identifié est : {top['Risque']} ({top['Niveau']}, score {int(top['Score'])})."
            if is_synthetic:
                text = base + " Le rapport se concentre sur les risques prioritaires, les décisions à prendre et les actions immédiates."
            elif is_pedagogical:
                text = base + " Le rapport privilégie une lecture accessible : il explique les risques, les raisons des outils proposés et les actions à mener pour rendre la stratégie compréhensible pour le client et sa famille."
            elif is_technical:
                text = base + " Le rapport détaille les signaux retenus, la pondération des objectifs, le scoring, les points de vigilance et les validations professionnelles nécessaires."
            else:
                text = base + " Les recommandations constituent des pistes de travail à valider avec les professionnels compétents."
        p.add_run(text)
        if answers.get("objectif_libre"):
            doc.add_paragraph(f"Objectif exprimé par le client : {answers.get('objectif_libre')}")

    def add_context_section() -> None:
        key_rows = [
            ("Client / dossier", str(answers.get("client_name") or "Non renseigné")),
            ("Âge du dirigeant", str(answers.get("client_age") or "Non renseigné")),
            ("Entreprise", str(answers.get("company_name") or "Non renseigné")),
            ("Activité / secteur", str(answers.get("company_activity") or "Non renseigné")),
            ("Forme juridique", str(answers.get("company_form", "Non renseigné"))),
            ("Maturité du projet", str(answers.get("maturite_projet", "Non renseigné"))),
            ("Horizon de transmission", str(answers.get("delai_transmission", "Non renseigné"))),
            ("Qualité des informations", str(answers.get("qualite_information", "Non renseigné"))),
            ("Objectifs validés", ", ".join(answers.get("objectifs") or []) or "Non renseigné"),
        ]
        if is_synthetic:
            doc.add_heading("2. Informations clés du dossier", level=1)
            add_label_value_table("Éléments essentiels", key_rows, compact=True)
            return

        doc.add_heading("2. Informations recueillies", level=1)
        full_rows = key_rows + [
            ("Nombre d’enfants", str(answers.get("nb_enfants", "Non renseigné"))),
            ("Conjoint présent", str(answers.get("conjoint_present", "Non renseigné"))),
            ("Famille recomposée", str(answers.get("famille_recomposee", "Non renseigné"))),
            ("Dialogue familial", str(answers.get("dialogue_familial", "Non renseigné"))),
            ("Héritier repreneur identifié", str(answers.get("heritier_repreneur", "Non renseigné"))),
            ("Volonté des héritiers non repreneurs", str(answers.get("volonte_non_repreneurs", "Non renseigné"))),
            ("Valeur estimée de l’entreprise", f"{answers.get('valeur_entreprise', 0):,.0f} €".replace(",", " ")),
            ("Poids de l’entreprise dans le patrimoine", f"{answers.get('poids_entreprise', 0)} %"),
            ("Actifs liquides hors entreprise", str(answers.get("actifs_liquides", "Non renseigné"))),
            ("Besoin de revenus réguliers", str(answers.get("besoin_revenus_famille", "Non renseigné"))),
            ("Protection du conjoint prévue", str(answers.get("protection_conjoint_prevue", "Non renseigné"))),
            ("Gouvernance formalisée", str(answers.get("gouvernance_formalisee", "Non renseigné"))),
            ("Pacte Dutreil envisagé", str(answers.get("pacte_dutreil", "Non renseigné"))),
            ("Suivi annuel prévu", str(answers.get("suivi_annuel", "Non renseigné"))),
        ]
        if is_technical:
            full_rows += [
                ("Événement d’urgence ou de fragilité", str(answers.get("urgence_evenement", "Non renseigné"))),
                ("Accord du conjoint", str(answers.get("accord_conjoint", "Non renseigné"))),
                ("Soulte envisagée", str(answers.get("soulte_envisagee", "Non renseigné"))),
                ("Capacité de financement de la soulte", str(answers.get("capacite_financement_soulte", "Non renseigné"))),
                ("Valorisation indépendante", str(answers.get("valorisation_independante", "Non renseigné"))),
                ("Audit civil", str(answers.get("audit_civil", "Non renseigné"))),
                ("Endettement familial", str(answers.get("endettement_familial", "Non renseigné"))),
                ("Entreprise dépendante du dirigeant", str(answers.get("entreprise_dependante_dirigeant", "Non renseigné"))),
                ("Qualification holding animatrice", str(answers.get("holding_animatrice", "Non renseigné"))),
                ("Audit Dutreil", str(answers.get("audit_dutreil", "Non renseigné"))),
                ("Suivi des engagements", str(answers.get("suivi_engagements", "Non renseigné"))),
            ]
        add_label_value_table("Synthèse des réponses validées", full_rows, compact=not is_technical)

    def add_personalization_section() -> None:
        if is_synthetic:
            return
        doc.add_heading("3. Personnalisation du rapport", level=1)
        rows = [
            ("Orientation souhaitée", orientation),
            ("Niveau de détail attendu", str(answers.get("niveau_detail", "Détaillé"))),
            ("Objectif exprimé par le client", str(answers.get("objectif_libre") or "Non renseigné")),
            ("Attentes particulières", str(answers.get("attentes_client") or "Non renseigné")),
            ("Contraintes ou préférences", str(answers.get("contraintes_client") or "Non renseigné")),
            ("Personnes à associer", str(answers.get("personnes_a_associer") or "Non renseigné")),
            ("Observations libres du CGP", str(answers.get("observations") or "Non renseigné")),
        ]
        add_label_value_table("Éléments de contexte et de personnalisation", rows, compact=True)

    def add_methodology_section() -> None:
        if not is_technical:
            return
        doc.add_heading("3. Méthodologie et scoring", level=1)
        doc.add_paragraph(
            "Le prototype fonctionne comme un système expert fondé sur des règles. Les réponses validées activent ou non des signaux d’alerte. "
            "Ces signaux alimentent une probabilité de survenance du risque, ensuite combinée à une gravité de base propre à chaque risque. "
            "La pondération des objectifs renforce le score lorsque le risque est directement lié à un objectif jugé très important ou prioritaire."
        )
        table = doc.add_table(rows=1, cols=4)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, h in enumerate(["Niveau", "Lecture", "Type d’attention", "Suite logique"]):
            set_cell_text(table.rows[0].cells[i], h, bold=True, color="FFFFFF", size=9)
            shade_cell(table.rows[0].cells[i], "0F4C81")
        rows = [
            ("Inexistant", "Aucun signal détecté", "Pas d’action prioritaire", "Surveillance simple"),
            ("Faible", "Signal limité", "Point à surveiller", "Compléter si nécessaire"),
            ("Moyen", "Risque plausible", "Analyse complémentaire", "Valider avec professionnel"),
            ("Élevé", "Risque structurant", "Traitement prioritaire", "Plan d’action à organiser"),
            ("Critique", "Risque majeur", "Action immédiate", "Validation et coordination rapides"),
        ]
        for row in rows:
            cells = table.add_row().cells
            for idx, val in enumerate(row):
                cells[idx].text = val

    def add_objective_weights_section() -> None:
        if is_synthetic:
            return
        doc.add_heading("4. Pondération des objectifs", level=1)
        weights = answers.get("objective_weights") or {}
        objectives = answers.get("objectifs") or []
        if not objectives:
            doc.add_paragraph("Aucun objectif n’a été validé.")
            return
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"
        set_cell_text(table.rows[0].cells[0], "Objectif", bold=True, color="FFFFFF")
        set_cell_text(table.rows[0].cells[1], "Ordre d’importance", bold=True, color="FFFFFF")
        shade_cell(table.rows[0].cells[0], "0F4C81")
        shade_cell(table.rows[0].cells[1], "0F4C81")
        for obj in objectives:
            cells = table.add_row().cells
            cells[0].text = obj
            cells[1].text = weights.get(obj, "Important")

    def add_top_risks_section() -> None:
        doc.add_heading("3. Risques prioritaires" if is_synthetic else "5. Top 3 des risques prioritaires", level=1)
        if top3.empty:
            doc.add_paragraph("Aucun risque prioritaire n’a été détecté au vu des réponses validées.")
            return
        table = doc.add_table(rows=1, cols=5)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"
        headers = ["Rang", "Risque", "Niveau", "Score", "Première réponse à envisager"]
        for i, h in enumerate(headers):
            set_cell_text(table.rows[0].cells[i], h, bold=True, color="FFFFFF", size=9)
            shade_cell(table.rows[0].cells[i], "0F4C81")
        for idx, (_, row) in enumerate(top3.iterrows(), start=1):
            definition = RISK_DEFINITIONS[row["Code"]]
            cells = table.add_row().cells
            cells[0].text = str(idx)
            cells[1].text = str(row["Risque"])
            cells[2].text = str(row["Niveau"])
            cells[3].text = str(int(row["Score"]))
            cells[4].text = definition.actions_preventives[0] if definition.actions_preventives else "À approfondir"
            color = {"Critique": "FEE2E2", "Élevé": "FFEDD5", "Moyen": "DBEAFE", "Faible": "DCFCE7"}.get(str(row["Niveau"]), "F8FAFC")
            shade_cell(cells[2], color)

    def add_probability_gravity_matrix() -> None:
        if is_synthetic or is_pedagogical:
            return
        doc.add_heading("6. Matrice probabilité / gravité", level=1)
        doc.add_paragraph("Cette matrice positionne les risques détectés selon leur probabilité estimée et leur gravité.")
        table = doc.add_table(rows=6, cols=6)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"
        set_cell_text(table.cell(0, 0), "Gravité \\ Probabilité", bold=True, color="FFFFFF", size=8)
        shade_cell(table.cell(0, 0), "0F4C81")
        for prob in range(1, 6):
            set_cell_text(table.cell(0, prob), score_to_label(prob), bold=True, color="FFFFFF", size=8)
            shade_cell(table.cell(0, prob), "0F4C81")
        for grav in range(1, 6):
            set_cell_text(table.cell(grav, 0), score_to_label(grav), bold=True, color="FFFFFF", size=8)
            shade_cell(table.cell(grav, 0), "0F4C81")
        for _, row in detected.iterrows():
            prob = int(row["Probabilité"])
            grav = int(row["Gravité"])
            if 1 <= prob <= 5 and 1 <= grav <= 5:
                cell = table.cell(grav, prob)
                current = cell.text.strip()
                value = str(row["Risque"])
                cell.text = value if not current else current + "\n" + value
                color = {"Critique": "FEE2E2", "Élevé": "FFEDD5", "Moyen": "DBEAFE", "Faible": "DCFCE7"}.get(str(row["Niveau"]), "F8FAFC")
                shade_cell(cell, color)

    def add_risk_analysis_section() -> None:
        if is_synthetic:
            doc.add_heading("4. Solutions et actions prioritaires", level=1)
        elif is_pedagogical:
            doc.add_heading("3. Lecture pédagogique des risques", level=1)
        elif is_technical:
            doc.add_heading("6. Analyse technique par risque", level=1)
        else:
            doc.add_heading("6. Analyse et solutions par risque", level=1)

        if detected.empty:
            doc.add_paragraph("Aucune solution prioritaire n’est proposée à ce stade. Le diagnostic doit être complété si de nouveaux objectifs ou signaux apparaissent.")
            return

        risks_to_show = detected.head(5) if is_synthetic else detected
        for _, row in risks_to_show.iterrows():
            definition = RISK_DEFINITIONS[row["Code"]]
            doc.add_heading(f"{definition.libelle} – niveau {row['Niveau']}", level=2)
            if is_synthetic:
                table = doc.add_table(rows=1, cols=2)
                table.style = "Table Grid"
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                set_cell_text(table.rows[0].cells[0], "Lecture", bold=True, color="FFFFFF", size=9)
                set_cell_text(table.rows[0].cells[1], "Réponse prioritaire", bold=True, color="FFFFFF", size=9)
                shade_cell(table.rows[0].cells[0], "0F4C81")
                shade_cell(table.rows[0].cells[1], "0F4C81")
                cells = table.add_row().cells
                cells[0].text = f"Score {int(row['Score'])} – {row['Niveau']}. Signaux principaux : {row['Signaux retenus']}"
                cells[1].text = f"Outils à étudier : {', '.join(definition.outils[:3])}. Action immédiate : {definition.actions_preventives[0] if definition.actions_preventives else 'À approfondir.'}"
                continue

            if is_pedagogical:
                doc.add_paragraph(
                    f"Ce risque concerne l’objectif suivant : {definition.objectif}. Il apparaît car certains éléments du diagnostic peuvent fragiliser l’atteinte de cet objectif."
                )
                doc.add_heading("Ce que cela peut provoquer", level=3)
                add_bullets(definition.consequences)
                doc.add_heading("Pourquoi ce risque ressort dans le dossier", level=3)
                add_bullets(evidence.get(row["Code"], []), max_items=5)
                doc.add_heading("Outils à expliquer au client", level=3)
                add_tool_justification_table(row["Code"], max_tools=4)
                doc.add_heading("Actions préventives formulées simplement", level=3)
                add_bullets(definition.actions_preventives)
                continue

            # Équilibré ou technique
            doc.add_paragraph(f"Objectif concerné : {definition.objectif}")
            doc.add_paragraph(f"Score : {int(row['Score'])} | Probabilité : {score_to_label(int(row['Probabilité']))} | Gravité : {score_to_label(int(row['Gravité']))} | Pondération : {row['Pondération objectif']}")
            doc.add_heading("Signaux d’alerte retenus", level=3)
            add_bullets(evidence.get(row["Code"], []))
            doc.add_heading("Conséquences possibles", level=3)
            add_bullets(definition.consequences)
            doc.add_heading("Outils à étudier", level=3)
            add_bullets(definition.outils if is_technical else definition.outils[:4])
            doc.add_heading("Justification des outils proposés", level=3)
            add_tool_justification_table(row["Code"], None if is_technical else 4)
            doc.add_heading("Actions préventives", level=3)
            add_bullets(definition.actions_preventives)
            doc.add_heading("Professionnels à mobiliser", level=3)
            doc.add_paragraph(", ".join(definition.professionnels))

    def add_action_plan_section() -> None:
        title = "4. Plan d’action" if is_synthetic else "7. Recommandations hiérarchisées en trois temps"
        if is_pedagogical:
            title = "5. Plan d’action expliqué"
        doc.add_heading(title, level=1)
        plan_df = build_action_plan(df)
        if plan_df.empty:
            doc.add_paragraph("Aucune recommandation hiérarchisée n’est générée tant qu’aucun risque n’est détecté.")
            return
        for horizon in ["Actions immédiates", "Actions à moyen terme", "Suivi annuel"]:
            horizon_df = plan_df[plan_df["Horizon"] == horizon]
            if horizon_df.empty:
                continue
            doc.add_heading(horizon, level=2)
            if is_pedagogical:
                if horizon == "Actions immédiates":
                    doc.add_paragraph("Ces actions visent à sécuriser les points les plus sensibles avant toute mise en œuvre juridique ou fiscale.")
                elif horizon == "Actions à moyen terme":
                    doc.add_paragraph("Ces actions permettent de construire progressivement une stratégie cohérente et acceptable pour la famille.")
                else:
                    doc.add_paragraph("Le suivi annuel évite que la stratégie devienne inadaptée avec le temps.")
            table = doc.add_table(rows=1, cols=4)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.style = "Table Grid"
            for i, h in enumerate(["Priorité", "Action", "Risques concernés", "Professionnels"]):
                set_cell_text(table.rows[0].cells[i], h, bold=True, color="FFFFFF", size=9)
                shade_cell(table.rows[0].cells[i], "0B7285")
            max_rows = 6 if is_synthetic else None
            iterable = horizon_df.head(max_rows) if max_rows else horizon_df
            for _, plan_row in iterable.iterrows():
                cells = table.add_row().cells
                cells[0].text = str(plan_row["Priorité"])
                cells[1].text = str(plan_row["Action"])
                cells[2].text = str(plan_row["Risques concernés"])
                cells[3].text = str(plan_row["Professionnels"])

    def add_professional_validation_points() -> None:
        title = "5. Points à valider" if is_synthetic else "8. Points à valider avec les professionnels compétents"
        if is_technical:
            title = "8. Points à valider et informations manquantes"
        doc.add_heading(title, level=1)
        points = []
        if detected.empty:
            points.append("Compléter le diagnostic avant toute recommandation structurée.")
        else:
            for _, row in detected.iterrows():
                risk_code = row["Code"]
                risk = RISK_DEFINITIONS[risk_code]
                if risk_code in ["contestation", "conflit_heritiers", "conjoint"]:
                    points.append(f"Validation notariale à prévoir pour le risque : {risk.libelle}.")
                if risk_code in ["dilution", "tiers", "blocage_gouvernance", "successeur"]:
                    points.append(f"Validation juridique des statuts, pactes ou clauses à prévoir pour le risque : {risk.libelle}.")
                if risk_code in ["fiscalite", "dutreil"]:
                    points.append(f"Validation fiscale spécialisée à prévoir pour le risque : {risk.libelle}.")
                if risk_code in ["liquidite", "dependance"]:
                    points.append(f"Analyse financière et patrimoniale à approfondir pour le risque : {risk.libelle}.")
        # Déduplication en conservant l’ordre
        seen = set()
        uniq = []
        for point in points:
            if point not in seen:
                seen.add(point)
                uniq.append(point)
        add_bullets(uniq)
        if is_technical:
            gaps = missing_data(answers)
            doc.add_heading("Informations complémentaires à recueillir", level=2)
            add_bullets(gaps if gaps else ["Aucune information indispensable manquante n’a été identifiée par l’outil."])

    def add_zero_risks_section() -> None:
        if is_synthetic or is_pedagogical:
            return
        doc.add_heading("Risques non détectés au vu des réponses validées", level=1)
        zero = df[df["Score"] == 0]["Risque"].tolist()
        doc.add_paragraph(", ".join(zero) if zero else "Aucun.")

    def add_pedagogical_lexicon() -> None:
        if not is_pedagogical:
            return
        doc.add_heading("4. Repères pédagogiques sur les outils", level=1)
        entries = [
            ("Donation-partage", "Mécanisme permettant d’organiser de son vivant la répartition de tout ou partie du patrimoine entre les héritiers."),
            ("Soulte", "Compensation financière versée à un héritier lorsque les lots transmis ne sont pas équivalents en valeur."),
            ("Pacte d’associés", "Convention destinée à organiser les relations entre associés, notamment les décisions, les sorties et certaines règles de gouvernance."),
            ("Pacte Dutreil", "Dispositif fiscal permettant, sous conditions, de réduire fortement la base taxable d’une transmission d’entreprise."),
            ("Démembrement", "Séparation entre l’usufruit et la nue-propriété, utile pour organiser une transmission progressive."),
        ]
        table = doc.add_table(rows=1, cols=2)
        table.style = "Table Grid"
        set_cell_text(table.rows[0].cells[0], "Outil", bold=True, color="FFFFFF")
        set_cell_text(table.rows[0].cells[1], "Explication simple", bold=True, color="FFFFFF")
        shade_cell(table.rows[0].cells[0], "0F4C81")
        shade_cell(table.rows[0].cells[1], "0F4C81")
        for tool, explanation in entries:
            cells = table.add_row().cells
            cells[0].text = tool
            cells[1].text = explanation

    def add_caution_section() -> None:
        title = "6. Clause de prudence" if is_synthetic else "9. Clause de prudence et limites du rapport"
        if is_pedagogical:
            title = "7. Clause de prudence et limites du rapport"
        doc.add_heading(title, level=1)
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"
        cell = table.rows[0].cells[0]
        shade_cell(cell, "FFF7ED")
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run("Clause de prudence\n")
        r.bold = True
        r.font.color.rgb = RGBColor(154, 52, 18)
        p.add_run(
            "Ce rapport constitue un support d’aide à la décision pour le conseiller en gestion de patrimoine. "
            "Il ne constitue pas une consultation juridique, fiscale, comptable ou notariale. "
            "Il ne remplace pas l’intervention du notaire, de l’avocat, de l’expert-comptable ou de tout autre professionnel compétent. "
            "Les solutions évoquées sont des pistes d’analyse qui doivent être vérifiées, adaptées et validées avant toute mise en œuvre. "
            "La pertinence de la stratégie dépend de la qualité des informations recueillies et de l’évolution de la situation familiale, patrimoniale, professionnelle et fiscale."
        )

    add_cover_page()
    add_static_toc()
    add_executive_summary()
    add_context_section()
    add_personalization_section()
    add_methodology_section()
    add_objective_weights_section()
    add_top_risks_section()
    add_probability_gravity_matrix()
    add_risk_analysis_section()
    add_pedagogical_lexicon()
    add_action_plan_section()
    add_professional_validation_points()
    add_zero_risks_section()
    add_caution_section()

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

def create_xlsx_matrix(df: pd.DataFrame) -> bytes:
    """Génère une matrice Excel lisible et directement exploitable."""
    if Workbook is None:
        raise RuntimeError("La dépendance openpyxl n'est pas installée.")

    export_columns = [
        "Objectif", "Risque", "Probabilité", "Gravité", "Score", "Niveau", "Pondération objectif",
        "Signaux retenus", "Outils à étudier", "Justification des outils", "Actions préventives", "Professionnels",
    ]
    wb = Workbook()
    ws = wb.active
    ws.title = "Matrice des risques"

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(export_columns))
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = "Matrice des risques – Diagnostic holding familiale"
    title_cell.font = Font(bold=True, size=16, color="FFFFFF")
    title_cell.fill = PatternFill("solid", fgColor="0F4C81")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(export_columns))
    subtitle_cell = ws.cell(row=2, column=1)
    subtitle_cell.value = "Document de travail exporté depuis le prototype de système expert CGP"
    subtitle_cell.font = Font(italic=True, color="475569")
    subtitle_cell.alignment = Alignment(horizontal="center", vertical="center")

    header_row = 4
    for col_idx, col_name in enumerate(export_columns, start=1):
        cell = ws.cell(row=header_row, column=col_idx)
        cell.value = col_name
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0B7285")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    priority_fills = {
        "Critique": "FEE2E2",
        "Élevé": "FFEDD5",
        "Moyen": "DBEAFE",
        "Faible": "DCFCE7",
        "Inexistant": "F1F5F9",
    }

    sorted_df = df[export_columns].copy()
    for row_idx, (_, row) in enumerate(sorted_df.iterrows(), start=header_row + 1):
        priority = row.get("Niveau", "")
        fill_color = priority_fills.get(priority, "FFFFFF")
        for col_idx, col_name in enumerate(export_columns, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = row[col_name]
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
            if col_name == "Niveau":
                cell.fill = PatternFill("solid", fgColor=fill_color)
                cell.font = Font(bold=True)
            elif row_idx % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F8FAFC")

    widths = {
        "A": 24, "B": 34, "C": 12, "D": 10, "E": 10, "F": 18, "G": 22,
        "H": 55, "I": 42, "J": 70, "K": 48, "L": 34,
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width
    for row in range(header_row + 1, header_row + 1 + len(sorted_df)):
        ws.row_dimensions[row].height = 60
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(export_columns))}{header_row + len(sorted_df)}"

    plan_df = build_action_plan(df)
    ws_plan = wb.create_sheet("Plan d'action")
    plan_columns = ["Horizon", "Priorité", "Action", "Risques concernés", "Professionnels"]
    ws_plan.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(plan_columns))
    plan_title = ws_plan.cell(row=1, column=1)
    plan_title.value = "Plan d’action hiérarchisé – actions immédiates, moyen terme et suivi annuel"
    plan_title.font = Font(bold=True, size=15, color="FFFFFF")
    plan_title.fill = PatternFill("solid", fgColor="0F4C81")
    plan_title.alignment = Alignment(horizontal="center", vertical="center")
    for col_idx, col_name in enumerate(plan_columns, start=1):
        cell = ws_plan.cell(row=3, column=col_idx)
        cell.value = col_name
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0B7285")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row_idx, (_, row) in enumerate(plan_df.iterrows(), start=4):
        for col_idx, col_name in enumerate(plan_columns, start=1):
            cell = ws_plan.cell(row=row_idx, column=col_idx)
            cell.value = row[col_name]
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
            if row_idx % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F8FAFC")
    for col_letter, width in {"A": 24, "B": 16, "C": 70, "D": 36, "E": 42}.items():
        ws_plan.column_dimensions[col_letter].width = width
    for row in range(4, 4 + len(plan_df)):
        ws_plan.row_dimensions[row].height = 55
    ws_plan.freeze_panes = "A4"

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

init_app()

VIEWS = ["Questionnaire adaptatif", "Résultats et solutions", "Exporter", "Règles de décision", "Debug validation"]
st.session_state.setdefault("app_page", VIEWS[0])

def render_view_navigation(current_view: str) -> None:
    """Affiche des boutons de navigation entre les vues principales."""
    idx = VIEWS.index(current_view)
    st.divider()
    c1, c2, c3 = st.columns([1.6, 4, 1.6])
    with c1:
        if idx > 0:
            if st.button(f"← {VIEWS[idx - 1]}", key=f"prev_view_{current_view}"):
                st.session_state.app_page = VIEWS[idx - 1]
                st.rerun()
    with c2:
        st.caption("Navigation entre les onglets de l’outil")
    with c3:
        if idx < len(VIEWS) - 1:
            if st.button(f"{VIEWS[idx + 1]} →", key=f"next_view_{current_view}", type="primary"):
                st.session_state.app_page = VIEWS[idx + 1]
                st.rerun()

st.markdown("""
<style>
    .main .block-container { padding-top: 1.2rem; }
    .hero {background: linear-gradient(135deg, #0f4c81, #0b7285); color: white; padding: 24px 28px; border-radius: 16px; margin-bottom: 18px;}
    .hero h1 {margin: 0; color: white; font-size: 2rem;}
    .hero p {margin: 8px 0 0 0; color: #e6f7ff; font-size: 1.03rem;}
    .card {background: #ffffff !important; color: #0f172a !important; border-radius: 14px; padding: 18px 20px; margin: 12px 0; box-shadow: 0 1px 4px rgba(15, 23, 42, .12); border: 1px solid #e5e7eb;}
    .card h3, .card p, .card li, .card ul, .card div, .card strong {color: #0f172a !important;}
    .card h3 {margin-top: 0; margin-bottom: 8px;}
    .card .muted {color:#475569 !important;}
    .badge {color: white !important; padding: 5px 9px; border-radius: 999px; font-weight: 700; display: inline-block;}
    .small-note {color:#475569 !important; font-size:.92rem;}
    .step-pill {background:#eff6ff; color:#1e3a8a; border:1px solid #bfdbfe; border-radius:999px; padding:5px 10px; font-weight:600; display:inline-block; margin-right:6px; margin-bottom:6px;}
    .step-pill.off {background:#f8fafc;color:#64748b;border-color:#e2e8f0;}
    .step-pill.done {background:#ecfdf5;color:#047857;border-color:#a7f3d0;}
    .importance-card {background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; padding:14px 16px; margin:10px 0; box-shadow:0 1px 3px rgba(15,23,42,.08);}
    .importance-help {background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:10px 12px; color:#334155; margin:8px 0 12px 0;}
    .solution-grid {display:grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin-top: 10px;}
    .justification-table {width:100%; border-collapse: collapse; margin-top: 10px; background:#ffffff; color:#0f172a;}
    .justification-table th {background:#e0f2fe; color:#0f172a; text-align:left; padding:8px; border:1px solid #cbd5e1;}
    .justification-table td {padding:8px; border:1px solid #cbd5e1; vertical-align:top; color:#0f172a;}
    @media (max-width: 1000px) {.solution-grid {grid-template-columns: 1fr;}}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
    <h1>Système expert interactif – Diagnostic holding familiale</h1>
    <p>Chaque étape est validée explicitement. Les résultats utilisent uniquement les réponses validées.</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("Navigation")
    page = st.session_state.get("app_page", VIEWS[0])
    for view in VIEWS:
        active = view == page
        label = f"● {view}" if active else view
        if st.button(label, key=f"nav_{safe_key(view)}", use_container_width=True, type="primary" if active else "secondary"):
            st.session_state.app_page = view
            st.rerun()
    st.divider()
    st.write("Étapes validées")
    for i in range(1, 6):
        label = "validée" if i in st.session_state.validated_steps else "non validée"
        st.caption(f"Étape {i} : {label}" + (f" à {st.session_state.last_validation.get(i)}" if i in st.session_state.last_validation else ""))
    st.progress(int(len(st.session_state.validated_steps.intersection({1,2,3,4,5})) / 5 * 100))
    st.divider()
    if st.button("Réinitialiser le diagnostic", type="secondary"):
        reset_app()
        st.rerun()
    st.info("Les réponses saisies ne modifient pas les résultats tant qu’elles n’ont pas été validées.")

# Les résultats sont calculés uniquement avec les réponses validées.
validated_answers = st.session_state.answers
df, evidence = calculate_risks(validated_answers)
detected_df = df[df["Score"] > 0].copy()
zero_df = df[df["Score"] == 0].copy()

if page == "Questionnaire adaptatif":
    st.subheader("Questionnaire adaptatif")
    pills = []
    for i in range(1, 7):
        cls = "step-pill"
        if i in st.session_state.validated_steps:
            cls += " done"
        elif i != st.session_state.current_step:
            cls += " off"
        pills.append(f"<span class='{cls}'>Étape {i}</span>")
    st.markdown(" ".join(pills), unsafe_allow_html=True)
    st.caption("Les résultats se mettent à jour uniquement après validation de l’étape. Cela évite les réponses partielles ou non enregistrées.")

    step = st.session_state.current_step

    if step == 1:
        st.markdown("### Étape 1 – Cadrage du dossier, objectifs et personnalisation")
        st.caption("Cette étape sert à cadrer l’entretien, à personnaliser le rapport et à préciser les objectifs poursuivis. Certaines informations sont utilisées pour l’analyse ; d’autres servent uniquement à rendre le rapport plus contextualisé.")

        st.markdown("#### Identification du dossier")
        c1, c2, c3 = st.columns(3)
        with c1:
            text_input_field("Nom du dossier / client", "client_name", "Ex. Famille Martin")
            number_input("Âge du dirigeant", "client_age", min_value=0, max_value=100, step=1)
        with c2:
            text_input_field("Nom de l’entreprise", "company_name", "Ex. Martin Industrie")
            selectbox("Forme juridique de la société", "company_form", [UNKNOWN, "SAS", "SARL", "SA", "SNC", "Entreprise individuelle", "Autre"])
        with c3:
            text_input_field("Activité / secteur", "company_activity", "Ex. industrie, bâtiment, conseil...")
            text_input_field("Nom du CGP / cabinet", "cgp_name", "Ex. Cabinet Dupont")

        c4, c5 = st.columns(2)
        with c4:
            text_input_field("Date de l’entretien", "entretien_date", "Ex. 16/06/2026")
        with c5:
            selectbox("Qualité des informations recueillies", "qualite_information", [UNKNOWN, "Confirmées", "Partiellement confirmées", "À vérifier"])

        st.markdown("#### Maturité du projet")
        c1, c2, c3 = st.columns(3)
        with c1:
            selectbox("Niveau de maturité du projet", "maturite_projet", [UNKNOWN, "Simple réflexion", "Projet envisagé", "Transmission prévue", "Transmission urgente"])
        with c2:
            selectbox("Horizon de transmission envisagé", "delai_transmission", [UNKNOWN, "Moins de 12 mois", "1 à 3 ans", "Plus de 3 ans", "Pas défini"])
        with c3:
            yes_no_unknown("Existe-t-il un événement d’urgence ou de fragilité ?", "urgence_evenement")

        st.markdown("#### Objectifs poursuivis")
        sync_widget_from_draft("objectifs")
        st.multiselect(
            "Quels objectifs le dirigeant poursuit-il ?",
            OBJECTIVE_DISPLAY_ORDER,
            key="w_objectifs",
            placeholder="Sélectionner un ou plusieurs objectifs",
            on_change=sync_draft_key,
            args=("objectifs",),
        )
        selected_objectives = list(get_draft("objectifs") or [])
        if selected_objectives:
            st.markdown("#### Importance des objectifs")
            st.caption("Sélectionne le niveau d’importance de chaque objectif. Le classement des risques est renforcé lorsque l’objectif est jugé très important ou prioritaire.")
            st.markdown(
                """
                <div class="importance-help">
                    <strong>Important</strong> : pondération normale ·
                    <strong>Très important</strong> : renforcement modéré ·
                    <strong>Prioritaire</strong> : renforcement fort des risques associés
                </div>
                """,
                unsafe_allow_html=True,
            )
            for objective in selected_objectives:
                st.markdown('<div class="importance-card">', unsafe_allow_html=True)
                objective_weight_buttons(objective)
                st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.warning("Aucun risque ne sera activé tant qu’au moins un objectif n’aura pas été sélectionné et validé.")

        st.markdown("#### Personnalisation du rapport")
        c1, c2 = st.columns(2)
        with c1:
            selectbox("Orientation souhaitée du rapport", "rapport_orientation", ["Équilibré", "Très pédagogique", "Synthétique et décisionnel", "Approfondi et technique"])
            st.caption("Le choix modifie réellement le rapport exporté : longueur, vocabulaire, niveau d’explication et place du scoring.")
        with c2:
            selectbox("Niveau de détail attendu", "niveau_detail", ["Synthétique", "Détaillé", "Très détaillé"])
        text_area_field("Objectif exprimé par le client avec ses propres mots", "objectif_libre", "Ex. transmettre progressivement à mon fils sans léser mes deux autres enfants")
        text_area_field("Attentes particulières du client", "attentes_client", "Ex. éviter les conflits familiaux, protéger le conjoint, conserver le contrôle...")
        text_area_field("Contraintes ou préférences exprimées", "contraintes_client", "Ex. refus d’ouvrir le capital, souhait de simplicité, besoin de revenus...")
        text_area_field("Personnes à associer à la réflexion", "personnes_a_associer", "Ex. conjoint, enfants, notaire, expert-comptable, avocat...")
        text_area_field("Observations libres du CGP", "observations", "Notes utiles pour contextualiser le rapport", height=110)
        validate_buttons(step)

    elif step == 2:
        st.markdown("### Étape 2 – Diagnostic familial")
        objectifs = selected_objectives_from_draft()
        c1, c2, c3 = st.columns(3)
        with c1:
            number_input("Nombre d’enfants", "nb_enfants", min_value=0, max_value=12, step=1)
        with c2:
            yes_no_unknown("Conjoint présent ?", "conjoint_present")
        with c3:
            yes_no_unknown("Famille recomposée ?", "famille_recomposee")

        if get_draft("conjoint_present") == YES:
            yes_no_unknown("Le conjoint adhère-t-il au principe de transmission envisagé ?", "accord_conjoint", [UNKNOWN, YES, NO, UNCERTAIN])

        if get_draft("nb_enfants") > 0 and "Transmettre l’entreprise" in objectifs:
            st.markdown("#### Questions activées : transmission familiale")
            yes_no_unknown("Un héritier repreneur est-il identifié ?", "heritier_repreneur", [UNKNOWN, YES, NO, UNCERTAIN])
            if get_draft("heritier_repreneur") == YES and get_draft("nb_enfants") >= 2:
                yes_no_unknown("Les autres héritiers sont-ils aussi impliqués dans l’entreprise ?", "autres_heritiers_actifs")
                selectbox("Quelle est la volonté probable des héritiers non repreneurs ?", "volonte_non_repreneurs", [UNKNOWN, "Rester associés", "Sortir du capital", "Recevoir principalement une compensation", "Incertain / non abordé"])
                yes_no_unknown("Une soulte (compensation financière) est-elle envisagée pour compenser les héritiers non repreneurs ?", "soulte_envisagee")
                if get_draft("soulte_envisagee") == YES:
                    yes_no_unknown("La capacité de financement de cette soulte est-elle validée ?", "capacite_financement_soulte")
        else:
            st.info("Les questions relatives au repreneur apparaissent uniquement si l’objectif « Transmettre l’entreprise » est sélectionné et si au moins un enfant est renseigné.")

        if get_draft("nb_enfants") >= 2:
            st.markdown("#### Dialogue familial")
            yes_no_unknown("Un dialogue familial a-t-il déjà été organisé ?", "dialogue_familial")

        if get_draft("conjoint_present") == YES:
            st.markdown("#### Protection du conjoint et des proches")
            st.caption("Ces questions sont prises en compte dans l’analyse dès qu’un conjoint est présent, même si l’objectif de protection n’a pas été sélectionné comme prioritaire.")
            yes_no_unknown("Le conjoint dépend-il financièrement du dirigeant ou de l’entreprise ?", "conjoint_dependant")
            yes_no_unknown("Un dispositif de protection du conjoint est-il déjà prévu ?", "protection_conjoint_prevue")
            yes_no_unknown("Le régime matrimonial a-t-il été analysé ou adapté ?", "regime_matrimonial_adapte")

        if get_draft("nb_enfants") >= 2 or get_draft("famille_recomposee") == YES:
            st.markdown("#### Questions activées : sécurisation successorale")
            yes_no_unknown("Une valorisation indépendante des titres a-t-elle été réalisée ou prévue ?", "valorisation_independante")
            yes_no_unknown("Un audit civil et successoral a-t-il été réalisé avec un notaire ?", "audit_civil")
        validate_buttons(step)

    elif step == 3:
        st.markdown("### Étape 3 – Diagnostic patrimonial")
        objectifs = selected_objectives_from_draft()
        c1, c2, c3 = st.columns(3)
        with c1:
            number_input("Valeur estimée de l’entreprise (€)", "valeur_entreprise", min_value=0, step=100_000)
        with c2:
            slider("Poids estimé de l’entreprise dans le patrimoine familial (%)", "poids_entreprise", 0, 100)
        with c3:
            selectbox("Niveau d’actifs liquides hors entreprise", "actifs_liquides", [UNKNOWN, "Faible", "Moyen", "Élevé"])

        c4, c5 = st.columns(2)
        with c4:
            selectbox("Niveau d’endettement personnel ou familial", "endettement_familial", [UNKNOWN, "Faible", "Moyen", "Élevé"])
        with c5:
            selectbox("Besoin de revenus réguliers pour la famille", "besoin_revenus_famille", [UNKNOWN, "Faible", "Moyen", "Élevé"])

        show_diversification = get_draft("poids_entreprise") >= 40 or "Diversifier le patrimoine" in objectifs
        show_prevoyance = (
            show_diversification
            or get_draft("conjoint_present") == YES
            or "Protéger le conjoint et les proches" in objectifs
            or get_draft("besoin_revenus_famille") in ["Moyen", "Élevé"]
            or get_draft("urgence_evenement") == YES
        )
        if show_diversification or show_prevoyance:
            st.markdown("#### Questions activées : sécurité patrimoniale")
            if show_diversification:
                yes_no_unknown("Une diversification patrimoniale est-elle déjà organisée ?", "diversification")
            if show_prevoyance:
                yes_no_unknown("Une prévoyance ou une assurance décès est-elle prévue ?", "prevoyance")
        else:
            st.info("Les questions relatives à la diversification et à la prévoyance apparaissent selon le poids de l’entreprise, le besoin de revenus, la présence d’un conjoint ou l’existence d’une fragilité particulière.")
        validate_buttons(step)

    elif step == 4:
        st.markdown("### Étape 4 – Diagnostic professionnel et gouvernance")
        objectifs = selected_objectives_from_draft()
        show_governance = (
            "Conserver le contrôle familial" in objectifs
            or get_draft("nb_enfants") >= 2
            or get_draft("besoin_revenus_famille") in ["Moyen", "Élevé"]
            or "Préserver l’équité entre les héritiers" in objectifs
        )
        if show_governance:
            yes_no_unknown("Des clauses d’entrée, de sortie ou de cession des titres sont-elles prévues ?", "clauses_entree_sortie")
            yes_no_unknown("La gouvernance de la holding est-elle formalisée ?", "gouvernance_formalisee")
            yes_no_unknown("La holding réunira-t-elle des associés actifs et des associés non actifs ?", "associes_actifs_passifs")
            yes_no_unknown("Une politique de distribution de dividendes est-elle déjà définie ?", "politique_dividendes_definie")
        else:
            st.info("Les questions de gouvernance renforcée apparaissent si l’objectif de contrôle, l’équité familiale, le besoin de revenus ou la configuration familiale l’exigent.")

        st.markdown("#### Continuité opérationnelle")
        yes_no_unknown("L’entreprise est-elle fortement dépendante du dirigeant actuel ?", "entreprise_dependante_dirigeant")

        if get_draft("heritier_repreneur") in [YES, UNCERTAIN]:
            st.markdown("#### Questions activées : succession managériale")
            yes_no_unknown("Le successeur est-il préparé à reprendre la direction ?", "successeur_prepare")
            yes_no_unknown("Un calendrier de transmission du pouvoir est-il prévu ?", "calendrier_transmission")
        validate_buttons(step)

    elif step == 5:
        st.markdown("### Étape 5 – Fiscalité, Dutreil et suivi")
        objectifs = selected_objectives_from_draft()
        if "Optimiser la fiscalité" in objectifs or get_draft("valeur_entreprise") >= 1_000_000:
            yes_no_unknown("Une simulation fiscale préalable a-t-elle été réalisée ou prévue ?", "simulation_fiscale")
            yes_no_unknown("Un Pacte Dutreil est-il envisagé ou déjà mis en place ?", "pacte_dutreil")
            if get_draft("pacte_dutreil") == YES:
                st.markdown("#### Questions activées : vigilance Dutreil")
                yes_no_unknown("La qualification de holding animatrice est-elle établie ?", "holding_animatrice", [UNKNOWN, YES, NO, UNCERTAIN])
                yes_no_unknown("Un audit Dutreil a-t-il été réalisé ou programmé ?", "audit_dutreil")
                yes_no_unknown("Un suivi des engagements de conservation est-il prévu ?", "suivi_engagements")
        else:
            st.info("Les questions fiscales détaillées apparaissent si l’objectif fiscal est sélectionné ou si la valeur de l’entreprise atteint 1 M€.")

        st.markdown("#### Suivi de la stratégie")
        yes_no_unknown("Une revue annuelle de la stratégie patrimoniale est-elle prévue ?", "suivi_annuel")
        yes_no_unknown("Un rapport écrit de diagnostic et de recommandation est-il prévu ?", "formalisation_rapport")
        validate_buttons(step)

    elif step == 6:
        st.markdown("### Étape 6 – Synthèse")
        st.info("Cette synthèse est calculée uniquement à partir des réponses validées.")
        gaps = missing_data(validated_answers)
        if gaps:
            st.warning("Informations complémentaires à recueillir :")
            for g in gaps:
                st.write(f"- {g}")
        if detected_df.empty:
            st.success("Aucun risque significatif n’est détecté à partir des réponses actuellement validées.")
        else:
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Risques détectés", len(detected_df))
            c2.metric("Risques critiques", int((detected_df["Niveau"] == "Critique").sum()))
            c3.metric("Score maximal", int(detected_df["Score"].max()))
            c4.metric("Risque principal", detected_df.iloc[0]["Risque"])
            st.dataframe(detected_df[["Objectif", "Risque", "Probabilité", "Gravité", "Score", "Niveau", "Pondération objectif"]], use_container_width=True, hide_index=True)
        c1, c2, _ = st.columns([1.4, 2.2, 4.4])
        with c1:
            if st.button("← Étape précédente", key="questionnaire_step6_prev"):
                st.session_state.current_step = 5
                st.rerun()
        with c2:
            if st.button("Voir les résultats et solutions →", key="questionnaire_to_results", type="primary"):
                st.session_state.app_page = "Résultats et solutions"
                st.rerun()

elif page == "Résultats et solutions":
    st.subheader("Résultats et solutions proposées")
    st.caption("Résultats calculés uniquement à partir des réponses validées. Pour intégrer une modification, retourne dans le questionnaire et valide l’étape concernée.")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Risques détectés", len(detected_df))
    k2.metric("Risques inexistants", len(zero_df))
    k3.metric("Critiques", int((detected_df["Niveau"] == "Critique").sum()) if not detected_df.empty else 0)
    k4.metric("Risques élevés", int((detected_df["Niveau"] == "Élevé").sum()) if not detected_df.empty else 0)

    gaps = missing_data(validated_answers)
    if gaps:
        with st.expander("Informations manquantes pouvant affiner le diagnostic", expanded=True):
            for g in gaps:
                st.write(f"- {g}")

    if not detected_df.empty:
        st.markdown("### Matrice des risques")
        chart_df = detected_df[["Risque", "Score"]].sort_values("Score", ascending=True)
        st.bar_chart(chart_df, x="Risque", y="Score", horizontal=True)

    st.markdown("### Classement détaillé")
    st.dataframe(df[["Objectif", "Risque", "Probabilité", "Gravité", "Score", "Niveau", "Pondération objectif", "Signaux retenus"]], use_container_width=True, hide_index=True, column_config={"Signaux retenus": st.column_config.TextColumn(width="large")})

    st.markdown("### Solutions envisagées par risque détecté")
    if detected_df.empty:
        st.success("Aucun risque détecté. Sélectionne les objectifs, complète le questionnaire, puis valide les étapes concernées.")
    else:
        for _, row in detected_df.iterrows():
            definition = RISK_DEFINITIONS[row["Code"]]
            color = PRIORITY_COLORS[row["Niveau"]]
            body = f"""
            <p>{priority_badge(row['Niveau'])} &nbsp; <strong>Score :</strong> {int(row['Score'])} &nbsp; <strong>Probabilité :</strong> {score_to_label(int(row['Probabilité']))} &nbsp; <strong>Gravité :</strong> {score_to_label(int(row['Gravité']))}</p>
            <p><strong>Objectif concerné :</strong> {definition.objectif}</p>
            <p><strong>Pourquoi ce risque est activé :</strong></p>
            {render_list(evidence.get(row['Code'], []))}
            <div class="solution-grid">
              <div><strong>Conséquences possibles</strong>{render_list(definition.consequences)}</div>
              <div><strong>Outils à étudier</strong>{render_list(definition.outils)}</div>
              <div><strong>Actions préventives</strong>{render_list(definition.actions_preventives)}</div>
            </div>
            <p><strong>Pourquoi ces outils sont proposés :</strong></p>
            {render_tool_justifications(row['Code'])}
            <p><strong>Professionnels à mobiliser :</strong> {', '.join(definition.professionnels)}</p>
            """
            card(definition.libelle, body, accent=color)

    st.markdown("### Recommandations hiérarchisées en trois temps")
    plan_df = build_action_plan(df)
    if plan_df.empty:
        st.info("Aucun plan d’action n’est généré tant qu’aucun risque n’est détecté.")
    else:
        for horizon in ["Actions immédiates", "Actions à moyen terme", "Suivi annuel"]:
            horizon_df = plan_df[plan_df["Horizon"] == horizon]
            if not horizon_df.empty:
                with st.expander(horizon, expanded=(horizon == "Actions immédiates")):
                    st.dataframe(
                        horizon_df[["Priorité", "Action", "Risques concernés", "Professionnels"]],
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Action": st.column_config.TextColumn(width="large"),
                            "Professionnels": st.column_config.TextColumn(width="medium"),
                        },
                    )

    with st.expander("Risques inexistants — non détectés au vu des réponses validées"):
        st.dataframe(zero_df[["Objectif", "Risque", "Niveau"]], use_container_width=True, hide_index=True)

    render_view_navigation("Résultats et solutions")

elif page == "Exporter":
    st.subheader("Exporter")
    st.caption("Cette page génère un rapport Word structuré et une matrice Excel lisible à partir des réponses validées, des risques détectés et des solutions envisagées.")

    if len(st.session_state.validated_steps.intersection({1, 2, 3, 4, 5})) == 0:
        st.warning("Aucune étape n’a encore été validée. Le rapport peut être généré, mais il sera vide ou peu exploitable.")

    c1, c2 = st.columns(2)
    with c1:
        try:
            docx_data = create_docx_report(validated_answers.get("client_name") or "", validated_answers, df, evidence)
            st.download_button(
                "Télécharger le rapport",
                data=docx_data,
                file_name="rapport_diagnostic_holding_familiale.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )
        except Exception as exc:
            st.error(f"Export Word indisponible : {exc}")
    with c2:
        try:
            xlsx_data = create_xlsx_matrix(df)
            st.download_button(
                "Télécharger la matrice des risques",
                data=xlsx_data,
                file_name="matrice_risques_holding_familiale.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(f"Export Excel indisponible : {exc}")

    st.markdown("### Contenu du rapport")
    orientation = validated_answers.get("rapport_orientation", "Équilibré")
    st.write(f"Orientation sélectionnée : **{orientation}**")
    st.info("Le contenu du rapport varie selon l’orientation : le format synthétique privilégie les décisions et le top des risques ; le format pédagogique explique les outils et les enjeux avec un vocabulaire plus accessible ; le format technique détaille le scoring, les signaux retenus et les validations professionnelles.")

    plan_preview_df = build_action_plan(df)
    if not plan_preview_df.empty:
        st.markdown("### Aperçu du plan d’action hiérarchisé")
        st.dataframe(
            plan_preview_df,
            use_container_width=True,
            hide_index=True,
            column_config={"Action": st.column_config.TextColumn(width="large")},
        )

    if not detected_df.empty:
        st.markdown("### Aperçu synthétique des risques détectés")
        st.dataframe(
            detected_df[["Objectif", "Risque", "Score", "Niveau", "Outils à étudier", "Justification des outils", "Actions préventives"]],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Outils à étudier": st.column_config.TextColumn(width="large"),
                "Justification des outils": st.column_config.TextColumn(width="large"),
                "Actions préventives": st.column_config.TextColumn(width="large"),
            },
        )
    else:
        st.info("Aucun risque n’est détecté à partir des réponses validées.")

    render_view_navigation("Exporter")

elif page == "Règles de décision":
    st.subheader("Règles de décision")
    st.write("L’algorithme est un système expert transparent : chaque réponse validée ajoute des points à certains risques. Aucun signal = niveau Inexistant.")
    st.code("""
SI une étape n’est pas validée
ALORS ses réponses ne sont pas utilisées dans le scoring.

SI aucun objectif n’est validé
ALORS tous les risques restent à 0.

SI un objectif est pondéré comme « Très important » ou « Prioritaire »
ALORS les risques associés à cet objectif sont renforcés dans le scoring.

SI le contrôle familial est recherché ET plusieurs héritiers existent
ALORS le risque de dilution augmente.

SI un enfant reprend ET les autres héritiers ne sont pas impliqués
ALORS le risque de conflit repreneur / non repreneurs augmente.

SI une soulte est envisagée MAIS son financement n’est pas validé
ALORS le risque de liquidité augmente fortement.

SI un Pacte Dutreil est envisagé ET la holding animatrice est incertaine
ALORS le risque de remise en cause du Dutreil augmente fortement.

SI le conjoint dépend financièrement du dirigeant ET qu’aucune protection n’est prévue
ALORS le risque de fragilisation du conjoint augmente.
    """.strip(), language="text")
    rules_export = {code: {"objectif": d.objectif, "risque": d.libelle, "gravite_base": d.gravite_base, "outils": d.outils, "justification_des_outils": get_tool_justifications(code), "actions_preventives": d.actions_preventives, "professionnels": d.professionnels} for code, d in RISK_DEFINITIONS.items()}
    st.download_button("Télécharger le dictionnaire des risques JSON", data=json.dumps(rules_export, ensure_ascii=False, indent=2).encode("utf-8"), file_name="dictionnaire_risques_holding_familiale.json", mime="application/json")
    st.json(rules_export)

    render_view_navigation("Règles de décision")

elif page == "Debug validation":
    st.subheader("Debug validation : vérifier l’enregistrement des réponses")
    st.write("Cette vue distingue les réponses en cours de saisie des réponses validées et réellement utilisées pour le scoring.")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Réponses en cours de saisie")
        st.json(st.session_state.get("draft_answers", {}))
    with c2:
        st.markdown("### Réponses validées utilisées pour les résultats")
        st.json(validated_answers)
    st.markdown("### Étapes validées")
    st.write(sorted(list(st.session_state.validated_steps)))
    st.markdown("### Scores calculés sur réponses validées")
    st.dataframe(df[["Risque", "Probabilité", "Gravité", "Score", "Niveau", "Signaux retenus"]], use_container_width=True, hide_index=True, column_config={"Signaux retenus": st.column_config.TextColumn(width="large")})


if page == "Debug validation":
    render_view_navigation("Debug validation")
