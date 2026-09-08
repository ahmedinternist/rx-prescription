"""Major therapeutic-group taxonomy for browsing the local drug database.

Classes organise medicines; they do not provide prescribing, dosing, or safety
advice.  Explicit database metadata always takes precedence over the compact
built-in mapping used for the supplied seed database.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class DrugClass:
    code: str
    detail: str


GROUPS = (
    "gastrointestinal", "endocrine_nutrition", "cardiovascular_blood",
    "blood", "antiinfectives", "pain_musculoskeletal", "neuro_mental_health",
    "respiratory_allergy_ent", "skin_eye_ear", "genitourinary_reproductive",
    "pediatric_fluids",
)

SUBCLASSES = {
    "gastrointestinal": (
        "PPI (Proton Pump Inhibitor)",
        "H2-Receptor Antagonist (H2RA)",
        "Antacid & Mucosal Protectant",
        "Antiflatulent & Digestive Enzymes",
        "Antispasmodic & IBS Agent",
        "Antiemetic & Prokinetic",
        "Laxative: Osmotic & Bulking",
        "Laxative: Stimulant & Stool Softener",
        "Antidiarrheal & Motility Inhibitor",
        "Probiotic, Prebiotic & ORS",
        "Intestinal Anti-inflammatory (IBD)",
    ),
    "endocrine_nutrition": (
        "Thyroid Replacement Hormone",
        "Antithyroid Agent (Thionamides)",
        "Systemic Glucocorticoids",
        "Antidiabetic: Biguanides",
        "Antidiabetic: Sulfonylureas",
        "Antidiabetic: DPP-4 Inhibitors (Gliptins)",
        "Antidiabetic: SGLT2 Inhibitors (Gliflozins)",
        "Antidiabetic: GLP-1 Receptor Agonists",
        "Antidiabetic: Thiazolidinediones (TZD)",
        "Insulin: Rapid & Short-Acting",
        "Insulin: Intermediate & Long-Acting (Basal)",
        "Vitamins & Mineral Supplements",
    ),
    "cardiovascular_blood": (
        "ACE Inhibitor (ACEI)",
        "ARB (Angiotensin Receptor Blocker)",
        "ARNI (Angiotensin Receptor Neprilysin Inhibitor)",
        "CCB: Dihydropyridine (Peripheral Vasodilator)",
        "CCB: Non-Dihydropyridine (Rate-Limiting)",
        "Beta-Blocker: Cardioselective (Beta-1)",
        "Beta-Blocker: Non-Selective / Alpha-Beta",
        "Diuretic: Thiazide & Thiazide-Like",
        "Diuretic: Loop Diuretic",
        "Diuretic: MRA (Potassium-Sparing)",
        "Lipid-Lowering: HMG-CoA Reductase Inhibitor (Statin)",
        "Lipid-Lowering: Cholesterol Absorption Inhibitor",
        "Lipid-Lowering: Fibrate",
        "Nitrates & Direct Vasodilators",
        "Cardiac Glycosides & Antiarrhythmics",
    ),
    "blood": (
        "Antiplatelet Agent (Cyclooxygenase / ADP)",
        "Anticoagulant: DOAC",
        "Anticoagulant: Parenteral (Heparins)",
        "Anticoagulant: Vitamin K Antagonist (VKA)",
        "Hemostatic & Antifibrinolytic",
        "Antianemic Agent (Iron & Erythropoietin)",
    ),
    "antiinfectives": (
        "Aminopenicillins & Beta-Lactamase Inhibitors",
        "Antipseudomonal & Penicillinase-Resistant",
        "Cephalosporins: 1st & 2nd Generation",
        "Cephalosporins: 3rd & 4th Generation",
        "Carbapenems & Monobactams",
        "Macrolides & Ketolides",
        "Fluoroquinolones",
        "Tetracyclines & Glycylcyclines",
        "Aminoglycosides",
        "Glycopeptides & Lipoglycopeptides",
        "Nitroimidazoles & Antiseptics",
        "Systemic Antifungal: Azoles & Echinocandins",
        "Antiviral: Anti-Herpetic & Anti-Influenza",
        "Antiviral: Direct-Acting Anti-HCV & Antiretroviral",
        "Anthelmintics & Scabicides",
        "Antimalarial Chemotherapy",
    ),
    "pain_musculoskeletal": (
        "NSAID: Non-Selective",
        "NSAID: COX-2 Selective Inhibitor (Coxibs)",
        "Muscle Relaxants",
        "Anti-Gout: Acute & Urate-Lowering (ULT)",
        "Bisphosphonates & Bone Metabolism Agents",
        "Analgesic & Antipyretic (Non-Opioid)",
        "Opioid Analgesic:",
    ),
    "neuro_mental_health": (
        "Benzodiazepines & Z-Drugs",
        "Antipsychotic: Second Generation (Atypical)",
        "SSRI",
        "SNRI & NaSSA",
        "Tricyclic",
        "Dopaminergics & Cognitive Enhancers",
    ),
    "respiratory_allergy_ent": (
        "Nasal Decongestant & Saline Wash",
        "Intranasal Corticosteroids (INCS)",
        "Antihistamines: 2nd Generation (Non-Sedating)",
        "Antihistamines: 1st Gen & Antivertigo",
        "SABA (Short-Acting Beta-2 Agonist)",
        "LABA (Long-Acting Beta-2 Agonist)",
        "ICS (Inhaled Corticosteroids)",
        "Anticholinergic Bronchodilators (SAMA & LAMA)",
        "Leukotriene Receptor Antagonists (LTRA)",
        "Cough Suppressants & Antitussives",
        "Mucolytics & Expectorants",
    ),
    "skin_eye_ear": (
        "Topical Antifungal & Antibacterial",
        "Topical Corticosteroid (Mild to Superpotent)",
        "Topical Retinoid, Benzoyl Peroxide & Azelaic Acid",
        "Emollient, Barrier Protectant & Cicatrizant",
        "Ophthalmic Antibacterial & Antiviral",
        "Ophthalmic Anti-Allergic & Artificial Tears",
        "Antiglaucoma Preparations (Prostaglandin & Beta-blocker)",
        "Otic Analgesic, Antibiotic & Ceruminolytic",
    ),
    "genitourinary_reproductive": (
        "BPH Agent: Alpha-1 Blocker",
        "BPH Agent: 5-Alpha Reductase Inhibitor (5-ARI)",
        "Erectile Dysfunction (PDE-5 Inhibitor)",
        "Urinary Antispasmodic / Overactive Bladder (OAB)",
        "Combined Oral Contraceptive (COC)",
        "Progestin-Only Formulation (POP & Implant/Depot)",
        "Emergency Contraceptive",
        "Vaginal Antifungal & Antimicrobial",
    ),
}

# Seed-database mapping. New or imported data can supply therapeutic_group and
# detailed_class columns instead; they take precedence in classify().
_NAMES = {
    "amoxicillin": ("antiinfectives", "Aminopenicillins & Beta-Lactamase Inhibitors"),
    "amoxicillin/clavulanate": ("antiinfectives", "Aminopenicillins & Beta-Lactamase Inhibitors"),
    "azithromycin": ("antiinfectives", "Macrolides & Ketolides"),
    "ciprofloxacin": ("antiinfectives", "Fluoroquinolones"),
    "ceftriaxone": ("antiinfectives", "Cephalosporins: 3rd & 4th Generation"),
    "metronidazole": ("antiinfectives", "Nitroimidazoles & Antiseptics"),
    "ibuprofen": ("pain_musculoskeletal", "NSAID: Non-Selective"),
    "diclofenac": ("pain_musculoskeletal", "NSAID: Non-Selective"),
    "paracetamol": ("pain_musculoskeletal", "Analgesic & Antipyretic (Non-Opioid)"),
    "omeprazole": ("gastrointestinal", "PPI (Proton Pump Inhibitor)"),
    "ondansetron": ("gastrointestinal", "Antiemetic & Prokinetic"),
    "cetirizine": ("respiratory_allergy_ent", "Antihistamines: 2nd Generation (Non-Sedating)"),
    "loratadine": ("respiratory_allergy_ent", "Antihistamines: 2nd Generation (Non-Sedating)"),
    "metformin": ("endocrine_nutrition", "Antidiabetic: Biguanides"),
    "insulin glargine": ("endocrine_nutrition", "Insulin: Intermediate & Long-Acting (Basal)"),
    "levothyroxine": ("endocrine_nutrition", "Thyroid Replacement Hormone"),
    "prednisolone": ("endocrine_nutrition", "Systemic Glucocorticoids"),
    "vitamin d3": ("endocrine_nutrition", "Vitamins & Mineral Supplements"),
    "amlodipine": ("cardiovascular_blood", "CCB: Dihydropyridine (Peripheral Vasodilator)"),
    "captopril": ("cardiovascular_blood", "ACE Inhibitor (ACEI)"),
    "losartan": ("cardiovascular_blood", "ARB (Angiotensin Receptor Blocker)"),
    "atorvastatin": ("cardiovascular_blood", "Lipid-Lowering: HMG-CoA Reductase Inhibitor (Statin)"),
    "aspirin": ("blood", "Antiplatelet Agent (Cyclooxygenase / ADP)"),
    "furosemide": ("cardiovascular_blood", "Diuretic: Loop Diuretic"),
    "ferrous sulfate": ("blood", "Antianemic Agent (Iron & Erythropoietin)"),
    "diazepam": ("neuro_mental_health", "Benzodiazepines & Z-Drugs"),
    "salbutamol": ("respiratory_allergy_ent", "SABA (Short-Acting Beta-2 Agonist)"),
}

_CATEGORY_HINTS = {
    "antibiotic": ("antiinfectives", "Aminopenicillins & Beta-Lactamase Inhibitors"),
    "antiprotozoal": ("antiinfectives", "Nitroimidazoles & Antiseptics"),
    "nsaid": ("pain_musculoskeletal", "NSAID: Non-Selective"),
    "analgesic": ("pain_musculoskeletal", "Analgesic & Antipyretic (Non-Opioid)"),
    "ppi": ("gastrointestinal", "PPI (Proton Pump Inhibitor)"),
    "antiemetic": ("gastrointestinal", "Antiemetic & Prokinetic"),
    "antihistamine": ("respiratory_allergy_ent", "Antihistamines: 2nd Generation (Non-Sedating)"),
    "bronchodilator": ("respiratory_allergy_ent", "SABA (Short-Acting Beta-2 Agonist)"),
    "antidiabetic": ("endocrine_nutrition", "Antidiabetic: Biguanides"),
    "corticosteroid": ("endocrine_nutrition", "Systemic Glucocorticoids"),
    "vitamin": ("endocrine_nutrition", "Vitamins & Mineral Supplements"),
    "hormone": ("endocrine_nutrition", "Thyroid Replacement Hormone"),
    "antihypertensive": ("cardiovascular_blood", "CCB: Dihydropyridine (Peripheral Vasodilator)"),
    "ace inhibitor": ("cardiovascular_blood", "ACE Inhibitor (ACEI)"),
    "arb": ("cardiovascular_blood", "ARB (Angiotensin Receptor Blocker)"),
    "statin": ("cardiovascular_blood", "Lipid-Lowering: HMG-CoA Reductase Inhibitor (Statin)"),
    "antiplatelet": ("blood", "Antiplatelet Agent (Cyclooxygenase / ADP)"),
    "diuretic": ("cardiovascular_blood", "Diuretic: Loop Diuretic"),
    "iron supplement": ("blood", "Antianemic Agent (Iron & Erythropoietin)"),
    "benzodiazepine": ("neuro_mental_health", "Benzodiazepines & Z-Drugs"),
}


def classify(name: str, category: str = "", therapeutic_group: str = "",
             detailed_class: str = "") -> Optional[DrugClass]:
    """Return a group/class for a medicine, or None when it is not mapped."""
    explicit_group = therapeutic_group.strip()
    if explicit_group:
        return DrugClass(explicit_group, detailed_class.strip() or category.strip() or "Class not specified")
    mapped = _NAMES.get(name.casefold().strip())
    if mapped:
        return DrugClass(*mapped)
    mapped = _CATEGORY_HINTS.get(category.casefold().strip())
    return DrugClass(*mapped) if mapped else None


def group_for(drug) -> Optional[DrugClass]:
    """Classify a drug_db.Drug without importing drug_db (avoids a cycle)."""
    return classify(drug.generic_name, drug.category, drug.therapeutic_group, drug.detailed_class)


def subclasses_for(group_code: str) -> tuple[str, ...]:
    return SUBCLASSES.get(group_code, ())


def all_subclasses() -> tuple[tuple[str, str], ...]:
    """Return detailed classes in the same stable order as the dashboard."""
    return tuple((group, detail) for group in GROUPS for detail in subclasses_for(group))
