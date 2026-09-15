"""Major therapeutic-group taxonomy for browsing the local drug database.

Classes organise medicines; they do not provide prescribing, dosing, or safety
advice.  Explicit database metadata always takes precedence over the compact
built-in mapping used for the supplied seed database.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional
import unicodedata


@dataclass(frozen=True)
class DrugClass:
    code: str
    detail: str
    confidence: str = "confirmed"


GROUPS = (
    "gastrointestinal", "endocrine_nutrition", "cardiovascular_blood",
    "blood", "antiinfectives", "pain_musculoskeletal", "neuro_mental_health",
    "respiratory_allergy_ent", "skin_eye_ear", "genitourinary_reproductive",
    "pediatric_preparations", "iv_fluids_devices",
)

# Older app builds stored these combined group codes in imported databases and
# favorites.  Keep them readable while presenting the new, clearer groups.
GROUP_ALIASES = {
    "pediatric_fluids": "pediatric_preparations",
}

# English labels used by the class browser. Import files commonly contain a
# visible label rather than the internal code, so keep one canonical resolver
# here instead of letting arbitrary group text enter the database.
GROUP_LABELS = {
    "gastrointestinal": "Gastrointestinal",
    "endocrine_nutrition": "Endocrine, diabetes & nutrition",
    "cardiovascular_blood": "Cardiovascular",
    "blood": "Blood",
    "antiinfectives": "Anti-infectives",
    "pain_musculoskeletal": "Analgesic & musculoskeletal",
    "neuro_mental_health": "Neurology & mental health",
    "respiratory_allergy_ent": "Respiratory, allergy & ENT",
    "skin_eye_ear": "Skin, eye & ear",
    "genitourinary_reproductive": "Genitourinary & reproductive",
    "pediatric_preparations": "Pediatric preparations",
    "iv_fluids_devices": "IV fluids & devices",
}


def _classification_key(value: str) -> str:
    """Return a punctuation-insensitive key for imported classification text."""
    value = unicodedata.normalize("NFKD", str(value or "")).casefold()
    value = value.replace("&", " and ")
    return " ".join("".join(
        character if character.isalnum() else " " for character in value
    ).split())


_GROUP_IMPORT_ALIASES = {
    "anti infective": "antiinfectives",
    "anti infectives": "antiinfectives",
    "ant infective": "antiinfectives",
    "analgesia and musculoskeletal": "pain_musculoskeletal",
    "analgesic and musculoskeletal": "pain_musculoskeletal",
    "pain and musculoskeletal": "pain_musculoskeletal",
    "neurology": "neuro_mental_health",
    "respiratory and allergy": "respiratory_allergy_ent",
    "skin": "skin_eye_ear",
    "genitourinary": "genitourinary_reproductive",
    "pediatric": "pediatric_preparations",
    "iv fluid": "iv_fluids_devices",
    "iv fluids": "iv_fluids_devices",
}


def resolve_group_code(value: str) -> str:
    """Resolve an imported code or visible English label to a valid group code."""
    raw = str(value or "").strip()
    if not raw:
        return ""
    raw = GROUP_ALIASES.get(raw, raw)
    if raw in GROUPS:
        return raw
    key = _classification_key(raw)
    candidates = {_classification_key(code): code for code in GROUPS}
    candidates.update({
        _classification_key(label): code for code, label in GROUP_LABELS.items()
    })
    candidates.update(_GROUP_IMPORT_ALIASES)
    return candidates.get(key, "")

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
        "Bile Acid Sequestrant",
        "Herbal",
        "Hemorrhoid and Fissure",
        "Other",
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
        "Dopamine Receptor Antagonist",
        "Obesity Drugs",
        "Other",
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
        "Central Alpha-2 Agonist",
        "Venotonic & Vasoprotective",
        "Carbonic Anhydrase Inhibitor",
        "Other",
    ),
    "blood": (
        "Antiplatelet Agent (Cyclooxygenase / ADP)",
        "Anticoagulant: DOAC",
        "Anticoagulant: Parenteral (Heparins)",
        "Anticoagulant: Vitamin K Antagonist (VKA)",
        "Hemostatic & Antifibrinolytic",
        "Antianemic Agent (Iron & Erythropoietin)",
        "Other",
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
        "Other",
    ),
    "pain_musculoskeletal": (
        "NSAID: Non-Selective",
        "NSAID: COX-2 Selective Inhibitor (Coxibs)",
        "Muscle Relaxants",
        "Anti-Gout: Acute & Urate-Lowering (ULT)",
        "Bisphosphonates & Bone Metabolism Agents",
        "Analgesic & Antipyretic (Non-Opioid)",
        "Opioid Analgesic:",
        "DMARD",
        "Joint Supplement",
        "Topical Analgesics",
        "Other",
    ),
    "neuro_mental_health": (
        "Benzodiazepines & Z-Drugs",
        "Antipsychotic: Second Generation (Atypical)",
        "SSRI",
        "SNRI & NaSSA",
        "Tricyclic",
        "Dopaminergics & Cognitive Enhancers",
        "Antiepileptic",
        "Vitamins & Supplements",
        "AntiMigraine",
        "Other",
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
        "Theophylline (Methylxanthine)",
        "Antifibrotic",
        "Other",
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
        "Oral Retinoid",
        "Scabicidal",
        "Hair Tonics",
        "Other",
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
        "Vitamins & Mineral Supplements",
        "Chemolytic",
        "Sex Hormones",
        "Other",
    ),
    "pediatric_preparations": (
        "Pediatric Preparations",
    ),
    "iv_fluids_devices": (
        "IV Fluids & Devices",
    ),
}


def resolve_detailed_class(group_code: str, value: str) -> str:
    """Return the exact configured class name for imported class text."""
    group_code = resolve_group_code(group_code)
    key = _classification_key(value)
    if not group_code or not key:
        return ""
    return next((detail for detail in SUBCLASSES.get(group_code, ())
                 if _classification_key(detail) == key), "")


def group_for_detailed_class(value: str) -> tuple[str, str]:
    """Resolve a detailed class when it uniquely identifies its major group."""
    key = _classification_key(value)
    if not key:
        return "", ""
    matches = [(group, detail) for group, details in SUBCLASSES.items()
               for detail in details if _classification_key(detail) == key]
    return matches[0] if len(matches) == 1 else ("", "")

# Every class picker and class-browser page uses the same predictable A-Z order.
SUBCLASSES = {
    code: tuple(sorted(dict.fromkeys(details), key=str.casefold))
    for code, details in SUBCLASSES.items()
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
    supplied_group = therapeutic_group.strip()
    legacy_detail = ""
    if supplied_group == "pediatric_fluids":
        combined_text = " ".join((name, category, detailed_class)).casefold()
        explicit_group = (
            "iv_fluids_devices"
            if any(token in combined_text for token in ("iv ", "fluid", "device", "cannula"))
            else "pediatric_preparations"
        )
        legacy_detail = SUBCLASSES[explicit_group][0]
    else:
        explicit_group = resolve_group_code(supplied_group)
    if supplied_group:
        if not explicit_group:
            return None
        explicit_detail = (
            resolve_detailed_class(explicit_group, detailed_class)
            or resolve_detailed_class(explicit_group, category)
            or legacy_detail
        )
        if not explicit_detail:
            return None
        return DrugClass(explicit_group, explicit_detail)
    mapped = _NAMES.get(name.casefold().strip())
    if mapped:
        return DrugClass(*mapped, confidence="suggested")
    mapped = _CATEGORY_HINTS.get(category.casefold().strip())
    return DrugClass(*mapped, confidence="suggested") if mapped else None


def groups_for(drug) -> tuple[DrugClass, ...]:
    """Return every explicit mapping, or one built-in suggestion when unmapped."""
    if str(getattr(drug, "mapping_status", "") or "").casefold() == "unrecognized":
        return ()
    pairs: list[tuple[str, str]] = []
    for item in str(getattr(drug, "class_mappings", "") or "").split("|"):
        group, separator, detail = item.partition("::")
        group, detail = group.strip(), detail.strip()
        if not separator or not group or not detail:
            continue
        normalized = classify(
            getattr(drug, "generic_name", ""), getattr(drug, "category", ""),
            group, detail)
        pair = (normalized.code, normalized.detail) if normalized else (group, detail)
        if pair not in pairs:
            pairs.append(pair)
    legacy_group = str(getattr(drug, "therapeutic_group", "") or "").strip()
    legacy_detail = str(getattr(drug, "detailed_class", "") or "").strip()
    if legacy_group:
        normalized = classify(
            getattr(drug, "generic_name", ""), getattr(drug, "category", ""),
            legacy_group, legacy_detail)
        if normalized:
            pair = (normalized.code, normalized.detail)
            if pair not in pairs:
                pairs.insert(0, pair)
    if pairs:
        status = str(getattr(drug, "mapping_status", "") or "confirmed").casefold()
        confidence = status if status in {"confirmed", "suggested"} else "confirmed"
        return tuple(DrugClass(group, detail, confidence) for group, detail in pairs)
    suggested = classify(
        getattr(drug, "generic_name", ""), getattr(drug, "category", ""))
    return (suggested,) if suggested else ()


def group_for(drug) -> Optional[DrugClass]:
    """Classify a drug_db.Drug without importing drug_db (avoids a cycle)."""
    found = groups_for(drug)
    return found[0] if found else None


def subclasses_for(group_code: str) -> tuple[str, ...]:
    return SUBCLASSES.get(group_code, ())


def all_subclasses() -> tuple[tuple[str, str], ...]:
    """Return detailed classes in the same stable order as the dashboard."""
    return tuple((group, detail) for group in GROUPS for detail in subclasses_for(group))
