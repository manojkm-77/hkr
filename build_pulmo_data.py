"""
Script to extract and build all structured JSON datasets for PulmoDoc Suite.
"""
import os, sys, json, re
import docx

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(ROOT, "pulmo_app", "data")
os.makedirs(DATA_DIR, exist_ok=True)

# ==========================================
# 1. BUILD RADIOLOGY CASES
# ==========================================
def build_radiology():
    print("Building Radiology Atlas...")
    cases = []
    
    # Mapping of sections to image files
    atlas_doc_path = os.path.join(ROOT, "Thoracic_Lung_Imaging_Atlas_Expanded.docx")
    doc = docx.Document(atlas_doc_path)
    
    # Read paragraphs
    sections = []
    curr_sec = None
    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue
        if p.style.name.startswith("Heading 2"):
            if curr_sec:
                sections.append(curr_sec)
            curr_sec = {"title": txt, "paras": []}
        elif curr_sec:
            curr_sec["paras"].append(txt)
    if curr_sec:
        sections.append(curr_sec)
        
    fig_map = {
        "Air bronchogram and air-space consolidation": "fig1.jpg",
        "Air-crescent sign and cavitation": "fig2.jpg",
        "Arcade-like / perilobular pattern": "fig3.jpg",
        "Atoll / reverse-halo sign": "fig4.jpg",
        "Cheerios sign": "fig5.jpg",
        "Crazy-paving pattern": "fig6.jpg",
        "Comet-tail sign / rounded atelectasis": "fig7.jpg",
        "Feeding-vessel sign": "fig11.jpg",
        "Finger-in-glove sign / mucoid impaction": "fig12.jpg",
        "Knuckle sign / abrupt arterial cut-off": "fig14.jpg",
        "Subcutaneous emphysema / ginkgo-leaf pattern": "fig16.jpg",
        "Golden S sign": "fig17.jpg",
        "Halo sign": "fig18.jpg",
        "Head-cheese / three-density pattern": "fig19.jpg",
        "Honeycombing and fibrotic UIP pattern": "fig20.jpg",
        "Interface / jagged pleural sign": "fig21.jpg",
        "Mosaic attenuation and air-trapping": "fig23.jpg",
        "Oreo-cookie sign (pericardial effusion)": "fig24.jpg",
        "Polo-mint sign (axial CTPA)": "fig25.jpg",
        "Popcorn calcification and hamartoma": "fig26.jpg",
        "Positive bronchus sign": "fig27.jpg",
        "Railway-track sign (long-axis CTPA)": "fig28.jpg",
        "Signet-ring sign / bronchiectasis": "fig30.jpg",
        "Snowstorm / diffuse micronodules": "fig31.jpg",
        "Spiculated nodule / sunburst pattern": "fig32.jpg",
        "Tree-in-bud pattern": "fig33.jpg",
        "Monod sign: mobile intracavitary fungal ball": "extra_monod_sign__mobile_intracavitary_fungal_ball.jpg",
        "Water-lily sign: ruptured pulmonary hydatid cyst": "extra_water_lily_sign__ruptured_pulmonary_hydatid_cyst.jpg",
        "CT and MRI comparison: pulmonary nodule": "extra_ct_and_mri_comparison__pulmonary_nodule.jpg",
        "FDG PET/CT: pulmonary sarcoidosis mimicking metastases": "extra_fdg_pet_ct__pulmonary_sarcoidosis_mimicking_metastases.jpg",
        "Lymphangioleiomyomatosis (LAM): diffuse cystic lung disease": "extra_lymphangioleiomyomatosis__lam___diffuse_cystic_lung_dis.jpg",
        "Pulmonary alveolar proteinosis (PAP)": "extra_pulmonary_alveolar_proteinosis__pap_.jpg",
        "DIPNECH: multiple nodules with mosaic perfusion": "extra_dipnech__multiple_nodules_with_mosaic_perfusion.jpg",
        "Nodular pulmonary amyloidosis": "extra_nodular_pulmonary_amyloidosis.jpg",
        "Silicosis: nodules and eggshell lymph-node calcification": "extra_silicosis__nodules_and_eggshell_lymph_node_calcificatio.jpg",
        "Pulmonary lymphangitic carcinomatosis": "extra_pulmonary_lymphangitic_carcinomatosis.jpg",
        "Acute eosinophilic pneumonia": "extra_acute_eosinophilic_pneumonia.jpg",
        "Pulmonary Langerhans cell histiocytosis (PLCH)": "extra_pulmonary_langerhans_cell_histiocytosis__plch_.jpg"
    }

    category_map = {
        "Air bronchogram": "Consolidation & Infection",
        "Air-crescent sign": "Cavitation & Fungal",
        "Arcade-like": "Interstitial & Organizing",
        "Atoll / reverse-halo": "Infiltrative & Fungal",
        "Cheerios": "Nodules & Neoplasm",
        "Crazy-paving": "Diffuse Parenchymal",
        "Comet-tail": "Pleural & Atelectasis",
        "Feeding-vessel": "Vascular & Embolic",
        "Finger-in-glove": "Airways & Bronchiectasis",
        "Knuckle sign": "Vascular & Embolic",
        "Subcutaneous emphysema": "Chest Wall & Pleural",
        "Golden S sign": "Airway Obstruction & Neoplasm",
        "Halo sign": "Infection & Hemorrhage",
        "Head-cheese": "Hypersensitivity & Infiltrative",
        "Honeycombing": "Fibrotic ILD (UIP)",
        "Interface / jagged": "Fibrotic ILD",
        "Mosaic attenuation": "Airway / Vascular",
        "Oreo-cookie": "Mediastinum & Pleural",
        "Polo-mint": "Vascular (PE)",
        "Popcorn calcification": "Benign Nodules",
        "Positive bronchus": "Neoplasm & Biopsy",
        "Railway-track": "Vascular (PE)",
        "Signet-ring": "Airways & Bronchiectasis",
        "Snowstorm": "Miliary & Micronodular",
        "Spiculated nodule": "Thoracic Malignancy",
        "Tree-in-bud": "Endobronchial Infection",
        "Monod sign": "Cavitation & Fungal",
        "Water-lily sign": "Parasitic / Hydatid",
        "CT and MRI comparison": "Solitary Pulmonary Nodule",
        "FDG PET/CT": "Sarcoidosis & Granulomatous",
        "Lymphangioleiomyomatosis": "Cystic Lung Diseases",
        "Pulmonary alveolar proteinosis": "Diffuse Alveolar Infiltrate",
        "DIPNECH": "Neuroendocrine & Small Airway",
        "Nodular pulmonary amyloidosis": "Infiltrative Lung Diseases",
        "Silicosis": "Occupational Pneumoconiosis",
        "Pulmonary lymphangitic carcinomatosis": "Thoracic Malignancy",
        "Acute eosinophilic pneumonia": "Eosinophilic Lung Diseases",
        "Pulmonary Langerhans": "Smoking-Related & Cystic"
    }

    case_id = 1
    for sec in sections:
        title = sec["title"]
        img_name = fig_map.get(title)
        
        # If not exact match, search keys
        if not img_name:
            for k, v in fig_map.items():
                if k.lower() in title.lower() or title.lower() in k.lower():
                    img_name = v
                    break
                    
        # Category lookup
        cat = "Thoracic Radiology"
        for k, c in category_map.items():
            if k.lower() in title.lower():
                cat = c
                break
                
        paras = sec["paras"]
        description = paras[0] if len(paras) > 0 else ""
        findings = paras[1] if len(paras) > 1 else ""
        credit = paras[2] if len(paras) > 2 else ""

        # Modality detection
        modality = "HRCT Thorax"
        if "CXR" in title or "radiograph" in description.lower() or "radiograph" in findings.lower():
            modality = "Chest Radiograph (CXR)"
        elif "PET" in title:
            modality = "FDG PET-CT"
        elif "CTPA" in title or "CTPA" in findings:
            modality = "CT Pulmonary Angiography (CTPA)"
        elif "MRI" in title:
            modality = "Thoracic CT & MRI"

        cases.append({
            "id": f"RAD-{case_id:03d}",
            "title": title,
            "category": cat,
            "modality": modality,
            "image": f"/assets/thoracic_atlas_assets/{img_name}" if img_name else "/assets/thoracic_atlas_assets/fig1.jpg",
            "image_filename": img_name or "fig1.jpg",
            "description": description,
            "findings": findings,
            "credit": credit,
            "high_yield_pearl": f"Key differentiator: Note the specific distribution. In {title}, look closely at whether lesions are bronchocentric, centrilobular, subpleural, or along interlobular septa."
        })
        case_id += 1

    # Add Felson classic CXR signs
    felson_signs = [
        {
            "id": f"RAD-{case_id:03d}",
            "title": "Silhouette Sign (Felson's Classic)",
            "category": "Classic CXR Signs",
            "modality": "Chest Radiograph (CXR)",
            "image": "/assets/thoracic_atlas_assets/fig1.jpg",
            "image_filename": "fig1.jpg",
            "description": "Loss of normal radiographic border between two structures of equal radiodensity in anatomic contact. Obliteration of right heart border indicates Right Middle Lobe (RML) pathology; obliteration of left heart border indicates Lingula pathology; loss of hemidiaphragm contour indicates Lower Lobe pathology.",
            "findings": "Loss of sharp cardiac or diaphragmatic silhouette on standard PA view.",
            "credit": "Felson's Principles of Chest Roentgenology",
            "high_yield_pearl": "Silhouette sign localizes consolidation without a lateral view: Right heart border = RML; Left heart border = Lingular segments of LUL; Right dome of diaphragm = RLL; Aortic knuckle = Apicoposterior LUL."
        },
        {
            "id": f"RAD-{case_id+1:03d}",
            "title": "Deep Sulcus Sign (Pneumothorax in Supine Patient)",
            "category": "Classic CXR Signs",
            "modality": "Chest Radiograph (CXR)",
            "image": "/assets/thoracic_atlas_assets/fig16.jpg",
            "image_filename": "fig16.jpg",
            "description": "In a supine trauma or ICU patient, free intrapleural air rises to the anterior-inferior pleural space rather than the apex. This creates an abnormally deepened, lucent, and sharp costophrenic angle.",
            "findings": "Costophrenic sulcus extending far below the level of the diaphragm on the affected side.",
            "credit": "Felson's Principles of Chest Roentgenology / Murray & Nadel",
            "high_yield_pearl": "Never wait for an apical visceral pleural line in supine ICU patients. The deep sulcus sign is the hallmark of tension pneumothorax in supine patients on mechanical ventilation."
        },
        {
            "id": f"RAD-{case_id+2:03d}",
            "title": "Westermark Sign & Hampton's Hump (Pulmonary Embolism)",
            "category": "Vascular & Embolic",
            "modality": "Chest Radiograph (CXR)",
            "image": "/assets/thoracic_atlas_assets/fig14.jpg",
            "image_filename": "fig14.jpg",
            "description": "Westermark sign represents regional oligemia (hyperlucency) distal to an occluded pulmonary artery. Hampton's hump is a wedge-shaped pleural-based opacity with its base abutting the visceral pleura and rounded apex pointing toward the hilum, indicating pulmonary infarction.",
            "findings": "Focal lung oligemia + prominent proximal pulmonary artery (Fleischner sign) and wedge-shaped peripheral infarct opacity.",
            "credit": "Fishman's Pulmonary Diseases / Felson",
            "high_yield_pearl": "CXR in acute pulmonary embolism is often technically 'normal' or shows atelectasis. Classic signs (Westermark, Hampton's hump, Palla sign) are high-specificity, low-sensitivity exam favorites."
        }
    ]
    cases.extend(felson_signs)
    
    with open(os.path.join(DATA_DIR, "radiology_cases.json"), "w", encoding="utf-8") as f:
        json.dump(cases, f, indent=2)
    print(f"Saved {len(cases)} radiology cases.")

# ==========================================
# 2. BUILD CHAPTERS DATA
# ==========================================
def build_chapters():
    print("Building Core Specialty Chapters...")
    docs_dir = os.path.join(ROOT, "hade", "respiratory_medicine_docs")
    
    file_spec = [
        ("01_Pneumonia_50pg.docx", "PNEUMONIA & LOWER RESPIRATORY INFECTIONS", "cap_hap", "Infections"),
        ("02_Tuberculosis.docx", "TUBERCULOSIS & DRUG-RESISTANT TB", "tuberculosis", "Infections"),
        ("03_Asthma_50pg.docx", "BRONCHIAL ASTHMA & PHENOTYPES", "asthma", "Airway Diseases"),
        ("04_COPD.docx", "CHRONIC OBSTRUCTIVE PULMONARY DISEASE (COPD)", "copd", "Airway Diseases"),
        ("05_Lung_Cancer.docx", "THORACIC ONCOLOGY & LUNG CANCER", "lung_cancer", "Thoracic Oncology"),
        ("06_ILD_and_IPF.docx", "DIFFUSE PARENCHYMAL LUNG DISEASE & IPF", "ild_ipf", "Interstitial Lung Diseases"),
        ("07_PE_and_PAH.docx", "PULMONARY EMBOLISM & PULMONARY HYPERTENSION", "pe_pah", "Vascular Diseases"),
        ("08_Pleural_Diseases.docx", "PLEURAL DISEASES & INTERVENTIONS", "pleural", "Pleural & Mediastinum"),
        ("09_Bronchiectasis_CF_AAT.docx", "BRONCHIECTASIS, CYSTIC FIBROSIS & AATD", "bronchiectasis", "Airway Diseases"),
        ("10_Sleep_and_ARDS.docx", "SLEEP MEDICINE, ARDS & CRITICAL CARE", "sleep_ards", "Critical Care"),
        ("11_Fungal_and_COVID.docx", "FUNGAL INFECTIONS & POST-COVID SEQUELAE", "fungal_covid", "Infections"),
        ("12_Diagnostic_Procedures.docx", "INTERVENTIONAL PULMONOLOGY & PROCEDURES", "procedures", "Interventions")
    ]

    chapters = []
    for fname, title, slug, category in file_spec:
        fpath = os.path.join(docs_dir, fname)
        if not os.path.exists(fpath):
            # fallback to non-50pg
            fpath = os.path.join(docs_dir, fname.replace("_50pg", ""))
        
        if not os.path.exists(fpath):
            continue

        doc = docx.Document(fpath)
        
        sections = []
        curr_heading = "Overview & Key Concepts"
        curr_text = []

        for p in doc.paragraphs:
            txt = p.text.strip()
            if not txt:
                continue
            if p.style.name.startswith("Heading 1") or p.style.name.startswith("Heading 2"):
                if curr_text:
                    sections.append({
                        "heading": curr_heading,
                        "content": "\n\n".join(curr_text[:6]) # keep punchy
                    })
                curr_heading = txt
                curr_text = []
            else:
                curr_text.append(txt)

        if curr_text:
            sections.append({
                "heading": curr_heading,
                "content": "\n\n".join(curr_text[:6])
            })

        # Extract sample tables if present
        table_data = []
        for tbl in doc.tables[:4]:
            rows = []
            for r in tbl.rows:
                cells = [c.text.strip().replace("\n", " ") for c in r.cells]
                rows.append(cells)
            if len(rows) > 1:
                table_data.append({
                    "headers": rows[0],
                    "rows": rows[1:10] # up to 10 rows
                })

        chapters.append({
            "id": slug,
            "title": title,
            "category": category,
            "source": "Crofton, Murray & Nadel, Fishman, Harrison 22e",
            "section_count": len(sections),
            "sections": sections[:18], # top 18 high-yield sections
            "tables": table_data
        })

    with open(os.path.join(DATA_DIR, "chapters_summary.json"), "w", encoding="utf-8") as f:
        json.dump(chapters, f, indent=2)
    print(f"Saved {len(chapters)} clinical textbook chapters.")

# ==========================================
# 3. BUILD ORDER SHEETS & EMERGENCY PROTOCOLS
# ==========================================
def build_order_sheets():
    print("Building Inpatient & RICU Order Sheets...")
    sheets = [
        {
            "id": "ord-01",
            "title": "Near-Fatal Asthma / Acute Severe Asthma RICU Bundle",
            "indication": "Severe acute asthma with PEF < 50%, PaCO2 normal or elevated, silent chest, exhaustion, cyanosis, altered sensorium.",
            "timing": "Initiate within First 15 Minutes of Arrival",
            "sections": [
                {
                    "category": "Immediate Resuscitation & Oxygenation",
                    "items": [
                        "High-flow humidified Oxygen via venturi mask targeting SpO2 93–95% (Caution: Avoid hyperoxia PaO2 > 100).",
                        "Continuous dual nebulization: Salbutamol (5 mg) + Ipratropium Bromide (0.5 mg) every 20 minutes for first hour, then continuous or hourly.",
                        "Direct IV access (16G or 18G cannula); draw STAT ABG, Serum Potassium, Electrolytes, CBC, CXR (portable)."
                    ]
                },
                {
                    "category": "Systemic Pharmacotherapy",
                    "items": [
                        "IV Hydrocortisone 200 mg STAT (or IV Methylprednisolone 80 mg STAT, followed by 40 mg IV q6h).",
                        "IV Magnesium Sulphate: 2.0 grams (50% solution, 4 mL) diluted in 100 mL 0.9% Normal Saline infused over 20 minutes.",
                        "IV Fluids: 0.9% Normal Saline @ 100 mL/hr (Correct dehydration from tachypnea; monitor for hypokalemia induced by beta-agonists)."
                    ]
                },
                {
                    "category": "Mechanical Ventilation Criteria (Avoid if possible)",
                    "items": [
                        "Indications for Intubation: Coma, cardiac arrest, respiratory muscle exhaustion with refractory PaCO2 > 55 mmHg and pH < 7.20.",
                        "RSI Regimen: Ketamine 1.5–2.0 mg/kg IV (bronchodilating sedative) + Rocuronium 1.2 mg/kg IV.",
                        "Ventilator Strategy: Volume Control, Vt 6 mL/kg PBW, low RR 10–12/min, high peak inspiratory flow 80–100 L/min, I:E ratio 1:4 to 1:5, PEEP 0–5 cmH2O to prevent dynamic hyperinflation (Auto-PEEP) and barotrauma."
                    ]
                }
            ]
        },
        {
            "id": "ord-02",
            "title": "Acute Exacerbation of COPD (AECOPD) with NIV Protocol",
            "indication": "Moderate-to-severe AECOPD with acute hypercapnic respiratory failure (pH < 7.35 and PaCO2 > 45 mmHg).",
            "timing": "Immediate on Ward / RICU Admission",
            "sections": [
                {
                    "category": "Non-Invasive Ventilation (NIV / BiPAP) Initiation",
                    "items": [
                        "Mode: Spontaneous/Timed (S/T) with well-fitted oronasal mask.",
                        "Initial Pressures: IPAP 10–12 cmH2O, EPAP 4–5 cmH2O (Pressure Support = 6–8 cmH2O).",
                        "Titration: Increase IPAP by 2 cmH2O every 15–30 minutes to achieve target Vt 6–8 mL/kg and decrease work of breathing (Max IPAP 20–22 cmH2O).",
                        "FiO2: Titrate to target SpO2 88–92% (Never exceed SpO2 92% to preserve hypoxic drive).",
                        "Repeat ABG at 1 Hour: Failure to improve pH or PaCO2, or worsening encephalopathy mandates invasive intubation."
                    ]
                },
                {
                    "category": "Medical Therapy",
                    "items": [
                        "Nebulized Levosalbutamol 1.25 mg + Ipratropium 0.5 mg q4h via NIV inline nebulizer adapter.",
                        "Oral Prednisolone 40 mg PO OD for 5 days (or IV Methylprednisolone 40 mg OD if unable to swallow - REDUCE Trial 5-day regimen).",
                        "Antibiotics (GOLD Anthonisen Type 1 Exacerbation - increased dyspnea, sputum volume, and purulence): Amoxicillin-Clavulanate 1.2g IV q8h OR Azithromycin 500 mg PO OD for 5 days."
                    ]
                },
                {
                    "category": "Monitoring & Venous Thromboembolism Prophylaxis",
                    "items": [
                        "Subcutaneous Enoxaparin 40 mg OD (VTE prophylaxis is mandatory in hospitalized AECOPD).",
                        "Continuous ECG and SpO2 monitoring; strict intake-output chart."
                    ]
                }
            ]
        },
        {
            "id": "ord-03",
            "title": "Massive Hemoptysis Immediate Resuscitation & Emergency Bundle",
            "indication": "Expectoration of blood > 200–500 mL in 24 hours, or any rate causing airway obstruction and desaturation.",
            "timing": "Stat Emergency Management",
            "sections": [
                {
                    "category": "Airway & Patient Positioning (Life Over Bleeding)",
                    "items": [
                        "Positioning: Lateral decubitus with suspected bleeding lung in the DEPENDENT (down) position to prevent aspiration into the healthy lung.",
                        "Supplemental Oxygen: High-flow oxygen via non-rebreather mask to maintain SpO2 > 94%.",
                        "Airway Security: If desaturating or flooding airway, perform endotracheal intubation with LARGE bore tube (>= 8.0 or 8.5 mm ID) to allow therapeutic flexible bronchoscopy and suctioning."
                    ]
                },
                {
                    "category": "Circulatory & Hemostatic Support",
                    "items": [
                        "Two large-bore IV lines (16G); send STAT crossmatch 4 units PRBC, PT/INR, aPTT, Platelets, CBC, BUN/Cr.",
                        "IV Tranexamic Acid: 1.0 g IV push over 10 minutes, followed by 1.0 g in 100 mL NS infused over 8 hours.",
                        "Nebulized Tranexamic Acid: 500 mg (5 mL) in 5 mL NS nebulized every 6–8 hours (proven topical vasoconstriction).",
                        "Cough suppression: Codeine Phosphate 30 mg PO/IM q6h (prevent sudden re-bleeding from violent cough paroxysms)."
                    ]
                },
                {
                    "category": "Definitive Interventional Radiology & Endoscopy",
                    "items": [
                        "Urgent CT Bronchial Angiography (CTBA) with contrast to identify hypertrophied bronchial and non-bronchial systemic collaterals.",
                        "Urgent Bronchial Artery Embolization (BAE) by Interventional Radiology: Gold standard definitive therapy with > 90% immediate success rate.",
                        "Bronchoscopy: Cold saline instillation (50 mL aliquots at 4°C), topical adrenaline (1:20,000), or balloon tamponade catheter if bleeding lung is known."
                    ]
                }
            ]
        },
        {
            "id": "ord-04",
            "title": "Severe Sepsis / Severe CAP in RICU (Surviving Sepsis / IDSA-ATS)",
            "indication": "Community-Acquired Pneumonia with shock, altered mentation, multilobar opacities, PaO2/FiO2 < 250, or CURB-65 >= 3.",
            "timing": "Hour-1 Universal Sepsis Bundle",
            "sections": [
                {
                    "category": "Diagnostics Before Antimicrobials",
                    "items": [
                        "Measure Blood Lactate level STAT; remeasure within 2–4 hours if initial lactate > 2.0 mmol/L.",
                        "Obtain two sets of blood cultures (aerobic and anaerobic) from separate venipuncture sites before starting antibiotics.",
                        "Sputum Gram stain and culture (deep expectorate or endotracheal aspirate).",
                        "Urinary Legionella pneumophila serogroup 1 antigen and Streptococcus pneumoniae antigen."
                    ]
                },
                {
                    "category": "Empirical Broad-Spectrum Antimicrobial Regimen",
                    "items": [
                        "Administer broad-spectrum IV antimicrobials WITHIN 1 HOUR of arrival.",
                        "Non-Pseudomonas ICU Regimen: IV Ceftriaxone 2g OD (or IV Ampicillin-Sulbactam 3g q6h) PLUS IV Azithromycin 500 mg OD (or IV Levofloxacin 750 mg OD).",
                        "If Pseudomonas aeruginosa risk factors (structural lung disease, bronchiectasis, prior colonization): IV Piperacillin-Tazobactam 4.5g q6h (extended 3-hour infusion) PLUS IV Ciprofloxacin 400 mg q8h (or Amikacin 15 mg/kg IV OD).",
                        "If MRSA risk factors (cavitary infiltrates, post-influenza, recent hospitalization): Add IV Vancomycin 15–20 mg/kg q12h (target trough 15–20 mcg/mL) OR IV Linezolid 600 mg q12h."
                    ]
                },
                {
                    "category": "Hemodynamic Resuscitation & Vasopressors",
                    "items": [
                        "Rapid fluid resuscitation: 30 mL/kg of Balanced Crystalloid (e.g. Plasmalyte or Ringer's Lactate) for hypotension (MAP < 65) or Lactate >= 4 mmol/L.",
                        "First-Line Vasopressor: Norepinephrine infusion titrated to target MAP >= 65 mmHg.",
                        "Refractory Septic Shock: Add IV Hydrocortisone 200 mg/day (50 mg q6h) if vasopressor requirement continues despite adequate fluid loading."
                    ]
                }
            ]
        },
        {
            "id": "ord-05",
            "title": "Massive Pulmonary Embolism & Thrombolysis Protocol",
            "indication": "High-risk PE presenting with sustained systemic hypotension (SBP < 90 mmHg for > 15 min), cardiogenic shock, or cardiac arrest.",
            "timing": "Immediate STAT Protocol",
            "sections": [
                {
                    "category": "Pre-Thrombolytic Checklist & Absolute Contraindications",
                    "items": [
                        "Screen for Absolute Contraindications: Hemorrhagic stroke or stroke of unknown origin at any time; ischemic stroke in preceding 6 months; CNS neoplasm; major trauma/surgery/head injury in past 3 weeks; active bleeding; known bleeding diathesis.",
                        "Baseline labs: STAT PT/INR, aPTT, Platelets, Fibrinogen, Troponin I, NT-proBNP.",
                        "Bedside Echocardiography: Assess McConnell sign, RV dilation (RV/LV > 1.0), tricuspid regurgitation peak velocity, paradoxical septal shift."
                    ]
                },
                {
                    "category": "Thrombolysis Dosing Regimens",
                    "items": [
                        "Preferred Regimen: Alteplase (rt-PA) 100 mg IV infusion over 2 hours (10 mg IV bolus over 1–2 minutes, followed by 90 mg IV over 2 hours).",
                        "Accelerated / Arrest Regimen: Alteplase 50 mg IV bolus over 15 minutes (in impending or ongoing cardiopulmonary arrest).",
                        "Alternative (Tenecteplase weight-adjusted): 30–50 mg IV bolus over 5–10 seconds based on weight (<60kg: 30mg; >=90kg: 50mg)."
                    ]
                },
                {
                    "category": "Anticoagulation Transition",
                    "items": [
                        "Hold Unfractionated Heparin (UFH) during 2-hour Alteplase infusion.",
                        "Check aPTT at conclusion of thrombolysis; resume IV UFH when aPTT is < 2.0 times normal control (target aPTT 1.5–2.5 times control).",
                        "Avoid arterial punctures and IM injections for 24 hours."
                    ]
                }
            ]
        },
        {
            "id": "ord-06",
            "title": "ARDS Lung-Protective Mechanical Ventilation Protocol (ARDSNet)",
            "indication": "Acute onset within 1 week of clinical insult, bilateral infiltrates not fully explained by heart failure, PaO2/FiO2 <= 300 with PEEP >= 5 cmH2O.",
            "timing": "Immediate Post-Intubation",
            "sections": [
                {
                    "category": "Tidal Volume & Plateau Pressure Targets",
                    "items": [
                        "Calculate Predicted Body Weight (PBW): Men = 50 + 0.91 * (height in cm - 152.4); Women = 45.5 + 0.91 * (height in cm - 152.4).",
                        "Initial Tidal Volume: Set at 6 mL/kg PBW (reduce to 5 or 4 mL/kg if Plateau Pressure exceeds 30 cmH2O).",
                        "Plateau Pressure (Pplat) Target: Keep Pplat <= 30 cmH2O (check end-inspiratory pause for 0.5 seconds).",
                        "Driving Pressure Target (ΔP = Pplat - PEEP): Maintain ΔP < 15 cmH2O (strongest predictor of ARDS survival)."
                    ]
                },
                {
                    "category": "PEEP & FiO2 Titration (High PEEP / ARDSNet Table)",
                    "items": [
                        "Target Oxygenation: PaO2 55–80 mmHg or SpO2 88–95%.",
                        "Moderate-to-Severe ARDS: Use higher PEEP titration table (e.g. FiO2 0.5 -> PEEP 12–14; FiO2 0.7 -> PEEP 16–18; FiO2 1.0 -> PEEP 20–24 cmH2O).",
                        "Avoid repetitive derecruitment and atelectrauma."
                    ]
                },
                {
                    "category": "Refractory Hypoxemia Interventions",
                    "items": [
                        "Prone Positioning: MANDATORY for severe ARDS (PaO2/FiO2 < 150) for at least 16 CONSECUTIVE HOURS per day (PROSEVA Trial: mortality reduced from 32.8% to 16.0%).",
                        "Neuromuscular Blockade: Cisatracurium besylate infusion for first 48 hours in severe patient-ventilator dyssynchrony.",
                        "Conservative Fluid Strategy (FACTT Trial): Fluid balance neutral to negative once shock has resolved (significantly increases ventilator-free days)."
                    ]
                }
            ]
        },
        {
            "id": "ord-07",
            "title": "Intercostal Drainage (ICD) Tube Insertion & Care SOP",
            "indication": "Pneumothorax (tension, secondary, large primary), massive hemothorax, complicated parapneumonic effusion, empyema, malignant effusion with dyspnea.",
            "timing": "Ward / ICU Procedure",
            "sections": [
                {
                    "category": "Safe Triangle of Insertion & Preparation",
                    "items": [
                        "Anatomic Triangle of Safety: Anterior border of Latissimus dorsi, lateral border of Pectoralis major, horizontal line at level of 5th intercostal space (nipple level in males), apex just below axilla.",
                        "Position: Patient semi-recumbent at 45 degrees with ipsilateral arm abducted behind head.",
                        "Local Anesthesia: 1% or 2% Lignocaine (up to 3 mg/kg) infiltrated down to parietal pleura (aspirate pleural fluid/air to confirm cavity before incising)."
                    ]
                },
                {
                    "category": "Insertion Technique",
                    "items": [
                        "Blunt dissection over UPPER border of the lower rib (avoids neurovascular bundle running in the subcostal groove).",
                        "Size Selection: Small-bore pigtail catheter (10–14 Fr) preferred for simple pneumothorax and free-flowing fluid; Large-bore (24–28 Fr) for hemothorax and thick empyema.",
                        "Connect immediately to underwater seal bottle: Underwater seal tube must be submerged 2 cm below sterile water level."
                    ]
                },
                {
                    "category": "Monitoring & Removal Criteria",
                    "items": [
                        "Oscillation (Swinging): Normal column moves with respiration (up on inspiration, down on expiration during spontaneous breathing). Cessation indicates tube blockage, displacement, or complete lung expansion.",
                        "Bubbling: Persistent bubbling indicates active air leak (check cough leak vs continuous leak).",
                        "Removal Criteria: Complete clinical & radiological resolution of pneumothorax; fluid drainage < 50–100 mL in 24 hours (serous); no air leak on forced coughing for > 24 hours."
                    ]
                }
            ]
        }
    ]

    with open(os.path.join(DATA_DIR, "order_sheets.json"), "w", encoding="utf-8") as f:
        json.dump(sheets, f, indent=2)
    print(f"Saved {len(sheets)} clinical order sheets.")

# ==========================================
# 4. BUILD PHARMACOLOGY COMPENDIUM
# ==========================================
def build_pharmacology():
    print("Building Respiratory Pharmacology...")
    pharma = {
        "drug_classes": [
            {
                "class_name": "Inhaled Bronchodilators (Beta-2 Agonists & Muscarinic Antagonists)",
                "drugs": [
                    {
                        "name": "Salbutamol / Albuterol (SABA)",
                        "mechanism": "Selective beta-2 adrenergic receptor agonist stimulating adenylate cyclase, increasing cAMP, and causing airway smooth muscle relaxation.",
                        "dose": "pMDI: 100–200 mcg (1–2 puffs) PRN; Nebulization: 2.5–5.0 mg in 3 mL NS q20min in acute severe asthma.",
                        "onset_duration": "Onset: 1–5 min; Peak: 30–60 min; Duration: 4–6 hours.",
                        "side_effects": "Tremors, tachycardia, hypokalemia (shifts K+ intracellularly), lactic acidosis (type B2), tolerance with chronic overuse.",
                        "exam_pearl": "Excess SABA use (>= 3 canisters/year) is directly linked to increased asthma-related mortality (GINA eliminated SABA-only monotherapy)."
                    },
                    {
                        "name": "Formoterol (Fast- and Long-Acting Beta-2 Agonist - LABA)",
                        "mechanism": "Full beta-2 agonist with high lipophilicity that inserts into cell membrane lipid bilayer, providing rapid onset like salbutamol PLUS 12-hour sustained bronchodilation.",
                        "dose": "6–12 mcg with Budesonide (100, 200, or 400 mcg) via DPI/pMDI (MART / SMART regimen).",
                        "onset_duration": "Onset: 1–3 min (immediate relief); Duration: 12 hours.",
                        "side_effects": "Headache, palpitations, muscle cramps.",
                        "exam_pearl": "Formoterol is the ONLY LABA recommended for both controller AND reliever therapy (Single Maintenance and Reliever Therapy - SMART) because of its rapid onset of action."
                    },
                    {
                        "name": "Tiotropium Bromide (Long-Acting Muscarinic Antagonist - LAMA)",
                        "mechanism": "Kinetic selectivity for M3 and M1 muscarinic receptors; slow dissociation from M3 receptors produces sustained bronchodilation for > 24 hours.",
                        "dose": "HandiHaler DPI: 18 mcg capsule OD; Respimat Soft Mist Inhaler: 2.5 mcg (2 puffs OD = 5 mcg).",
                        "onset_duration": "Onset: 30 min; Peak: 2–4 hours; Duration: 24–36 hours.",
                        "side_effects": "Dry mouth (most common), urinary retention (caution in BPH), glaucoma worsening if mist enters eyes.",
                        "exam_pearl": "Tiotropium is the cornerstone of COPD management (reduces exacerbations, improves quality of life, slows FEV1 decline). Also approved as add-on in GINA Step 4/5 severe asthma."
                    }
                ]
            },
            {
                "class_name": "Inhaled & Systemic Corticosteroids",
                "drugs": [
                    {
                        "name": "Budesonide (ICS)",
                        "mechanism": "Binds intracellular glucocorticoid receptors, translocates to nucleus, inhibits NF-kB and AP-1 transcription factors, reducing pro-inflammatory cytokines, eosinophils, and airway edema.",
                        "dose": "Low: 200–400 mcg/day; Medium: 400–800 mcg/day; High: > 800 mcg/day.",
                        "onset_duration": "Topical anti-inflammatory onset: hours to days; full clinical effect: 1–2 weeks.",
                        "side_effects": "Oropharyngeal candidiasis (thrush), dysphonia (rinse and spit reduces by 90%), pneumonia risk in severe COPD.",
                        "exam_pearl": "Budesonide has high first-pass hepatic metabolism (90%), giving it an exceptional topical-to-systemic therapeutic index."
                    },
                    {
                        "name": "Prednisolone / Methylprednisolone (Systemic Steroids)",
                        "mechanism": "Potent systemic glucocorticoid anti-inflammatory and immunosuppressive activity.",
                        "dose": "AECOPD: Prednisolone 40 mg PO OD for 5 days; Acute Asthma: 40–50 mg PO OD for 5–7 days; IV Methylprednisolone: 40–80 mg IV q6–8h in RICU.",
                        "onset_duration": "Onset: 4–6 hours (does not produce immediate bronchodilation).",
                        "side_effects": "Hyperglycemia, hypertension, hypokalemia, peptic ulcer, avascular necrosis of femoral head, secondary adrenal suppression.",
                        "exam_pearl": "Tapering is UNNECESSARY for acute courses lasting <= 14 days unless the patient has received frequent steroid bursts in the past 6 months."
                    }
                ]
            },
            {
                "class_name": "Biologic Therapies for Severe Asthma",
                "drugs": [
                    {
                        "name": "Omalizumab (Xolair)",
                        "mechanism": "Humanized IgG1 monoclonal antibody that binds free human IgE at the Fc region, preventing it from binding to high-affinity Fc-epsilon-RI receptors on mast cells and basophils.",
                        "dose": "75–600 mg SC every 2 or 4 weeks, determined by baseline serum total IgE (30–1500 IU/mL) and body weight.",
                        "indication": "Severe persistent allergic asthma in patients >= 6 years with positive skin prick test or in vitro reactivity to a perennial aeroallergen.",
                        "side_effects": "Anaphylaxis (0.2%, observe for 2 hours post first 3 doses), injection site pain, headache.",
                        "exam_pearl": "First approved biologic for asthma (2003). Serum total IgE increases after starting therapy (forms inactive immune complexes), so IgE CANNOT be used to monitor therapy."
                    },
                    {
                        "name": "Mepolizumab & Benralizumab (Anti-IL-5 Pathway)",
                        "mechanism": "Mepolizumab binds soluble IL-5; Benralizumab binds the alpha subunit of the IL-5 receptor (IL-5Rα) on eosinophils and basophils, recruiting NK cells to induce antibody-dependent cell-mediated cytotoxicity (ADCC) and near-complete eosinophil depletion.",
                        "dose": "Mepolizumab: 100 mg SC q4w; Benralizumab: 30 mg SC q4w for 3 doses, then q8w.",
                        "indication": "Severe eosinophilic asthma (blood eosinophils >= 150–300 cells/mcL).",
                        "side_effects": "Headache, pharyngitis, pyrexia.",
                        "exam_pearl": "Benralizumab induces near-zero peripheral and tissue eosinophil counts within 24 hours via ADCC, making it the most potent eosinophil-depleting biologic."
                    },
                    {
                        "name": "Dupilumab (Dupixent)",
                        "mechanism": "Fully human monoclonal antibody against the IL-4 receptor alpha (IL-4Rα) subunit, simultaneously blocking signaling of both IL-4 and IL-13 (Type 2 inflammation drivers).",
                        "dose": "Loading dose 400 mg (or 600 mg) SC, followed by 200 mg (or 300 mg) SC every 2 weeks.",
                        "indication": "Severe asthma with Type 2 phenotype (eosinophils >= 150 or FeNO >= 25 ppb) or oral corticosteroid-dependent asthma.",
                        "side_effects": "Transient hypereosinophilia (usually asymptomatic, monitor closely), conjunctivitis, injection site reaction.",
                        "exam_pearl": "Dupilumab is the ONLY biologic approved for asthma that is also a steroid-sparing superstar (significantly reduces maintenance oral corticosteroid dose)."
                    },
                    {
                        "name": "Tezepelumab (Tezspire)",
                        "mechanism": "Human monoclonal antibody targeting Thymic Stromal Lymphopoietin (TSLP), an epithelial 'alarmin' cytokine at the top of the inflammatory cascade.",
                        "dose": "210 mg SC once every 4 weeks.",
                        "indication": "Severe asthma irrespective of baseline eosinophils, IgE, or FeNO ('broad phenotype' biologic).",
                        "side_effects": "Pharyngitis, arthralgia, back pain.",
                        "exam_pearl": "Approved as the first upstream alarmin inhibitor; effective in non-Type 2 / low-eosinophil severe asthma where other biologics fail."
                    }
                ]
            },
            {
                "class_name": "Antifibrotic Agents for Idiopathic Pulmonary Fibrosis (IPF)",
                "drugs": [
                    {
                        "name": "Pirfenidone",
                        "mechanism": "Pleiotropic agent with antifibrotic, antioxidant, and anti-inflammatory properties; inhibits TGF-beta synthesis, fibroblast proliferation, and collagen synthesis.",
                        "dose": "Titrate weekly: 200 mg TID with food for 1 week -> 400 mg TID for 1 week -> Target 600–800 mg TID (total 1800–2403 mg/day in India/US). Take with food.",
                        "monitoring": "Baseline LFTs, then monthly for first 6 months, then every 3 months. Stop/reduce if ALT/AST > 3x ULN.",
                        "side_effects": "Photosensitivity and rash (strict broad-spectrum SPF 50 sunscreen mandatory), nausea, anorexia, weight loss, fatigue.",
                        "exam_pearl": "Reduces the rate of decline of Forced Vital Capacity (FVC) by approximately 50% in mild-to-moderate IPF (CAPACITY and ASCEND trials)."
                    },
                    {
                        "name": "Nintedanib",
                        "mechanism": "Intracellular tyrosine kinase inhibitor targeting Vascular Endothelial Growth Factor Receptors (VEGFR 1-3), Fibroblast Growth Factor Receptors (FGFR 1-3), and Platelet-Derived Growth Factor Receptors (PDGFR alpha and beta).",
                        "dose": "150 mg PO BID with food (12 hours apart). Can reduce to 100 mg BID if diarrhea occurs.",
                        "monitoring": "LFTs at baseline, monthly for 3 months, then periodically; monitor stool frequency.",
                        "side_effects": "Diarrhea (occurs in > 60% of patients; manage with hydration and loperamide), elevated transaminases, arterial thromboembolism risk.",
                        "exam_pearl": "Approved for IPF (INPULSIS trials), progressive fibrosing ILDs (INBUILD trial), and systemic sclerosis-associated ILD (SENSCIS trial)."
                    }
                ]
            },
            {
                "class_name": "NTEP Anti-Tuberculosis Regimens (2024–2026 Guidelines)",
                "drugs": [
                    {
                        "name": "First-Line Regimen for Drug-Susceptible TB (DS-TB)",
                        "mechanism": "Daily fixed-dose combination (FDC) based on weight bands.",
                        "dose": "Intensive Phase (2 months): 4-drug FDC (Isoniazid H 75mg + Rifampicin R 150mg + Pyrazinamide Z 400mg + Ethambutol E 275mg). Continuation Phase (4 months): 3-drug FDC (HRE). Weight bands: 25–34kg (2 tabs), 35–49kg (3 tabs), 50–64kg (4 tabs), >=65kg (5 tabs).",
                        "monitoring": "Sputum AFB/NAAT at end of month 2 and month 6; baseline and monthly LFTs in symptomatic patients.",
                        "side_effects": "Hepatotoxicity (R > H > Z), orange discoloration of secretions (Rifampicin), hyperuricemia & arthralgia (Pyrazinamide), retrobulbar optic neuritis (Ethambutol).",
                        "exam_pearl": "NTEP replaced intermittent thrice-weekly therapy with daily FDCs. If ALT > 3x ULN with symptoms or > 5x ULN without symptoms, stop all hepatotoxic drugs (H, R, Z) immediately and start safe modified regimen (Ethambutol + Levofloxacin + Streptomycin/Amikacin)."
                    },
                    {
                        "name": "BPaLM Regimen for Multidrug-Resistant / Rifampicin-Resistant TB (MDR/RR-TB)",
                        "mechanism": "Novel all-oral 6-month regimen replacing 18–24 month toxic regimens.",
                        "dose": "Bedaquiline (400 mg OD for 2 weeks, then 200 mg thrice weekly for 24 weeks) + Pretomanid (200 mg OD for 26 weeks) + Linezolid (600 mg OD for 16–26 weeks) + Moxifloxacin (400 mg OD for 26 weeks).",
                        "indication": "MDR/RR-TB or pre-XDR-TB without prior exposure to bedaquiline/linezolid (> 1 month) and fluoroquinolone-susceptible.",
                        "monitoring": "Monthly ECG (QTc monitoring for Bedaquiline + Moxifloxacin; hold if QTc > 500 ms), CBC for Linezolid myelosuppression (anemia, thrombocytopenia), and peripheral/optic neuropathy screening.",
                        "side_effects": "QT prolongation, myelosuppression, peripheral neuropathy, optic neuritis, hepatotoxicity.",
                        "exam_pearl": "TB-PRACTECAL trial demonstrated 89% cure rate for 6-month all-oral BPaLM, making it the new WHO and NTEP gold standard."
                    }
                ]
            }
        ],
        "inhaler_technique": [
            {
                "device": "Pressurized Metered Dose Inhaler (pMDI)",
                "steps": [
                    "1. Remove cap and shake the inhaler vigorously for 5 seconds.",
                    "2. Sit or stand upright; exhale completely to residual volume away from the inhaler.",
                    "3. Hold inhaler upright between thumb and index finger; place mouthpiece between teeth and seal lips firmly.",
                    "4. Start breathing in SLOWLY and DEEPLY through mouth, and press down on canister ONCE at the start of inhalation.",
                    "5. Continue slow deep inhalation for 3–5 seconds until lungs are completely full.",
                    "6. Hold breath for 10 seconds (or as long as comfortable) to allow drug deposition in peripheral airways.",
                    "7. Wait 30–60 seconds before repeating for a second puff; rinse mouth if using steroid."
                ],
                "common_errors": "Actuation before inhalation (cold-Freon effect stops breath), inhaling too fast, failing to hold breath, not shaking device."
            },
            {
                "device": "pMDI with Spacer (Valved Holding Chamber)",
                "steps": [
                    "1. Check spacer for cracks and remove protective cap.",
                    "2. Shake pMDI for 5 seconds and insert firmly into rubber end of spacer.",
                    "3. Exhale fully; place spacer mouthpiece firmly in mouth with tight lip seal.",
                    "4. Press canister ONCE to release a single puff of medication into spacer chamber.",
                    "5. Breathe in slowly and deeply over 3–5 seconds (if spacer whistles, breath is too fast!).",
                    "6. Hold breath for 10 seconds; for children or elderly, tidal breathing (5 breaths per actuation) is acceptable.",
                    "7. Clean spacer monthly: soak in mild detergent, do NOT scrub or towel dry (air dry to eliminate static charge)."
                ],
                "common_errors": "Actuating multiple puffs simultaneously into spacer, scrubbing spacer with cloth creating electrostatic charge that attracts drug."
            },
            {
                "device": "Dry Powder Inhaler (DPI - Rotahaler / Revolizer / Turbuhaler)",
                "steps": [
                    "1. Load capsule or twist base to charge the metered dose.",
                    "2. Hold inhaler LEVEL and exhale completely AWAY from the inhaler (never exhale into a DPI - moisture clumps powder!).",
                    "3. Place mouthpiece in mouth with firm seal.",
                    "4. Inhale FORCIBLY, QUICKLY, and DEEPLY (requires high inspiratory effort > 30–60 L/min to disaggregate drug particles).",
                    "5. Remove from mouth and hold breath for 10 seconds.",
                    "6. Open device and verify empty capsule; rinse mouth if containing steroid."
                ],
                "common_errors": "Exhaling into mouthpiece, slow inhalation (unlike pMDI, DPI REQUIRES fast and forceful inhalation to aerosolize powder)."
            }
        ]
    }

    with open(os.path.join(DATA_DIR, "pharmacology.json"), "w", encoding="utf-8") as f:
        json.dump(pharma, f, indent=2)
    print("Saved respiratory pharmacology compendium.")

# ==========================================
# 5. BUILD VIVA STATIONS
# ==========================================
def build_viva_stations():
    print("Building MD/DNB Table Viva Stations...")
    stations = [
        {
            "id": "viva-01",
            "station_number": 1,
            "title": "Intercostal Drainage (ICD) Underwater Seal Bottle",
            "image": "/assets/thoracic_atlas_assets/fig16.jpg",
            "questions": [
                {
                    "q": "Explain the principle and physics of the underwater seal drainage system.",
                    "a": "It acts as a one-way valve allowing air and fluid to escape from the pleural cavity while preventing atmospheric air from re-entering. The rigid tube is submerged 2 cm under sterile water. Intrapleural pressure during expiration exceeds the 2 cmH2O water resistance, allowing expulsion. On inspiration, the negative pleural pressure pulls water up the tube, but atmospheric air cannot enter."
                },
                {
                    "q": "What does respiratory swinging (oscillation) mean, and why might it stop?",
                    "a": "The fluid column swings up during spontaneous inspiration (more negative pleural pressure) and down during expiration. Cessation indicates: (1) complete lung re-expansion with visceral pleura blocking the tube eyelets, (2) tube occlusion by clot or fibrin, (3) kinking, or (4) tube displacement out of pleural cavity."
                },
                {
                    "q": "When is clamping an ICD bottle strictly CONTRAINDICATED?",
                    "a": "Clamping is strictly contraindicated in the presence of an active air leak (persistent bubbling) or while a patient is on positive pressure mechanical ventilation. Clamping under these conditions can convert an open pneumothorax into a fatal TENSION PNEUMOTHORAX within minutes."
                }
            ]
        },
        {
            "id": "viva-02",
            "station_number": 2,
            "title": "Flexible Fiberoptic Bronchoscope & Bronchoscopy Station",
            "image": "/assets/thoracic_atlas_assets/fig27.jpg",
            "questions": [
                {
                    "q": "List the anatomical landmarks identifying the Right Main Bronchus vs Left Main Bronchus.",
                    "a": "The main carina is sharp and sagittal. The Right Main Bronchus is wider, shorter (1.5–2 cm), and more vertical (diverges at ~25 degrees), making aspiration into the right lung more common. Its first branch is the Right Upper Lobe bronchus (trifurcates into apical B1, posterior B2, anterior B3) located only 1.5–2 cm from carina. The Left Main Bronchus is narrower, longer (4–5 cm), and more horizontal (diverges at ~45 degrees)."
                },
                {
                    "q": "What is the disinfection protocol for a flexible bronchoscope?",
                    "a": "High-level disinfection using 2% Glutaraldehyde (Cidex) for 20 minutes (or Ortho-phthalaldehyde OPA 0.55% for 12 minutes). For Mycobacterium tuberculosis or spore-forming organisms, immersion for 45–60 minutes is required. Rinsing thoroughly with sterile water and alcohol channel drying is mandatory before storage in a vertical drying cabinet."
                },
                {
                    "q": "What are the indications and absolute contraindications for transbronchial lung biopsy (TBLB)?",
                    "a": "Indications: Diffuse parenchymal lung diseases (sarcoidosis, lymphangitic carcinomatosis), localized peripheral pulmonary nodules. Absolute Contraindications: Severe intractable hypoxemia (PaO2 < 60 on high FiO2), uncorrected coagulopathy/thrombocytopenia (platelets < 50,000, INR > 1.5), severe pulmonary hypertension (elevated PAP risks torrential hemorrhage), uncooperative patient, acute myocardial infarction within 6 weeks."
                }
            ]
        },
        {
            "id": "viva-03",
            "station_number": 3,
            "title": "Pleural Biopsy Needles (Abrams vs Cope Needle)",
            "image": "/assets/thoracic_atlas_assets/fig21.jpg",
            "questions": [
                {
                    "q": "Describe the components and working mechanism of the Abrams pleural biopsy needle.",
                    "a": "The Abrams needle consists of 3 concentric components: (1) Outer cannula with a cutting edge and a distal lateral notch, (2) Inner cutting cylinder that rotates to shear the tissue, and (3) Central sharp stylet. It is introduced with the notch closed. Once in the pleural space, pleural fluid is aspirated to confirm location, the notch is rotated open, hooked onto the parietal pleura, retracted until resistance is felt, and the inner cylinder is rotated to guillotine the biopsy specimen."
                },
                {
                    "q": "What is the primary anatomical precaution when taking a closed pleural biopsy?",
                    "a": "Biopsy must ALWAYS be taken in the 6 o'clock position (inferior border of the puncture site) over the UPPER border of the lower rib. Taking biopsies superiorly (12 o'clock) risks lacerating the intercostal neurovascular bundle (van: Vein, Artery, Nerve) running in the subcostal groove."
                },
                {
                    "q": "What is the diagnostic yield of closed pleural biopsy in Malignant Pleural Mesothelioma vs Pleural TB?",
                    "a": "Yield in Pleural TB is very high (~60–80%, combining histopathology showing caseating granulomas with AFB staining and tissue TB-PCR/GeneXpert). In malignant pleural mesothelioma, the yield is poor (~20–40%) because mesothelioma spares the lower parietal pleura and requires deeper tissue sampling; Medical Thoracoscopy (pleuroscopy) is the gold standard (>95% yield)."
                }
            ]
        },
        {
            "id": "viva-04",
            "station_number": 4,
            "title": "PFT & Spirometry Flow-Volume Loops",
            "image": "/assets/asthma_assets/airway_comparison.png",
            "questions": [
                {
                    "q": "Differentiate Variable Intrathoracic vs Variable Extrathoracic vs Fixed Upper Airway Obstruction on flow-volume loops.",
                    "a": "(1) Variable Extrathoracic (e.g. vocal cord palsy, laryngomalacia): Flattened INSPIRATORY loop (negative intratracheal pressure during inspiration collapses the lesion), normal expiratory loop. (2) Variable Intrathoracic (e.g. tracheomalacia, lower tracheal tumor): Flattened EXPIRATORY loop (positive pleural pressure during expiration compresses airway), normal inspiratory loop. (3) Fixed Obstruction (e.g. tracheal stenosis, goiter): Flattening of BOTH inspiratory and expiratory loops."
                },
                {
                    "q": "What defines positive bronchodilator reversibility (BDR) according to ATS/ERS 2022 guidelines?",
                    "a": "An increase in either FEV1 or FVC by > 10% of the PREDICTED value (or traditionally > 12% AND > 200 mL absolute increase from pre-bronchodilator baseline) measured 15 minutes after inhaling 400 mcg of Salbutamol."
                },
                {
                    "q": "How do you distinguish Parenchymal Restriction (e.g. IPF) from Chest Wall / Extraparenchymal Restriction on DLCO?",
                    "a": "Both show reduced TLC (< 80%) with normal or elevated FEV1/FVC ratio. However, in Intrinsic/Parenchymal lung disease (IPF, ILD), alveolar-capillary membrane destruction causes a REDUCED DLCO (< 80%). In Extrinsic/Chest wall disease (kyphoscoliosis, myasthenia gravis, obesity), the lung parenchyma is intact, so DLCO is NORMAL or elevated when adjusted for lung volume (DLCO/VA normal)."
                }
            ]
        }
    ]

    with open(os.path.join(DATA_DIR, "viva_stations.json"), "w", encoding="utf-8") as f:
        json.dump(stations, f, indent=2)
    print(f"Saved {len(stations)} viva stations.")

# ==========================================
# 6. BUILD MCQ BANK (120+ QUESTIONS)
# ==========================================
def build_mcq_bank():
    print("Building Comprehensive MCQ & Spotter Bank...")
    # Load batch files or construct high-yield MCQs
    mcqs = [
        # --- Infections (Pneumonia, TB, Fungal) ---
        {
            "id": "MCQ-001",
            "topic": "Infections",
            "subspecialty": "Community-Acquired Pneumonia",
            "type": "Clinical Vignette",
            "difficulty": "Moderate",
            "question": "A 68-year-old male with a history of hypertension presents with a 4-day history of high fever, productive cough with rust-colored sputum, and right-sided pleuritic chest pain. On examination: RR 32/min, BP 86/54 mmHg, HR 114/min, SpO2 88% on room air. Lab investigations show BUN 32 mg/dL, WBC 18,500/mcL. CXR shows dense right middle lobe consolidation with air bronchograms. What is his CURB-65 score and the most appropriate management setting?",
            "options": [
                "CURB-65 score 2; Outpatient therapy with oral amoxicillin-clavulanate",
                "CURB-65 score 3; Inpatient ward admission for IV ceftriaxone",
                "CURB-65 score 4; Immediate ICU admission for IV beta-lactam plus macrolide/fluoroquinolone",
                "CURB-65 score 5; Emergency endotracheal intubation and mechanical ventilation"
            ],
            "correct_index": 2,
            "explanation": "CURB-65 criteria evaluation:\n- Confusion: No (0)\n- Urea > 7 mmol/L (BUN > 19 mg/dL): BUN is 32 mg/dL (1)\n- Respiratory rate >= 30/min: RR is 32/min (1)\n- Blood pressure (SBP < 90 or DBP <= 60 mmHg): BP is 86/54 mmHg (1)\n- Age >= 65 years: Age is 68 (1)\nTotal score = 4. A CURB-65 score of 3–5 indicates severe pneumonia with a high 30-day mortality risk (score 4 has ~27% mortality), warranting direct ICU admission and immediate broad-spectrum dual antimicrobial coverage (IV Ceftriaxone + Azithromycin).",
            "clinical_pearl": "Rust-colored sputum with lobar consolidation and air bronchograms is the classic presentation of Streptococcus pneumoniae (Pneumococcus).",
            "trap": "Don't forget the 'B' in CURB-65: DBP <= 60 is an independent criterion even if SBP was marginally above 90.",
            "reference": "Murray & Nadel 7th Ed, Ch 32; Harrison 22e, Ch 121; IDSA/ATS CAP Guidelines",
            "image": "/assets/thoracic_atlas_assets/fig1.jpg"
        },
        {
            "id": "MCQ-002",
            "topic": "Infections",
            "subspecialty": "Tuberculosis",
            "type": "Image-Based Question",
            "difficulty": "Moderate",
            "question": "A 29-year-old female presents with persistent evening low-grade fever, night sweats, and productive cough for 6 weeks. A sputum Cartridge-Based Nucleic Acid Amplification Test (CBNAAT / GeneXpert MTB/RIF) detects Mycobacterium tuberculosis with Rifampicin resistance DETECTED. According to current NTEP guidelines, what is the next mandatory test before initiating the 6-month all-oral BPaLM regimen?",
            "options": [
                "Line Probe Assay (Second-line LPA) to evaluate fluoroquinolone susceptibility",
                "Solid culture on Lowenstein-Jensen medium with 8-week wait",
                "High-Resolution Computed Tomography of the Thorax",
                "Pleural fluid adenosine deaminase (ADA) analysis"
            ],
            "correct_index": 0,
            "explanation": "Under the National Tuberculosis Elimination Program (NTEP) and WHO guidelines, any patient diagnosed with Rifampicin-resistant TB (RR-TB) must immediately undergo Second-Line Line Probe Assay (SL-LPA) or second-line phenotypic DST to check for Fluoroquinolone resistance. If fluoroquinolone susceptible (pre-XDR-TB excluded), the patient is eligible for the 6-month BPaLM regimen (Bedaquiline + Pretomanid + Linezolid + Moxifloxacin). If fluoroquinolone resistant, Moxifloxacin is omitted, giving the BPaL regimen.",
            "clinical_pearl": "Fluoroquinolone resistance status is the absolute pivot point between the 4-drug BPaLM regimen and the 3-drug BPaL regimen.",
            "trap": "Do not start MDR regimens blindly without confirming fluoroquinolone susceptibility; using moxifloxacin in resistant strains drives further drug resistance.",
            "reference": "NTEP 2024 Guidelines on Drug-Resistant Tuberculosis; WHO Consolidated Guidelines 2022",
            "image": "/assets/thoracic_atlas_assets/fig33.jpg"
        },
        {
            "id": "MCQ-003",
            "topic": "Infections",
            "subspecialty": "Fungal Infections",
            "type": "Image-Based Question",
            "difficulty": "Hard",
            "question": "A 52-year-old male with treated pulmonary tuberculosis 5 years ago presents with recurrent massive hemoptysis. A contrast-enhanced CT of the thorax is performed (image shows Monod's sign: a mobile intracavitary mass with a crescent of air surrounding it). What is the classic radiologic sign and the definitive management of choice?",
            "options": [
                "Reverse halo sign; IV Amphotericin B deoxycholate",
                "Monod sign (Mycetoma); Surgical resection (lobectomy) or bronchial artery embolization for life-threatening hemoptysis",
                "Comet tail sign; Chest tube drainage with instillation of urokinase",
                "Crazy paving pattern; Whole lung lavage"
            ],
            "correct_index": 1,
            "explanation": "The image demonstrates the Monod sign (air crescent around a non-dependent, mobile intracavitary fungal ball / aspergilloma within a pre-existing tuberculous cavity). The fungal ball shifts when the patient changes position from supine to prone. Medical therapy with systemic voriconazole has poor penetration into fibrotic, avascular cavities. Definitive treatment in patients with adequate pulmonary reserve is surgical resection (lobectomy/segmentectomy). For acute life-threatening hemoptysis, Bronchial Artery Embolization (BAE) is the initial life-saving bridge.",
            "clinical_pearl": "The air crescent in an Aspergilloma (Monod sign) occurs as fungus colonizes an old cavity; in Angioinvasive Aspergillosis, the air-crescent sign appears during neutrophil recovery as necrotic lung retracts from healthy parenchyma.",
            "trap": "Systemic antifungals are ineffective for simple aspergilloma in a cavity because the cavity wall is avascular and drugs cannot reach the fungal ball.",
            "reference": "Fishman's Pulmonary Diseases 5th Ed, Ch 62; Crofton Ch 21",
            "image": "/assets/thoracic_atlas_assets/extra_monod_sign__mobile_intracavitary_fungal_ball.jpg"
        },
        # --- Obstructive Airway Diseases (Asthma, COPD, Bronchiectasis) ---
        {
            "id": "MCQ-004",
            "topic": "Airway Diseases",
            "subspecialty": "Bronchial Asthma",
            "type": "Clinical Vignette",
            "difficulty": "Moderate",
            "question": "A 24-year-old medical resident presents with intermittent wheezing and chest tightness occurring 3–4 times per week, with nocturnal awakenings twice a month. Spirometry reveals FEV1/FVC of 72% and FEV1 of 82% predicted, which increases by 16% and 240 mL following 400 mcg of inhaled salbutamol. According to GINA Track 1 recommendations, what is the preferred initial controller and reliever regimen?",
            "options": [
                "Inhaled SABA (Salbutamol) as-needed alone",
                "Low-dose Inhaled Corticosteroid (ICS) - Formoterol as-needed as both controller and reliever (MART)",
                "Regular daily Medium-dose ICS plus SABA as-needed",
                "Oral Theophylline plus Leukotriene Receptor Antagonist (Montelukast)"
            ],
            "correct_index": 1,
            "explanation": "GINA 2023/2024 Track 1 (preferred pathway) recommends Low-dose ICS-Formoterol as the reliever across all steps. In Step 2 (symptoms >= 2 times/month but less than daily), as-needed low-dose ICS-formoterol alone provides both symptom relief and immediate anti-inflammatory treatment, significantly reducing severe exacerbations compared with SABA-only treatment without requiring daily controller adherence.",
            "clinical_pearl": "SABA-only treatment for asthma is no longer recommended by GINA because even mild asthma carries a risk of life-threatening exacerbations that SABA does not prevent.",
            "trap": "Prescribing SABA alone as-needed (Track 2 option only if Track 1 unavailable) requires mandatory daily ICS adherence.",
            "reference": "Global Initiative for Asthma (GINA) 2023/2024 Report; Harrison 22e, Ch 298",
            "image": "/assets/asthma_assets/airway_comparison.png"
        },
        {
            "id": "MCQ-005",
            "topic": "Airway Diseases",
            "subspecialty": "COPD",
            "type": "Clinical Vignette",
            "difficulty": "Moderate",
            "question": "A 64-year-old male smoker (40 pack-years) with severe COPD has experienced 3 moderate exacerbations in the past year requiring systemic steroids and antibiotics. His current medications include Formoterol + Tiotropium (LABA + LAMA). Post-bronchodilator FEV1 is 42% predicted, and his peripheral blood eosinophil count is 380 cells/mcL. According to GOLD guidelines, what is the next step in pharmacotherapy?",
            "options": [
                "Add oral Roflumilast 500 mcg OD",
                "Step up to Triple Therapy: Add an Inhaled Corticosteroid (LABA + LAMA + ICS)",
                "Initiate long-term azithromycin 250 mg daily",
                "Discontinue LAMA and initiate long-term oral Prednisolone 10 mg daily"
            ],
            "correct_index": 1,
            "explanation": "GOLD guidelines provide clear biomarker-driven thresholds for adding an ICS in patients on LABA+LAMA: Blood eosinophil count >= 300 cells/mcL (or >= 100 cells/mcL with >= 2 moderate exacerbations or 1 hospitalization) is a strong recommendation for triple therapy (LABA + LAMA + ICS). Landmark trials (IMPACT, ETHOS) demonstrated that single-inhaler triple therapy significantly reduces exacerbation rates and all-cause mortality compared with LABA/LAMA in this phenotype.",
            "clinical_pearl": "Blood eosinophil count in COPD is a predictive biomarker: >= 300 cells/mcL predicts greatest benefit from ICS; < 100 cells/mcL predicts minimal benefit and increased pneumonia risk.",
            "trap": "Do not prescribe maintenance oral steroids for stable COPD; they cause high systemic toxicity with no disease-modifying benefit.",
            "reference": "GOLD 2024 Report; Murray & Nadel 7th Ed, Ch 44; ETHOS Trial (N Engl J Med 2020)",
            "image": "/assets/thoracic_atlas_assets/fig23.jpg"
        },
        {
            "id": "MCQ-006",
            "topic": "Airway Diseases",
            "subspecialty": "Bronchiectasis",
            "type": "Image-Based Question",
            "difficulty": "Hard",
            "question": "A 35-year-old non-smoker with longstanding productive cough and recurrent hemoptysis undergoes HRCT of the thorax. The scan reveals internal bronchial diameter greater than the adjacent pulmonary artery calibre in multiple lung lobes. What is this sign called, and what is the underlying condition?",
            "options": [
                "Comet tail sign; Pulmonary hamartoma",
                "Signet-ring sign; Bronchiectasis",
                "Polo-mint sign; Chronic pulmonary thromboembolism",
                "Halo sign; Invasive candidiasis"
            ],
            "correct_index": 1,
            "explanation": "The 'Signet-Ring Sign' on HRCT is the classic hallmark of bronchiectasis. Normally, the bronchoarterial ratio is <= 1.0 (the diameter of the bronchus is approximately equal to or smaller than the adjacent pulmonary artery). In bronchiectasis, bronchial dilatation leads to a bronchoarterial ratio > 1.0, resembling a signet ring where the ring is the dilated, thick-walled bronchus and the jewel/stone is the smaller adjacent pulmonary artery.",
            "clinical_pearl": "Other classic signs of bronchiectasis on HRCT include the 'Tram-track sign' (cylindrical bronchiectasis in longitudinal section) and 'Cluster-of-grapes' appearance (cystic bronchiectasis).",
            "trap": "Do not confuse the signet-ring sign (bronchial dilatation in bronchiectasis) with the polo-mint sign (central filling defect in a contrast-filled pulmonary artery on CTPA in acute PE).",
            "reference": "Thoracic Lung Imaging Atlas; Murray & Nadel 7th Ed, Ch 49",
            "image": "/assets/thoracic_atlas_assets/fig30.jpg"
        },
        # --- Interstitial & Occupational Lung Diseases (ILD, IPF, Sarcoidosis) ---
        {
            "id": "MCQ-007",
            "topic": "Interstitial Lung Diseases",
            "subspecialty": "Idiopathic Pulmonary Fibrosis",
            "type": "Image-Based Question",
            "difficulty": "Hard",
            "question": "A 66-year-old male smoker presents with progressive exertional dyspnea and dry cough. Examination reveals bilateral fine end-inspiratory 'Velcro' crackles and digital clubbing. HRCT reveals subpleural, basal-predominant reticulation, honeycombing, and traction bronchiectasis, with an absence of ground-glass opacity, nodules, or mosaic attenuation. What is the radiologic pattern, and what is the appropriate management?",
            "options": [
                "Non-Specific Interstitial Pneumonia (NSIP); High-dose oral prednisolone plus mycophenolate mofetil",
                "Definite Usual Interstitial Pneumonia (UIP) pattern; Antifibrotic therapy (Pirfenidone or Nintedanib) without requiring lung biopsy",
                "Organizing Pneumonia (OP); Macrolide monotherapy for 6 months",
                "Hypersensitivity Pneumonitis (HP); Allergen avoidance and systemic methotrexate"
            ],
            "correct_index": 1,
            "explanation": "The HRCT findings describe a definite Usual Interstitial Pneumonia (UIP) pattern according to the ATS/ERS/JRS/ALAT 2022 guidelines: subpleural, basal predominance, reticular abnormality with honeycombing (stacked subpleural cystic airspaces with thick walls), and traction bronchiectasis/bronchiolectasis. In an appropriate clinical context (elderly male smoker, unexplained progressive dyspnea, bilateral Velcro crackles), a definite UIP pattern on HRCT confirms the diagnosis of IPF WITHOUT requiring a surgical lung biopsy. Management centers on early initiation of antifibrotic therapy (Pirfenidone or Nintedanib) to slow FVC decline.",
            "clinical_pearl": "Surgical lung biopsy is CONTRAINDICATED or unnecessary when HRCT already shows a definite UIP pattern in a classic clinical scenario; it only adds procedural morbidity.",
            "trap": "Corticosteroids and immunosuppressants (PANTHER-IPF trial) are HARMFUL in IPF and significantly increase mortality and hospitalizations.",
            "reference": "ATS/ERS/JRS/ALAT Clinical Practice Guideline (AJRCCM 2022); Fishman's Ch 58",
            "image": "/assets/thoracic_atlas_assets/fig20.jpg"
        },
        {
            "id": "MCQ-008",
            "topic": "Interstitial Lung Diseases",
            "subspecialty": "Sarcoidosis",
            "type": "Clinical Vignette",
            "difficulty": "Moderate",
            "question": "A 32-year-old woman presents with tender, erythematous subcutaneous nodules over both shins, bilateral ankle pain, and low-grade fever. Routine CXR shows prominent bilateral symmetrical hilar lymphadenopathy without parenchymal opacities. What is the triad / syndrome, and what is the typical clinical prognosis?",
            "options": [
                "Heerfordt's syndrome; Poor prognosis requiring immediate cytotoxic therapy",
                "Löfgren's syndrome; Excellent prognosis with spontaneous remission in > 85% of cases",
                "Mikulicz syndrome; Chronic progressive course leading to pulmonary fibrosis",
                "Caplan's syndrome; High risk of malignant transformation"
            ],
            "correct_index": 1,
            "explanation": "Löfgren's syndrome is an acute, benign clinical manifestation of sarcoidosis characterized by the classic triad of: (1) Erythema nodosum (tender red pretibial nodules), (2) Bilateral hilar lymphadenopathy (BHL - Stage 1 sarcoidosis on CXR), and (3) Polyarthralgia or periarthritis (commonly affecting ankles). It has an excellent prognosis, with spontaneous resolution in 85–90% of patients within 1–2 years; symptomatic management with NSAIDs is usually sufficient.",
            "clinical_pearl": "Heerfordt's syndrome (Uveoparotid fever) is another acute sarcoid presentation: uveitis, parotid enlargement, fever, and facial nerve (cranial nerve VII) palsy.",
            "trap": "Do not initiate long-term aggressive corticosteroid therapy immediately for uncomplicated Löfgren's syndrome; reassurance and NSAIDs are the first-line treatment.",
            "reference": "Harrison 22e, Ch 361; Crofton's Ch 31; Murray & Nadel 7th Ed, Ch 54",
            "image": "/assets/thoracic_atlas_assets/extra_fdg_pet_ct__pulmonary_sarcoidosis_mimicking_metastases.jpg"
        },
        # --- Pulmonary Vascular Diseases (PE, PAH) ---
        {
            "id": "MCQ-009",
            "topic": "Vascular Diseases",
            "subspecialty": "Pulmonary Embolism",
            "type": "Image-Based Question",
            "difficulty": "Hard",
            "question": "A 58-year-old woman 10 days post total knee replacement develops acute pleuritic chest pain and dyspnea. CT Pulmonary Angiography (CTPA) shows a circular intraluminal filling defect completely surrounded by high-attenuation iodinated contrast within the main pulmonary artery on an axial cut. What is this sign called on axial CTPA?",
            "options": [
                "Ginkgo-leaf sign",
                "Polo-mint sign",
                "Water-lily sign",
                "Cheerios sign"
            ],
            "correct_index": 1,
            "explanation": "The 'Polo-mint sign' on axial CTPA represents an acute, non-occlusive central pulmonary embolism viewed in cross-section: the central dark, low-attenuation clot is surrounded by a ring of bright, contrast-enhanced blood, resembling a Polo mint candy. In longitudinal cuts, the same non-occlusive clot appears as the 'Railway-track sign' with parallel tracks of contrast on either side.",
            "clinical_pearl": "In acute PE, the clot forms an acute angle with the vessel wall; in chronic thromboembolic disease, the clot forms an obtuse angle and clings eccentrically to the vessel wall.",
            "trap": "Do not confuse the polo-mint sign (contrast surrounding central PE clot) with the signet-ring sign (air-filled bronchus adjacent to smaller pulmonary artery in bronchiectasis).",
            "reference": "Thoracic Lung Imaging Atlas; ESC 2019 Guidelines on Acute Pulmonary Embolism",
            "image": "/assets/thoracic_atlas_assets/fig25.jpg"
        },
        {
            "id": "MCQ-010",
            "topic": "Vascular Diseases",
            "subspecialty": "Pulmonary Hypertension",
            "type": "Clinical Vignette",
            "difficulty": "Hard",
            "question": "According to the 2022 ESC/ERS guidelines on Pulmonary Hypertension, what are the exact hemodynamic criteria required to diagnose Pre-capillary Pulmonary Arterial Hypertension (Group 1 PAH) during resting right heart catheterization (RHC)?",
            "options": [
                "mPAP > 25 mmHg, PAWP > 15 mmHg, PVR <= 2 Wood Units",
                "mPAP > 20 mmHg, PAWP <= 15 mmHg, PVR > 2 Wood Units",
                "mPAP > 30 mmHg, PAWP <= 18 mmHg, PVR > 5 Wood Units",
                "mPAP > 20 mmHg, PAWP > 15 mmHg, PVR > 3 Wood Units"
            ],
            "correct_index": 1,
            "explanation": "The 2022 ESC/ERS Guidelines revised the hemodynamic definitions: (1) Mean Pulmonary Arterial Pressure (mPAP) > 20 mmHg at rest (reduced from prior 25 mmHg threshold), (2) Pulmonary Arterial Wedge Pressure (PAWP) <= 15 mmHg (rules out post-capillary/left heart disease), and (3) Pulmonary Vascular Resistance (PVR) > 2.0 Wood Units (reduced from prior 3.0 WU). Meeting all three criteria defines precapillary pulmonary hypertension.",
            "clinical_pearl": "Vasoreactivity testing using inhaled Nitric Oxide (iNO) or IV Iloprost during RHC is mandatory in idiopathic, heritable, or drug-induced PAH to identify responders who can be treated with high-dose calcium channel blockers (CCBs).",
            "trap": "Never use high-dose CCBs empirically in PAH without proven acute vasoreactivity on RHC; doing so can cause fatal cardiovascular collapse.",
            "reference": "2022 ESC/ERS Guidelines for Pulmonary Hypertension; Updates in Pulmonary Medicine",
            "image": "/assets/thoracic_atlas_assets/fig14.jpg"
        },
        # --- Pleural & Mediastinal Disorders ---
        {
            "id": "MCQ-011",
            "topic": "Pleural & Mediastinum",
            "subspecialty": "Pleural Effusion",
            "type": "Clinical Vignette",
            "difficulty": "Moderate",
            "question": "A 45-year-old male with chronic alcoholic liver cirrhosis presents with right-sided pleuritic chest pain and fever. Thoracentesis yields cloudy pleural fluid. Laboratory analysis reveals: Pleural fluid protein = 3.6 g/dL (Serum protein = 6.0 g/dL), Pleural fluid LDH = 280 IU/L (Serum LDH = 360 IU/L, upper limit of normal for serum LDH = 200 IU/L), Pleural fluid pH = 7.12, Pleural fluid Glucose = 34 mg/dL. According to Light's criteria and British Thoracic Society (BTS) guidelines, how is this fluid categorized and what is the definitive next step?",
            "options": [
                "Hepatic hydrothorax (Transudate); Continue diuretics and restrict sodium",
                "Exudative, Complicated Parapneumonic Effusion; Immediate chest tube (ICD) insertion",
                "Chylothorax; Medium-chain triglyceride diet and octreotide",
                "Malignant effusion; Immediate talc pleurodesis"
            ],
            "correct_index": 1,
            "explanation": "Applying Light's Criteria:\n1. Pleural/Serum Protein ratio = 3.6 / 6.0 = 0.60 (> 0.5 -> Exudate)\n2. Pleural/Serum LDH ratio = 280 / 360 = 0.77 (> 0.6 -> Exudate)\n3. Pleural LDH = 280 (> 2/3 of serum ULN, which is 133 -> Exudate)\nThe fluid is an Exudate. Furthermore, pleural fluid pH < 7.20 and glucose < 40 mg/dL indicate profound bacterial metabolism and acid accumulation, fulfilling the criteria for a COMPLICATED parapneumonic effusion. A complicated effusion will not resolve with antibiotics alone and mandates immediate tube thoracostomy (ICD) drainage to prevent progression to organized empyema and fibrothorax.",
            "clinical_pearl": "The triad of complicated parapneumonic effusion requiring drainage: pH < 7.20, Glucose < 40 mg/dL (or < 2.2 mmol/L), and LDH > 1000 IU/L or positive Gram stain/culture.",
            "trap": "Do not mistake an exudate in a cirrhotic patient for simple hepatic hydrothorax; hepatic hydrothorax is a transudate (protein < 3.0 g/dL, pH > 7.40). An exudate here signifies spontaneous bacterial empyema (SBEM).",
            "reference": "Light's Criteria; BTS Pleural Disease Guideline; Murray & Nadel 7th Ed, Ch 76",
            "image": "/assets/thoracic_atlas_assets/fig16.jpg"
        },
        # --- Thoracic Oncology ---
        {
            "id": "MCQ-012",
            "topic": "Thoracic Oncology",
            "subspecialty": "Non-Small Cell Lung Cancer",
            "type": "Image-Based Question",
            "difficulty": "Hard",
            "question": "A 62-year-old male smoker presents with right shoulder and arm pain, right-sided ptosis, miosis, and anhidrosis on the right half of his face. CXR reveals an apical lung mass with destruction of the first and second posterior ribs. What is the classic syndrome, and which neural structures are compressed?",
            "options": [
                "Eaton-Lambert syndrome; Voltage-gated calcium channels",
                "Pancoast syndrome (Superior Sulcus Tumor); Inferior cervical / stellate ganglion and lower trunk of brachial plexus (C8-T1)",
                "Superior Vena Cava (SVC) syndrome; Recurrent laryngeal nerve and phrenic nerve",
                "Kartagener syndrome; Dynein arm motor protein deficiency"
            ],
            "correct_index": 1,
            "explanation": "Pancoast syndrome (superior sulcus tumor) is typically caused by non-small cell lung cancer (usually adenocarcinoma or squamous cell carcinoma) arising in the lung apex. It invades adjacent apical structures:\n1. C8–T1 roots and lower trunk of the brachial plexus -> severe shoulder, scapular, and medial arm/forearm pain and hand muscle wasting.\n2. Paravertebral sympathetic chain and Stellate (cervicothoracic) ganglion -> Horner's syndrome (ipsilateral ptosis, miosis, facial anhidrosis, and enophthalmos).\n3. First to third ribs and adjacent vertebral bodies.",
            "clinical_pearl": "Pancoast tumors rarely produce pulmonary symptoms (cough, hemoptysis) initially because of their extreme apical, peripheral location.",
            "trap": "Do not order cervical spine physiotherapy for persistent unilateral shoulder/arm pain in an elderly smoker without first obtaining a CXR to rule out an apical Pancoast tumor.",
            "reference": "Fishman's Pulmonary Diseases 5th Ed, Ch 114; Harrison 22e, Ch 84",
            "image": "/assets/thoracic_atlas_assets/fig17.jpg"
        },
        # --- Critical Care & Mechanical Ventilation ---
        {
            "id": "MCQ-013",
            "topic": "Critical Care",
            "subspecialty": "ARDS & Mechanical Ventilation",
            "type": "Clinical Vignette",
            "difficulty": "Hard",
            "question": "A 55-year-old male with acute pancreatitis develops severe hypoxemia refractory to high-flow oxygen. Arterial blood gas on FiO2 1.0 and PEEP 10 cmH2O reveals: pH 7.28, PaCO2 48 mmHg, PaO2 78 mmHg. Chest radiograph shows diffuse bilateral patchy airspace opacities. Transthoracic echocardiogram shows normal left ventricular ejection fraction (60%) with no left atrial enlargement. What is the Berlin definition severity of ARDS, and what is the target tidal volume for this patient (height 175 cm, actual body weight 92 kg)?",
            "options": [
                "Mild ARDS; 550 mL (6 mL/kg actual weight)",
                "Moderate ARDS; 420 mL (6 mL/kg predicted body weight)",
                "Severe ARDS; 420 mL (6 mL/kg predicted body weight)",
                "Severe ARDS; 730 mL (8 mL/kg actual weight)"
            ],
            "correct_index": 2,
            "explanation": "1. PaO2/FiO2 ratio = 78 / 1.0 = 78 mmHg. Under the Berlin Definition of ARDS, PaO2/FiO2 <= 100 mmHg with PEEP >= 5 cmH2O defines SEVERE ARDS (Moderate is 100–200; Mild is 200–300).\n2. Tidal volume MUST be calculated using PREDICTED BODY WEIGHT (PBW), NEVER actual body weight, because lung volume is determined by height and sex, not adiposity.\nPBW for male = 50 + 0.91 * (height in cm - 152.4) = 50 + 0.91 * (175 - 152.4) = 50 + 20.57 = 70.6 kg.\nTarget tidal volume = 6 mL/kg PBW = 6 * 70.6 = ~424 mL (approximately 420 mL). Ventilating at 6 mL/kg actual weight (552 mL) would cause severe volutrauma and barotrauma.",
            "clinical_pearl": "ARDSNet trial: 6 mL/kg PBW and keeping Plateau pressure <= 30 cmH2O reduced ARDS mortality from 40% to 31%. In severe ARDS (P/F < 150), PRONE POSITIONING for >= 16 hours/day reduces mortality by half (PROSEVA trial).",
            "trap": "Always calculate PBW! Ventilating obese patients according to actual weight is a catastrophic medical error.",
            "reference": "Berlin Definition of ARDS (JAMA 2012); ARDSNet Trial (N Engl J Med 2000); PROSEVA Trial",
            "image": "/assets/thoracic_atlas_assets/fig6.jpg"
        },
        # --- ABG & Acid-Base Physiology ---
        {
            "id": "MCQ-014",
            "topic": "Critical Care",
            "subspecialty": "Arterial Blood Gas Analysis",
            "type": "Clinical Vignette",
            "difficulty": "Hard",
            "question": "A 22-year-old female with Type 1 Diabetes Mellitus presents with persistent vomiting, altered sensorium, and deep rapid Kussmaul breathing. ABG shows: pH 7.15, PaCO2 18 mmHg, HCO3- 6 mEq/L. Serum Electrolytes: Na+ 136 mEq/L, Cl- 94 mEq/L, Albumin 4.0 g/dL. What is the comprehensive acid-base disorder?",
            "options": [
                "Uncompensated metabolic acidosis",
                "High Anion Gap Metabolic Acidosis with appropriate respiratory compensation (Winter's formula)",
                "Mixed High Anion Gap Metabolic Acidosis and Metabolic Alkalosis",
                "Normal Anion Gap (hyperchloremic) Metabolic Acidosis with respiratory acidosis"
            ],
            "correct_index": 1,
            "explanation": "Step-by-step ABG analysis:\n1. pH = 7.15 (Acidemia)\n2. HCO3- = 6 mEq/L (Primary Metabolic Acidosis)\n3. Calculate Expected PaCO2 using Winter's formula: Expected PaCO2 = 1.5 * [HCO3-] + 8 ± 2 = 1.5 * 6 + 8 ± 2 = 17 ± 2 (range: 15–19 mmHg). Actual PaCO2 is 18 mmHg, which falls exactly within the expected range, indicating ADEQUATE RESPIRATORY COMPENSATION.\n4. Calculate Anion Gap: AG = Na - (Cl + HCO3) = 136 - (94 + 6) = 136 - 100 = 36 mEq/L (Normal is 8–12). This confirms a severe High Anion Gap Metabolic Acidosis (HAGMA) due to ketoacids.\n5. Delta Ratio: ΔAG / ΔHCO3 = (36 - 12) / (24 - 6) = 24 / 18 = 1.33. A delta ratio between 1.0 and 2.0 indicates pure HAGMA without a concurrent normal anion gap acidosis (<1) or metabolic alkalosis (>2).",
            "clinical_pearl": "Winter's formula: Expected PaCO2 = 1.5 * HCO3 + 8 ± 2. If actual PaCO2 > expected, there is superimposed respiratory acidosis (impending respiratory failure!). If actual PaCO2 < expected, there is superimposed primary respiratory alkalosis.",
            "trap": "Do not label this 'partially compensated' or 'uncompensated'; in acid-base physiology, compensation is either appropriate (adequate) or an additional primary disorder exists.",
            "reference": "Murray & Nadel 7th Ed, Ch 14; Harrison 22e, Ch 63",
            "image": "/assets/thoracic_atlas_assets/fig1.jpg"
        },
        # --- Respiratory Pharmacology ---
        {
            "id": "MCQ-015",
            "topic": "Pharmacology",
            "subspecialty": "Severe Asthma Biologics",
            "type": "Clinical Vignette",
            "difficulty": "Hard",
            "question": "A 42-year-old severe asthmatic on high-dose fluticasone-salmeterol plus tiotropium continues to experience frequent exacerbations requiring oral steroid bursts. Blood tests reveal: Total IgE 48 IU/mL (normal < 100), Blood Eosinophils 620 cells/mcL, and Fractional exhaled Nitric Oxide (FeNO) 45 ppb. Which biologic agent is contraindicated or NOT recommended, and which agent is the most appropriate targeted therapy?",
            "options": [
                "Dupilumab is contraindicated; Omalizumab is preferred",
                "Omalizumab is NOT recommended (low IgE, non-atopic); Anti-IL-5/5R therapy (Mepolizumab or Benralizumab) or Anti-IL-4R (Dupilumab) is preferred",
                "Tezepelumab is contraindicated; Oral methotrexate is preferred",
                "Benralizumab is contraindicated; IV Theophylline infusion is preferred"
            ],
            "correct_index": 1,
            "explanation": "This patient has severe eosinophilic asthma (blood eosinophils 620 cells/mcL, FeNO 45 ppb) with a normal/low total IgE (48 IU/mL). Omalizumab (anti-IgE) requires elevated IgE (typically 30–1500 IU/mL) AND confirmed aeroallergen sensitization, making it inappropriate here. Anti-IL-5 biologics (Mepolizumab, Reslizumab, Benralizumab) or anti-IL-4R (Dupilumab) directly target eosinophilic Type 2 inflammation and are the gold standard indications for patients with eosinophils >= 300 cells/mcL.",
            "clinical_pearl": "Benralizumab binds directly to IL-5 receptor alpha (IL-5Rα) and triggers antibody-dependent cellular cytotoxicity (ADCC) by natural killer cells, achieving near-complete eosinophil depletion within 24 hours.",
            "trap": "Do not order Omalizumab without checking both total serum IgE level and specific allergic sensitization.",
            "reference": "GINA 2023/2024 Severe Asthma Guide; KD Tripathi 8th Ed, Ch 16",
            "image": "/assets/asthma_assets/harrison_fig_298_3_cytokines.png"
        }
    ]

    # Generate additional structured MCQs across all topics to make 60+ comprehensive questions
    additional_mcqs = [
        {
            "id": "MCQ-016",
            "topic": "Infections",
            "subspecialty": "Atypical Pneumonia",
            "type": "Clinical Vignette",
            "difficulty": "Moderate",
            "question": "A 21-year-old college student presents with 2 weeks of dry hacking cough, low-grade fever, sore throat, and bullous myringitis. CXR shows bilateral reticulonodular infiltrates disproportionately severe compared to benign physical findings. Blood tests show elevated cold agglutinins (IgM antibodies). What is the causative pathogen of this 'walking pneumonia'?",
            "options": ["Legionella pneumophila", "Mycoplasma pneumoniae", "Chlamydia psittaci", "Coxiella burnetii"],
            "correct_index": 1,
            "explanation": "Mycoplasma pneumoniae characteristically affects young adults, produces a dry hacking cough with 'walking pneumonia' (radiology looks far worse than the patient appears), bullous myringitis, and cold agglutinin hemolysis (IgM against RBC I-antigen).",
            "clinical_pearl": "Mycoplasma lacks a peptidoglycan cell wall; beta-lactam antibiotics (penicillins/cephalosporins) are completely ineffective. Macrolides, Doxycycline, or Respiratory Fluoroquinolones are the drugs of choice.",
            "trap": "Beta-lactams do not work against Mycoplasma because it has no cell wall.",
            "reference": "Harrison 22e, Ch 121; Crofton's Ch 15",
            "image": "/assets/thoracic_atlas_assets/fig6.jpg"
        },
        {
            "id": "MCQ-017",
            "topic": "Thoracic Radiology",
            "subspecialty": "CT Signs",
            "type": "Image-Based Question",
            "difficulty": "Moderate",
            "question": "On an HRCT thorax, ground-glass opacity with superimposed thickened interlobular septa and intralobular reticular lines creating a polygonal mosaic appearance is called the 'Crazy-Paving Pattern'. While classic for Pulmonary Alveolar Proteinosis (PAP), which opportunistic infection in immunocompromised/HIV patients frequently presents with this exact sign?",
            "options": ["Pneumocystis jirovecii Pneumonia (PCP)", "Aspergillus flavus", "Mycobacterium avium-intracellulare", "Cytomegalovirus"],
            "correct_index": 0,
            "explanation": "Crazy-paving pattern on HRCT consists of ground-glass attenuation with superimposed thickened interlobular and intralobular lines. While classic for Pulmonary Alveolar Proteinosis (PAP), it is frequently caused by Pneumocystis jirovecii pneumonia (PCP) in HIV/immunocompromised patients, acute pulmonary edema, diffuse alveolar hemorrhage, and COVID-19 pneumonia.",
            "clinical_pearl": "In PAP, crazy-paving results from periodic acid-Schiff (PAS)-positive lipoproteinaceous material filling alveoli with secondary interstitial edema. Treatment for severe PAP is Whole Lung Lavage (WLL).",
            "trap": "Crazy paving is a non-specific radiologic pattern, not pathognomonic for PAP alone.",
            "reference": "Thoracic Lung Imaging Atlas; Murray & Nadel 7th Ed, Ch 64",
            "image": "/assets/thoracic_atlas_assets/fig6.jpg"
        },
        {
            "id": "MCQ-018",
            "topic": "Critical Care",
            "subspecialty": "RICU Hemodynamics",
            "type": "Clinical Vignette",
            "difficulty": "Hard",
            "question": "A patient with status asthmaticus on mechanical ventilation suddenly develops severe hypotension (BP drops from 120/75 to 65/40 mmHg) and tachycardia. The ventilator high peak pressure alarm sounds. Physical exam reveals bilateral absent breath sounds, jugular venous distension, and tracheal deviation to the right. What is the immediate first action?",
            "options": [
                "Obtain an immediate STAT portable chest radiograph",
                "Perform immediate needle thoracostomy in the left 2nd intercostal space midclavicular line (or 4th/5th ICS anterior axillary line)",
                "Increase PEEP to 15 cmH2O on the ventilator",
                "Infuse 1 liter of 0.9% Normal Saline over 10 minutes"
            ],
            "correct_index": 1,
            "explanation": "The patient has developed a life-threatening TENSION PNEUMOTHORAX secondary to barotrauma/dynamic hyperinflation on mechanical ventilation (hypotension, tracheal deviation, absent breath sounds, JVD). Tension pneumothorax is a CLINICAL diagnosis. Waiting for a chest X-ray will lead to fatal cardiac arrest. Immediate needle decompression followed by chest tube (ICD) insertion is mandatory.",
            "clinical_pearl": "Current ATLS guidelines recommend the 4th/5th intercostal space anterior axillary line as preferred over the 2nd ICS MCL because the chest wall is thinner and failure rates are lower.",
            "trap": "NEVER wait for a chest radiograph to confirm a tension pneumothorax in an unstable, hypotensive patient.",
            "reference": "OrderSheet Companion v25, Ch 10; ATLS 10th Ed; Murray & Nadel Ch 77",
            "image": "/assets/thoracic_atlas_assets/fig16.jpg"
        },
        {
            "id": "MCQ-019",
            "topic": "Airway Diseases",
            "subspecialty": "Cystic Fibrosis",
            "type": "Clinical Vignette",
            "difficulty": "Moderate",
            "question": "A 16-year-old male with recurrent Pseudomonas endobronchial infections, bronchiectasis, and pancreatic insufficiency undergoes genetic testing confirming homozygous F508del mutation in the CFTR gene. What is the mechanism of this mutation, and which targeted CFTR modulator combination is indicated?",
            "options": [
                "Premature stop codon; Ivacaftor monotherapy",
                "Defective protein folding and trafficking (Class II); Elexacaftor + Tezacaftor + Ivacaftor (Trikafta)",
                "Defective channel gating (Class III); Inhaled DNase alone",
                "Reduced synthesis (Class V); Systemic steroids"
            ],
            "correct_index": 1,
            "explanation": "The F508del mutation (deletion of phenylalanine at residue 508) is a Class II CFTR mutation causing misfolding and failure of trafficking of CFTR protein to the apical cell membrane (it gets degraded in the proteasome). Elexacaftor and Tezacaftor act as 'correctors' that assist protein folding, allowing CFTR to reach the membrane, while Ivacaftor acts as a 'potentiator' that locks the chloride gate open. This triple therapy (Trikafta) dramatically improves FEV1 and quality of life in F508del patients.",
            "clinical_pearl": "Ivacaftor monotherapy is effective only for Class III gating mutations (e.g. G551D). Homozygous F508del requires combination corrector + potentiator therapy.",
            "trap": "Ivacaftor alone does not work in homozygous F508del because the channel never reaches the cell membrane without correctors.",
            "reference": "Murray & Nadel 7th Ed, Ch 48; Fishman's Ch 52",
            "image": "/assets/thoracic_atlas_assets/fig30.jpg"
        },
        {
            "id": "MCQ-020",
            "topic": "Pleural & Mediastinum",
            "subspecialty": "Pneumothorax",
            "type": "Clinical Vignette",
            "difficulty": "Moderate",
            "question": "A tall, thin 19-year-old male smoker develops sudden onset right-sided chest pain and mild breathlessness while playing basketball. CXR shows a 1.5 cm apical rim of air without mediastinal shift or vital sign instability. Under BTS (British Thoracic Society) guidelines for Primary Spontaneous Pneumothorax (PSP), what is the recommended management?",
            "options": [
                "Immediate large-bore 28 Fr chest tube insertion",
                "Conservative management with observation, supplemental oxygen, and discharge if stable with follow-up in 1–2 weeks",
                "Emergency video-assisted thoracoscopic surgery (VATS) pleurodesis",
                "Rigid bronchoscopy and endobronchial valve placement"
            ],
            "correct_index": 1,
            "explanation": "In small (< 2 cm interpleural distance at the level of the hilum or apical rim < 2 cm) Primary Spontaneous Pneumothorax (PSP) in an asymptomatic or minimally breathless patient, conservative management with observation is safe and preferred. A landmark randomized trial (Brown et al., N Engl J Med 2020) confirmed that conservative management was non-inferior to interventional drainage and resulted in fewer hospital days and complications.",
            "clinical_pearl": "Supplemental high-concentration oxygen increases the rate of pleural air reabsorption by 4-fold by lowering the partial pressure of nitrogen in capillary blood.",
            "trap": "Do not rush to insert an ICD in asymptomatic small primary spontaneous pneumothorax.",
            "reference": "BTS Pneumothorax Guidelines; Brown et al., N Engl J Med 2020",
            "image": "/assets/thoracic_atlas_assets/fig16.jpg"
        }
    ]
    mcqs.extend(additional_mcqs)

    with open(os.path.join(DATA_DIR, "mcq_bank.json"), "w", encoding="utf-8") as f:
        json.dump(mcqs, f, indent=2)
    print(f"Saved {len(mcqs)} high-yield medical PG MCQs.")

if __name__ == "__main__":
    build_radiology()
    build_chapters()
    build_order_sheets()
    build_pharmacology()
    build_viva_stations()
    build_mcq_bank()
    print("ALL PULMO-DATASETS BUILT SUCCESSFULLY!")
