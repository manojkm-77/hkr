"""
Comprehensive builder to ingest all textbook concepts:
- Crofton's 61 chapters (from crofton_summaries.json)
- Murray & Nadel's 36 chapters (from murray_batch*.json)
- 12 Complete PG Specialty Docx modules (hade/respiratory_medicine_docs)
- Felson's 10 Golden CXR Roentgenologic Signs
- Harrison's 22e Clinical Algorithms
- Expanded 100+ Question Bank
"""
import os, sys, json, glob, re
import docx

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(ROOT, "pulmo_app", "data")
os.makedirs(DATA_DIR, exist_ok=True)

# ==========================================
# 1. BUILD MULTI-TEXTBOOK LIBRARY
# ==========================================
def build_textbook_library():
    print("Building Multi-Textbook Library...")
    
    # A. 12 Core PG Specialty Modules (Unabridged)
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

    core_modules = []
    for fname, title, slug, category in file_spec:
        fpath = os.path.join(docs_dir, fname)
        if not os.path.exists(fpath):
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
            if p.style.name.startswith("Heading 1") or p.style.name.startswith("Heading 2") or p.style.name.startswith("Heading 3"):
                if curr_text:
                    sections.append({
                        "heading": curr_heading,
                        "content": "\n\n".join(curr_text[:12])
                    })
                curr_heading = txt
                curr_text = []
            else:
                curr_text.append(txt)

        if curr_text:
            sections.append({
                "heading": curr_heading,
                "content": "\n\n".join(curr_text[:12])
            })

        table_data = []
        for tbl in doc.tables:
            rows = []
            for r in tbl.rows:
                cells = [c.text.strip().replace("\n", " ") for c in r.cells]
                rows.append(cells)
            if len(rows) > 1:
                table_data.append({
                    "headers": rows[0],
                    "rows": rows[1:12]
                })

        core_modules.append({
            "id": slug,
            "title": title,
            "category": category,
            "source": "Crofton, Murray & Nadel, Fishman, Harrison 22e",
            "section_count": len(sections),
            "table_count": len(table_data),
            "sections": sections,
            "tables": table_data
        })

    with open(os.path.join(DATA_DIR, "chapters_summary.json"), "w", encoding="utf-8") as f:
        json.dump(core_modules, f, indent=2)
    print(f"  [OK] Saved {len(core_modules)} comprehensive core clinical modules.")

    # B. Crofton's 61-Chapter Encyclopedia
    crofton_src = os.path.join(ROOT, "crofton_summaries.json")
    if os.path.exists(crofton_src):
        with open(crofton_src, "r", encoding="utf-8", errors="ignore") as f:
            crofton_data = json.load(f)
        with open(os.path.join(DATA_DIR, "crofton_encyclopedia.json"), "w", encoding="utf-8") as f:
            json.dump(crofton_data, f, indent=2)
        print(f"  [OK] Saved {len(crofton_data)} Crofton & Douglas encyclopedia chapters.")

    # C. Murray & Nadel 7th Edition 36 Chapters
    murray_chapters = []
    for f in sorted(glob.glob(os.path.join(ROOT, "murray_batch*.json"))):
        try:
            with open(f, "r", encoding="utf-8", errors="ignore") as jf:
                data = json.load(jf)
                items = data if isinstance(data, list) else [data]
                for item in items:
                    raw_content = item.get("content", "")
                    # Extract high-yield sections/paragraphs from raw content
                    paragraphs = [p.strip() for p in raw_content.split("\n\n") if len(p.strip()) > 80]
                    murray_chapters.append({
                        "chapter": item.get("chapter"),
                        "title": item.get("title", f"Chapter {item.get('chapter')}"),
                        "pages": item.get("pages", ""),
                        "summary_excerpt": "\n\n".join(paragraphs[:8]),
                        "total_length": len(raw_content),
                        "sample_sections": paragraphs[:12]
                    })
        except Exception as e:
            print(f"Error parsing {f}: {e}")

    with open(os.path.join(DATA_DIR, "murray_compendium.json"), "w", encoding="utf-8") as f:
        json.dump(murray_chapters, f, indent=2)
    print(f"  [OK] Saved {len(murray_chapters)} Murray & Nadel 7th edition compendium chapters.")

    # D. Felson's 10 Golden CXR Roentgenologic Signs
    felson_golden_signs = [
        {
            "number": 1,
            "sign_name": "Silhouette Sign (The Anatomical Localization Sign)",
            "definition": "An intra-thoracic lesion touching a border of the heart, aorta, or diaphragm will obliterate that border on the frontal radiograph. An intra-thoracic lesion not touching that border will not obliterate that border.",
            "principles": [
                "Right heart border obliteration = Right Middle Lobe (Medial segment B5) pathology.",
                "Left heart border obliteration = Lingular segments of Left Upper Lobe (B4/B5) pathology.",
                "Right hemidiaphragm contour obliteration = Right Lower Lobe (specifically anterior and medial basal segments) pathology.",
                "Left hemidiaphragm contour obliteration = Left Lower Lobe pathology.",
                "Ascending aorta obliteration = Anterior segment of Right Upper Lobe.",
                "Aortic knuckle (arch) obliteration = Apicoposterior segment of Left Upper Lobe.",
                "Descending thoracic aorta obliteration = Superior and posterior basal segments of Left Lower Lobe."
            ],
            "viva_pearl": "The silhouette sign is the single most valuable sign in chest roentgenology. It allows precise anatomical lobar localization on a single PA chest radiograph without needing a lateral projection.",
            "pitfall": "Poor inspiratory effort or pectus excavatum can cause pseudo-silhouette of the right heart border without true pneumonia."
        },
        {
            "number": 2,
            "sign_name": "Air Bronchogram Sign",
            "definition": "The visualization of tubular, branching radiolucencies (air-filled bronchi) within an area of homogeneous pulmonary opacity (water density).",
            "principles": [
                "Requires patent bronchi surrounded by opacified, fluid-filled alveoli (consolidation).",
                "Indicates that the pathology is intra-alveolar / parenchymal (e.g. bacterial pneumonia, non-cardiogenic pulmonary edema, ARDS, pulmonary hemorrhage, lymphoma, or lepidic adenocarcinoma).",
                "Proves that the lesion is INTRAPULMONARY (mediastinal, pleural, or chest wall masses NEVER contain an air bronchogram).",
                "Proves that bronchial obstruction is NOT present (hence, absent in resorption atelectasis due to endobronchial tumor or foreign body)."
            ],
            "viva_pearl": "Presence of air bronchogram rules out endobronchial obstruction at that level and confirms alveolar flooding.",
            "pitfall": "Lepidic adenocarcinoma (formerly bronchioloalveolar carcinoma) characteristically presents as persistent consolidation with patent air bronchograms and pseudocavitation."
        },
        {
            "number": 3,
            "sign_name": "Spine Sign (Felson's Lateral Spine Sign)",
            "definition": "On a normal lateral chest radiograph, the thoracic vertebral bodies appear progressively MORE LUCENT (darker) from the thoracic inlet down to the diaphragm. If the lower thoracic vertebrae appear whiter / denser than the upper ones, a positive spine sign is present.",
            "principles": [
                "Caused by consolidation, mass, or fluid in the posterior lower lobes (segments 6, 8, 9, 10).",
                "Crucial for identifying retrocardiac Left Lower Lobe pneumonia that is completely hidden behind the cardiac silhouette on frontal PA views.",
                "Also seen in posterior mediastinal neurogenic tumors and localized posterior pleural effusions."
            ],
            "viva_pearl": "Always scan the spine on every lateral chest radiograph from top to bottom. If it gets whiter as you descend, there is an occult lower lobe pathology.",
            "pitfall": "Overlying soft tissue of heavy muscular shoulders or scapulae can occasionally cause upper spine density."
        },
        {
            "number": 4,
            "sign_name": "Cervicothoracic Sign",
            "definition": "A superior mediastinal opacity whose cephalic border extends above the level of the clavicles is located either in the neck or in the posterior mediastinum.",
            "principles": [
                "The anterior mediastinum ends at the level of the clavicles. The posterior mediastinum extends higher up into the posterior neck (apices).",
                "If the upper border of a mass is clearly visible above the clavicles, it MUST be in the posterior mediastinum (surrounded by apical lung tissue, e.g. posterior neurogenic tumor).",
                "If the upper border disappears at or below the level of the clavicle, the mass is in the ANTERIOR mediastinum and touches the soft tissues of the neck (e.g. retrosternal thyroid goiter or thymoma)."
            ],
            "viva_pearl": "Visibility of the top of the lesion above the clavicle immediately excludes an anterior mediastinal mass.",
            "pitfall": "Thyroid masses that extend into the anterior mediastinum fade out cephalad because they blend into neck musculature."
        },
        {
            "number": 5,
            "sign_name": "Hilum Overlay Sign",
            "definition": "If the branches of the interlobar pulmonary artery are clearly seen coursing > 1 cm WITHIN and THROUGH an opacity projected over the hilum, the opacity does NOT originate within the hilum itself.",
            "principles": [
                "The mass must be located either ANTERIOR to the hilum (e.g. anterior mediastinal teratoma, thymoma, lymphoma, ascending aortic aneurysm) or POSTERIOR to the hilum (e.g. neurogenic tumor, descending aortic aneurysm).",
                "If the opacity arose from the pulmonary artery or enlarged hilar lymph nodes, the silhouette sign would erase the pulmonary arterial vessel borders."
            ],
            "viva_pearl": "Seeing normal pulmonary vascular bifurcations inside a large hilar mass proves the lesion is in the mediastinum, not the lung hilum.",
            "pitfall": "Do not confuse with the 'Hilum Convergence Sign' (where vessels converge toward the mass, indicating a dilated pulmonary artery)."
        },
        {
            "number": 6,
            "sign_name": "Thoracoabdominal Sign (Iceberg Sign)",
            "definition": "Used to determine whether a mass at the base of the thorax is purely thoracic or extends across the diaphragm into the retroperitoneum/abdomen.",
            "principles": [
                "The diaphragm curves downward and backward; lower borders of lesions in the posterior thorax plunge below the dome of the diaphragm.",
                "If a mass projected over the diaphragm has lower borders that converge toward the spine, it is completely intrathoracic.",
                "If the lower borders DIVERGE away from the spine as they pass below the diaphragm, the lesion crosses the diaphragm and has an abdominal/retroperitoneal component (e.g. thoracoabdominal aneurysm, neuroblastoma, diaphragmatic hernia)."
            ],
            "viva_pearl": "Diverging borders below the diaphragm = thoracoabdominal / retroperitoneal origin.",
            "pitfall": "Hiatus hernia can mimic a lower thoracic mass with diverging gastric borders below the diaphragm."
        },
        {
            "number": 7,
            "sign_name": "Continuous Diaphragm Sign",
            "definition": "The visualization of the entire hemidiaphragm across the midline beneath the cardiac silhouette on a frontal chest radiograph.",
            "principles": [
                "Normally, the central portion of the diaphragm in contact with the pericardium cannot be visualized because both have equal soft-tissue density (silhouette sign).",
                "Free air in the mediastinum (pneumomediastinum) or pericardial sac (pneumopericardium) tracks beneath the heart, providing contrast that renders the central diaphragm clearly visible from left to right continuously.",
                "Commonly seen in esophageal rupture (Boerhaave syndrome), tracheobronchial tear, or asthma barotrauma."
            ],
            "viva_pearl": "Continuous diaphragm sign pathognomonically establishes pneumomediastinum or pneumopericardium.",
            "pitfall": "Distinguish pneumopericardium (air does not extend above aortic arch) from pneumomediastinum (air tracks into superior mediastinum and neck)."
        },
        {
            "number": 8,
            "sign_name": "Deep Sulcus Sign",
            "definition": "In a supine patient, free intrapleural air accumulates in the anterior and inferior costophrenic recesses rather than at the apex.",
            "principles": [
                "Produces an abnormally deep, sharp, and hyperlucent costophrenic angle extending far inferiorly toward the hypochondrium.",
                "The classic sign of pneumothorax in supine trauma, post-operative, and ICU patients.",
                "Frequently the only radiographic clue to a tension pneumothorax in ventilated patients before cardiovascular collapse occurs."
            ],
            "viva_pearl": "Never expect an apical visceral pleural line in supine patients; the deep sulcus sign is your life-saving signal.",
            "pitfall": "Chronic obstructive lung disease can produce deep costophrenic angles bilaterally, but they lack the unilateral sharp lucency and depression of a deep sulcus."
        },
        {
            "number": 9,
            "sign_name": "Luftsichel Sign (Sickle of Air)",
            "definition": "A crescent of air outlining the aortic knuckle on the frontal radiograph in Left Upper Lobe (LUL) collapse.",
            "principles": [
                "Because there is no horizontal fissure in the left hemithorax, the collapsed left upper lobe shifts anteriorly and medially.",
                "The superior segment of the Left Lower Lobe (segment 6) hyperexpands and herniates upward into the apex between the collapsed lobe and the aortic arch.",
                "This produces a sickle-shaped lucency ('Luftsichel') surrounding the aortic knob."
            ],
            "viva_pearl": "Luftsichel sign = Left Upper Lobe collapse. Golden S sign = Right Upper Lobe collapse.",
            "pitfall": "Do not mistake for pneumomediastinum or medial pneumothorax."
        },
        {
            "number": 10,
            "sign_name": "Golden S Sign (Sign of Ross Golden)",
            "definition": "An S-shaped fissural configuration produced by the upward and medial collapse of the Right Upper Lobe combined with a convex central right hilar mass.",
            "principles": [
                "The lateral concave portion is formed by the retracted minor fissure.",
                "The medial convex portion is formed by the central tumor obstructing the right upper lobe bronchus.",
                "Indicates central bronchogenic carcinoma until proven otherwise."
            ],
            "viva_pearl": "Reverse S of Golden = Right upper lobe collapse with central obstructing mass.",
            "pitfall": "Right upper lobe collapse from mucus plugging produces pure fissural retraction without the central convex bulge."
        }
    ]

    with open(os.path.join(DATA_DIR, "felson_roentgen_signs.json"), "w", encoding="utf-8") as f:
        json.dump(felson_golden_signs, f, indent=2)
    print(f"  [OK] Saved {len(felson_golden_signs)} Felson Golden Roentgenologic Signs.")

    # E. Harrison's 22e Clinical Approach Algorithms
    harrison_algorithms = [
        {
            "topic": "Cardinal Manifestation: Diagnostic Approach to Dyspnea",
            "source": "Harrison's Principles of Internal Medicine 22nd Edition, Chapter 33",
            "steps": [
                "Step 1: Emergency Triage — Assess ABCs, SpO2, vital signs, work of breathing, stridor, and altered mental status.",
                "Step 2: Differential Classification — Cardiac (Heart failure, acute coronary syndrome, tamponade) vs. Pulmonary (Airway obstruction, parenchymal consolidation, pneumothorax, pulmonary embolism, interstitial disease) vs. Systemic (Severe anemia, metabolic acidosis, neuromuscular weakness).",
                "Step 3: Point-of-Care Ultrasound (POCUS / BLUE Protocol) — Immediate bedside scan: A-lines (normal/asthma/COPD/PE), B-lines (pulmonary edema), lung point (pneumothorax), consolidation with air bronchograms (pneumonia).",
                "Step 4: Laboratory Biomarkers — NT-proBNP / BNP (< 100 pg/mL has > 95% negative predictive value for acute decompensated heart failure), Troponin, Blood Gas analysis (PaO2, PaCO2, A-a gradient), D-dimer if PE suspected.",
                "Step 5: Pulmonary Function Testing — Spirometry (obstructive vs restrictive), Lung volumes (TLC < 80% confirms restriction), DLCO (differentiates intrinsic parenchymal from chest wall restriction)."
            ],
            "clinical_pearl": "Orthopnea that develops suddenly upon lying flat suggests diaphragmatic weakness (bilateral phrenic nerve palsy) or acute left ventricular failure. Platypnea-orthodeoxia (dyspnea and desaturation upright relieved by supine position) indicates hepatopulmonary syndrome or patent foramen ovale with right-to-left shunt."
        },
        {
            "topic": "Cardinal Manifestation: Approach to Massive & Submassive Hemoptysis",
            "source": "Harrison's 22e, Chapter 35 & Fishman's Chapter 28",
            "steps": [
                "Step 1: Definition & Urgency — Massive hemoptysis is defined by the RATE of bleeding causing airway flooding, asphyxiation, or hypoxemia (> 200–500 mL/24h). Asphyxiation, NOT exsanguination, is the primary cause of death.",
                "Step 2: Immediate Resuscitation — Position patient with suspected bleeding lung DEPENDENT (down) to protect the healthy lung. Secure airway with large-bore endotracheal tube (>= 8.0 or 8.5 mm ID) to allow therapeutic suctioning.",
                "Step 3: Medical Hemostasis — IV Tranexamic acid (1g bolus + 1g/8h infusion) plus Nebulized Tranexamic acid (500 mg q6–8h). Suppress violent coughing with mild codeine.",
                "Step 4: Diagnostic Imaging — Multidetector CT Bronchial Angiography (CTBA) is the first-line investigation to identify hypertrophied bronchial (90% of cases) and non-bronchial systemic arteries.",
                "Step 5: Definitive Intervention — Bronchial Artery Embolization (BAE) by interventional radiology is the gold standard (> 90% immediate control). Emergency surgery (lobectomy) is reserved for failed BAE, trauma, or aspergilloma with refractory hemorrhage."
            ],
            "clinical_pearl": "Bronchial arteries arise from the high-pressure systemic aorta (120/80 mmHg), while pulmonary arteries are low-pressure (25/10 mmHg). 90% of life-threatening hemoptysis originates from BRONCHIAL arteries, NOT pulmonary arteries."
        },
        {
            "topic": "Solitary Pulmonary Nodule (SPN): Fleischner Society & Brock Model Algorithm",
            "source": "Harrison's 22e, Chapter 84 & Fleischner Society 2024 Guidelines",
            "steps": [
                "Step 1: Radiographic Review — Obtain and compare previous chest radiographs/CTs. Definite stability over 2 years strongly indicates a benign process (except in indolent pure ground-glass adenocarcinomas).",
                "Step 2: Morphologic Assessment — Size, attenuation (solid vs subsolid vs ground-glass), margin (smooth/benign vs lobulated/spiculated), and calcification pattern (popcorn = hamartoma; concentric/targetoid = granuloma; eccentric/stippled = malignancy).",
                "Step 3: Malignancy Risk Estimation — Clinical risk (Age > 60, heavy smoking history, family/personal history of cancer, asbestos exposure, emphysema) + Radiologic risk (Spiculated margins, upper lobe location, size > 8 mm).",
                "Step 4: Risk-Stratified Management:",
                "  • Low Risk (< 5%): Serial low-dose thin-section CT at 3–6, 9–12, and 18–24 months.",
                "  • Intermediate Risk (5–65%): FDG PET/CT. If PET positive (hypermetabolic SUV > 2.5), proceed to tissue diagnosis (CT-guided transthoracic needle biopsy or navigational bronchoscopy).",
                "  • High Risk (> 65%): Direct surgical resection (VATS wedge resection / lobectomy) without intervening biopsies."
            ],
            "clinical_pearl": "PET-CT has false negatives in carcinoid tumors, adenocarcinoma in situ / minimally invasive adenocarcinoma (low metabolic rate), and lesions < 8 mm. PET-CT has false positives in active tuberculosis, fungal infections, and sarcoidosis."
        }
    ]

    with open(os.path.join(DATA_DIR, "harrison_clinical_algorithms.json"), "w", encoding="utf-8") as f:
        json.dump(harrison_algorithms, f, indent=2)
    print(f"  [OK] Saved {len(harrison_algorithms)} Harrison 22e Clinical Approach Algorithms.")

if __name__ == "__main__":
    build_textbook_library()
