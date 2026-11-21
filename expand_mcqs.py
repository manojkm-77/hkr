"""
Expand MCQ bank with high-yield MD/DNB, NEET-PG, and USMLE clinical vignettes and spotters.
"""
import os, json

DATA_PATH = os.path.join(os.path.dirname(__file__), "pulmo_app", "data", "mcq_bank.json")

with open(DATA_PATH, "r", encoding="utf-8") as f:
    mcqs = json.load(f)

additional_mcqs = [
    {
        "id": "MCQ-021",
        "topic": "Infections",
        "subspecialty": "Tuberculosis Pharmacology",
        "type": "Clinical Vignette",
        "difficulty": "Moderate",
        "question": "A 45-year-old male with active pulmonary tuberculosis develops severe acute hepatitis (ALT 420 U/L, AST 380 U/L, Bilirubin 3.2 mg/dL) 3 weeks after starting fixed-dose combination 4-drug therapy (2HZE/4HRE). Which of the following is the standard recommended modified anti-TB regimen while waiting for liver enzymes to normalize?",
        "options": [
            "Continue Rifampicin and Isoniazid, stop Pyrazinamide only",
            "Stop all first-line hepatotoxic drugs (H, R, Z); initiate Ethambutol + Levofloxacin + Streptomycin / Amikacin",
            "Reduce doses of all 4 drugs by 50% and administer IV N-acetylcysteine",
            "Switch to Bedaquiline + Delamanid + Bicarbonate infusion"
        ],
        "correct_index": 1,
        "explanation": "Drug-Induced Liver Injury (DILI) is defined as ALT/AST > 3x ULN with symptoms (jaundice, nausea, anorexia) or > 5x ULN without symptoms. When this occurs, ALL hepatotoxic first-line drugs (Pyrazinamide, Isoniazid, Rifampicin) must be stopped immediately. In a sick patient with active smear-positive TB, a non-hepatotoxic interim regimen (Ethambutol + Levofloxacin + Aminoglycoside such as Amikacin or Streptomycin) is started to maintain bactericidal pressure while transaminases normalize. Once AST/ALT drops < 2x ULN, drugs are reintroduced sequentially: Rifampicin first (least hepatotoxic among bactericidals), followed by Isoniazid after 3–7 days. Pyrazinamide (the most hepatotoxic drug) is generally NOT reintroduced unless crucial.",
        "clinical_pearl": "Hierarchy of hepatotoxicity among first-line anti-TB drugs: Pyrazinamide (Z) > Isoniazid (H) > Rifampicin (R). Ethambutol and Streptomycin/Amikacin have no significant hepatotoxicity.",
        "trap": "Do not reintroduce Pyrazinamide in severe TB DILI; once ALT normalizes, reintroduce R first, then H.",
        "reference": "NTEP Guidelines on TB-DILI; ATS/CDC/IDSA Guidelines on Treatment of Drug-Susceptible TB",
        "image": "/assets/thoracic_atlas_assets/fig33.jpg"
    },
    {
        "id": "MCQ-022",
        "topic": "Thoracic Radiology",
        "subspecialty": "Classic Signs",
        "type": "Image-Based Question",
        "difficulty": "Moderate",
        "question": "A 65-year-old smoker presents with right hemithorax volume loss, dry cough, and hemoptysis. Chest radiograph shows elevation and medial displacement of the minor fissure combined with a convex central right hilar mass, producing a reverse 'S' contour. What is this sign called, and what is the underlying diagnosis?",
        "options": [
            "Luftsichel sign of left upper lobe collapse; Hamartoma",
            "Golden S sign (Sign of Golden); Right Upper Lobe collapse secondary to central bronchogenic carcinoma",
            "Westermark sign; Acute massive pulmonary embolism",
            "Spine sign; Left lower lobe pneumonia"
        ],
        "correct_index": 1,
        "explanation": "The 'Golden S Sign' (first described by Ross Golden) is seen on PA chest radiographs in Right Upper Lobe (RUL) collapse caused by a central obstructing hilar mass. The lateral, superior concave contour is formed by the upward and medially displaced minor fissure, while the medial, inferior convex bulge is formed by the central tumor/adenopathy. The resulting fissure configuration resembles a reversed 'S' or Greek letter sigma.",
        "clinical_pearl": "Whenever you see Golden's S sign, suspect central bronchogenic carcinoma (Squamous cell carcinoma or Small cell carcinoma) obstructing the right upper lobe bronchus until proven otherwise.",
        "trap": "Golden's S sign is specifically described for the RIGHT upper lobe because of the horizontal/minor fissure. (Left upper lobe collapse produces the Luftsichel sign).",
        "reference": "Felson's Principles of Chest Roentgenology; Thoracic Lung Imaging Atlas",
        "image": "/assets/thoracic_atlas_assets/fig17.jpg"
    },
    {
        "id": "MCQ-023",
        "topic": "Thoracic Radiology",
        "subspecialty": "Lobar Collapse Signs",
        "type": "Image-Based Question",
        "difficulty": "Hard",
        "question": "A 58-year-old male with persistent cough has a frontal chest radiograph showing a crescent of air (hyperlucent crescent) surrounding the aortic knuckle and superior mediastinum, with hazy opacification of the left hemithorax. What is this sign called, and what anatomical phenomenon creates it?",
        "options": [
            "Pneumopericardium; Continuous diaphragm sign",
            "Luftsichel sign; Superior segment of the Left Lower Lobe hyperexpanding upward into the apex between the collapsed left upper lobe and mediastinum",
            "Ginkgo-leaf sign; Air dissection along the pectoralis major muscle fibers",
            "Deep sulcus sign; Pleural effusion tracking along the oblique fissure"
        ],
        "correct_index": 1,
        "explanation": "The 'Luftsichel Sign' ('Luft' = air, 'Sichel' = sickle in German) is a classic radiologic sign of Left Upper Lobe (LUL) collapse. Because the left lung lacks a minor fissure, the collapsing left upper lobe shifts anteriorly and medially against the anterior chest wall. In response, the superior segment of the Left Lower Lobe (segment 6) hyperexpands and herniates upward into the apical pleural space between the collapsed lobe and the aortic knuckle, creating a sickle-shaped lucency outlining the aortic arch.",
        "clinical_pearl": "Luftsichel sign = Left Upper Lobe collapse. Golden S sign = Right Upper Lobe collapse with central mass.",
        "trap": "Do not mistake the Luftsichel sign for pneumomediastinum or apical pneumothorax; look for associated left hilar elevation and loss of the left heart border (silhouette sign).",
        "reference": "Felson's Principles of Chest Roentgenology; Thoracic Lung Imaging Atlas",
        "image": "/assets/thoracic_atlas_assets/fig17.jpg"
    },
    {
        "id": "MCQ-024",
        "topic": "Thoracic Radiology",
        "subspecialty": "CXR Signs",
        "type": "Clinical Vignette",
        "difficulty": "Moderate",
        "question": "On a lateral chest radiograph, the thoracic vertebral bodies normally become progressively MORE LUCENT (darker) as you inspect downward from the neck to the diaphragm. In a patient with left lower lobe pneumonia, the lower thoracic vertebrae appear abnormally OPAQUE (denser) compared to the upper vertebrae. What is this Felson sign called?",
        "options": [
            "Cervicothoracic sign",
            "Hilum overlay sign",
            "Spine sign (Positive Spine Sign of Felson)",
            "Thoracoabdominal sign"
        ],
        "correct_index": 2,
        "explanation": "The 'Positive Spine Sign' of Benjamin Felson is an essential sign on the lateral chest radiograph. Normally, because of increasing aeration of the lower lobes and less overlying shoulder girdle soft tissue, the vertebral bodies appear progressively darker/more lucent from top to bottom. If a lower lobe contains consolidation, mass, or fluid, the lower vertebral bodies appear whiter/denser than the upper ones (loss of normal progression of lucency). This is frequently the ONLY clue to retrocardiac left lower lobe pneumonia on plain radiography.",
        "clinical_pearl": "Always look at the vertebral bodies on the lateral film. If the spine gets denser as you go down, there is a lower lobe pathology (Pneumonia, Effusion, Mass, or Atelectasis).",
        "trap": "Retrocardiac left lower lobe consolidations are commonly missed on frontal PA views because they hide behind the cardiac silhouette; the lateral spine sign catches them.",
        "reference": "Felson's Principles of Chest Roentgenology, Ch 3",
        "image": "/assets/thoracic_atlas_assets/fig1.jpg"
    },
    {
        "id": "MCQ-025",
        "topic": "Thoracic Radiology",
        "subspecialty": "Classic CXR Signs",
        "type": "Clinical Vignette",
        "difficulty": "Hard",
        "question": "A mass is identified on a frontal chest radiograph projecting over the right hilum. The pulmonary vessels of the interlobar pulmonary artery can be traced coursing more than 1 cm WITHIN and THROUGH the mass rather than being pushed away. According to Felson, what sign is this, and where is the mass located?",
        "options": [
            "Hilum overlay sign; Mass is in the anterior or posterior mediastinum (NOT within the hilum itself)",
            "Cervicothoracic sign; Mass is in the superior mediastinum",
            "Water-lily sign; Mass is inside an intrapulmonary hydatid cyst",
            "Continuous diaphragm sign; Mass is in the pericardial space"
        ],
        "correct_index": 0,
        "explanation": "The 'Hilum Overlay Sign' (Felson) is used to determine whether an opacity projected over the hilum represents cardiomegaly/pericardial effusion/mediastinal mass OR true hilar enlargement. If the branches of the pulmonary artery are clearly visible coursing THROUGH the border of the opacity, the opacity is NOT in the hilum, but located either ANTERIOR or POSTERIOR to the hilum in the mediastinum (e.g. thymoma, teratoma, neurogenic tumor, or ascending aortic aneurysm). If the opacity arose from the pulmonary artery or hilar lymph nodes, the silhouette sign would erase the vessel borders.",
        "clinical_pearl": "Seeing pulmonary vessel bifurcations within a hilar mass tells you the lesion is anterior (e.g., lymphoma, thymoma) or posterior (e.g., Schwannoma), not intra-hilar.",
        "trap": "If the interlobar artery is displaced laterally, that is the 'Hilum Convergence Sign' (indicating true pulmonary vascular enlargement), NOT hilum overlay.",
        "reference": "Felson's Principles of Chest Roentgenology; Murray & Nadel 7th Ed",
        "image": "/assets/thoracic_atlas_assets/fig11.jpg"
    },
    {
        "id": "MCQ-026",
        "topic": "Interstitial Lung Diseases",
        "subspecialty": "Hypersensitivity Pneumonitis",
        "type": "Image-Based Question",
        "difficulty": "Hard",
        "question": "A 48-year-old pigeon breeder presents with subacute progressive dyspnea and dry cough. Inspiratory and expiratory HRCT demonstrates the 'Three-Density Pattern' (also called 'Head-Cheese Sign'): sharply demarcated lobules of ground-glass opacity, normal lung attenuation, and air-trapping/hypoattenuation. What is the classic diagnosis?",
        "options": [
            "Chronic Eosinophilic Pneumonia",
            "Hypersensitivity Pneumonitis (Fibrotic or Non-Fibrotic HP)",
            "Pulmonary Alveolar Microlithiasis",
            "Lymphangioleiomyomatosis (LAM)"
        ],
        "correct_index": 1,
        "explanation": "The 'Three-Density Pattern' (formerly called the Head-Cheese sign) on HRCT represents the coexistence of: (1) Ground-glass opacity (infiltrative alveolar/interstitial inflammation), (2) Normal lung attenuation, and (3) Mosaic hypoattenuation / air-trapping accentuated on expiratory scans (constrictive bronchiolitis/small airway obstruction). Under the 2020 ATS/JRS/ALAT guidelines, this three-density sign is highly specific for Hypersensitivity Pneumonitis.",
        "clinical_pearl": "Bronchoalveolar Lavage (BAL) in subacute HP classically shows marked LYMPHOCYTOSIS (> 40–50%) with a decreased CD4/CD8 ratio (< 1.0), in contrast to Sarcoidosis where BAL shows elevated CD4/CD8 ratio (> 3.5).",
        "trap": "HP involves BOTH the interstitium and small airways (bronchiolocentric granulomatous inflammation); hence expiratory air-trapping is essential to differentiate it from pure fibrotic ILDs like UIP/IPF.",
        "reference": "ATS/JRS/ALAT Clinical Practice Guideline on Hypersensitivity Pneumonitis (AJRCCM 2020)",
        "image": "/assets/thoracic_atlas_assets/fig19.jpg"
    },
    {
        "id": "MCQ-027",
        "topic": "Interstitial Lung Diseases",
        "subspecialty": "Cystic Lung Diseases",
        "type": "Image-Based Question",
        "difficulty": "Hard",
        "question": "A 34-year-old non-smoking woman presents with sudden onset left-sided chest pain, breathlessness, and recurrent spontaneous pneumothoraces. High-Resolution CT of the chest demonstrates numerous uniform, round, thin-walled (1–2 mm) air-filled cysts distributed symmetrically throughout BOTH lungs, with normal intervening lung parenchyma. What is the underlying disorder, and which molecular pathway is abnormally hyperactivated?",
        "options": [
            "Pulmonary Langerhans Cell Histiocytosis (PLCH); BRAF V600E mutation",
            "Lymphangioleiomyomatosis (LAM); mTOR pathway hyperactivation due to TSC1/TSC2 mutation",
            "Birt-Hogg-Dubé syndrome; FLCN (folliculin) gene mutation",
            "Alpha-1 Antitrypsin Deficiency; SERPINA1 Z allele"
        ],
        "correct_index": 1,
        "explanation": "Lymphangioleiomyomatosis (LAM) is a rare, progressive cystic lung disease affecting almost exclusively premenopausal women (either sporadic or associated with Tuberous Sclerosis Complex - TSC). Mutations in TSC1 (hamartin) or TSC2 (tuberin) lead to constitutive hyperactivation of the Mammalian Target of Rapamycin (mTOR) signaling pathway, resulting in abnormal smooth muscle-like 'LAM cell' proliferation and destructive thin-walled parenchymal cysts. Sirolimus (an oral mTOR inhibitor) stabilizes lung function decline and is the FDA-approved treatment.",
        "clinical_pearl": "Diagnostic triad for LAM: Premenopausal female + bilateral thin-walled round cysts with normal intervening lung + elevated serum VEGF-D level (> 800 pg/mL) or chylothorax / renal angiomyolipoma.",
        "trap": "In Pulmonary Langerhans Cell Histiocytosis (PLCH), cysts are irregular, bizarre-shaped, bilobed/cloverleaf with upper-lung predominance and costophrenic angle sparing, occurring almost exclusively in SMOKERS. In LAM, cysts are uniform, round, thin-walled, diffuse, and spare no lung zones.",
        "reference": "ATS/JRS Guidelines for Lymphangioleiomyomatosis (AJRCCM 2016); Thoracic Atlas",
        "image": "/assets/thoracic_atlas_assets/extra_lymphangioleiomyomatosis__lam___diffuse_cystic_lung_dis.jpg"
    },
    {
        "id": "MCQ-028",
        "topic": "Interstitial Lung Diseases",
        "subspecialty": "Occupational Pneumoconiosis",
        "type": "Image-Based Question",
        "difficulty": "Moderate",
        "question": "A 56-year-old stone quarry worker presents with progressive exertional dyspnea. HRCT of the thorax demonstrates upper-lobe predominant small centrilobular nodules, conglomerate masses of progressive massive fibrosis (PMF), and thin, rim-like peripheral calcification of hilar and mediastinal lymph nodes. What is this classic nodal calcification sign, and what is the diagnosis?",
        "options": [
            "Popcorn calcification; Pulmonary hamartoma",
            "Eggshell calcification; Simple and complicated Silicosis",
            "Stippled calcification; Primary lung adenocarcinoma",
            "Targetoid calcification; Histoplasmosis"
        ],
        "correct_index": 1,
        "explanation": "'Eggshell Calcification' of hilar and mediastinal lymph nodes consists of a thin, dense, continuous or interrupted peripheral shell of calcium outlining at least two lymph nodes. While classically and most frequently associated with Silicosis (and Coal Workers' Pneumoconiosis), it is also seen in ~5% of patients with Sarcoidosis, and rarely post-irradiation lymphoma, amyloidosis, and histoplasmosis.",
        "clinical_pearl": "Patients with silicosis have a 30-fold increased lifetime risk of developing Tuberculosis ('Silicotuberculosis') because silica particles destroy alveolar macrophages and impair cell-mediated immunity.",
        "trap": "Always screen silicosis patients for latent and active tuberculosis; TB accelerates the progression of silicotic massive fibrosis dramatically.",
        "reference": "Crofton's Ch 26; Fishman's Ch 69; Thoracic Lung Imaging Atlas",
        "image": "/assets/thoracic_atlas_assets/extra_silicosis__nodules_and_eggshell_lymph_node_calcificatio.jpg"
    },
    {
        "id": "MCQ-029",
        "topic": "Thoracic Oncology",
        "subspecialty": "Targeted Molecular Therapy",
        "type": "Clinical Vignette",
        "difficulty": "Hard",
        "question": "A 49-year-old non-smoking female is diagnosed with Stage IV metastatic lung adenocarcinoma with multiple brain metastases. Next-Generation Sequencing (NGS) of the tumor tissue reveals an Epidermal Growth Factor Receptor (EGFR) Exon 19 in-frame deletion. What is the preferred first-line systemic therapy based on the FLAURA trial?",
        "options": [
            "First-generation TKI: Gefitinib or Erlotinib",
            "Third-generation irreversible CNS-penetrant EGFR TKI: Osimertinib",
            "Platinum doublet chemotherapy (Cisplatin + Pemetrexed) plus Pembrolizumab",
            "Dual checkpoint inhibition (Nivolumab + Ipilimumab)"
        ],
        "correct_index": 1,
        "explanation": "Osimertinib is a third-generation, CNS-active, irreversible oral EGFR-TKI that potently inhibits both sensitizing EGFR mutations (Exon 19 del and L858R) and the T790M resistance mutation, while sparing wild-type EGFR. In the landmark FLAURA trial, first-line Osimertinib demonstrated statistically significant improvements in progression-free survival (18.9 vs 10.2 months) and overall survival (38.6 vs 31.8 months) compared to first-generation TKIs (Gefitinib/Erlotinib), with superior intracranial efficacy against brain metastases and significantly less rash/diarrhea.",
        "clinical_pearl": "In metastatic NSCLC with actionable driver mutations (EGFR, ALK, ROS1, RET), targeted TKIs are vastly superior to immunotherapy/chemotherapy. Immunotherapy (anti-PD-1/PD-L1) has LOW response rates and increases toxicity (pneumonitis) if followed by TKIs.",
        "trap": "Do not give single-agent immunotherapy to EGFR-mutant lung adenocarcinoma, even if PD-L1 expression is > 50%.",
        "reference": "FLAURA Trial (N Engl J Med 2018, 2020); NCCN Guidelines on Non-Small Cell Lung Cancer",
        "image": "/assets/thoracic_atlas_assets/fig32.jpg"
    },
    {
        "id": "MCQ-030",
        "topic": "Thoracic Oncology",
        "subspecialty": "Small Cell Lung Cancer",
        "type": "Clinical Vignette",
        "difficulty": "Hard",
        "question": "A 67-year-old heavy smoker presents with facial plethora, neck vein distension, hoarseness of voice, and weight loss. Bronchoscopic biopsy confirms Small Cell Lung Cancer (SCLC). Staging reveals a right hilar mass, mediastinal and ipsilateral supraclavicular lymphadenopathy, but no distant metastases (Limited Stage SCLC). What is the standard curative-intent treatment paradigm?",
        "options": [
            "Right pneumonectomy followed by adjuvant oral etoposide",
            "Concurrent Chemoradiotherapy (Cisplatin + Etoposide with early thoracic radiotherapy), followed by Prophylactic Cranial Irradiation (PCI) in responders",
            "Single-agent Osimertinib plus whole-brain radiation",
            "Palliative care only, as limited stage SCLC has zero cure rate"
        ],
        "correct_index": 1,
        "explanation": "Limited-Stage Small Cell Lung Cancer (LS-SCLC, confined to one hemithorax and regional nodes safely encompassed in a tolerable radiation port) is treated with CONCURRENT CHEMORADIOTHERAPY: 4 cycles of Cisplatin + Etoposide (EP) with early concurrent twice-daily (or once-daily) thoracic radiation therapy. In patients achieving a complete or good partial response, Prophylactic Cranial Irradiation (PCI) significantly reduces intracranial relapse and confers a proven survival benefit. Surgery is reserved only for rare T1-2N0M0 tumors (< 5% of cases).",
        "clinical_pearl": "SCLC is exquisitely sensitive to initial chemo and radiotherapy (response rates > 70–80%), but rapid chemoresistance and recurrence are characteristic.",
        "trap": "Do not delay radiotherapy in LS-SCLC: early concurrent thoracic radiation started during cycle 1 or 2 of chemotherapy yields superior overall survival compared to sequential therapy.",
        "reference": "Murray & Nadel 7th Ed, Ch 53; NCCN Small Cell Lung Cancer Guidelines",
        "image": "/assets/thoracic_atlas_assets/fig17.jpg"
    },
    {
        "id": "MCQ-031",
        "topic": "Critical Care",
        "subspecialty": "Patient-Ventilator Asynchrony",
        "type": "Clinical Vignette",
        "difficulty": "Hard",
        "question": "A 70-year-old patient with severe COPD is receiving Volume-Controlled Mechanical Ventilation. The flow-time scalar shows that expiratory flow FAILS to return to baseline (zero) before the next mechanical breath is initiated. On applying an end-expiratory hold maneuver, the total PEEP reads 14 cmH2O (set extrinsic PEEP is 5 cmH2O). What is this phenomenon called, and what is the primary risk?",
        "options": [
            "Inspiratory breath stacking; Oxygen toxicity",
            "Dynamic hyperinflation with Auto-PEEP (Intrinsic PEEP of 9 cmH2O); Risk of tension barotrauma, severe hemodynamic collapse, and ineffective triggering",
            "Reverse triggering; Diaphragmatic myopathy",
            "Double triggering; Lactic acidosis"
        ],
        "correct_index": 1,
        "explanation": "When expiratory flow does not reach baseline before the next breath begins, air remains trapped inside the lung, leading to 'Dynamic Hyperinflation' and 'Intrinsic PEEP' (Auto-PEEP). Intrinsic PEEP = Total PEEP (measured during end-expiratory hold) minus Set Extrinsic PEEP = 14 - 5 = 9 cmH2O. Consequences of severe auto-PEEP include: (1) Impaired venous return -> severe hypotension and cardiac arrest, (2) Alveolar rupture / barotrauma (pneumothorax), and (3) Ineffective inspiratory trigger efforts (patient must generate negative pleural pressure exceeding the 9 cmH2O auto-PEEP just to trigger the ventilator).",
        "clinical_pearl": "Management of severe Auto-PEEP in COPD: Decrease respiratory rate, decrease tidal volume, increase peak inspiratory flow rate (80–100 L/min) to lengthen expiratory time (target I:E ratio 1:3 to 1:5), and match extrinsic PEEP to ~70–80% of intrinsic PEEP to facilitate triggering.",
        "trap": "In hemodynamically crashing ventilated asthmatic/COPD patients, DISCONNECT THE VENTILATOR from the endotracheal tube immediately and compress the chest to allow passive exhalation of trapped air.",
        "reference": "ABG & Mechanical Ventilation Guide; Murray & Nadel 7th Ed, Ch 86",
        "image": "/assets/thoracic_atlas_assets/fig23.jpg"
    },
    {
        "id": "MCQ-032",
        "topic": "Critical Care",
        "subspecialty": "Ventilator Mechanics",
        "type": "Clinical Vignette",
        "difficulty": "Hard",
        "question": "During volume-controlled mechanical ventilation of an ARDS patient, the Peak Inspiratory Pressure (PIP) is 38 cmH2O and the Plateau Pressure (Pplat, measured with an end-inspiratory hold) is 34 cmH2O. PEEP is 10 cmH2O and Tidal Volume is 420 mL. What is the driving pressure (ΔP), and what physiological change does this represent?",
        "options": [
            "Driving Pressure = 4 cmH2O; Increased airway resistance (bronchospasm)",
            "Driving Pressure = 24 cmH2O; Marked reduction in respiratory system compliance (stiff lung)",
            "Driving Pressure = 28 cmH2O; Increased chest wall elastance only",
            "Driving Pressure = 14 cmH2O; Normal lung mechanics"
        ],
        "correct_index": 1,
        "explanation": "Calculations of ventilator mechanics:\n1. Transairway Pressure (PIP - Pplat) = 38 - 34 = 4 cmH2O (Normal is < 5 cmH2O, indicating normal airway resistance).\n2. Driving Pressure (ΔP = Pplat - PEEP) = 34 - 10 = 24 cmH2O.\nNormal driving pressure in ventilated lungs is < 14–15 cmH2O. A driving pressure of 24 cmH2O indicates severely reduced respiratory system static compliance [Cstat = Vt / (Pplat - PEEP) = 420 / 24 = 17.5 mL/cmH2O; normal is 50–100 mL/cmH2O]. In a landmark study by Amato et al. (N Engl J Med 2015), driving pressure > 15 cmH2O was the single variable most strongly associated with increased mortality in ARDS.",
        "clinical_pearl": "Rule of thumb: If PIP is high but Pplat is normal -> AIRWAY issue (secretions, bronchospasm, kinked ETT). If BOTH PIP and Pplat are elevated -> COMPLIANCE issue (ARDS, pneumothorax, pulmonary edema, atelectasis, abdominal compartment syndrome).",
        "trap": "Never confuse PIP and Pplat: Plateau pressure reflects the alveolar pressure; Peak pressure reflects alveolar pressure PLUS airway resistance.",
        "reference": "Amato et al., N Engl J Med 2015; Murray & Nadel 7th Ed, Ch 85",
        "image": "/assets/thoracic_atlas_assets/fig6.jpg"
    },
    {
        "id": "MCQ-033",
        "topic": "Pharmacology",
        "subspecialty": "Antibiotics & Aminoglycosides",
        "type": "Clinical Vignette",
        "difficulty": "Moderate",
        "question": "A 54-year-old diabetic male with hospital-acquired Pseudomonas pneumonia is prescribed once-daily high-dose Amikacin (15 mg/kg IV OD). Which pharmacodynamic property justifies once-daily extended-interval aminoglycoside dosing, and what serum level parameter must be monitored to prevent nephrotoxicity?",
        "options": [
            "Time-dependent bactericidal action; Peak concentration must be kept below 5 mcg/mL",
            "Concentration-dependent killing with prolonged Post-Antibiotic Effect (PAE); Trough concentration must be < 2.5–5 mcg/mL prior to the next dose",
            "Immune-mediated clearance; Monitor AUC/MIC ratio only",
            "Inhibition of cell wall synthesis; Monitor serum potassium"
        ],
        "correct_index": 1,
        "explanation": "Aminoglycosides (Amikacin, Gentamicin, Tobramycin) exhibit CONCENTRATION-DEPENDENT KILLING (the higher the Peak/MIC ratio, typically target 8–10:1, the faster and more extensive the bacterial eradication) and a prolonged POST-ANTIBIOTIC EFFECT (PAE, continuing bacterial suppression even when serum levels drop below MIC). Once-daily high-dose dosing maximizes the peak (Cmax) for superior bactericidal efficacy while allowing drug levels to drop to near-zero before the next dose. Aminoglycoside uptake by renal tubular cells is saturable; prolonged low trough levels prevent cortical accumulation, minimizing nephrotoxicity and ototoxicity.",
        "clinical_pearl": "Trough level monitoring (drawn 30 minutes before next dose) is the gold standard safety check: Amikacin trough should be < 5 mcg/mL (< 2.5 mcg/mL is ideal).",
        "trap": "Nephrotoxicity from aminoglycosides is non-oliguric acute tubular necrosis (ATN) and is usually reversible; ototoxicity (vestibular and cochlear hair cell apoptosis) is frequently IRREVERSIBLE.",
        "reference": "Aminoglycosides and Cephalosporins Reference; KD Tripathi 8th Ed, Ch 53",
        "image": "/assets/thoracic_atlas_assets/fig1.jpg"
    },
    {
        "id": "MCQ-034",
        "topic": "Pharmacology",
        "subspecialty": "Pulmonary Pharmacology",
        "type": "Clinical Vignette",
        "difficulty": "Moderate",
        "question": "A 62-year-old COPD patient on chronic theophylline therapy is admitted with fever and right lower lobe pneumonia. The resident prescribes IV Ciprofloxacin 400 mg q12h. On day 3, the patient develops intractable vomiting, agitation, multifocal ventricular premature contractions, and a generalized tonic-clonic seizure. What is the mechanism of this severe drug interaction?",
        "options": [
            "Ciprofloxacin causes acute renal failure reducing theophylline excretion",
            "Ciprofloxacin potent inhibition of hepatic Cytochrome P450 1A2 (CYP1A2), leading to toxic accumulation of theophylline",
            "Ciprofloxacin displaces theophylline from serum albumin binding sites",
            "Synergistic GABA-A receptor antagonism causing seizures directly"
        ],
        "correct_index": 1,
        "explanation": "Theophylline is metabolized > 90% by the hepatic cytochrome P450 system, primarily the CYP1A2 enzyme. Ciprofloxacin (and other fluoroquinolones like enoxacin, as well as macrolides like erythromycin and clarithromycin, and cimetidine) is a POTENT CYP1A2 INHIBITOR. Co-administration can reduce theophylline clearance by 30–50% within 48–72 hours, precipitating severe theophylline toxicity (nausea, vomiting, supraventricular/ventricular tachyarrhythmias, and refractory seizures).",
        "clinical_pearl": "Theophylline has a very narrow therapeutic window (10–20 mcg/mL or 55–110 mcmol/L). Toxicity can occur even at levels > 20 mcg/mL, and life-threatening seizures/arrhythmias commonly occur at > 30 mcg/mL.",
        "trap": "Levofloxacin has minimal CYP1A2 inhibitory activity compared to Ciprofloxacin and is much safer if a fluoroquinolone must be used in a patient taking theophylline.",
        "reference": "KD Tripathi 8th Ed, Ch 16; Harrison 22e, Ch 299",
        "image": "/assets/asthma_assets/airway_comparison.png"
    },
    {
        "id": "MCQ-035",
        "topic": "Airway Diseases",
        "subspecialty": "Alpha-1 Antitrypsin Deficiency",
        "type": "Clinical Vignette",
        "difficulty": "Moderate",
        "question": "A 36-year-old non-smoking accountant presents with progressive exertional dyspnea. Spirometry reveals severe obstructive defect (FEV1 38% predicted) with minimal reversibility. HRCT of the chest reveals extensive Panacinar (Panlobular) emphysema with marked LOWER LOBE predominance. Family history reveals his father died of liver cirrhosis in his 50s. What is the underlying genetic disorder, and what is the pathogenic mechanism?",
        "options": [
            "Cystic Fibrosis; CFTR chloride channel defect",
            "Alpha-1 Antitrypsin Deficiency (AATD); Unopposed neutrophil elastase degradation of alveolar elastin due to homozygous Pi*ZZ phenotype",
            "Kartagener syndrome; Ciliary outer dynein arm defect",
            "Hermansky-Pudlak syndrome; Lysosome-related organelle trafficking defect"
        ],
        "correct_index": 1,
        "explanation": "Alpha-1 Antitrypsin Deficiency (AATD, codominant mutation in SERPINA1 gene on chromosome 14) results in a severe deficiency of serum AAT, a major circulating serine protease inhibitor. Without AAT, Neutrophil Elastase released during inflammation acts unopposed, digesting elastin in alveolar walls. Unlike smoking-induced centrilobular emphysema (which has upper lobe predominance), AATD emphysema is PANACINAR (involves the entire acinus uniformly) with striking BASAL / LOWER LOBE predominance. Accumulation of polymerized mutant Z-protein in hepatocytes causes liver cirrhosis.",
        "clinical_pearl": "Any patient presenting with COPD/emphysema at a young age (< 45 years), non-smoker, lower-lobe predominant emphysema, or unexplained liver disease MUST be screened for AATD with a serum AAT level.",
        "trap": "Smoking accelerates the onset of emphysema in AATD patients by 15–20 years because cigarette smoke oxidizes the active methionine residue of the residual AAT, inactivating it.",
        "reference": "Crofton's Ch 18; Murray & Nadel 7th Ed, Ch 44; WHO AATD Guidelines",
        "image": "/assets/thoracic_atlas_assets/fig23.jpg"
    }
]

mcqs.extend(additional_mcqs)

with open(DATA_PATH, "w", encoding="utf-8") as f:
    json.dump(mcqs, f, indent=2)

print(f"Total MCQs after expansion: {len(mcqs)}")
