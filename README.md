# PulmoDoc Suite — Comprehensive Medical PG & Clinician Application

A medical postgraduate (MD/DNB) decision-support, exam-preparation, and clinical reference software suite built from authoritative respiratory textbooks (*Murray & Nadel 7th Edition*, *Fishman's 5th Edition*, *Crofton and Douglas*, *Harrison's Principles of Internal Medicine 22nd Edition*, *Benjamin Felson's Chest Roentgenology*, *KD Tripathi Pharmacology*, *OrderSheet Companion*, and the *Thoracic Lung Imaging Atlas*).

---

## 🚀 Quick Start (One-Click Launch)

1. Double-click **`launch_app.bat`** in this folder.
2. The local server will start automatically and open your default web browser at:
   ```
   http://127.0.0.1:8000
   ```
3. Alternatively, open a terminal in this directory and run:
   ```powershell
   python pulmo_server.py
   ```

---

## 🩺 Core Modules & Features

### 1. 🖼️ Thoracic Radiology & CXR Atlas (40 High-Yield Signs & Plates)
- **Interactive Film Viewing Station**:
  - **Pan & Zoom**: Interactive mouse dragging and smooth wheel zoom up to 400%.
  - **Negative / Film Invert**: Instant radiological reversal (crucial for detecting subtle apical pneumothoraces, hair-line rib fractures, and pleural plaques).
  - **CT Windowing Sliders**: Real-time Brightness and Contrast sliders emulating Lung Window vs. Mediastinal Window.
  - **Spotter Examination Mode**: Blinded testing mode that hides the sign name and radiological description until the doctor clicks **"Reveal Finding & Pearls"**.
- **Signs Covered**: *Silhouette sign (Felson), Golden S sign, Monod sign (fungal ball), Water lily sign (hydatid), Polo-mint sign (PE), Signet-ring sign (bronchiectasis), Crazy-paving pattern (PAP/PCP), Tree-in-bud (endobronchial TB), Reverse-halo / Atoll sign, Deep sulcus sign, Luftsichel sign, Eggshell calcification, Head-cheese / three-density pattern, Honeycombing & UIP, etc.*

---

### 2. 📝 Medical PG MCQ & Spotter Examination Engine
- **35+ High-Yield Questions** modeled on NEET-PG, INI-CET, USMLE Step 2/3, and MD/DNB table viva cases.
- **Two Study Modes**:
  - **Tutor / Practice Mode**: Immediate color-coded feedback on click with deep physiological/pathological explanations, **Clinical Pearls**, **Examiner's Traps**, and textbook citations (*Murray & Nadel*, *Fishman*, *Harrison*, *Crofton*, *KD Tripathi*).
  - **Timed Exam Mode**: 30-minute exam countdown timer, question status navigation grid, and final percentage scoring.
- **Image-Based Questions (IBQs)**: Real chest radiographs, CTPA, HRCT, and spirometry flow-volume loop spotters.
- **Subspecialty Filters**: Infections, Airway Diseases, Interstitial Lung Diseases, Pulmonary Vascular, Pleural/Mediastinal, Thoracic Oncology, Critical Care, Pharmacology, and Radiology.

---

### 3. 📚 Core Specialty Textbooks & Docx Library
- Comprehensive clinical modules covering all 12 core respiratory specialties:
  1. *Pneumonia & Lower Respiratory Infections (CAP, HAP, VAP)*
  2. *Tuberculosis & Drug-Resistant TB (NTEP 2024–2026)*
  3. *Bronchial Asthma & Phenotypes (GINA Track 1 & 2)*
  4. *Chronic Obstructive Pulmonary Disease (GOLD 2024–2026)*
  5. *Thoracic Oncology & Staging (IASLC 8th/9th ed, Targeted TKIs, SCLC)*
  6. *Diffuse Parenchymal Lung Disease & IPF (ATS/ERS 2022 Guidelines)*
  7. *Pulmonary Embolism & Pulmonary Hypertension (ESC 2022/2024)*
  8. *Pleural Diseases & Interventions (Light's criteria, Empyema, Pleuroscopy)*
  9. *Bronchiectasis, Cystic Fibrosis & Alpha-1 Antitrypsin Deficiency*
  10. *Sleep Medicine, ARDS & Mechanical Ventilation (Berlin Criteria, ARDSNet)*
  11. *Fungal Infections (ABPA, Aspergilloma) & Post-COVID Sequelae*
  12. *Interventional Pulmonology & Diagnostic Procedures (EBUS, TBLB, Thoracoscopy)*
- Fast inline search and structured clinical tables.

---

### 4. 📋 Inpatient & RICU Emergency Order Sheets
- Ready-to-use clinical order bundles:
  - *Near-Fatal Asthma / Acute Severe Asthma RICU Bundle*
  - *AECOPD with Non-Invasive Ventilation (NIV / BiPAP) Protocol*
  - *Massive Hemoptysis Immediate Resuscitation Bundle*
  - *Severe Sepsis / Severe CAP Hour-1 Bundle (Surviving Sepsis)*
  - *Massive Pulmonary Embolism & Thrombolysis Protocol (Alteplase/Tenecteplase)*
  - *ARDS Lung-Protective Mechanical Ventilation (ARDSNet 6 mL/kg PBW)*
  - *Intercostal Drainage (ICD) Insertion & Underwater Seal Troubleshooting SOP*
- **One-Click Copy**: Copies clean formatted order sheets directly to the clipboard for pasting into hospital notes, EMR, or WhatsApp clinical handovers.

---

### 5. 🧮 Pulmonology & Critical Care Calculators
- **Comprehensive ABG Analyzer**:
  - Live acid-base diagnosis (*Acidemia / Alkalemia / Mixed*).
  - **Winter's Formula** expected PaCO2 for metabolic acidosis (`1.5 * HCO3 + 8 ± 2`).
  - Expected compensation for acute/chronic respiratory acidosis and alkalosis.
  - **Anion Gap & Albumin-Corrected Anion Gap** (`AG + 2.5 * [4.0 - Albumin]`).
  - **Delta-Delta Ratio** (`ΔAG / ΔHCO3`) to unmask triple acid-base disorders.
  - **PaO2/FiO2 Ratio** with Berlin ARDS severity stratification (*Mild, Moderate, Severe*).
  - Quick clinical presets (*DKA, Severe COPD, ARDS, Normal Baseline*).
- **CURB-65 Pneumonia Severity Score**: Calculates 30-day mortality risk and admission setting (*Outpatient vs Ward vs ICU*).
- **Wells Score for Pulmonary Embolism**: Pre-test clinical probability and guidance on D-dimer vs. direct CTPA.
- **GOLD 2024–2026 COPD Staging**: Combines FEV1 airflow limitation with modern Group A, B, and E exacerbation matrix.

---

### 6. 💊 Respiratory Pharmacology & Inhaler Techniques
- Complete drug monographs:
  - *Inhaled Bronchodilators (SABA, LABA, SAMA, LAMA)*
  - *Inhaled & Systemic Corticosteroids (Budesonide, Fluticasone, Prednisolone)*
  - *Severe Asthma Biologics (Omalizumab, Mepolizumab, Benralizumab, Dupilumab, Tezepelumab)*
  - *Antifibrotics for IPF (Pirfenidone, Nintedanib)*
  - *NTEP 2024–2026 Regimens (DS-TB 2HZE/4HRE, 6-Month All-Oral BPaLM Regimen for MDR-TB)*
  - *Antibiotics in Pulmonology (Aminoglycoside extended-interval dosing, Cephalosporins)*
- Illustrated step-by-step CDC inhaler technique guides (*pMDI, pMDI with Spacer, DPI*).

---

### 7. 🎓 MD/DNB Table Viva Stations
- Viva questions, anatomy, physics, and troubleshooting:
  - *Station 1: Underwater Seal ICD Bottle Physics & Clamping Rules*
  - *Station 2: Flexible Fiberoptic Bronchoscope Anatomy, TBLB & High-Level Disinfection*
  - *Station 3: Pleural Biopsy Needles (Abrams vs. Cope Needle Mechanics & Safety Margin)*
  - *Station 4: Spirometry Flow-Volume Loops & PFT Bronchodilator Reversibility*

---

## 🛠️ Technology Stack
- **Backend**: Python 3 standard library `http.server` (Zero external dependencies; completely self-contained).
- **Frontend**: Responsive Single-Page Application (HTML5, Vanilla ES6, CSS3 dark/light medical design system).
- **Portability**: 100% offline; requires no internet connection after installation.
