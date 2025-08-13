
# Clinical Pharmacy Inventory Automation (VBA + Python)

## Overview

During a visit to an IDHS clinical pharmacy, I observed pharmacists spending \~2 hours each week on manual inventory: hand-counting drugs on paper (“sheet 1”), then cross-checking each item against supplier quantities (“sheet 2”). I designed and delivered this automation to standardize supplier data, link it to canister locations, and turn the weekly reconciliation into a guided, one-click workflow inside Excel.

**What this demonstrates:** practical problem-solving in a regulated environment, stakeholder consulting, Python data engineering, entity matching, and advanced Excel/VBA UI/automation.

---

## Business Problem

* **Manual & error-prone:** Paper counts (sheet 1) + manual math against supplier output (sheet 2).
* **Search overhead:** Drug orders differ between sheets; finding each item wastes time.
* **Operational drag:** Pharmacists lose a big chunk of their day on non-clinical work.

**Goal:** eliminate paper, standardize inputs, and produce instant surplus/shortage reports while keeping the workflow familiar to staff.

---

## Constraints & Considerations

* **Canister locations only on paper (PDF)**
* **Supplier formats inconsistent**
* **No patient/PHI involved, but proprietary lists are sensitive**
* **Solution must run in Excel (lowest friction)**

---

## Approach (Consulting → Data Engineering → Automation)

### 1) On-site consulting & workflow mapping

* Shadowed pharmacists, documented current steps and pain points.
* Defined acceptance criteria: remove paper, standardize supplier data, generate clear shortage/surplus outputs, and keep Excel as the front end.

### 2) Extract canister locations from PDF (sheet 1)

* PDF lacked embedded text. Used **Amazon Textract** to convert scans into structured tables.
* Parsed Textract output with **Python (pandas)** to clean, normalize, and export **CSV** canister maps.

### 3) Standardize supplier data (sheet 2)

* Cleaned and normalized supplier spreadsheets in Python (types, casing, punctuation, unit artifacts).

### 4) Record linkage / entity resolution

* Drug names differed between sources. Implemented **fuzzy matching** (≥0.85 threshold) using Python (fuzzywuzzy) to attach the correct **canister** to each supplier row.
* Produced a **standardized inventory schema**: supplier format + new `canister` column.

### 5) Excel front end with advanced VBA

* Built an **.xlsm** workbook that:

  * Provides a **physical count entry sheet** keyed by canister (the canister number idicates where to find the drug). 
  * **Imports** weekly supplier quantities from file explorer (no hardcoded paths).
  * **Generates** a reconciled report: shortages, surpluses, physical and expected quantities.
* UX details: clear-inputs macro buttons, Navigation (hyperlinks), quick filters to focus on selected drug sets.

### 6) Controls, safety, and maintainability

* Parameterized matching threshold; guardrails around file imports.
* Separated any pharmacy-specific mappings from core logic.
* No PHI processed; demo uses fully **dummy data**.

---

## Outcomes

* **Paper eliminated** from the weekly count process.
* **Manual lookups removed** via standardization + canister linkage.
* **Reconciliation automated** in a familiar Excel interface (low training).


---

## Skills & Tools Demonstrated

* **Consulting & stakeholder alignment:** requirements gathering, iteration with practicing pharmacists, change management.
* **Data engineering:** OCR table extraction (Amazon Textract), pandas cleaning/normalization, CSV pipelines.
* **Entity resolution:** fuzzy matching for consolidating the same drugs with different name formats.
* **Advanced Excel/VBA:** macro-driven UI, dynamic imports, automated reporting, user-safe workflows.
* **Compliance mindset:** zero PHI, dummy demo assets.

---

## What I’d Improve Next

* **Replace OCR with first-party digital source** (eliminate PDFs entirely).
* **Power Query + parameterized transforms** for maintainable imports.
* **Optional Python service** (packaged) for matching at scale; add basic unit tests.
* **Dashboarding** (Power BI) for trendlines on shortages, wastage, and order cadence.
* **Extendability**: formal schema mapping to adapt across pharmacies/wholesalers.

---

## Usage Context

* Designed for **Microsoft Excel (desktop)** with macros enabled.
* Demo assets use **fully synthetic data**; no proprietary information is included.
