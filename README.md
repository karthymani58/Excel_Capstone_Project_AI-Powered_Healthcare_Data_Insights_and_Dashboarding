# 🏥 Healthcare Data Analysis & Executive Dashboarding Implementation

![Excel](https://img.shields.io/badge/Tool-Microsoft_Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)
![Data Analysis](https://img.shields.io/badge/Domain-Healthcare_Analytics-0284c7?style=flat)
![Status](https://img.shields.io/badge/Status-Completed-10b981?style=flat)

---

## 📌 Section 1: Executive Summary

**Project Objective:** To clean, transform, and analyze multi-source healthcare data (Medical Examinations, Hospitalization Details, and Customer Profiles) to evaluate cost drivers, risk factors, and patient demographics using Excel formulas, lookup functions, AI Quick Analysis tools, and dynamic dashboards.

### Scope of Work Overview:
* **Data Cleaning:** Audited missing values (`?`), executed mode/mean imputations, and reconciled data by removing 8 invalid/orphan records.
* **Data Transformation:** Split customer names, standardized surgical metrics, calculated patient age relative to baseline (08-Jun-2023), assigned BMI diagnostic categories, and standardized financial currency metrics.
* **Data Integration:** Merged source tables into a unified master worksheet (`Healthcare`) via relational lookups (`VLOOKUP`).
* **Manual & AI Analysis:** Built cross-tabulation Pivot Tables, Donut and Column charts, AI Quick Analysis summary metrics, and an executive dashboard with interactive slicers.

---

## 📊 Summary Key Metrics

| Metric | Calculated Value |
| :--- | :--- |
| **Total Customer Records** | `2,335` |
| **Total Hospitalization Records** | `2,335`|
| **Total Healthcare Expenditure** | `$31,592,358.61` |
| **Average Cost Per Patient** | `$13,529.92` |
| **Average Patient Age** | `39.2 Years` *(Baseline: 08-Jun-2023)* |
| **Data Purge / Reconciled Rows** | `8 Invalid/Orphan Rows Removed` |

---

## 🛠️ Section 2: Comprehensive Step-by-Step Implementation

### Phase 1: Data Cleaning

#### Step 1.1: Missing Value Audit & Reconciliation (Removal of 8 Invalid Rows)
* **What:** Audited all sheets for placeholder symbols (`?`) and un-linkable customer records.
* **Why:** Reconciled row discrepancies between source tables to ensure a clean 1:1 join across all datasets.
* **How:** Identified 8 invalid records in *Hospitalisation Details* (6 rows with `Customer ID = "?"`, and 2 orphan rows with `Customer ID = "id2444"` and `"id3444"`). Deleted these 8 rows, reducing *Hospitalisation Details* from 2,343 to 2,335 rows to perfectly match Customer Names and Medical Examinations[cite: 1].

#### Step 1.2: Imputing Date Values (Month & Year)
* **What:** Filled missing month entries with `"Sep"` and missing year entries with the rounded average dataset year.
* **Why:** Prepares complete temporal records for conversion into structured date objects.
* **Calculation:** Mean Year = $1982.84 \rightarrow \mathbf{1983}$.
* **Formula Implemented:**
  ```excel
  =IF(B2="?", ROUND(AVERAGE($B$2:$B$2336), 0), B2)
  ```

#### Step 1.3: Mode Imputation for Categorical Fields
* **What:** Identified the statistical mode (most frequent category) for `smoker`, `Hospital tier`, and `City tier`, replacing missing `?` entries.
* **Why:** Preserves categorical distributions without introducing bias.
* **Imputed Values:**
  * `smoker`: `"No"` (1,845 occurrences)
  * `Hospital tier`: `"tier - 2"` (1,339 occurrences)
  * `City tier`: `"tier - 2"` (812 occurrences)
* **Formula Implemented:**
  ```excel
  =INDEX($G$2:$G$2336, MODE(MATCH($G$2:$G$2336, $G$2:$G$2336, 0)))
  ```

#### Step 1.4: State ID Standardization
* **What:** Replaced missing State ID entries (`?`) with `"Unknown"`.
* **Why:** Prevents breaking lookup formulas, pivot grouping, and dashboard slicers.
* **Formula Implemented:**
  ```excel
  =IF(I2="?", "Unknown", I2)
  ```

---

### Phase 2: Data Transformation

#### Step 2.1: Splitting Customer Names
* **What:** Separated the full name string (e.g., `"Hawks, Ms. Kelly"`) in Customer Names into **Title**, **First Name**, and **Last Name**.
* **Why:** Normalizes customer identity metrics for structured relational queries.
* **Execution:** Used *Text to Columns* (delimited by comma and dot).

#### Step 2.2: Cleaning Major Surgeries Count
* **What:** Converted `NumberOfMajorSurgeries` into a clean numeric column by replacing `"No major surgery"` with `0`.
* **Why:** Enables numerical aggregations (such as sum, average, and risk-correlations) in pivot tables.
* **Formula Implemented:**
  ```excel
  =IF(G2="No major surgery", 0, VALUE(G2))
  ```

#### Step 2.3: Consistency Check on Categorical Flags
* **What:** Standardized text casing across `Heart Issues` and `smokers`.
* **Why:** Eliminates duplicate grouping rows in pivot tables caused by case variations (e.g., `"yes"`, `"YES"`, `"Yes"`).
* **Formula Implemented:**
  ```excel
  =PROPER(TRIM(Cell))
  ```

#### Step 2.4: Weight Status Classification
* **What:** Created a calculated column `"Weight Status"` based on Body Mass Index (BMI).
* **Why:** Categorizes continuous numerical BMI metrics into standard diagnostic categories.
* **Formula Implemented:**
  ```excel
  =IF(BMI_Cell<18.5, "Underweight", IF(BMI_Cell<=24.9, "Normal Weight", IF(BMI_Cell<=29.9, "Overweight", "Obesity")))
  ```

#### Step 2.5: Generating Date of Birth
* **What:** Concatenated date, month, and year into a formatted Date of Birth field (`DD-MMM-YYYY`).
* **Why:** Unifies individual date components into a standard date object.
* **Formula Implemented:**
  ```excel
  =DATEVALUE(CONCATENATE(D2, "-", C2, "-", B2))
  ``` *(Formatted cell via Format Cells $\rightarrow$ Custom $\rightarrow$ `DD-MMM-YYYY`)*

#### Step 2.6: Calculating Patient Age
* **What:** Calculated current patient age relative to the dataset baseline date (8th June 2023).
* **Why:** Enables demographic age grouping against healthcare expenditure.
* **Formula Implemented:**
  ```excel
  =DATEDIF(DOB_Cell, DATE(2023, 6, 8), "Y")
  ```

#### Step 2.7: Currency Formatting
* **What:** Formatted charges as Currency (`$#,##0.00`).
* **Why:** Ensures standardized financial presentation.

---

### Phase 3: Master Table Integration ("Healthcare")

#### Step 3.1: Unified "Healthcare" Master Table Creation
* **What:** Created a central worksheet named **`Healthcare`** merging all three source tables into a single master sheet containing 2,335 rows using `Customer ID` as the primary key.
* **Why:** Consolidates data to enable multi-dimensional pivot tables and unified dashboard slicers.
* **Formula Implemented (VLOOKUP Approach):**
  ```excel
  =VLOOKUP(A2, 'Customer Names'!$A:$D, 3, FALSE)
  ```[cite: 1]

#### Retained Master Schema (17 Fields):
* `Customer ID (Primary Key)`
* `First Name`
* `BMI`
* `HBA1C`
* `Heart Issues`
* `Any Transplants`
* `Cancer History`
* `Number of Major Surgeries`
* `Smoker`
* `Weight Status`
* `Diabetes Status`
* `Date of Birth`
* `Charges`
* `Hospital Tier`
* `City Tier`
* `State ID`
* `Age`

---

### Phase 4: AI & Manual Visual Analysis and Dashboard Assembly

#### Step 4.1: Manual Pivot Table & Custom Visual Insights
1. **Cancer History Distribution among Smokers vs. Non-Smokers (Donut Chart):**
   * **Objective:** Evaluate cancer rates between smokers and non-smokers.
   * **Insight:** Active smokers display a disproportionately high rate of recorded cancer cases compared to non-smokers, establishing smoking as a primary cost and health risk driver.
2. **Hospitalization Charges by Weight Status (Column Chart):**
   * **Objective:** Compare total and average medical charges across BMI weight categories.
   * **Insight:** Patients in the Obesity category incur significantly higher total and average hospitalization costs, driven by chronic comorbidities and extended hospital stays.

#### Step 4.2: AI-Driven Summary Statistics & Quick Analysis
* **What:** Highlighted the master dataset and used Excel Quick Analysis Tools (`Ctrl + Q`) to compute key metrics and generate summary charts.
* **Key Calculated Metrics:**
  * **Total Customers:** `2,335`
  * **Total Hospitalization Count:** `2,335`
  * **Total Healthcare Costs:** `$31,592,358.61`
  * **Average Healthcare Cost:** `$13,529.92`
  * **Average Patient Age:** `39.2 Years`

#### Step 4.3: Interactive Healthcare Dashboard Assembly
* **Layout Structure:** Positioned top-level **KPI Summary Cards** across the top and arranged analytical charts in an aligned, two-column grid.
* **Interactive Controls:** Added dynamic Excel **Slicers** connected across all pivot charts for **Weight Status** and **Diabetes Status**.
* **Formatting:** Formatted header banner as `"Healthcare Dashboard - AI Analysis"`, updated chart palettes for visual harmony, removed gridlines, and configured cross-filtering via *PivotTable Analyze $\rightarrow$ Filter Connections*.

---
