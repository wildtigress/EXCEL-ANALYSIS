

# 📊 GTM Employee Data Assessment

> **Excel Data Cleaning & Analysis Assessment**  
> Formula-driven employee data analysis with department summaries and bonus calculations

---

## 📌 Overview

This project involves **cleaning a messy employee dataset** and building a **fully formula-driven summary table** in Microsoft Excel. The assessment covers real-world data quality issues — inconsistent formatting, special characters in names, missing department information — and solves them using Excel functions and lookup formulas.

---

## 📂 Workbook Structure

| Sheet | Description |
|-------|-------------|
| `Assessment` | Task instructions and Q1/Q2 requirements |
| `data Employee` | Raw employee data (dirty) + cleaned version |
| `data Department` | Department ID to Name mapping |
| `data Salary Bonus Calculation` | Per-employee bonus % and calculated bonus amounts |
| `Summary` | Final formula-driven department summary table |

---

## 🧹 Q1 — Data Cleaning (`data Employee` tab)

### Issues Found & Fixed

| Issue | Example (Before) | Fixed (After) |
|-------|-----------------|---------------|
| Special characters in names | `Mark@Johnson` | `Mark Johnson` |
| Inconsistent capitalization | `jane smith` | `Jane Smith` |
| Comma in name | `Emily, Brown` | `Emily Brown` |
| Apostrophe in name | `Chris O'Brien` | `Chris OBrien` |
| Inconsistent date formats | Mixed formats | Standardized `YYYY-MM-DD` |

### Cleaned Employee Data

| Employee ID | Original Name | Cleaned Name | Department | Join Date | Salary |
|-------------|--------------|-------------|------------|-----------|--------|
| 101 | John Doe | John Doe | D001 | 2020-02-01 | ₹50,000 |
| 102 | jane smith | Jane Smith | D004 | 2021-03-12 | ₹55,000 |
| 103 | Mark@Johnson | Mark Johnson | D003 | 2020-05-10 | ₹60,000 |
| 104 | Sam Patel | Sam Patel | D001 | 2022-06-10 | ₹48,000 |
| 105 | Emily, Brown | Emily Brown | D004 | 2021-07-15 | ₹52,000 |
| 106 | Chris O'Brien | Chris OBrien | D005 | 2022-08-15 | ₹49,500 |
| 107 | Robert Brown | Robert Brown | D002 | 2021-11-25 | ₹57,500 |
| 108 | Linda White | Linda White | Unknown | 2022-04-30 | ₹54,000 |
| 109 | James Black | James Black | D003 | 2020-09-13 | ₹61,500 |
| 110 | Sarah Miller | Sarah Miller | D001 | 2021-12-09 | ₹58,000 |

---

## 📋 Q2 — Department Summary Table (`Summary` tab)

All values are **formula-driven** — no hardcoded calculations.

### Final Summary

| Dept ID | Department | Employees | Employee List | Avg Salary | Total Bonus |
|---------|-----------|-----------|--------------|-----------|-------------|
| D001 | Sales | 3 | John Doe, Sam Patel, Sarah Miller | ₹52,000 | ₹15,600 |
| D002 | Marketing | 1 | Robert Brown | ₹57,500 | ₹5,750 |
| D003 | IT | 2 | Mark Johnson, James Black | ₹60,750 | ₹12,150 |
| D004 | HR | 2 | Jane Smith, Emily Brown | ₹53,500 | ₹10,700 |
| D005 | Finance | 1 | Chris OBrien | ₹49,500 | ₹4,950 |
| Unknown | Unknown | 1 | Linda White | ₹54,000 | ₹5,400 |

### Key Formula Logic

```excel
# Department Name (with Unknown fallback)
=IFERROR(VLOOKUP(A2,'data Department'!$A:$B,2,0),"Unknown")

# Number of Employees
=COUNTIF('data Employee'!$D:$D, A2)

# Average Salary
=AVERAGEIF('data Employee'!$D:$D, A2, 'data Employee'!$F:$F)

# Bonus % (default 10% if missing)
=IFERROR(IF(C2="",0.1,C2), 0.1)

# Total Bonus
=SUMPRODUCT(
  ('data Employee'!$D$2:$D$11=A2)*
  ('data Employee'!$F$2:$F$11)*
  ('data Salary Bonus Calculation'!$C$2:$C$11)
)
```

---

## 💰 Bonus Calculation Rules

From `data Salary Bonus Calculation` sheet:

| Rule | Logic |
|------|-------|
| Bonus % provided | Use the given percentage |
| Bonus % is missing/blank | Default to **10%** |
| Bonus Amount | `= Salary × Bonus %` |

### Bonus Details Per Employee

| Employee ID | Salary | Bonus % | Bonus Amount |
|-------------|--------|---------|-------------|
| 101 – John Doe | ₹50,000 | 10% | ₹5,000 |
| 102 – Jane Smith | ₹55,000 | 10% (default) | ₹5,500 |
| 103 – Mark Johnson | ₹60,000 | 12% | ₹7,200 |
| 104 – Sam Patel | ₹48,000 | 8% | ₹3,840 |
| 105 – Emily Brown | ₹52,000 | 10% (default) | ₹5,200 |
| 106 – Chris OBrien | ₹49,500 | 9% | ₹4,455 |
| 107 – Robert Brown | ₹57,500 | 11% | ₹6,325 |
| 108 – Linda White | ₹54,000 | 10% (default) | ₹5,400 |
| 109 – James Black | ₹61,500 | 13% | ₹7,995 |
| 110 – Sarah Miller | ₹58,000 | 10% | ₹5,800 |

---

## 🛠️ Excel Skills Demonstrated

| Skill | Used For |
|-------|---------|
| `VLOOKUP` | Department name lookup from ID |
| `COUNTIF` | Count employees per department |
| `AVERAGEIF` | Average salary per department |
| `SUMPRODUCT` | Total bonus per department |
| `IFERROR` | Handle missing/null values gracefully |
| `IF` | Default bonus % fallback logic |
| Text functions (`TRIM`, `PROPER`, `SUBSTITUTE`) | Name cleaning |
| Data validation | Ensuring clean inputs |
| Formula-only approach | No hardcoded calculated values |

---

## 📁 Project Structure

```
gtm-employee-assessment/
│
├── GTM_Assessment.xlsx           # Main workbook with all 5 sheets
│   ├── Assessment                # Task instructions (Q1 & Q2)
│   ├── data Employee             # Raw + cleaned employee data
│   ├── data Department           # Department reference table
│   ├── data Salary Bonus...      # Bonus calculation table
│   └── Summary                   # Final formula-driven summary
│
└── README.md
```

---

## 🔑 Key Design Decisions

- **Everything formula-driven** — all summary values auto-update if source data changes
- **Unknown department fallback** — employees with no department ID mapped to "Unknown"
- **Default bonus logic** — missing bonus % treated as 10% automatically via `IFERROR`
- **VLOOKUP for department mapping** — avoids hardcoding department names

---

## 📬 Contact

**Samiksha Barnwal**  
📧 [099samiksha@gmail.com](mailto:099samiksha@gmail.com)  
🐙 [github.com/wildtigress](https://github.com/wildtigress)  
💼 [linkedin.com/in/samiksha4](https://linkedin.com/in/samiksha4)

---

*Built with Microsoft Excel — Data Cleaning, VLOOKUP, SUMPRODUCT, and Formula-Driven Analysis*
