# 📊 Excel Salary Dashboard — Data Jobs (2023)

[![Live Demo](https://img.shields.io/badge/Live%20Demo-View%20Here-4f9cf7?style=for-the-badge)](https://KrishnaSai315.github.io/Excel-Salary-Dashboard/excel_salary_dashboard_demo.html)
[![Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)]()
[![Dataset](https://img.shields.io/badge/Dataset-32K%20Records-orange?style=for-the-badge)]()

> An interactive Excel dashboard for exploring real-world data science job salaries across job titles, countries, and schedule types.

---

## 🧭 Table of Contents

- [Overview](#overview)
- [Live Demo](#live-demo)
- [Dashboard Features](#dashboard-features)
- [Dataset](#dataset)
- [Excel Skills Used](#excel-skills-used)
- [Key Formulas Explained](#key-formulas-explained)
- [How to Use the Dashboard](#how-to-use-the-dashboard)
- [Key Insights](#key-insights)
- [File Structure](#file-structure)
- [About](#about)

---

## Overview

This **Salary Dashboard** was built to help data professionals and job seekers:

- 🔍 Investigate median salaries for specific data roles
- 🌍 Compare compensation across countries and schedule types
- 💡 Quickly identify whether a job offer reflects market rates

The dataset contains **real-world data science job listings from 2023**, covering job titles, salary figures, locations, and required skills — all presented through an interactive, no-code Excel interface.

**Data scope:** 2023 · Global · ~32,000 job listings across 9 major data roles

---

## Live Demo

👉 **[View Interactive Demo](./excel_salary_dashboard_demo.html)** — open in any browser, no Excel required.

---

## Dashboard Features

### 1. 📊 Salary Bar Chart
A horizontal bar chart ranks all job titles by descending **median salary**. Color-coded and sorted, making it easy to compare Engineer, Scientist, and Analyst roles at a glance.

> **Key finding:** Senior roles and Engineers command significantly higher salaries than Analyst-level roles.

### 2. 🗺️ Country Median Salary Map
An Excel map chart plots median salary by country using color intensity. Countries with more data appear more defined; regions with no data are left blank.

> **Key finding:** The United States consistently shows the highest median salaries globally; European and Asian markets trail by 20–40%.

### 3. 🎛️ Three Interactive Filters
Users can slice salary data across **three dimensions simultaneously**:

| Filter | Options |
|--------|---------|
| **Job Title** | 9 standardized roles (Data Analyst → Senior Data Engineer) |
| **Country** | 50+ countries with sufficient data |
| **Schedule Type** | Full-time, Part-time, Contractor, Temp, Internship |

All three filters drive a single **Median Salary** result card that updates instantly.

---

## Dataset

| Field | Description |
|-------|-------------|
| `job_title_short` | Standardized role (e.g., Data Analyst, Data Scientist) |
| `salary_year_avg` | Average annual salary (USD) |
| `salary_hour_avg` | Average hourly rate |
| `job_country` | Country where the role is based |
| `job_schedule_type` | Full-time, Part-time, Contractor, etc. |
| `job_skills` | Required technical skills |
| `company_name` | Hiring company |
| `job_work_from_home` | Whether remote work is available |

---

## Excel Skills Used

| Skill | Purpose |
|-------|---------|
| **Bar Chart** | Compare median salaries across all job titles |
| **Map Chart** | Visualize geographic salary distribution |
| **MEDIAN + IF (Array Formula)** | Compute filtered median salary |
| **FILTER Function** | Generate a clean unique list of schedule types |
| **Data Validation** | Restrict dropdowns to valid, filtered options |
| **Named Ranges** | Power the dynamic filtering logic |

---

## Key Formulas Explained

### Median Salary — Multi-Criteria Array Formula

\`\`\`excel
=MEDIAN(
  IF(
    (jobs[job_title_short]=A2)*
    (jobs[job_country]=country)*
    (ISNUMBER(SEARCH(type,jobs[job_schedule_type])))*
    (jobs[salary_year_avg]<>0),
    jobs[salary_year_avg]
  )
)
\`\`\`

**What it does:**
- Filters rows where job title, country, and schedule type all match
- Excludes records with zero or blank salary
- Returns the **median** (not average) of matching salary values
- The `*` between conditions acts as AND — all must be TRUE

> ⚠️ **Array formula:** In older Excel versions, press `Ctrl + Shift + Enter` to enter this formula.

---

### Unique Schedule Type List — FILTER Formula

\`\`\`excel
=FILTER(
  J2#,
  (NOT(ISNUMBER(SEARCH("and",J2#))+ISNUMBER(SEARCH(",",J2#))))*(J2#<>0)
)
\`\`\`

**What it does:**
- Removes combined entries like "Full-time and Part-time"
- Removes blank/zero entries
- Returns a **clean list** used to populate the Schedule Type dropdown

---

## How to Use the Dashboard

| Step | Action | Result |
|------|--------|--------|
| **1** | Open `1_Salary_Dashboard.xlsx` | Dashboard loads with all jobs visible |
| **2** | Click the **Job Title** dropdown | Select any of the 9 data roles |
| **3** | Click the **Country** dropdown | Filter to a specific country or keep "All" |
| **4** | Click the **Schedule Type** dropdown | Choose Full-time, Part-time, Contractor, etc. |
| **5** | Read the **Median Salary** card | Result updates automatically |
| **6** | Read the **Bar Chart** | Bars update to reflect your selection |

> 💡 **Tip:** Start with "All" in Country and Type to see the broadest market view. Then narrow down to get a precise benchmark.

---

## Key Insights

- 🏆 **Senior Data Engineers** earn the highest median salary (~$147,500/yr in the US)
- 🇺🇸 **Full-time US roles** consistently pay 15–25% above the global median for the same title
- 📈 **Data Analysts** have the widest salary spread — from ~$50K entry-level to $150K+ senior
- 🌐 **Remote roles** show a slightly higher median, skewing toward senior-level hires
- ⏱️ **Contractor/hourly** top rates ($65–$85/hr) annualize to over $130K

---

## File Structure

\`\`\`
📁 Excel-Salary-Dashboard/
│
├── 📊 1_Salary_Dashboard.xlsx          ← Main dashboard file (open this)
├── 📄 README.md                         ← This file
├── 🌐 excel_salary_dashboard_demo.html  ← Browser-based interactive demo
│
└── 📁 0_Resources/
    └── 📁 Images/
        ├── 1_Salary_Dashboard_Final_Dashboard.gif
        ├── 1_Salary_Dashboard_Chart1.png
        ├── 1_Salary_Dashboard_Country_Map.gif
        ├── 1_Salary_Dashboard_Data_Validation.gif
        └── ...screenshots
\`\`\`

---

## About

**Built by:** Loknadh Venkata Krishna Sai Kona
**GitHub:** [KrishnaSai315](https://github.com/KrishnaSai315)
**LinkedIn:** [lvkrishna3](https://linkedin.com/in/lvkrishna3/)

**Tools used:** Microsoft Excel 2021 / Microsoft 365
**Dataset:** 2023 Data Science Job Listings (real-world, ~32,000 entries)
**Project type:** Data Analytics Portfolio Project

---

> ⭐ If this dashboard helped you understand data job salaries, consider **starring** the repository!
