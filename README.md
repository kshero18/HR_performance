*A workforce analytics and data governance mini-project*

This project is a structured HR Monthly Reporting Pack built using Excel, showcasing workforce analytics, HR KPI reporting and data governance practices. It includes a dashboard summarising headcount, tenure and attrition across departments, along with supporting documentation such as a Data Issues Log, Data Dictionary, and KPI Definitions—mirroring real workforce reporting workflows in People & Culture teams.

**Full Report (PDF)**
You can view the full exported reporting pack:
>> HR_Monthly_Report_2025_11_24.pdf

**Tools Used**
1. Microsoft Excel (PivotTables, Charting, Table Design)
2. VBA for automated PDF export
3. Data validation and documentation
4. Power Query (optional, for cleaning/extending the dataset)
5. Kaggle dataset (IBM HR Analytics)

**Project Overview**

This report provides a monthly snapshot of workforce composition across three departments:
1. Human Resources
2. Research & Development
3. Sales
The dashboard summarises key HR metrics and presents them clearly for leadership review. Supporting documentation ensures the reporting process is transparent, repeatable and aligned with data governance standards.

**Key Workforce Metrics**
1. Headcount
>> HR: 63
>> R&D: 961
>> Sales: 446

2. Average Tenure (Years)
>> HR: 7.24 years
>> R&D: 6.86 years
>> Sales: 7.28 years

3. Attrition Rate (% by Department)
>> HR: 19.05%
>> R&D: 13.84%
>> Sales: 20.63%
These metrics were calculated using pivot tables and visualised with comparative bar and column charts.

**Features & Contributions**
*Excel Dashboard Design*
1. Built pivot-table-driven KPIs for headcount, tenure and attrition
2. Designed clean visual comparisons across departments
3. Created a monthly template that can be refreshed with new data

*Data Issues Log (DQ-001 to DQ-007)*
Documented real data problems found during analysis, including:
1. Incorrect KPI aggregation (e.g., sum of averages)
2. Missing reporting period
3. Misleading pivot labels (e.g., “Sum of AttritionRatePct”)
4. Limited departmental coverage
5. Need for standardised KPI definitions
This demonstrates strong data quality awareness.

*Data Dictionary*
Comprehensive metadata documentation covering:
1. Employee Master Data
2. Job & Position Fields
3. Compensation Fields
4. Performance & Engagement Fields
5. Attrition Fields

Structured with:
1. Field Name
2. Description
3. Data Type
4. Allowed Values
5. Validation Rules
6. Notes
This mirrors real-world data governance documentation.

*KPI Definitions Sheet*
Defined formulas, HR interpretations and use cases for key metrics:
1. Headcount
2. Attrition Rate
3. Tenure
4. Performance Rating
5. Overtime Rate
6. Compensation KPIs

*Automated PDF Export (VBA)*
A VBA script exports the following into a multi-page PDF:
1. Dashboard
2. Data Issues Log
3. Data Dictionary
4. KPI Definitions
Creates a professional monthly reporting pack.

**Why This Project Matters**
This project demonstrates practical experience in:
>> Workforce analytics & P&C reporting
>> Data quality detection and documentation
>> KPI formulation and validation
>> Metadata and governance principles
>> Clear communication of insights
>> Dashboard building and automation
>> Turning raw HR data into actionable insights

It closely aligns with roles involving:
>> Workforce Systems
>> HR Reporting
>> People Analytics
>> Data Governance
>> Data Quality
>> Public Sector Workforce Reporting

**Dataset**
*Kaggle*: IBM HR Analytics Attrition Dataset
https://www.kaggle.com/datasets/pavansubhasht/ibm-hr-analytics-attrition-dataset

**Future Improvements**
>> Adding month-over-month trends
>> Adding forecasting (tenure decline, attrition prediction)
>> Integrating Power Query for automated load/cleaning
>> Building a Power BI version of the dashboard
>> Adding a Department and Role drill-down view
