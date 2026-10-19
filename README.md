# Employee-Attrition-Analysis


An Excel-based analysis of employee attrition to find **who leaves, why they leave, and what HR can do about it**.

## Business Question

> How can the organization reduce employee attrition and improve retention?

Every employee who leaves costs recruitment, onboarding and lost-productivity money. Knowing which groups are most at risk lets HR focus retention effort where it saves the most.

## Dataset

- **2,925 employees**, one row each (`Emp ID` is the unique key)
- **473 ex-employees** and 2,452 current employees
- Fields cover department, job role, age, gender, marital status, education, distance from home, overtime, travel, job and environment satisfaction, work-life balance, performance rating, monthly income, salary hike, stock options, tenure and promotion history

## Workbook Structure

File: `Employee_Attrition_Analysis_Solution.xlsx`

| Sheet | What it contains |
|---|---|
| **Dataset** | Cleaned data plus calculated columns: Attrition, Attrition Flag, Age Group, Distance Band, Income Band, Tenure Band |
| **Data Cleaning Metrics** | Cleaning notes, current vs attrition table and pie chart |
| **Job & Organizational Factors** | Department, job role, years at company (line chart), tenure band, education level |
| **Demographic Factors** | Age group, gender, marital status, distance from home |
| **Work Experience & Satisfaction** | Job satisfaction, environment satisfaction, overtime, performance rating |
| **Compensation** | Income band, average salary hike for leavers vs current employees |
| **Analysis** | Dashboard with all charts, conclusions, at-risk employee profile and HR recommendations |

## Approach

1. **Data preparation:** checked for duplicates and missing values, standardized the age bands, and added calculated columns.
2. **Attrition rate, not just counts:** every table shows leavers, total employees and **attrition rate (leavers ÷ total)**. Raw counts make large groups look worse than they are, so all charts plot the rate.
3. **Excel formulas:** all numbers are live `COUNTIFS` / `AVERAGEIFS` formulas on named ranges, so the workbook updates if the data changes.
4. **Visualization:** bar and column charts for group comparisons, a line chart for tenure, and a pie chart for the overall split.
5. **Insights:** findings are linked to the analysis sheets and turned into recommendations.

## Key Findings

Overall attrition rate: **16.2%**

| Factor | Finding |
|---|---|
| **Overtime** | 30.6% attrition with overtime vs 10.5% without |
| **Age** | Under 25: 39.8%, 25–34: 20.3%, 35–44: 10.0% |
| **Job role** | Sales Representative 40.5%, Laboratory Technician 24.0%, Human Resources 23.1% |
| **Department** | Sales 20.7%, HR 19.0%, R&D 13.9% |
| **Tenure** | 34.9% in the first two years vs 6.7% for 11–20 years |
| **Marital status** | Single 25.4% vs married 12.5% |
| **Income** | Under 5,000: 21.9%, 5,000–9,999: 11.1%, 10,000+: 8.9% |
| **Satisfaction** | Job satisfaction 1: 22.9% vs 4: 11.5%. Environment satisfaction 1: 25.5% vs 4: 13.4% |
| **Commute** | 21–30 km: 22.1% vs 1–5 km: 13.7% |
| **Gender** | Men 17.1% vs women 14.8%, not a strong driver |
| **Performance rating** | Rating 3: 16.0% vs rating 4: 16.9%, almost no difference |
| **Salary hike** | Average hike of 15.1% for leavers vs 15.2% for stayers, so hikes alone do not reduce attrition |

## Recommendations for HR

1. Hold regular stay-interviews and check-ins with disengaged employees, especially those with low job or environment satisfaction.
2. Improve the work environment by collecting feedback quarterly and acting on it visibly.
3. Cap overtime, rebalance workloads, and reward employees who regularly work extra hours.
4. Build development paths, mentoring and learning options for employees under 35.
5. Strengthen onboarding with 30/60/90-day check-ins, because the first two years carry the highest attrition.
6. Review pay for employees earning under 5,000 and for high-attrition roles (Sales Representative, Laboratory Technician, HR).
7. Offer flexible or remote-work options to employees with long commutes.

## Limitations

- The dataset has **no date column**, so yearly, quarterly and seasonal trends could not be calculated. Years at company is used as the time-based view.
- The dataset has **no reason-for-leaving column**, so voluntary vs involuntary attrition could not be analysed.
- Performance ratings contain only 3 and 4, so there are no low performers to compare.
- The analysis is descriptive. It shows which groups leave more often, not what causes it.

## Tools Used

- Microsoft Excel (formulas, named ranges, charts)
- Functions: `COUNTIFS`, `AVERAGEIFS`, `SUMPRODUCT`, `IF`, `IFERROR`, `TEXT`

GitHub: [Keerthanahegde08](https://github.com/Keerthanahegde08)
