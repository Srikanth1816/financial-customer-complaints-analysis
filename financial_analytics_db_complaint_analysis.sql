/*
PROJECT: Financial Customer Complaints Analysis
DIALECT: MySQL 8.0+
SOURCE: Cleaned Excel workbook -> Cleaned Data worksheet
TABLE: complaints

IMPORT SETUP
1. In MySQL Workbench, create/select your project database.
2. Export the workbook's "Cleaned Data" sheet as CSV (UTF-8).
3. Use Table Data Import Wizard to import it into a table named `complaints`.
4. Keep the original 14 source columns. Dates should be DATE values.
5. The two calculated columns `response_time_days` and
   `timely_response_validated` are not required for these queries; they can
   be calculated from the source dates/status as shown below.

IMPORTANT DEFINITIONS
- One complaint = one unique complaint_id.
- Open complaints are excluded from response-time and timely-response KPIs.
- A closed complaint is eligible for response KPIs only when its response
  date is present and its received/due dates are valid.
- Timely = date_company_responded <= response_due_date.
- Complaint counts by company are volumes, NOT company performance rates.
*/

-- Optional database and table definition. If you use the Workbench Import
-- Wizard, let it create the table instead, or create this table first.
CREATE DATABASE IF NOT EXISTS financial_analytics_db;
USE financial_analytics_db;

CREATE TABLE IF NOT EXISTS complaints (
    complaint_id VARCHAR(30) PRIMARY KEY,
    date_received DATE NOT NULL,
    financial_product VARCHAR(100),
    issue_category VARCHAR(150),
    issue_detail VARCHAR(200),
    company_name VARCHAR(150),
    state VARCHAR(20),
    submission_channel VARCHAR(50),
    response_due_date DATE,
    date_company_responded DATE NULL,
    response_status VARCHAR(30),
    company_response VARCHAR(100),
    timely_response VARCHAR(10),
    short_complaint_summary TEXT
);

-- Q1. Total complaints
-- Counts unique complaint IDs rather than rows.
SELECT COUNT(DISTINCT complaint_id) AS total_complaints
FROM complaints;

-- Q2. Monthly complaint trend
-- Groups complaint volume by calendar month of receipt.
SELECT
    DATE_FORMAT(date_received, '%Y-%m') AS complaint_month,
    COUNT(DISTINCT complaint_id) AS total_complaints
FROM complaints
WHERE date_received IS NOT NULL
GROUP BY DATE_FORMAT(date_received, '%Y-%m')
ORDER BY complaint_month;

-- Q3. Complaints by financial product
-- Shows which product categories account for the largest complaint volumes.
SELECT
    financial_product,
    COUNT(DISTINCT complaint_id) AS total_complaints
FROM complaints
GROUP BY financial_product
ORDER BY total_complaints DESC;

-- Q4. Most frequently reported issues
-- Uses the primary issue category for a consistent issue-level comparison.
SELECT
    issue_category,
    COUNT(DISTINCT complaint_id) AS total_complaints
FROM complaints
GROUP BY issue_category
ORDER BY total_complaints DESC
LIMIT 10;

-- Q5. Complaints by company
-- This is a raw volume comparison only; company sizes/customer bases may differ.
SELECT
    company_name,
    COUNT(DISTINCT complaint_id) AS total_complaints,
    SUM(CASE WHEN LOWER(response_status) = 'open' THEN 1 ELSE 0 END) AS open_complaints,
    SUM(CASE WHEN LOWER(response_status) = 'closed' THEN 1 ELSE 0 END) AS closed_complaints
FROM complaints
GROUP BY company_name
ORDER BY total_complaints DESC;

-- Q6. Complaints by customer state
-- Missing source locations are represented as 'Unknown' in the cleaned data.
SELECT
    COALESCE(NULLIF(TRIM(state), ''), 'Unknown') AS customer_state,
    COUNT(DISTINCT complaint_id) AS total_complaints
FROM complaints
GROUP BY COALESCE(NULLIF(TRIM(state), ''), 'Unknown')
ORDER BY total_complaints DESC;

-- Q7. Timely-response percentage by company
-- Denominator includes only closed complaints with both response and due dates.
-- Open complaints and incomplete response records are excluded.
SELECT
    company_name,
    COUNT(*) AS eligible_closed_complaints,
    SUM(CASE WHEN date_company_responded <= response_due_date THEN 1 ELSE 0 END)
        AS timely_closed_complaints,
    ROUND(
        100.0 * SUM(CASE WHEN date_company_responded <= response_due_date THEN 1 ELSE 0 END)
        / NULLIF(COUNT(*), 0),
        2
    ) AS timely_response_percentage
FROM complaints
WHERE LOWER(response_status) = 'closed'
  AND date_company_responded IS NOT NULL
  AND response_due_date IS NOT NULL
GROUP BY company_name
ORDER BY timely_response_percentage DESC;

-- Q8. Average response time
-- Uses calendar days from receipt to response for eligible closed complaints.
-- Open complaints are explicitly excluded.
SELECT
    ROUND(AVG(DATEDIFF(date_company_responded, date_received)), 2)
        AS average_response_time_days,
    COUNT(*) AS eligible_closed_complaints
FROM complaints
WHERE LOWER(response_status) = 'closed'
  AND date_company_responded IS NOT NULL
  AND date_received IS NOT NULL;

-- Q9. Overall timely-response KPI
-- A portfolio-level percentage using the same eligible closed-complaint rule.
SELECT
    COUNT(*) AS eligible_closed_complaints,
    SUM(CASE WHEN date_company_responded <= response_due_date THEN 1 ELSE 0 END)
        AS timely_closed_complaints,
    ROUND(
        100.0 * SUM(CASE WHEN date_company_responded <= response_due_date THEN 1 ELSE 0 END)
        / NULLIF(COUNT(*), 0),
        2
    ) AS overall_timely_response_percentage
FROM complaints
WHERE LOWER(response_status) = 'closed'
  AND date_company_responded IS NOT NULL
  AND response_due_date IS NOT NULL;

-- Q10. Open complaint backlog by age band
-- Helps operations prioritize unresolved complaints. Age is measured as of
-- the current date; use a fixed date instead if reproducing a historical snapshot.
SELECT
    CASE
        WHEN DATEDIFF(CURDATE(), date_received) <= 7 THEN '0-7 days'
        WHEN DATEDIFF(CURDATE(), date_received) <= 30 THEN '8-30 days'
        WHEN DATEDIFF(CURDATE(), date_received) <= 60 THEN '31-60 days'
        ELSE 'Over 60 days'
    END AS open_complaint_age_band,
    COUNT(DISTINCT complaint_id) AS open_complaints
FROM complaints
WHERE LOWER(response_status) = 'open'
GROUP BY open_complaint_age_band
ORDER BY
    CASE open_complaint_age_band
        WHEN '0-7 days' THEN 1
        WHEN '8-30 days' THEN 2
        WHEN '31-60 days' THEN 3
        ELSE 4
    END;

-- Q11. Source timeliness flag vs date-based validation
-- Audits whether the supplied timely_response field agrees with the dates.
-- Only closed complaints with both response and due dates are evaluated.
SELECT
    LOWER(TRIM(timely_response)) AS source_timely_flag,
    CASE
        WHEN date_company_responded <= response_due_date THEN 'Yes'
        ELSE 'No'
    END AS date_validated_flag,
    COUNT(*) AS complaint_count
FROM complaints
WHERE LOWER(response_status) = 'closed'
  AND date_company_responded IS NOT NULL
  AND response_due_date IS NOT NULL
GROUP BY
    LOWER(TRIM(timely_response)),
    CASE
        WHEN date_company_responded <= response_due_date THEN 'Yes'
        ELSE 'No'
    END
ORDER BY complaint_count DESC;

-- Q12. Monthly timely-response trend
-- Tracks timeliness by month the complaint was received, not the month closed.
-- Open complaints are excluded from each month's denominator.
SELECT
    DATE_FORMAT(date_received, '%Y-%m') AS complaint_month,
    COUNT(*) AS eligible_closed_complaints,
    ROUND(
        100.0 * SUM(CASE WHEN date_company_responded <= response_due_date THEN 1 ELSE 0 END)
        / NULLIF(COUNT(*), 0),
        2
    ) AS timely_response_percentage
FROM complaints
WHERE LOWER(response_status) = 'closed'
  AND date_company_responded IS NOT NULL
  AND response_due_date IS NOT NULL
GROUP BY DATE_FORMAT(date_received, '%Y-%m')
ORDER BY complaint_month;
