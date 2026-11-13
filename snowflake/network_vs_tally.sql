-- =====================================================================================
-- Network charges vs Tally customer bills, per NMI  (network charges only for now)
-- Needs: the NETWORK_LINES view (network_lines.sql) and the Tally extract table TALLY_MONTHLY_DETAILED_BILLING.
-- Same method as the old tool: every line is spread evenly per day over its period (average per day),
-- the days are added up per NMI per month on both sides, then compared. Difference = Tally - network.
-- Run USE DATABASE ... / USE SCHEMA ... first. Create the 4 views once; then just query them.
-- =====================================================================================

-- ---------- A) Network side: per NMI, per day, per charge group ----------
CREATE OR REPLACE VIEW NETWORK_DAILY AS
WITH lines AS (
    SELECT
        UPPER(TRIM(NMI)) AS NMI,
        CHARGE_GROUP,
        -- lines without a usable period go on the invoice date, full amount (same as the old tool)
        CASE WHEN INVOICE_START_DATE IS NOT NULL AND INVOICE_END_DATE >= INVOICE_START_DATE
             THEN INVOICE_START_DATE ELSE INVOICE_DATE END AS FROM_DATE,
        CASE WHEN INVOICE_START_DATE IS NOT NULL AND INVOICE_END_DATE >= INVOICE_START_DATE
             THEN INVOICE_END_DATE ELSE INVOICE_DATE END AS TO_DATE,
        COALESCE(AVERAGE_PERDAY, TOTAL_CHARGE_EXC_GST) AS PER_DAY
    FROM NETWORK_LINES
)
SELECT l.NMI, DATEADD(day, d.value::INT, l.FROM_DATE) AS CHARGE_DAY, l.CHARGE_GROUP, SUM(l.PER_DAY) AS NETWORK_PER_DAY
FROM lines l,
     LATERAL FLATTEN(INPUT => ARRAY_GENERATE_RANGE(0, DATEDIFF(day, l.FROM_DATE, l.TO_DATE) + 1)) d
GROUP BY 1, 2, 3;

-- ---------- B) Tally side: per NMI, per day, per charge group ----------
CREATE OR REPLACE VIEW TALLY_DAILY AS
WITH
t AS (                   -- *** column names of your Tally table: change ONLY here if yours differ ***
    SELECT
        UPPER(TRIM(NMI))                     AS NMI,
        INVOICENUMBER                        AS INVOICE_NUMBER,
        LOWER(TRIM(CHARGE_CATEGORIES))       AS CATEGORY,
        TRIM(COALESCE(BILLTEXT, ''))         AS BILL_TEXT,
        TRIM(COALESCE(PRODUCTITEMID, ''))    AS PRODUCT,
        TOTAL::FLOAT                         AS TOTAL_EX,          -- ex GST
        PERIODFROMDATE                       AS FROM_TXT,          -- text, e.g. 1/03/2025
        PERIODTODATE                         AS TO_TXT,
        LOWER(TRIM(COALESCE(INVOICESTATUS, 'Authorised')))  AS STATUS,
        LOWER(TRIM(COALESCE(SERVICE_TYPE, 'Electricity')))  AS SERVICE
    FROM TALLY_MONTHLY_DETAILED_BILLING
),
ok AS (                  -- authorised electricity bills only; dates from text
    SELECT *,
        COALESCE(TRY_TO_DATE(FROM_TXT, 'DD/MM/YYYY'), TRY_TO_DATE(FROM_TXT, 'YYYY-MM-DD')) AS FROM0,
        COALESCE(TRY_TO_DATE(TO_TXT, 'DD/MM/YYYY'), TRY_TO_DATE(TO_TXT, 'YYYY-MM-DD'))     AS TO0,
        -- revisions say which period they really belong to: "... for period 01-11-2025 to 30-11-2025"
        TRY_TO_DATE(REPLACE(REGEXP_SUBSTR(BILL_TEXT,
            'period[[:space:]]+([0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{4})[[:space:]]+to[[:space:]]+([0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{4})',
            1, 1, 'ie', 1), '-', '/'), 'DD/MM/YYYY') AS REV_FROM,
        TRY_TO_DATE(REPLACE(REGEXP_SUBSTR(BILL_TEXT,
            'period[[:space:]]+([0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{4})[[:space:]]+to[[:space:]]+([0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{4})',
            1, 1, 'ie', 2), '-', '/'), 'DD/MM/YYYY') AS REV_TO
    FROM t
    WHERE STATUS = 'authorised' AND SERVICE IN ('electricity', '')
),
net_products AS (        -- product codes used on "Network Charges" lines: an adjustment to one is a network charge too
    SELECT PRODUCT, MAX(CASE WHEN BILL_TEXT ILIKE '%meter%' THEN 0 ELSE 1 END) AS IS_NETWORK_CHARGE
    FROM t
    WHERE CATEGORY = 'network charges' AND PRODUCT <> '' AND UPPER(PRODUCT) NOT LIKE 'GLOBAL.%'
    GROUP BY 1
),
grouped AS (
    SELECT o.*,
        CASE
            WHEN o.CATEGORY = 'network charges'
                THEN CASE WHEN o.BILL_TEXT ILIKE '%meter%' THEN 'Metering' ELSE 'Network charges' END
            WHEN REGEXP_LIKE(o.BILL_TEXT, 'network[[:space:]]*-.*', 'is')            -- "Network - ..." fee lines
                THEN CASE WHEN o.BILL_TEXT ILIKE '%meter%' THEN 'Metering'
                          WHEN REGEXP_LIKE(o.BILL_TEXT, '.*(gsl|guaranteed service).*', 'is') THEN 'GSL'
                          ELSE 'Fees' END
            WHEN p.PRODUCT IS NOT NULL                                                -- adjustment to a network product
                THEN CASE WHEN p.IS_NETWORK_CHARGE = 1 THEN 'Network charges' ELSE 'Metering' END
        END AS CHARGE_GROUP
    FROM ok o LEFT JOIN net_products p ON p.PRODUCT = o.PRODUCT
),
periods AS (
    SELECT NMI, CHARGE_GROUP, TOTAL_EX,
        CASE WHEN REV_FROM IS NOT NULL AND REV_TO >= REV_FROM THEN REV_FROM ELSE FROM0 END AS FROM_DATE,
        CASE WHEN REV_FROM IS NOT NULL AND REV_TO >= REV_FROM THEN REV_TO
             WHEN TO0 >= FROM0 THEN TO0 ELSE FROM0 END AS TO_DATE
    FROM grouped
    WHERE CHARGE_GROUP IS NOT NULL AND FROM0 IS NOT NULL      -- not a network pass-through / no start date: left out
)
SELECT p.NMI, DATEADD(day, d.value::INT, p.FROM_DATE) AS CHARGE_DAY, p.CHARGE_GROUP,
       SUM(p.TOTAL_EX / (DATEDIFF(day, p.FROM_DATE, p.TO_DATE) + 1)) AS TALLY_PER_DAY
FROM periods p,
     LATERAL FLATTEN(INPUT => ARRAY_GENERATE_RANGE(0, DATEDIFF(day, p.FROM_DATE, p.TO_DATE) + 1)) d
GROUP BY 1, 2, 3;

-- ---------- C) Network charges: per NMI per month ----------
-- COUNTED = month is in the compared period: from the first to the last month (Jan 2025 onwards)
-- where both sides have bills (network at least half of Tally), same rule as the old tool.
CREATE OR REPLACE VIEW NETWORK_VS_TALLY_MONTH AS
WITH
net AS (
    SELECT NMI, DATE_TRUNC(month, CHARGE_DAY) AS MONTH, SUM(NETWORK_PER_DAY) AS NETWORK
    FROM NETWORK_DAILY WHERE CHARGE_GROUP = 'Network charges' GROUP BY 1, 2
),
tal AS (
    SELECT NMI, DATE_TRUNC(month, CHARGE_DAY) AS MONTH, SUM(TALLY_PER_DAY) AS BILLED
    FROM TALLY_DAILY WHERE CHARGE_GROUP = 'Network charges' GROUP BY 1, 2
),
both_sides AS (
    SELECT COALESCE(n.NMI, t.NMI) AS NMI, COALESCE(n.MONTH, t.MONTH) AS MONTH,
           COALESCE(n.NETWORK, 0) AS NETWORK, COALESCE(t.BILLED, 0) AS BILLED
    FROM net n FULL OUTER JOIN tal t ON t.NMI = n.NMI AND t.MONTH = n.MONTH
),
month_totals AS (
    SELECT MONTH, SUM(NETWORK) AS NETWORK, SUM(BILLED) AS BILLED FROM both_sides GROUP BY 1
),
counted AS (
    SELECT MIN(MONTH) AS COUNTED_FROM, MAX(MONTH) AS COUNTED_TO FROM month_totals
    WHERE MONTH >= '2025-01-01' AND BILLED > 0 AND NETWORK >= 0.5 * BILLED
)
SELECT b.NMI, b.MONTH,
       ROUND(b.NETWORK, 2)            AS NETWORK,
       ROUND(b.BILLED, 2)             AS BILLED,
       ROUND(b.BILLED - b.NETWORK, 2) AS DIFFERENCE,
       (b.MONTH BETWEEN c.COUNTED_FROM AND c.COUNTED_TO) AS COUNTED,
       (ABS(b.BILLED - b.NETWORK) > 100)                AS OVER_100
FROM both_sides b CROSS JOIN counted c
WHERE b.MONTH >= '2025-01-01';

-- ---------- D) Network charges: one row per NMI (counted months only) ----------
CREATE OR REPLACE VIEW NETWORK_VS_TALLY_NMI AS
SELECT NMI,
       ROUND(SUM(NETWORK), 2)                AS NETWORK,
       ROUND(SUM(BILLED), 2)                 AS BILLED,
       ROUND(SUM(BILLED) - SUM(NETWORK), 2)  AS DIFFERENCE,
       COUNT_IF(OVER_100)                    AS MONTHS_OVER_100,
       CASE WHEN SUM(BILLED) - SUM(NETWORK) < -100 THEN 'Under-billed'
            WHEN SUM(BILLED) - SUM(NETWORK) >  100 THEN 'Over-billed'
            ELSE 'OK' END                    AS RESULT,
       MIN(MONTH) AS FIRST_MONTH, MAX(MONTH) AS LAST_MONTH
FROM NETWORK_VS_TALLY_MONTH
WHERE COUNTED
GROUP BY NMI;

-- ---------- Use it ----------
SELECT * FROM NETWORK_VS_TALLY_NMI ORDER BY ABS(DIFFERENCE) DESC;                 -- biggest differences first
SELECT * FROM NETWORK_VS_TALLY_MONTH WHERE NMI = '<NMI>' ORDER BY MONTH;           -- drill into one NMI

-- Totals: under-billed and over-billed shown separately (never just the net)
SELECT RESULT, COUNT(*) AS NMIS, ROUND(SUM(DIFFERENCE), 2) AS DIFFERENCE
FROM NETWORK_VS_TALLY_NMI GROUP BY 1 ORDER BY 1;

-- Months compared and per-month totals
SELECT MONTH, COUNTED, ROUND(SUM(NETWORK), 2) AS NETWORK, ROUND(SUM(BILLED), 2) AS BILLED,
       ROUND(SUM(DIFFERENCE), 2) AS DIFFERENCE
FROM NETWORK_VS_TALLY_MONTH GROUP BY 1, 2 ORDER BY 1;
