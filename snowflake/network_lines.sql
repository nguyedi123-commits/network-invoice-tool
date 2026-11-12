-- =====================================================================================
-- Network Lines in Snowflake
-- Load every "Network Lines.csv" from the app into NETWORK_LINES_RAW (just keep adding them).
-- The NETWORK_LINES view gives the same result as the old tool's "Network Lines" tab:
--   * the same Network Lines.csv loaded twice counts once
--   * a V2 file replaces its V1
--   * the same invoice in more than one file: only the latest file counts (no double counting)
--   * interest is left out (ours to pay, not passed to customers)
--   * CHARGE_GROUP = Network charges / Metering / Fees / GSL
--     (network charge lines with "meter" in the description count as Metering)
-- Run USE DATABASE ... / USE SCHEMA ... first.
-- =====================================================================================

-- ---------- 1) Table (run once) ----------
CREATE TABLE IF NOT EXISTS NETWORK_LINES_RAW (
  INVOICE_LINE          NUMBER,
  INVOICE_NUMBER        VARCHAR,
  INVOICE_DATE          DATE,
  INVOICE_TYPE          VARCHAR,
  NMI                   VARCHAR,
  STATE                 VARCHAR,
  INVOICE_START_DATE    DATE,
  INVOICE_END_DATE      DATE,
  TARIFF_CODE           VARCHAR,
  TARIFF_TYPE           VARCHAR,
  UNIT                  FLOAT,
  UNIT_MEASURE          VARCHAR,
  UNIT_RATE             FLOAT,
  TOTAL_CHARGE_EXC_GST  FLOAT,
  GST_AMOUNTS           FLOAT,
  REQUIRED_PAYMENT_DATE DATE,
  NETWORKS              VARCHAR,
  INVOICE_NOTE          VARCHAR,
  AVERAGE_PERDAY        FLOAT,
  NETWORK_TARIFF        VARCHAR,
  LINE_ID               VARCHAR,
  OLD_INVOICE           VARCHAR,
  CHARGE_GROUP          VARCHAR,
  FILE_NAME             VARCHAR,
  LINE_SEQ              NUMBER
);

-- ---------- 2) Load each Network Lines.csv ----------
-- Snowsight: Data > Add data > Load data into a table > NETWORK_LINES_RAW > pick the file >
--   header: "First line contains header", field optionally enclosed by: double quote, dates YYYY-MM-DD.
-- Or with SQL:
-- CREATE FILE FORMAT IF NOT EXISTS NETWORK_LINES_CSV TYPE = CSV PARSE_HEADER = FALSE SKIP_HEADER = 1
--   FIELD_OPTIONALLY_ENCLOSED_BY = '"' EMPTY_FIELD_AS_NULL = TRUE DATE_FORMAT = 'YYYY-MM-DD';
-- PUT 'file://C:/path/Network Lines.csv' @~/network_lines;
-- COPY INTO NETWORK_LINES_RAW FROM @~/network_lines FILE_FORMAT = (FORMAT_NAME = NETWORK_LINES_CSV) PURGE = TRUE;

-- ---------- 3) The view (run once; it always reads everything loaded) ----------
CREATE OR REPLACE VIEW NETWORK_LINES AS
WITH
one_load AS (            -- the same file loaded twice -> one copy
    SELECT * FROM NETWORK_LINES_RAW
    QUALIFY ROW_NUMBER() OVER (PARTITION BY FILE_NAME, LINE_SEQ ORDER BY FILE_NAME) = 1
),
files AS (
    SELECT DISTINCT FILE_NAME, SPLIT_PART(FILE_NAME, '#', 5) AS FILE_TS FROM one_load
),
latest_version AS (      -- ...#20250909010311V2 replaces ...#20250909010311V1
    SELECT FILE_NAME FROM files
    QUALIFY ROW_NUMBER() OVER (
        PARTITION BY LEFT(FILE_NAME, LENGTH(FILE_NAME) - LENGTH(FILE_TS) + 14)
        ORDER BY TRY_CAST(SUBSTR(FILE_TS, 16) AS INTEGER) DESC NULLS LAST) = 1
),
latest_invoice AS (      -- same invoice in more than one file -> latest file only
    SELECT l.* FROM one_load l JOIN latest_version v ON v.FILE_NAME = l.FILE_NAME
    QUALIFY RIGHT(l.FILE_NAME, 17) = MAX(RIGHT(l.FILE_NAME, 17)) OVER (PARTITION BY l.NETWORKS, l.INVOICE_NUMBER)
)
SELECT
    INVOICE_LINE, INVOICE_NUMBER, INVOICE_DATE, INVOICE_TYPE, NMI, STATE,
    INVOICE_START_DATE, INVOICE_END_DATE, TARIFF_CODE, TARIFF_TYPE, UNIT, UNIT_MEASURE, UNIT_RATE,
    TOTAL_CHARGE_EXC_GST, GST_AMOUNTS, REQUIRED_PAYMENT_DATE, NETWORKS, INVOICE_NOTE, AVERAGE_PERDAY,
    NETWORK_TARIFF, LINE_ID, OLD_INVOICE,
    CASE
        WHEN COALESCE(CHARGE_GROUP, 'Network charges') = 'Network charges'
             AND UPPER(COALESCE(TARIFF_TYPE, '')) LIKE '%METER%'               THEN 'Metering'
        WHEN COALESCE(CHARGE_GROUP, 'Network charges') IN ('Network charges', 'Metering', 'Fees', 'GSL')
                                                                              THEN COALESCE(CHARGE_GROUP, 'Network charges')
        ELSE 'Fees'
    END AS CHARGE_GROUP,
    FILE_NAME, LINE_SEQ
FROM latest_invoice
WHERE COALESCE(CHARGE_GROUP, '') <> 'Interest';

-- ---------- 4) Use it ----------
SELECT * FROM NETWORK_LINES ORDER BY UPPER(NMI), INVOICE_START_DATE NULLS FIRST, INVOICE_END_DATE NULLS FIRST;

-- Checks
SELECT COUNT(*) AS LINES, ROUND(SUM(TOTAL_CHARGE_EXC_GST), 2) AS TOTAL_EX_GST FROM NETWORK_LINES;
SELECT NETWORKS, CHARGE_GROUP, COUNT(*) AS LINES, ROUND(SUM(TOTAL_CHARGE_EXC_GST), 2) AS TOTAL_EX_GST
FROM NETWORK_LINES GROUP BY 1, 2 ORDER BY 1, 2;
-- Files loaded so far
SELECT FILE_NAME, NETWORKS, COUNT(*) AS LINES, ROUND(SUM(TOTAL_CHARGE_EXC_GST), 2) AS TOTAL_EX_GST
FROM NETWORK_LINES_RAW GROUP BY 1, 2 ORDER BY 2, 1;
