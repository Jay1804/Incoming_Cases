import datetime

import pandas as pd

from db import get_engine

INCOMING_CASE_COUNT_SQL = """
SELECT
    ec.client_external_id AS "Client_Id",
    emc.company_name AS "Client",
    ecm.received_date AS "received_date",
    COUNT(*) AS "Case Count"
FROM ec_client ec

LEFT JOIN ec_master_company emc
    ON ec.client_id = emc.company_id

LEFT JOIN ec_case_master ecm
    ON ec.client_id = ecm.client_id

LEFT JOIN ec_case_candidates ecc1
    ON ecc1.candidate_id = ecm.candidate_id

WHERE ec.client_id NOT IN (1000, 88031)
  AND ecm.case_status NOT IN (8, 14)
  AND ecm.received_date >= %(start_date)s
  AND ecm.received_date <= %(end_date)s
  AND ecc1.first_name <> 'dummy'
  AND ecc1.first_name <> 'test'

GROUP BY
    ec.client_external_id,
    emc.company_name,
    ecm.received_date

ORDER BY
    ecm.received_date,
    ec.client_external_id;
"""

SENT_CASE_SQL = """
SELECT
    ec.client_external_id AS "Client Code",
    company_name AS Client_name,
    DATE(report_sent_on) AS "sent_date",
    COUNT(case_ars_no) AS case_count
FROM ec_case_reports ecr
LEFT JOIN ec_case_master ecm
    ON ecr.case_id = ecm.case_id
LEFT JOIN ec_client ec
    ON ec.client_id = ecm.client_id
LEFT JOIN ec_master_company emc
    ON ecr.CLIENT_ID = emc.COMPANY_ID
WHERE report_sent_on >= %(start_date)s
  AND report_sent_on < %(end_date_exclusive)s
  AND report_type = 1
GROUP BY
    company_name,
    ec.client_external_id,
    DATE(report_sent_on);
"""


def _exclude_test_clients(df, name_col):
    """Drop rows whose client name looks like a test/dummy account (case-insensitive)."""
    is_test = df[name_col].astype(str).str.contains(r"test|dummy", case=False, na=False, regex=True)
    return df[~is_test].reset_index(drop=True)


def fetch_incoming_case_count(start_date, end_date):
    df = pd.read_sql(
        INCOMING_CASE_COUNT_SQL,
        get_engine(),
        params={"start_date": start_date, "end_date": end_date},
    )
    return _exclude_test_clients(df, "Client")


def fetch_sent_case(start_date, end_date):
    end_date_exclusive = end_date + datetime.timedelta(days=1)
    df = pd.read_sql(
        SENT_CASE_SQL,
        get_engine(),
        params={"start_date": start_date, "end_date_exclusive": end_date_exclusive},
    )
    return _exclude_test_clients(df, "Client_name")
