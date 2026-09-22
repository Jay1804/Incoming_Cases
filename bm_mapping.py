import pandas as pd

MAPPING_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vRYUU3j9ZL7Te2Ij9KYNE5fwFG0mHKxXajocPkoytErZ3IL1zv7zhCHmJjUo-8_9_aIN8FdGUQA1oDj"
    "/pub?gid=0&single=true&output=csv"
)
SHEET_COLUMN_RENAME = {
    "Client Code": "Client Id",
    "BM": "BM Team Member",
    "BM Lead": "BM Team Lead",
}
MAPPING_COLUMNS = ["Client Id", "Client Name", "BM Team Member", "BM Team Lead"]


def load_bm_mapping(source=MAPPING_SHEET_URL):
    """Client -> BM Team Member/Lead mapping, sourced from a published Google
    Sheet rather than the DB, since the case-count queries don't carry BM info.
    'Client Code' in the sheet also carries non-numeric values (e.g. 'AUTH-10')
    that never match the DB's numeric client ids, so those rows are dropped."""
    df = pd.read_csv(source)
    df = df.rename(columns=SHEET_COLUMN_RENAME)[MAPPING_COLUMNS]
    df["Client Id"] = pd.to_numeric(df["Client Id"], errors="coerce")
    df = df.dropna(subset=["Client Id"])
    df["Client Id"] = df["Client Id"].astype(int)
    df = df.drop_duplicates(subset=["Client Id"])
    return df


def attach_bm_lead(df, id_col, mapping=None):
    """Left-join df (keyed on id_col, matching mapping's 'Client Id') with the
    BM mapping. Unmatched clients are labelled 'Unmapped' rather than dropped."""
    if mapping is None:
        mapping = load_bm_mapping()

    merged = df.merge(
        mapping, left_on=id_col, right_on="Client Id", how="left", suffixes=("", "_map")
    )
    merged["BM Team Lead"] = merged["BM Team Lead"].fillna("Unmapped")
    merged["BM Team Member"] = merged["BM Team Member"].fillna("Unmapped")
    return merged
