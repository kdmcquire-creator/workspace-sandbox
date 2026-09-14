"""Collapse the Enverus well export (one row per completion) to one row per API."""

import pandas as pd
from wellmatch import norm_county, norm_name, canon_name

ENV_COLS = [
    "Unformatted_API_UWI", "WellName", "WellNumber", "LeaseName", "County",
    "StateProvince", "ENVOperator", "RawOperator", "InitialOperator", "Field",
    "ENVInterval", "Formation", "ENVWellStatus", "Trajectory", "SpudDate",
    "CompletionDate", "FirstProdDate", "LateralLength_FT", "TVD_FT", "MD_FT",
    "Latitude", "Longitude", "Latitude_BH", "Longitude_BH",
    "Section_Township_Range",
]

# Fields where completion records disagree; keep the earliest/most original.
FIRST_NONNULL = [
    "WellName", "WellNumber", "LeaseName", "County", "StateProvince",
    "ENVOperator", "RawOperator", "InitialOperator", "Field", "ENVInterval",
    "Formation", "ENVWellStatus", "Trajectory", "LateralLength_FT", "TVD_FT",
    "MD_FT", "Latitude", "Longitude", "Latitude_BH", "Longitude_BH",
    "Section_Township_Range",
]


def load_enverus(path):
    raw = pd.read_csv(path, usecols=ENV_COLS, dtype=str, low_memory=False)

    agg = {c: "first" for c in FIRST_NONNULL}
    agg["SpudDate"] = "min"
    agg["FirstProdDate"] = "min"
    agg["CompletionDate"] = "min"

    wells = (
        raw.sort_values(["Unformatted_API_UWI", "CompletionDate"])
           .groupby("Unformatted_API_UWI", as_index=False)
           .agg(agg)
    )
    wells["county_n"] = wells["County"].map(norm_county)
    wells["name_n"] = wells["WellName"].map(norm_name)
    wells["name_c"] = wells["WellName"].map(canon_name)
    wells["lease_n"] = wells["LeaseName"].map(norm_name)
    wells["wellno_n"] = wells["WellNumber"].map(norm_name)

    return wells
