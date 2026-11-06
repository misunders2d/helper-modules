import os

from utils import mellanni_modules as mm
from connectors import gcloud as gc
from connectors import gdrive as gd
from common import user_folder, excluded_collections
import pandas as pd
import numpy as np
from concurrent.futures import ThreadPoolExecutor
from ctk_gui.ctk_windows import PopupError, filedialog

drive_id = "0AMdx9NlXacARUk9PVA"
MATRIX_FILE_ID = "1hnBbUMWQh4qgqowVIcgRll4zFvxAwc0Q"  # Master Product Library - https://docs.google.com/spreadsheets/d/1hnBbUMWQh4qgqowVIcgRll4zFvxAwc0Q/edit?usp=sharing&ouid=104139340250971501691&rtpof=true&sd=true

# Product size tiers page: https://sellercentral.amazon.com/help/hub/reference/GG5KW835AHDJCH8W
# Storage fee page: https://sellercentral.amazon.com/help/hub/reference/G3EDYEF6KUCFQTNM
# FBA fee page: https://sellercentral.amazon.com/help/hub/reference/GABBX6GZPA8MSZGW

SMALL_STANDARD = "Small standard size"
LARGE_STANDARD = "Large standard size"
SMALL_BULKY = "Small bulky"
LARGE_BULKY = "Large bulky"
EXTRA_LARGE_0_50 = "Extra large up to 50"
EXTRA_LARGE_50_70 = "Extra large 50 to 70"
EXTRA_LARGE_70_150 = "Extra large 70 to 150"
EXTRA_LARGE_150_PLUS = "Extra large 150+"
num_cols = [
    "l",
    "w",
    "h",
    "individual weight lbs",
    "sets in a box",
    "box length",
    "box width",
    "box depth",
    "box weight lbs",
    "dim_weight",
    "shipping_weight",
    "shipping_weight, oz",
    "fba_fee",
    "removal_fee",
    "liquidation_fee",
    "current_storage_fee",
    "storage_jan_sept",
    "storage_oct_dec",
    "avg_yearly_storage",
]
today = pd.to_datetime("today")
# today = pd.to_datetime("2026-04-30")


def get_prices_file(spreadsheet_id="1iB1CmY_XdOVA4FvLMPeiEGEcxiVEH3Bgp4FJs1iNmQs"):
    file = gd.download_gspread(spreadsheet_id=spreadsheet_id)
    file = file[["SKU", "Full price", "Sale price", "Status"]]
    file.columns = [x.lower() for x in file.columns]
    file["price"] = file["full price"].str.replace("$", "")
    file.loc[~file["status"].isin(["OK", "Stopped"]), "price"] = file[
        "sale price"
    ].str.replace("$", "")
    file["price"] = pd.to_numeric(file["price"], errors="coerce")
    file = file.drop_duplicates("sku")
    return file[["sku", "price"]]


def get_dims_file(
    folder_id="1zIHmbWcRRVyCTtuB9Atzam7IhAs8Ymx4", filename="DIMENSIONS.xlsx"
):
    file = pd.read_excel(
        gd.download_file(gd.find_file_id(folder_id, drive_id, filename)), skiprows=1
    )
    del file["Product Name"]
    file.columns = [x.lower() for x in file.columns]
    file = file[~file["collection"].isin(excluded_collections)]
    return file


# TODO Change to this function when Dictionary is stored in Google Sheets exclusively
# def get_dictionary(spreadsheet_id='1tezZ1Txml4E1YGYnO8-57lSnxlzsJ2boUYb4xmVxmtw', sheet_id='1749064367'):
#     file = gd.download_gspread(spreadsheet_id=spreadsheet_id, sheet_id=sheet_id)
#     file = file[['SKU','ASIN', 'Collection','Sub-collection', 'Size Map','Color','Actuality']]
#     file.columns = [x.lower() for x in file.columns]
#     file = file[~file['collection'].isin(excluded_collections)]
#     file = file.drop_duplicates('sku')
#     return file


def get_dictionary(
    folder_id="1zIHmbWcRRVyCTtuB9Atzam7IhAs8Ymx4", filename="Dictionary.xlsx"
):
    # file = pd.read_excel(
    #     gd.download_file(gd.find_file_id(folder_id, drive_id, filename))
    # )
    file = gd.download_gspread(spreadsheet_id="1Y4XhSBCXqmEVHHOnugEpzZZ3NQ5ZRGOlp-AsTE0KmRE", sheet_id="449289593")
    file = file[
        [
            "SKU",
            "ASIN",
            "Collection",
            "Sub-collection",
            "Size Map",
            "Color",
            "Actuality",
        ]
    ]
    file.columns = [x.lower() for x in file.columns]
    file = file[~file["collection"].isin(excluded_collections)]
    file = file.drop_duplicates("sku")
    file["country_code"] = "US"
    file["sales_channel"] = "Amazon.com"
    return file


def combine_files(dimensions, dictionary, prices):
    collections = dimensions["collection"].unique()
    dictionary["c"] = dictionary["sub-collection"].copy()
    dictionary.loc[~dictionary["sub-collection"].isin(collections), "c"] = dictionary[
        "collection"
    ]
    for c in ["collection", "sub-collection"]:
        del dictionary[c]
    dictionary = dictionary.rename(columns={"size map": "size", "c": "collection"})
    dictionary = dictionary[
        ["sku", "asin", "collection", "size", "color", "actuality", "sales_channel"]
    ]
    result = pd.merge(dictionary, dimensions, how="right", on=["collection", "size"])
    result = result.drop_duplicates("sku")
    result = pd.merge(result, prices, how="left", on="sku", validate="1:1")
    return result


def get_shipping_weight(df):
    df["dim_weight"] = (df["l"] * df["w"] * df["h"]) / 139
    df["shipping_weight"] = df[["dim_weight", "individual weight lbs"]].max(axis=1)
    return df


def get_size_tier(df):
    max_side = df[["l", "w", "h"]].max(axis=1)
    min_side = df[["l", "w", "h"]].min(axis=1)
    med_side = df[["l", "w", "h"]].median(axis=1)
    length_girth = (min_side + med_side) * 2
    weight = df["shipping_weight"].copy()

    conditions = [
        (
            SMALL_STANDARD,
            (weight <= 1) & (max_side <= 15) & (med_side <= 12) & (min_side <= 0.75),
        ),
        (
            LARGE_STANDARD,
            (weight <= 20) & (max_side <= 18) & (med_side <= 14) & (min_side <= 8),
        ),
        (
            SMALL_BULKY,
            (weight <= 50)
            & (max_side <= 37)
            & (med_side <= 28)
            & (min_side <= 20)
            & (length_girth <= 130),
        ),
        (
            LARGE_BULKY,
            (weight <= 50)
            & (max_side <= 59)
            & (med_side <= 33)
            & (min_side <= 33)
            & (length_girth <= 130),
        ),
        (
            EXTRA_LARGE_0_50,
            (weight <= 50)
            | (max_side > 59)
            | (med_side > 33)
            | (min_side > 33)
            | (length_girth > 130),
        ),
        (
            EXTRA_LARGE_50_70,
            50
            < (weight <= 70)
            | (max_side > 59)
            | (med_side > 33)
            | (min_side > 33)
            | (length_girth > 130),
        ),
        (
            EXTRA_LARGE_70_150,
            70
            < (weight <= 150)
            | (max_side > 59)
            | (med_side > 33)
            | (min_side > 33)
            | (length_girth > 130),
        ),
        (
            EXTRA_LARGE_150_PLUS,
            (weight > 150)
            | (max_side > 59)
            | (med_side > 33)
            | (min_side > 33)
            | (length_girth > 130),
        ),
    ]

    df["size_tier"] = np.select(
        [x[1] for x in conditions], [x[0] for x in conditions], "Unidentified"
    )
    return df


def get_storage_fee(df):
    this_month = today.month
    size_tier = df["size_tier"]
    volume = (df["l"] * df["w"] * df["h"]) / 1728
    fees_jan_sept = [
        (0.78 * volume, size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])),
        (0.56 * volume, ~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])),
    ]
    fees_oct_dec = [
        (2.40 * volume, size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])),
        (1.40 * volume, ~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])),
    ]

    fees_jan_sept_181_270 = [
        (1.56 * volume, size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])),
        (1.02 * volume, ~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])),
    ]

    fees_jan_sept_271_365 = [
        (1.81 * volume, size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])),
        (1.19 * volume, ~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])),
    ]

    fees_oct_dec_181_270 = [
        (3.09 * volume, size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])),
        (1.86 * volume, ~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])),
    ]

    fees_oct_dec_271_365 = [
        (3.34 * volume, size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])),
        (2.03 * volume, ~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])),
    ]

    if this_month in range(1, 10):
        storage = fees_jan_sept
        storage_181_270 = fees_jan_sept_181_270
        storage_271_365 = fees_jan_sept_271_365

    else:  # this_month in range(10,13):
        storage = fees_oct_dec
        storage_181_270 = fees_oct_dec_181_270
        storage_271_365 = fees_oct_dec_271_365

    df["current_storage_fee"] = np.select(
        [x[1] for x in storage], [x[0] for x in storage], np.nan
    ).astype("float")
    df["storage_jan_sept"] = np.select(
        [x[1] for x in fees_jan_sept], [x[0] for x in fees_jan_sept], np.nan
    ).astype("float")
    df["storage_oct_dec"] = np.select(
        [x[1] for x in fees_oct_dec], [x[0] for x in fees_oct_dec], np.nan
    ).astype("float")
    df["avg_yearly_storage"] = (
        (df["storage_jan_sept"] * 9) + (df["storage_oct_dec"] * 3)
    ) / 12

    df["current_storage_181-270"] = np.select(
        [x[1] for x in storage_181_270], [x[0] for x in storage_181_270], np.nan
    ).astype("float")
    df["current_storage_271-365"] = np.select(
        [x[1] for x in storage_271_365], [x[0] for x in storage_271_365], np.nan
    ).astype("float")

    df["jan_sept_storage_181-270"] = np.select(
        [x[1] for x in fees_jan_sept_181_270],
        [x[0] for x in fees_jan_sept_181_270],
        np.nan,
    ).astype("float")
    df["jan_sept_storage_271-365"] = np.select(
        [x[1] for x in fees_jan_sept_271_365],
        [x[0] for x in fees_jan_sept_271_365],
        np.nan,
    ).astype("float")

    df["oct_dec_storage_181-270"] = np.select(
        [x[1] for x in fees_oct_dec_181_270],
        [x[0] for x in fees_oct_dec_181_270],
        np.nan,
    ).astype("float")
    df["oct_dec_storage_271-365"] = np.select(
        [x[1] for x in fees_oct_dec_271_365],
        [x[0] for x in fees_oct_dec_271_365],
        np.nan,
    ).astype("float")
    return df


def get_fulfillment_fee(df, as_of=None, schedule=None):
    """Calculate non-apparel base fulfillment fees, excluding surcharge and SIPP.

    Default calls use published standard-size rates from January 15, 2026.
    Other 2026 size tiers retain legacy calculations, labeled in output.
    ``schedule="legacy"`` always uses original tables/formulas at ``as_of``.
    Explicit published standard-size schedules: 2026_non_peak, 2026_peak,
    2027_non_peak. Source: 2027 US FBA fee PDF, pp. 1-3, 12-13.
    Published schedules round excess weight UP per 4 oz interval, unlike legacy.
    New bulky/oversize classification requires the separate Product size tiers
    reference, absent from that PDF; explicit schedules reject those rows.
    2027 peak is unpublished and must never inherit the 2026 peak table.
    """
    fee_date = today if as_of is None else pd.Timestamp(as_of)
    if pd.isna(fee_date):
        raise ValueError("as_of must be a valid date")
    if schedule is None:
        if fee_date >= pd.Timestamp("2027-10-15"):
            raise ValueError("No published fulfillment schedule for this date")
        if pd.Timestamp("2026-01-15") <= fee_date < pd.Timestamp("2027-01-15"):
            # Correct standard-size active rates without changing unrelated tiers.
            result = get_fulfillment_fee(df, as_of=fee_date, schedule="legacy")
            standard = result["size_tier"].isin([SMALL_STANDARD, LARGE_STANDARD])
            result["fulfillment_schedule"] = "legacy (non-standard size)"
            selected = "2026_peak" if (fee_date.month, fee_date.day) >= (10, 15) else "2026_non_peak"
            if fee_date.month == 1 and fee_date.day <= 14:
                selected = "2026_peak"
            if standard.any():
                published = get_fulfillment_fee(result.loc[standard].copy(), schedule=selected)
                for column in published.columns:
                    if column not in result:
                        result[column] = np.nan
                    result.loc[standard, column] = published[column]
                result.loc[standard, "fulfillment_schedule"] = selected
            return result
        schedule = "2027_non_peak" if fee_date >= pd.Timestamp("2027-01-15") else "legacy"
    if schedule not in {"legacy", "2026_non_peak", "2026_peak", "2027_non_peak"}:
        raise ValueError("Unknown or unpublished fulfillment schedule: " + str(schedule))
    if schedule != "legacy":
        # Exact non-apparel rows, price columns <$10 / $10-$50 / >$50.
        small = {
            "2026_non_peak": [
                (2.43, 3.32, 3.58),
                (2.49, 3.42, 3.68),
                (2.56, 3.45, 3.71),
                (2.66, 3.54, 3.80),
                (2.77, 3.68, 3.94),
                (2.82, 3.78, 4.04),
                (2.92, 3.91, 4.17),
                (2.95, 3.96, 4.22),
            ],
            "2026_peak": [
                (2.62, 3.51, 3.77),
                (2.68, 3.61, 3.87),
                (2.76, 3.65, 3.91),
                (2.86, 3.74, 4.00),
                (2.98, 3.89, 4.15),
                (3.03, 3.99, 4.25),
                (3.14, 4.13, 4.39),
                (3.17, 4.18, 4.44),
            ],
            "2027_non_peak": [
                (2.64, 3.75, 4.31),
                (2.72, 3.78, 4.34),
                (2.80, 3.92, 4.47),
                (2.83, 3.97, 4.53),
                (2.87, 4.06, 4.62),
                (2.92, 4.16, 4.73),
                (2.96, 4.30, 4.85),
                (2.99, 4.32, 4.88),
            ],
        }
        large = {
            "2026_non_peak": [
                (2.91, 3.73, 3.99),
                (3.13, 3.95, 4.21),
                (3.38, 4.20, 4.46),
                (3.78, 4.60, 4.86),
                (4.22, 5.04, 5.30),
                (4.60, 5.42, 5.68),
                (4.75, 5.57, 5.83),
                (5.00, 5.82, 6.08),
                (5.10, 5.92, 6.18),
                (5.28, 6.10, 6.36),
                (5.44, 6.26, 6.52),
                (5.85, 6.67, 6.93),
                (6.15, 6.97, 7.23),
            ],
            "2026_peak": [
                (3.15, 3.97, 4.23),
                (3.39, 4.21, 4.47),
                (3.66, 4.48, 4.74),
                (4.07, 4.89, 5.15),
                (4.52, 5.34, 5.60),
                (4.91, 5.73, 5.99),
                (5.07, 5.89, 6.15),
                (5.33, 6.15, 6.41),
                (5.47, 6.29, 6.55),
                (5.67, 6.49, 6.75),
                (5.84, 6.66, 6.92),
                (6.26, 7.08, 7.34),
                (6.69, 7.51, 7.77),
            ],
            "2027_non_peak": [
                (2.91, 3.79, 4.35),
                (3.13, 4.01, 4.57),
                (3.38, 4.26, 4.83),
                (3.78, 4.67, 5.23),
                (4.22, 5.19, 5.75),
                (4.60, 5.57, 6.14),
                (4.75, 5.72, 6.29),
                (5.00, 5.98, 6.54),
                (5.10, 6.08, 6.64),
                (5.28, 6.26, 6.83),
                (5.44, 6.42, 6.99),
                (5.85, 6.85, 7.41),
                (6.15, 7.16, 7.73),
            ],
        }
        if not df["size_tier"].isin([SMALL_STANDARD, LARGE_STANDARD]).all():
            raise ValueError("Published schedules support standard-size only; new bulky/oversize tier definitions required")
        df.loc[df["size_tier"] == SMALL_STANDARD, "shipping_weight"] = df["individual weight lbs"]
        if not (np.isfinite(df[["shipping_weight", "price"]]).all().all()
                and df["shipping_weight"].gt(0).all() and df["price"].ge(0).all()):
            raise ValueError("Finite positive shipping weight and nonnegative price required")
        df["shipping_weight, oz"] = df["shipping_weight"] * 16
        price_band = np.select([df["price"] < 10, df["price"] <= 50], [0, 1], 2)
        values = {}
        for name in small:
            conditions, amounts = [], []
            for tier, bounds, rows in [
                (SMALL_STANDARD, np.arange(2, 17, 2) / 16, small[name]),
                (LARGE_STANDARD, [0.25, 0.5, 0.75, 1, 1.25, 1.5, 1.75, 2, 2.25, 2.5, 2.75, 3, 20], large[name]),
            ]:
                lower = 0
                for upper, rates in zip(bounds, rows):
                    conditions.append((df["size_tier"] == tier) & df["shipping_weight"].between(lower, upper, inclusive="right"))
                    amount = np.take(rates, price_band)
                    if tier == LARGE_STANDARD and upper == 20:
                        amount = amount + 0.08 * np.ceil((df["shipping_weight"] - 3) * 4)
                    amounts.append(amount)
                    lower = upper
            if not (np.sum(conditions, axis=0) == 1).all():
                raise ValueError("Each product must match exactly one published weight tier")
            values[name] = np.round(np.select(conditions, amounts, np.nan), 2)
            df["fba_" + name + "_fee"] = values[name]
        df["fba_2027_peak_fee"] = np.nan  # Not published in supplied PDF.
        year = schedule[:4]
        df["fba_non_peak_fee"] = values[year + "_non_peak"]
        df["fba_peak_fee"] = values["2026_peak"] if year == "2026" else np.nan
        df["fba_fee"] = values[schedule]
        return df
    today_date = fee_date.date()
    df.loc[
        df["size_tier"].isin([SMALL_STANDARD, EXTRA_LARGE_150_PLUS]), "shipping_weight"
    ] = df["individual weight lbs"]
    df["shipping_weight, oz"] = df["shipping_weight"] * 16
    size_tier = df["size_tier"]
    weight = df["shipping_weight"]
    weight_oz = df["shipping_weight, oz"]
    price = df["price"]

    price_tier1 = price < 10
    price_tier2 = price <= 50
    price_tier3 = price > 50

    fees_before_feb2024 = [  # obsolete
        (3.22, (size_tier == SMALL_STANDARD) & (weight_oz <= 4)),
        (
            3.40,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(4, 8, inclusive="right")),
        ),
        (
            3.58,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(8, 12, inclusive="right")),
        ),
        (
            3.77,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(12, 16, inclusive="right")),
        ),
        (3.86, (size_tier == LARGE_STANDARD) & (weight_oz <= 4)),
        (
            4.08,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(4, 8, inclusive="right")),
        ),
        (
            4.24,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(8, 12, inclusive="right")),
        ),
        (
            4.75,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(12, 16, inclusive="right")),
        ),
        (
            5.40,
            (size_tier == LARGE_STANDARD) & (weight.between(1, 1.5, inclusive="right")),
        ),
        (
            5.69,
            (size_tier == LARGE_STANDARD) & (weight.between(1.5, 2, inclusive="right")),
        ),
        (
            6.10,
            (size_tier == LARGE_STANDARD) & (weight.between(2, 2.5, inclusive="right")),
        ),
        (
            6.39,
            (size_tier == LARGE_STANDARD) & (weight.between(2.5, 3, inclusive="right")),
        ),
        (
            7.17 + 0.16 * ((weight - 3) * 2),
            (size_tier == LARGE_STANDARD) & (weight.between(3, 20, inclusive="right")),
        ),
        (9.73 + 0.42 * (weight - 1), (size_tier == "Small oversize") & (weight <= 70)),
        (
            19.05 + 0.42 * (weight - 1),
            (size_tier == "Medium oversize") & (weight <= 150),
        ),
        (
            89.98 + 0.83 * (weight - 90),
            (size_tier == "Large oversize") & (weight <= 150),
        ),
        (
            158.49 + 0.83 * (weight - 90),
            (size_tier == "Special oversize") & (weight > 150),
        ),
    ]
    fees_feb_apr_2024 = [  # obsolete
        (3.22, (size_tier == SMALL_STANDARD) & (weight_oz <= 2)),
        (
            3.31,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(2, 4, inclusive="right")),
        ),
        (
            3.40,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(4, 6, inclusive="right")),
        ),
        (
            3.49,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(6, 8, inclusive="right")),
        ),
        (
            3.58,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(8, 10, inclusive="right")),
        ),
        (
            3.68,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(10, 12, inclusive="right")),
        ),
        (
            3.77,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(12, 14, inclusive="right")),
        ),
        (
            3.82,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(14, 16, inclusive="right")),
        ),
        (3.86, (size_tier == LARGE_STANDARD) & (weight_oz <= 4)),
        (
            4.08,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(4, 8, inclusive="right")),
        ),
        (
            4.32,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(8, 12, inclusive="right")),
        ),
        (
            4.75,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(12, 16, inclusive="right")),
        ),
        (
            5.19,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1, 1.25, inclusive="right")),
        ),
        (
            5.57,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.25, 1.5, inclusive="right")),
        ),
        (
            5.75,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.5, 1.75, inclusive="right")),
        ),
        (
            6.00,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.75, 2, inclusive="right")),
        ),
        (
            6.10,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2, 2.25, inclusive="right")),
        ),
        (
            6.28,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.25, 2.5, inclusive="right")),
        ),
        (
            6.45,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.5, 2.75, inclusive="right")),
        ),
        (
            6.86,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.75, 3, inclusive="right")),
        ),
        (
            7.25 + 0.08 * ((weight - 3) * 4),
            (size_tier == LARGE_STANDARD) & (weight.between(3, 20, inclusive="right")),
        ),
        (9.73 + 0.42 * (weight - 1), (size_tier == LARGE_BULKY) & (weight <= 50)),
        (26.33 + 0.38 * (weight - 1), (size_tier == EXTRA_LARGE_0_50) & (weight <= 50)),
        (
            40.12 + 0.75 * (weight - 51),
            (size_tier == EXTRA_LARGE_50_70)
            & (weight.between(50, 70, inclusive="right")),
        ),
        (
            54.81 + 0.75 * (weight - 71),
            (size_tier == EXTRA_LARGE_70_150)
            & (weight.between(70, 150, inclusive="right")),
        ),
        (
            194.95 + 0.19 * (weight - 151),
            (size_tier == EXTRA_LARGE_70_150) & (weight > 150),
        ),
    ]
    fees_fba_non_peak_before_2026 = [
        (3.06, (size_tier == SMALL_STANDARD) & (weight_oz <= 2)),
        (
            3.15,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(2, 4, inclusive="right")),
        ),
        (
            3.24,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(4, 6, inclusive="right")),
        ),
        (
            3.33,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(6, 8, inclusive="right")),
        ),
        (
            3.43,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(8, 10, inclusive="right")),
        ),
        (
            3.53,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(10, 12, inclusive="right")),
        ),
        (
            3.60,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(12, 14, inclusive="right")),
        ),
        (
            3.65,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(14, 16, inclusive="right")),
        ),
        (3.68, (size_tier == LARGE_STANDARD) & (weight_oz <= 4)),
        (
            3.90,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(4, 8, inclusive="right")),
        ),
        (
            4.15,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(8, 12, inclusive="right")),
        ),
        (
            4.55,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(12, 16, inclusive="right")),
        ),
        (
            4.99,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1, 1.25, inclusive="right")),
        ),
        (
            5.37,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.25, 1.5, inclusive="right")),
        ),
        (
            5.52,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.5, 1.75, inclusive="right")),
        ),
        (
            5.77,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.75, 2, inclusive="right")),
        ),
        (
            5.87,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2, 2.25, inclusive="right")),
        ),
        (
            6.05,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.25, 2.5, inclusive="right")),
        ),
        (
            6.21,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.5, 2.75, inclusive="right")),
        ),
        (
            6.62,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.75, 3, inclusive="right")),
        ),
        (
            6.92 + 0.08 * ((weight - 3) * 4),
            (size_tier == LARGE_STANDARD) & (weight.between(3, 20, inclusive="right")),
        ),
        (9.61 + 0.38 * (weight - 1), (size_tier == LARGE_BULKY) & (weight <= 50)),
        (26.33 + 0.38 * (weight - 1), (size_tier == EXTRA_LARGE_0_50) & (weight <= 50)),
        (
            40.12 + 0.75 * (weight - 51),
            (size_tier == EXTRA_LARGE_50_70)
            & (weight.between(50, 70, inclusive="right")),
        ),
        (
            54.81 + 0.75 * (weight - 71),
            (size_tier == EXTRA_LARGE_70_150)
            & (weight.between(70, 150, inclusive="right")),
        ),
        (
            194.95 + 0.19 * (weight - 151),
            (size_tier == EXTRA_LARGE_150_PLUS) & (weight > 150),
        ),
    ]

    fees_fba_peak_before_2026 = [
        (3.25, (size_tier == SMALL_STANDARD) & (weight_oz <= 2)),
        (
            3.34,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(2, 4, inclusive="right")),
        ),
        (
            3.44,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(4, 6, inclusive="right")),
        ),
        (
            3.53,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(6, 8, inclusive="right")),
        ),
        (
            3.64,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(8, 10, inclusive="right")),
        ),
        (
            3.74,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(10, 12, inclusive="right")),
        ),
        (
            3.82,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(12, 14, inclusive="right")),
        ),
        (
            3.87,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(14, 16, inclusive="right")),
        ),
        (3.92, (size_tier == LARGE_STANDARD) & (weight_oz <= 4)),
        (
            4.16,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(4, 8, inclusive="right")),
        ),
        (
            4.43,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(8, 12, inclusive="right")),
        ),
        (
            4.84,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(12, 16, inclusive="right")),
        ),
        (
            5.29,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1, 1.25, inclusive="right")),
        ),
        (
            5.68,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.25, 1.5, inclusive="right")),
        ),
        (
            5.84,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.5, 1.75, inclusive="right")),
        ),
        (
            6.10,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.75, 2, inclusive="right")),
        ),
        (
            6.24,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2, 2.25, inclusive="right")),
        ),
        (
            6.44,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.25, 2.5, inclusive="right")),
        ),
        (
            6.61,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.5, 2.75, inclusive="right")),
        ),
        (
            7.03,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.75, 3, inclusive="right")),
        ),
        (
            7.46 + 0.08 * ((weight - 3) * 4),
            (size_tier == LARGE_STANDARD) & (weight.between(3, 20, inclusive="right")),
        ),
        (10.65 + 0.38 * (weight - 1), (size_tier == LARGE_BULKY) & (weight <= 50)),
        (29.06 + 0.38 * (weight - 1), (size_tier == EXTRA_LARGE_0_50) & (weight <= 50)),
        (
            42.93 + 0.75 * (weight - 51),
            (size_tier == EXTRA_LARGE_50_70)
            & (weight.between(50, 70, inclusive="right")),
        ),
        (
            59.23 + 0.75 * (weight - 71),
            (size_tier == EXTRA_LARGE_70_150)
            & (weight.between(70, 150, inclusive="right")),
        ),
        (
            203.46 + 0.19 * (weight - 151),
            (size_tier == EXTRA_LARGE_150_PLUS) & (weight > 150),
        ),
    ]

    fees_fba_non_peak = [
        (2.43, (size_tier == SMALL_STANDARD) & (weight_oz <= 2) & price_tier1),
        (3.32, (size_tier == SMALL_STANDARD) & (weight_oz <= 2) & price_tier2),
        (3.58, (size_tier == SMALL_STANDARD) & (weight_oz <= 2) & price_tier3),
        (
            2.49,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(2, 4, inclusive="right"))
            & price_tier1,
        ),
        (
            3.42,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(2, 4, inclusive="right"))
            & price_tier2,
        ),
        (
            3.68,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(2, 4, inclusive="right"))
            & price_tier3,
        ),
        (
            2.56,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(4, 6, inclusive="right"))
            & price_tier1,
        ),
        (
            3.45,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(4, 6, inclusive="right"))
            & price_tier2,
        ),
        (
            3.71,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(4, 6, inclusive="right"))
            & price_tier3,
        ),
        (
            2.66,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(6, 8, inclusive="right"))
            & price_tier1,
        ),
        (
            3.54,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(6, 8, inclusive="right"))
            & price_tier2,
        ),
        (
            3.80,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(6, 8, inclusive="right"))
            & price_tier3,
        ),
        (
            2.77,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(8, 10, inclusive="right"))
            & price_tier1,
        ),
        (
            3.68,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(8, 10, inclusive="right"))
            & price_tier2,
        ),
        (
            3.94,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(8, 10, inclusive="right"))
            & price_tier3,
        ),
        (
            2.82,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(10, 12, inclusive="right"))
            & price_tier1,
        ),
        (
            3.78,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(10, 12, inclusive="right"))
            & price_tier2,
        ),
        (
            4.04,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(10, 12, inclusive="right"))
            & price_tier3,
        ),
        (
            2.92,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(12, 14, inclusive="right"))
            & price_tier1,
        ),
        (
            3.91,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(12, 14, inclusive="right"))
            & price_tier2,
        ),
        (
            4.17,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(12, 14, inclusive="right"))
            & price_tier3,
        ),
        (
            2.95,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(14, 16, inclusive="right"))
            & price_tier1,
        ),
        (
            3.96,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(14, 16, inclusive="right"))
            & price_tier2,
        ),
        (
            4.22,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(14, 16, inclusive="right"))
            & price_tier3,
        ),
        (2.91, (size_tier == LARGE_STANDARD) & (weight_oz <= 4) & price_tier1),
        (3.73, (size_tier == LARGE_STANDARD) & (weight_oz <= 4) & price_tier2),
        (3.99, (size_tier == LARGE_STANDARD) & (weight_oz <= 4) & price_tier3),
        (
            3.13,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(4, 8, inclusive="right"))
            & price_tier1,
        ),
        (
            3.95,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(4, 8, inclusive="right"))
            & price_tier2,
        ),
        (
            4.21,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(4, 8, inclusive="right"))
            & price_tier3,
        ),
        (
            3.38,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(8, 12, inclusive="right"))
            & price_tier1,
        ),
        (
            4.20,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(8, 12, inclusive="right"))
            & price_tier2,
        ),
        (
            4.46,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(8, 12, inclusive="right"))
            & price_tier3,
        ),
        (
            3.78,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(12, 16, inclusive="right"))
            & price_tier1,
        ),
        (
            4.60,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(12, 16, inclusive="right"))
            & price_tier2,
        ),
        (
            4.86,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(12, 16, inclusive="right"))
            & price_tier3,
        ),
        (
            4.22,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1, 1.25, inclusive="right"))
            & price_tier1,
        ),
        (
            5.04,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1, 1.25, inclusive="right"))
            & price_tier2,
        ),
        (
            5.30,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1, 1.25, inclusive="right"))
            & price_tier3,
        ),
        (
            4.60,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.25, 1.5, inclusive="right"))
            & price_tier1,
        ),
        (
            5.42,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.25, 1.5, inclusive="right"))
            & price_tier2,
        ),
        (
            5.68,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.25, 1.5, inclusive="right"))
            & price_tier3,
        ),
        (
            4.75,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.5, 1.75, inclusive="right"))
            & price_tier1,
        ),
        (
            5.57,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.5, 1.75, inclusive="right"))
            & price_tier2,
        ),
        (
            5.83,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.5, 1.75, inclusive="right"))
            & price_tier3,
        ),
        (
            5.00,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.75, 2, inclusive="right"))
            & price_tier1,
        ),
        (
            5.82,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.75, 2, inclusive="right"))
            & price_tier2,
        ),
        (
            6.08,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.75, 2, inclusive="right"))
            & price_tier3,
        ),
        (
            5.10,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2, 2.25, inclusive="right"))
            & price_tier1,
        ),
        (
            5.92,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2, 2.25, inclusive="right"))
            & price_tier2,
        ),
        (
            6.18,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2, 2.25, inclusive="right"))
            & price_tier3,
        ),
        (
            5.28,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.25, 2.5, inclusive="right"))
            & price_tier1,
        ),
        (
            6.10,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.25, 2.5, inclusive="right"))
            & price_tier2,
        ),
        (
            6.36,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.25, 2.5, inclusive="right"))
            & price_tier3,
        ),
        (
            5.44,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.5, 2.75, inclusive="right"))
            & price_tier1,
        ),
        (
            6.26,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.5, 2.75, inclusive="right"))
            & price_tier2,
        ),
        (
            6.52,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.5, 2.75, inclusive="right"))
            & price_tier3,
        ),
        (
            5.85,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.75, 3, inclusive="right"))
            & price_tier1,
        ),
        (
            6.67,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.75, 3, inclusive="right"))
            & price_tier2,
        ),
        (
            6.93,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.75, 3, inclusive="right"))
            & price_tier3,
        ),
        (
            6.15 + 0.08 * ((weight - 3) * 4),
            (size_tier == LARGE_STANDARD)
            & (weight.between(3, 20, inclusive="right"))
            & price_tier1,
        ),
        (
            6.97 + 0.08 * ((weight - 3) * 4),
            (size_tier == LARGE_STANDARD)
            & (weight.between(3, 20, inclusive="right"))
            & price_tier2,
        ),
        (
            7.23 + 0.08 * ((weight - 3) * 4),
            (size_tier == LARGE_STANDARD)
            & (weight.between(3, 20, inclusive="right"))
            & price_tier3,
        ),
        (
            6.78 + 0.38 * (weight - 1),
            (size_tier == SMALL_BULKY) & (weight <= 50) & price_tier1,
        ),
        (
            7.55 + 0.38 * (weight - 1),
            (size_tier == SMALL_BULKY) & (weight <= 50) & price_tier2,
        ),
        (
            7.55 + 0.38 * (weight - 1),
            (size_tier == SMALL_BULKY) & (weight <= 50) & price_tier3,
        ),
        (
            8.58 + 0.38 * (weight - 1),
            (size_tier == LARGE_BULKY) & (weight <= 50) & price_tier1,
        ),
        (
            9.36 + 0.38 * (weight - 1),
            (size_tier == LARGE_BULKY) & (weight <= 50) & price_tier2,
        ),
        (
            9.35 + 0.38 * (weight - 1),
            (size_tier == LARGE_BULKY) & (weight <= 50) & price_tier3,
        ),
        (
            25.56 + 0.38 * (weight - 1),
            (size_tier == EXTRA_LARGE_0_50) & (weight <= 50) & price_tier1,
        ),
        (
            26.33 + 0.38 * (weight - 1),
            (size_tier == EXTRA_LARGE_0_50) & (weight <= 50) & price_tier2,
        ),
        (
            26.33 + 0.38 * (weight - 1),
            (size_tier == EXTRA_LARGE_0_50) & (weight <= 50) & price_tier3,
        ),
        (
            36.55 + 0.75 * (weight - 51),
            (size_tier == EXTRA_LARGE_50_70)
            & (weight.between(50, 70, inclusive="right"))
            & price_tier1,
        ),
        (
            37.32 + 0.75 * (weight - 51),
            (size_tier == EXTRA_LARGE_50_70)
            & (weight.between(50, 70, inclusive="right"))
            & price_tier2,
        ),
        (
            37.32 + 0.75 * (weight - 51),
            (size_tier == EXTRA_LARGE_50_70)
            & (weight.between(50, 70, inclusive="right"))
            & price_tier3,
        ),
        (
            50.55 + 0.75 * (weight - 71),
            (size_tier == EXTRA_LARGE_70_150)
            & (weight.between(70, 150, inclusive="right"))
            & price_tier1,
        ),
        (
            51.32 + 0.75 * (weight - 71),
            (size_tier == EXTRA_LARGE_70_150)
            & (weight.between(70, 150, inclusive="right"))
            & price_tier2,
        ),
        (
            51.32 + 0.75 * (weight - 71),
            (size_tier == EXTRA_LARGE_70_150)
            & (weight.between(70, 150, inclusive="right"))
            & price_tier3,
        ),
        (
            194.18 + 0.19 * (weight - 151),
            (size_tier == EXTRA_LARGE_150_PLUS) & (weight > 150) & price_tier1,
        ),
        (
            194.95 + 0.19 * (weight - 151),
            (size_tier == EXTRA_LARGE_150_PLUS) & (weight > 150) & price_tier2,
        ),
        (
            194.95 + 0.19 * (weight - 151),
            (size_tier == EXTRA_LARGE_150_PLUS) & (weight > 150) & price_tier3,
        ),
    ]

    fees_fba_peak = [
        (2.48, (size_tier == SMALL_STANDARD) & (weight_oz <= 2) & price_tier1),
        (3.25, (size_tier == SMALL_STANDARD) & (weight_oz <= 2) & price_tier2),
        (3.25, (size_tier == SMALL_STANDARD) & (weight_oz <= 2) & price_tier3),
        (
            2.57,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(2, 4, inclusive="right"))
            & price_tier1,
        ),
        (
            3.34,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(2, 4, inclusive="right"))
            & price_tier2,
        ),
        (
            3.34,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(2, 4, inclusive="right"))
            & price_tier3,
        ),
        (
            2.67,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(4, 6, inclusive="right"))
            & price_tier1,
        ),
        (
            3.44,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(4, 6, inclusive="right"))
            & price_tier2,
        ),
        (
            3.44,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(4, 6, inclusive="right"))
            & price_tier3,
        ),
        (
            2.76,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(6, 8, inclusive="right"))
            & price_tier1,
        ),
        (
            3.53,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(6, 8, inclusive="right"))
            & price_tier2,
        ),
        (
            3.53,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(6, 8, inclusive="right"))
            & price_tier3,
        ),
        (
            2.87,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(8, 10, inclusive="right"))
            & price_tier1,
        ),
        (
            3.64,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(8, 10, inclusive="right"))
            & price_tier2,
        ),
        (
            3.64,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(8, 10, inclusive="right"))
            & price_tier3,
        ),
        (
            2.97,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(10, 12, inclusive="right"))
            & price_tier1,
        ),
        (
            3.74,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(10, 12, inclusive="right"))
            & price_tier2,
        ),
        (
            3.74,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(10, 12, inclusive="right"))
            & price_tier3,
        ),
        (
            3.05,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(12, 14, inclusive="right"))
            & price_tier1,
        ),
        (
            3.82,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(12, 14, inclusive="right"))
            & price_tier2,
        ),
        (
            3.82,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(12, 14, inclusive="right"))
            & price_tier3,
        ),
        (
            3.10,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(14, 16, inclusive="right"))
            & price_tier1,
        ),
        (
            3.87,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(14, 16, inclusive="right"))
            & price_tier2,
        ),
        (
            3.87,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(14, 16, inclusive="right"))
            & price_tier3,
        ),
        (3.15, (size_tier == LARGE_STANDARD) & (weight_oz <= 4) & price_tier1),
        (3.92, (size_tier == LARGE_STANDARD) & (weight_oz <= 4) & price_tier2),
        (3.92, (size_tier == LARGE_STANDARD) & (weight_oz <= 4) & price_tier3),
        (
            3.39,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(4, 8, inclusive="right"))
            & price_tier1,
        ),
        (
            4.16,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(4, 8, inclusive="right"))
            & price_tier2,
        ),
        (
            4.16,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(4, 8, inclusive="right"))
            & price_tier3,
        ),
        (
            3.66,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(8, 12, inclusive="right"))
            & price_tier1,
        ),
        (
            4.43,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(8, 12, inclusive="right"))
            & price_tier2,
        ),
        (
            4.43,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(8, 12, inclusive="right"))
            & price_tier3,
        ),
        (
            4.07,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(12, 16, inclusive="right"))
            & price_tier1,
        ),
        (
            4.84,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(12, 16, inclusive="right"))
            & price_tier2,
        ),
        (
            4.84,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(12, 16, inclusive="right"))
            & price_tier3,
        ),
        (
            4.52,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1, 1.25, inclusive="right"))
            & price_tier1,
        ),
        (
            5.29,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1, 1.25, inclusive="right"))
            & price_tier2,
        ),
        (
            5.29,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1, 1.25, inclusive="right"))
            & price_tier3,
        ),
        (
            4.91,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.25, 1.5, inclusive="right"))
            & price_tier1,
        ),
        (
            5.68,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.25, 1.5, inclusive="right"))
            & price_tier2,
        ),
        (
            5.68,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.25, 1.5, inclusive="right"))
            & price_tier3,
        ),
        (
            5.07,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.5, 1.75, inclusive="right"))
            & price_tier1,
        ),
        (
            5.84,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.5, 1.75, inclusive="right"))
            & price_tier2,
        ),
        (
            5.84,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.5, 1.75, inclusive="right"))
            & price_tier3,
        ),
        (
            5.33,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.75, 2, inclusive="right"))
            & price_tier1,
        ),
        (
            6.10,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.75, 2, inclusive="right"))
            & price_tier2,
        ),
        (
            6.10,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.75, 2, inclusive="right"))
            & price_tier3,
        ),
        (
            5.47,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2, 2.25, inclusive="right"))
            & price_tier1,
        ),
        (
            6.24,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2, 2.25, inclusive="right"))
            & price_tier2,
        ),
        (
            6.24,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2, 2.25, inclusive="right"))
            & price_tier3,
        ),
        (
            5.67,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.25, 2.5, inclusive="right"))
            & price_tier1,
        ),
        (
            6.44,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.25, 2.5, inclusive="right"))
            & price_tier2,
        ),
        (
            6.44,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.25, 2.5, inclusive="right"))
            & price_tier3,
        ),
        (
            5.84,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.5, 2.75, inclusive="right"))
            & price_tier1,
        ),
        (
            6.61,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.5, 2.75, inclusive="right"))
            & price_tier2,
        ),
        (
            6.61,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.5, 2.75, inclusive="right"))
            & price_tier3,
        ),
        (
            6.26,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.75, 3, inclusive="right"))
            & price_tier1,
        ),
        (
            7.03,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.75, 3, inclusive="right"))
            & price_tier2,
        ),
        (
            7.03,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.75, 3, inclusive="right"))
            & price_tier3,
        ),
        (
            6.69 + 0.08 * ((weight - 3) * 4),
            (size_tier == LARGE_STANDARD)
            & (weight.between(3, 20, inclusive="right"))
            & price_tier1,
        ),
        (
            7.46 + 0.08 * ((weight - 3) * 4),
            (size_tier == LARGE_STANDARD)
            & (weight.between(3, 20, inclusive="right"))
            & price_tier2,
        ),
        (
            7.46 + 0.08 * ((weight - 3) * 4),
            (size_tier == LARGE_STANDARD)
            & (weight.between(3, 20, inclusive="right"))
            & price_tier3,
        ),
        (
            9.88 + 0.38 * (weight - 1),
            (size_tier == SMALL_BULKY) & (weight <= 50) & price_tier1,
        ),
        (
            10.65 + 0.38 * (weight - 1),
            (size_tier == SMALL_BULKY) & (weight <= 50) & price_tier2,
        ),
        (
            10.65 + 0.38 * (weight - 1),
            (size_tier == SMALL_BULKY) & (weight <= 50) & price_tier3,
        ),
        (
            9.88 + 0.38 * (weight - 1),
            (size_tier == LARGE_BULKY) & (weight <= 50) & price_tier1,
        ),
        (
            10.65 + 0.38 * (weight - 1),
            (size_tier == LARGE_BULKY) & (weight <= 50) & price_tier2,
        ),
        (
            10.65 + 0.38 * (weight - 1),
            (size_tier == LARGE_BULKY) & (weight <= 50) & price_tier3,
        ),
        (
            28.29 + 0.38 * (weight - 1),
            (size_tier == EXTRA_LARGE_0_50) & (weight <= 50) & price_tier1,
        ),
        (
            29.06 + 0.38 * (weight - 1),
            (size_tier == EXTRA_LARGE_0_50) & (weight <= 50) & price_tier2,
        ),
        (
            29.06 + 0.38 * (weight - 1),
            (size_tier == EXTRA_LARGE_0_50) & (weight <= 50) & price_tier3,
        ),
        (
            42.16 + 0.75 * (weight - 51),
            (size_tier == EXTRA_LARGE_50_70)
            & (weight.between(50, 70, inclusive="right"))
            & price_tier1,
        ),
        (
            42.93 + 0.75 * (weight - 51),
            (size_tier == EXTRA_LARGE_50_70)
            & (weight.between(50, 70, inclusive="right"))
            & price_tier2,
        ),
        (
            42.93 + 0.75 * (weight - 51),
            (size_tier == EXTRA_LARGE_50_70)
            & (weight.between(50, 70, inclusive="right"))
            & price_tier3,
        ),
        (
            58.46 + 0.75 * (weight - 71),
            (size_tier == EXTRA_LARGE_70_150)
            & (weight.between(70, 150, inclusive="right"))
            & price_tier1,
        ),
        (
            59.23 + 0.75 * (weight - 71),
            (size_tier == EXTRA_LARGE_70_150)
            & (weight.between(70, 150, inclusive="right"))
            & price_tier2,
        ),
        (
            59.23 + 0.75 * (weight - 71),
            (size_tier == EXTRA_LARGE_70_150)
            & (weight.between(70, 150, inclusive="right"))
            & price_tier3,
        ),
        (
            202.69 + 0.19 * (weight - 151),
            (size_tier == EXTRA_LARGE_150_PLUS) & (weight > 150) & price_tier1,
        ),
        (
            203.46 + 0.19 * (weight - 151),
            (size_tier == EXTRA_LARGE_150_PLUS) & (weight > 150) & price_tier2,
        ),
        (
            203.46 + 0.19 * (weight - 151),
            (size_tier == EXTRA_LARGE_150_PLUS) & (weight > 150) & price_tier3,
        ),
    ]

    actual_non_peak_fees = (
        fees_fba_non_peak_before_2026
        if fee_date < pd.to_datetime("2026-01-15")
        else fees_fba_non_peak
    )
    actual_peak_fees = (
        fees_fba_peak_before_2026
        if fee_date < pd.to_datetime("2026-01-15")
        else fees_fba_peak
    )

    peak_start, peak_end = (10, 15), (1, 14)
    if (
        peak_start <= (today_date.month, today_date.day)
        or (today_date.month, today_date.day) <= peak_end
    ):
        fees = actual_peak_fees
    else:
        fees = actual_non_peak_fees
    df["fba_fee"] = np.select([x[1] for x in fees], [x[0] for x in fees], np.nan)

    df["fba_non_peak_fee"] = np.select(
        [x[1] for x in actual_non_peak_fees],
        [x[0] for x in actual_non_peak_fees],
        np.nan,
    )

    df["fba_peak_fee"] = np.select(
        [x[1] for x in actual_peak_fees], [x[0] for x in actual_peak_fees], np.nan
    )
    return df


def get_removal_fee(df):
    size_tier = df["size_tier"]
    weight = df["shipping_weight"]
    removal_fee = [
        (
            1.04,
            (size_tier.isin([SMALL_STANDARD, LARGE_STANDARD]))
            & (weight.between(0, 0.5, inclusive="right")),
        ),
        (
            1.53,
            (size_tier.isin([SMALL_STANDARD, LARGE_STANDARD]))
            & (weight.between(0.5, 1, inclusive="right")),
        ),
        (
            2.27,
            (size_tier.isin([SMALL_STANDARD, LARGE_STANDARD]))
            & (weight.between(1, 2, inclusive="right")),
        ),
        (
            2.89 + 1.06 * (weight - 2),
            (size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])) & (weight > 2),
        ),
        (
            3.12,
            (~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD]))
            & (weight.between(0, 1, inclusive="right")),
        ),
        (
            4.30,
            (~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD]))
            & (weight.between(1, 2, inclusive="right")),
        ),
        (
            6.36,
            (~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD]))
            & (weight.between(2, 4, inclusive="right")),
        ),
        (
            10.04,
            (~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD]))
            & (weight.between(4, 10, inclusive="right")),
        ),
        (
            14.32 + 1.06 * (weight - 10),
            (~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])) & (weight > 10),
        ),
    ]
    df["removal_fee"] = np.select(
        [x[1] for x in removal_fee], [x[0] for x in removal_fee], np.nan
    )
    return df


def sipp_discount(df):
    df.loc[
        df["size_tier"].isin([SMALL_STANDARD, EXTRA_LARGE_150_PLUS]), "shipping_weight"
    ] = df["individual weight lbs"]
    df["shipping_weight, oz"] = df["shipping_weight"] * 16
    size_tier = df["size_tier"]
    weight = df["shipping_weight"]
    weight_oz = df["shipping_weight, oz"]

    sipp_discount = [
        (0.04, (size_tier == SMALL_STANDARD) & (weight_oz <= 2)),
        (
            0.04,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(2, 4, inclusive="right")),
        ),
        (
            0.05,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(4, 6, inclusive="right")),
        ),
        (
            0.05,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(6, 8, inclusive="right")),
        ),
        (
            0.06,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(8, 10, inclusive="right")),
        ),
        (
            0.06,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(10, 12, inclusive="right")),
        ),
        (
            0.07,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(12, 14, inclusive="right")),
        ),
        (
            0.07,
            (size_tier == SMALL_STANDARD)
            & (weight_oz.between(14, 16, inclusive="right")),
        ),
        (0.04, (size_tier == LARGE_STANDARD) & (weight_oz <= 4)),
        (
            0.04,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(4, 8, inclusive="right")),
        ),
        (
            0.07,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(8, 12, inclusive="right")),
        ),
        (
            0.08,
            (size_tier == LARGE_STANDARD)
            & (weight_oz.between(12, 16, inclusive="right")),
        ),
        (
            0.09,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1, 1.25, inclusive="right")),
        ),
        (
            0.09,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.25, 1.5, inclusive="right")),
        ),
        (
            0.10,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.5, 1.75, inclusive="right")),
        ),
        (
            0.11,
            (size_tier == LARGE_STANDARD)
            & (weight.between(1.75, 2, inclusive="right")),
        ),
        (
            0.12,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2, 2.25, inclusive="right")),
        ),
        (
            0.13,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.25, 2.5, inclusive="right")),
        ),
        (
            0.14,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.5, 2.75, inclusive="right")),
        ),
        (
            0.14,
            (size_tier == LARGE_STANDARD)
            & (weight.between(2.75, 3, inclusive="right")),
        ),
        (
            0.23,
            (size_tier == LARGE_STANDARD) & (weight.between(3, 20, inclusive="right")),
        ),
        (1.32, (size_tier == LARGE_BULKY) & (weight <= 50)),
        (0, (size_tier == EXTRA_LARGE_0_50) & (weight <= 50)),
        (
            0,
            (size_tier == EXTRA_LARGE_50_70)
            & (weight.between(50, 70, inclusive="right")),
        ),
        (
            0,
            (size_tier == EXTRA_LARGE_70_150)
            & (weight.between(70, 150, inclusive="right")),
        ),
        (0, (size_tier == EXTRA_LARGE_150_PLUS) & (weight > 150)),
    ]

    df["sipp_discount"] = np.select(
        [x[1] for x in sipp_discount], [x[0] for x in sipp_discount], np.nan
    )
    return df


def get_liquidation_fee(df):
    size_tier = df["size_tier"]
    weight = df["shipping_weight"]
    liquidation_fee = [
        (
            0.25,
            (size_tier.isin([SMALL_STANDARD, LARGE_STANDARD]))
            & (weight.between(0, 0.5, inclusive="right")),
        ),
        (
            0.3,
            (size_tier.isin([SMALL_STANDARD, LARGE_STANDARD]))
            & (weight.between(0.5, 1, inclusive="right")),
        ),
        (
            0.35,
            (size_tier.isin([SMALL_STANDARD, LARGE_STANDARD]))
            & (weight.between(1, 2, inclusive="right")),
        ),
        (
            0.4 + 0.2 * (weight - 2),
            (size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])) & (weight > 2),
        ),
        (
            0.6,
            (~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD]))
            & (weight.between(0, 1, inclusive="right")),
        ),
        (
            0.7,
            (~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD]))
            & (weight.between(1, 2, inclusive="right")),
        ),
        (
            0.9,
            (~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD]))
            & (weight.between(2, 4, inclusive="right")),
        ),
        (
            1.45,
            (~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD]))
            & (weight.between(4, 10, inclusive="right")),
        ),
        (
            1.9 + 0.2 * (weight - 10),
            (~size_tier.isin([SMALL_STANDARD, LARGE_STANDARD])) & (weight > 10),
        ),
    ]
    df["liquidation_fee"] = np.select(
        [x[1] for x in liquidation_fee], [x[0] for x in liquidation_fee], np.nan
    )
    return df


def export_to_excel(df):
    try:
        with pd.ExcelWriter(
            os.path.join(user_folder, "fees.xlsx"), engine="xlsxwriter"
        ) as writer:
            df.to_excel(writer, sheet_name="Fees", index=False)
            mm.format_header(df, writer, "Fees")
    except PermissionError:
        PopupError("Please close the file first")
        export_to_excel(df)
    mm.open_file_folder(user_folder)


def separate_file():
    file_path = filedialog.askopenfilename(
        title="Select file with dimensions",
        filetypes=[
            ("Excel files", "*.xls*"),
        ],
    )
    df = pd.read_excel(file_path)
    df = get_shipping_weight(df)
    df = get_size_tier(df)
    df = get_fulfillment_fee(df)
    df = get_removal_fee(df)
    df = get_liquidation_fee(df)
    df = get_storage_fee(df)
    df = sipp_discount(df)
    export_to_excel(df)


def pull_matrix_file(file_id=MATRIX_FILE_ID):
    """
    Pull the Master Product Library file from GDrive
    """
    try:
        matrix_file_obj = gd.download_file(file_id=file_id)
        matrix_file = pd.ExcelFile(matrix_file_obj)
        matrix_df = pd.read_excel(matrix_file, sheet_name="Mellanni Catalog")
        matrix_df_columns = matrix_df.columns.tolist()
        individual_dimensions_index = matrix_df_columns.index("Individual Dimensions (in)")
        individual_weight_index = matrix_df_columns.index("Individual Weight Lbs")
        sets_in_a_box_index = matrix_df_columns.index("Sets in a box")
        box_dimensions_index = matrix_df_columns.index("Box Dimensions (in)")
        box_weight_index = matrix_df_columns.index("Box Weight Lbs")

        matrix_df = pd.read_excel(matrix_file, sheet_name="Mellanni Catalog", skiprows=1)
        sku_index = matrix_df.columns.tolist().index("sku")

        columns_to_use = {
            sku_index: "sku",
            individual_dimensions_index: "target_l",
            individual_dimensions_index + 1: "target_w",
            individual_dimensions_index + 2: "target_h",
            individual_weight_index: "target_weight",
            sets_in_a_box_index: "target_qty_per_box",
            box_dimensions_index: "target_box_l",
            box_dimensions_index + 1: "target_box_w",
            box_dimensions_index + 2: "target_box_h",
            box_weight_index: "target_box_weight",
        }

        column_mapping = {
            matrix_df.columns.tolist()[i]: new_name
            for i, new_name in columns_to_use.items()
        }
        matrix_df = matrix_df.rename(columns=column_mapping)
        matrix_df = matrix_df[columns_to_use.values()]
        return matrix_df
    except Exception as e:
        raise BaseException(f"Error while pulling Matrix file: {str(e)}")


def main(out=True):
    executor = ThreadPoolExecutor()
    dimensions_future = executor.submit(get_dims_file)
    prices_future = executor.submit(get_prices_file)

    dimensions = dimensions_future.result()
    prices = prices_future.result()
    dictionary_future = executor.submit(get_dictionary)
    dictionary = dictionary_future.result()
    # dimensions = get_dims_file()
    # dictionary = get_dictionary()
    # prices = get_prices_file()
    combined = combine_files(dimensions.copy(), dictionary.copy(), prices.copy())
    combined = get_shipping_weight(combined)
    combined = get_size_tier(combined)
    combined = get_fulfillment_fee(combined)
    combined = get_removal_fee(combined)
    combined = get_liquidation_fee(combined)
    combined = get_storage_fee(combined)
    for nc in num_cols:
        combined[nc] = combined[nc].astype(float, errors="ignore")
    if not out:
        return combined
    export_to_excel(combined)

    dims_cloud = dimensions.copy().fillna(0)
    dims_cloud = gc.normalize_columns(dims_cloud)
    gc.push_to_cloud(
        dims_cloud, destination="auxillary_development.dimensions", if_exists="replace"
    )
    del dims_cloud


if __name__ == "__main__":
    try:
        mode = int(
            input("Select mode: 1 for Mellanni dimensions, or 2 for file upload\n\n")
        )
        if mode == 1:
            main(out=True)
        elif mode == 2:
            separate_file()
    except Exception as e:
        PopupError(e)
