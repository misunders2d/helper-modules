import os
from typing import List, Literal, Union, Dict, Any

FormattingType = Literal[
    "2-color",
    "3-color",
    "highlight",
    "number",
    "medium number",
    "decimal",
    "currency",
    "percent",
]
import pandas as pd
import re
import sys, subprocess
from connectors import gdrive as gd
import customtkinter
import datetime

from ctk_gui.ctk_windows import PopupError, PopupYesNo

excluded_collections = [
    "Cotton 300TC Percale Sheet Set",
    "Egyptian Cotton Striped Bed Sheet Set",
    "Cotton Flannel Pillowcases",
    "Decorative Pillow",
    "All-Year-Round Throws",
    "Down Alternative Comforter",
    "Faux Rabbit Fur Area Rug",
    "Comforter",
    "Bed Skirt",
    "3pc Microfiber Bed Sheet Set Full",
    "Cotton Pillowcases",
    "Blackout Curtains",
    "Cotton Flannel Fitted Sheet",
    "Faux Fur Throw",
    "Plush Coverlet Set",
    "3pc Microfiber Bed Sheet Set Queen",
    "6 PC Egyptian Cotton Striped Bed Sheet Set",
    "Faux Cachemire Acrylic Throw Blanket",
    "Cotton 300TC Percale Pillowcase",
    "Acrylic Knit Sherpa Throw Blanket",
    "Cotton Quilt",
    "Cotton 300TC Sateen Sheet Set",
    "Egyptian Cotton Bed Sheet Set",
    "Jersey Cotton Quilt",
    "Bundles",
    "Pillow inserts",
]
user_folder = os.path.join(os.path.expanduser("~"), "temp")


def week_number(date: datetime.date) -> int:
    """returns a week number for weeks starting with Sunday"""
    if not isinstance(date, datetime.date):
        try:
            date = datetime.datetime.strptime(date, "%Y-%m-%d")
        except Exception as e:
            raise BaseException(f"Date format not recognized:\n{e}")
    return (
        date.isocalendar().week + 1 if date.weekday() == 6 else date.isocalendar().week
    )


def open_file_folder(path: str) -> None:
    try:
        if hasattr(os, "startfile"):
            os.startfile(path)  # type: ignore
        else:
            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.Popen(
                [opener, path], stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL
            )
    except Exception as e:
        print(f"Uncaught exception occurred: {e}")
    return None


def export_to_excel(
    dfs: List[pd.DataFrame],
    sheet_names: List[str],
    filename: str = "test.xlsx",
    out_folder: str | None = None,
    column_formats: (
        Dict[str, Union[FormattingType, Dict[str, Any], List[Union[FormattingType, Dict[str, Any]]]]]
        | None
    ) = None,
) -> None:
    """
    Exports dataframes to multiple sheets in an Excel file with optional column formatting.

    Args:
        dfs: List of pandas DataFrames to export.
        sheet_names: List of sheet names corresponding to the DataFrames.
        filename: Name of the output file.
        out_folder: Directory to save the file. If None, a dialog will open.
        column_formats: A dictionary mapping column names to formatting configurations.

    Example Usage:
        column_formats = {
            "Total Sales": "currency",
            "Growth %": "percent",
            "Margin": {
                "type": "3-color",
                "min_color": "#F8696B", # Red
                "mid_color": "#FFEB84", # Yellow
                "max_color": "#63BE7B", # Green
                "mid_value": 0
            },
            "Stock Level": {
                "type": "2-color",
                "min_color": "#FFFFFF",
                "max_color": "#00FF00"
            }
        }
        export_to_excel([df], ["Report"], "sales.xlsx", column_formats=column_formats)
    """
    from customtkinter import filedialog

    if not out_folder:
        out_folder = filedialog.askdirectory(
            title="Select output folder", initialdir=os.path.expanduser("~")
        )
    full_output = os.path.join(out_folder, filename)
    try:
        with pd.ExcelWriter(full_output, engine="xlsxwriter") as writer:
            for df, sheet_name in list(zip(dfs, sheet_names)):
                if len(df) > 0:
                    df.to_excel(excel_writer=writer, sheet_name=sheet_name, index=False)
                    format_header(df, writer, sheet_name)
                    if column_formats:
                        apply_formatting(df, writer, sheet_name, column_formats)
    except PermissionError:
        print(f"{filename} is open, please close the file first")
        export_to_excel(dfs, sheet_names, filename, out_folder)
    except Exception as e:
        print(e)

    return None


def get_comments():
    from openpyxl import load_workbook
    import os

    file_path = customtkinter.filedialog.askopenfilename(
        title="Select the file processing report"
    )

    if file_path and file_path != "":
        try:
            wb = load_workbook(file_path, data_only=True)
            ws = wb["Template"]  # or whatever sheet name
            ws.delete_rows(0)
            ws.delete_rows(0)

            comments = []
            for row in ws.rows:
                comments.append(row[2].comment)
            comments = comments[1:]

            file = pd.read_excel(file_path, sheet_name="Template", skiprows=2)
            cols = file.columns.tolist()
            cols.insert(4, "comments")

            file["comments"] = comments
            file = file[cols]

            output = customtkinter.filedialog.askdirectory(title="Select output folder")
            with pd.ExcelWriter(os.path.join(output, "comments.xlsx")) as writer:
                file.to_excel(writer, sheet_name="Comments", index=False)
                format_header(file, writer, "Comments")
            open_file_folder(output)
        except Exception as e:
            PopupError(f"Sorry, error occurred:\n{e}")
    else:
        PopupError("Please select a file")
    return None


def password_generator(x):
    """
    Generates a password of 'x' lenght from letters, digits and punctuation marks

    Parameters
    ----------
    x : int
        number of symbols to use.

    Returns
    str
    password
    """
    import string
    import random

    text_part = x // 3 * 2
    num_part = x - text_part
    text_str = [random.choice(string.ascii_letters) for x in range(text_part)]
    num_str = [
        random.choice(string.digits * 2 + string.punctuation) for x in range(num_part)
    ]
    text = text_str + num_str
    random.shuffle(text)
    password = "".join(text)
    return password


def convert_to_pacific(db, columns):
    import pytz

    pacific = pytz.timezone("US/Pacific")
    db["pacific-date"] = pd.to_datetime(db[columns]).dt.tz_convert(pacific)
    db["pacific-date"] = pd.to_datetime(db["pacific-date"]).dt.tz_localize(None)
    return db["pacific-date"]


def format_header(df, writer, sheet):
    workbook = writer.book
    cell_format = workbook.add_format(
        {"bold": True, "text_wrap": True, "valign": "center", "font_size": 9}
    )
    worksheet = writer.sheets[sheet]
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, cell_format)
    max_row, max_col = df.shape
    worksheet.autofilter(0, 0, max_row, max_col - 1)
    worksheet.freeze_panes(1, 0)
    return None


def format_columns(df, writer, sheet, col_num):
    worksheet = writer.sheets[sheet]
    if not isinstance(col_num, list):
        col_num = [col_num]
    else:
        pass
    for c in col_num:
        width = max(df.iloc[:, c].astype(str).map(len).max(), len(df.iloc[:, c].name))
        worksheet.set_column(c, c, width)
    return None


def apply_formatting(
    df: pd.DataFrame,
    writer: pd.ExcelWriter,
    sheet: str,
    column_formats: Dict[
        str, Union[FormattingType, Dict[str, Any], List[Union[FormattingType, Dict[str, Any]]]]
    ],
):
    """
    Internal helper to apply conditional and number formatting to specific columns.

    Supported Types:
        - "2-color": Gradient between two colors.
        - "3-color": Gradient between three colors (min, mid, max).
        - "highlight": Highlight min or max value (target="min" or "max").
        - "number"/"medium number": #,##0
        - "decimal": #,##0.00
        - "currency": $#,##0.00
        - "percent": 0.0%

    Note: A list of formats can be passed for a single column (e.g., ["3-color", "decimal"]).
    """
    workbook = writer.book
    worksheet = writer.sheets[sheet]
    max_row = len(df)
    cols = list(df.columns)

    for col_name, format_config_raw in column_formats.items():
        if col_name not in cols:
            continue

        col_idx = cols.index(col_name)
        # Excel range (skipping header)
        cell_range = f"{chr(65 + col_idx)}2:{chr(65 + col_idx)}{max_row + 1}"

        # Normalize to a list of configs
        if isinstance(format_config_raw, list):
            configs = format_config_raw
        else:
            configs = [format_config_raw]

        for format_config in configs:
            if isinstance(format_config, str):
                format_config = {"type": format_config}

            fmt_type = format_config.get("type")

            if fmt_type == "2-color":
                worksheet.conditional_format(
                    cell_range,
                    {
                        "type": "2_color_scale",
                        "min_color": format_config.get("min_color", "#FFFFFF"),
                        "max_color": format_config.get("max_color", "#63BE7B"),
                        "min_type": format_config.get("min_type", "min"),
                        "max_type": format_config.get("max_type", "max"),
                        "min_value": format_config.get("min_value"),
                        "max_value": format_config.get("max_value"),
                    },
                )
            elif fmt_type == "3-color":
                worksheet.conditional_format(
                    cell_range,
                    {
                        "type": "3_color_scale",
                        "min_color": format_config.get("min_color", "#F8696B"),
                        "mid_color": format_config.get("mid_color", "#FFEB84"),
                        "max_color": format_config.get("max_color", "#63BE7B"),
                        "min_type": format_config.get("min_type", "min"),
                        "mid_type": format_config.get("mid_type", "percentile"),
                        "max_type": format_config.get("max_type", "max"),
                        "min_value": format_config.get("min_value"),
                        "mid_value": format_config.get("mid_value", 50),
                        "max_value": format_config.get("max_value"),
                    },
                )
            elif fmt_type == "highlight":
                target = format_config.get("target", "max")
                color = format_config.get("color", "#C6EFCE")
                font_color = format_config.get("font_color", "#006100")

                fmt = workbook.add_format({"bg_color": color, "font_color": font_color})

                # Using formula for min/max highlight
                col_letter = chr(65 + col_idx)
                if target == "max":
                    formula = (
                        f"={col_letter}2=MAX(${col_letter}$2:${col_letter}${max_row + 1})"
                    )
                else:
                    formula = (
                        f"={col_letter}2=MIN(${col_letter}$2:${col_letter}${max_row + 1})"
                    )

                worksheet.conditional_format(
                    cell_range, {"type": "formula", "criteria": formula, "format": fmt}
                )

            # Number formats
            num_fmt_str = None
            if fmt_type == "number" or fmt_type == "medium number":
                additional_decimals = "." + "0" * int(format_config.get("decimals", 0))
                num_fmt_str = "#,##0"
                if len(additional_decimals) > 1:
                    num_fmt_str += additional_decimals
            elif fmt_type == "decimal":
                num_fmt_str = "#,##0.00"
            elif fmt_type == "currency":
                num_fmt_str = "$#,##0.00"
            elif fmt_type == "percent":
                num_fmt_str = "0.0%"

            if num_fmt_str:
                num_fmt = workbook.add_format({"num_format": num_fmt_str})
                worksheet.set_column(col_idx, col_idx, None, num_fmt)

    return None


def encrypt_string(hash_string):
    """
    Create a hashed string from any input
    """
    import hashlib

    string_encoded = hashlib.sha256(str(hash_string).encode()).hexdigest()
    return string_encoded
