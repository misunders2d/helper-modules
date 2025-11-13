import os
from typing import List
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
            subprocess.Popen([opener, path])
    except Exception as e:
        print(f"Uncaught exception occurred: {e}")
    return None


def export_to_excel(
    dfs: List[pd.DataFrame],
    sheet_names: List[str],
    filename: str = "test.xlsx",
    out_folder: str | None = None,
) -> None:
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


def encrypt_string(hash_string):
    """
    Create a hashed string from any input
    """
    import hashlib

    string_encoded = hashlib.sha256(str(hash_string).encode()).hexdigest()
    return string_encoded
