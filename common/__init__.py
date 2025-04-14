import os

from .events import events, event_dates
from .global_params import excluded_collections, user_folder

from pandas import date_range, to_datetime
date_ranges = {key:date_range(value[0], value[1]) for key, value in events.items()}
excluded_dates = [x.date() for value in date_ranges.values() for x in value]

if not os.path.exists(user_folder):
    os.makedirs(user_folder)