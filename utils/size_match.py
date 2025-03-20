import os

from utils import mellanni_modules as mm
from connectors import gcloud as gc
from connectors import gdrive as gd
from common import user_folder, excluded_collections
import pandas as pd
import numpy as np
from ctk_gui.ctk_windows import PopupError, filedialog

drive_id = '0AMdx9NlXacARUk9PVA'

# Product size tiers page: https://sellercentral.amazon.com/help/hub/reference/GG5KW835AHDJCH8W
# Storage fee page: https://sellercentral.amazon.com/help/hub/reference/G3EDYEF6KUCFQTNM
# FBA fee page: https://sellercentral.amazon.com/help/hub/reference/GABBX6GZPA8MSZGW

SMALL_STANDARD = 'Small standard size'
LARGE_STANDARD = 'Large standard size'
LARGE_BULKY = 'Large bulky'
EXTRA_LARGE_0_50 = 'Extra large up to 50'
EXTRA_LARGE_50_70 = 'Extra large 50 to 70'
EXTRA_LARGE_70_150 = 'Extra large 70 to 150'
EXTRA_LARGE_150_PLUS = 'Extra large 150+'
num_cols = ['l', 'w', 'h', 'individual weight lbs', 'sets in a box',
            'box length', 'box width', 'box depth', 'box weight lbs',
            'dim_weight','shipping_weight', 'shipping_weight, oz', 'fba_fee',
            'removal_fee','current_storage_fee','storage_jan_sept',
            'storage_oct_dec','avg_yearly_storage']
today = pd.to_datetime('today')

def get_dims_file(folder_id='1zIHmbWcRRVyCTtuB9Atzam7IhAs8Ymx4',filename='DIMENSIONS.xlsx'):
    file = pd.read_excel(gd.download_file(gd.find_file_id(folder_id, drive_id, filename)), skiprows = 1)
    del file['Product Name']
    file.columns = [x.lower() for x in file.columns]
    file = file[~file['collection'].isin(excluded_collections)]
    return file

# TODO Change to this function when Dictionary is stored in Google Sheets exclusively
# def get_dictionary(spreadsheet_id='1tezZ1Txml4E1YGYnO8-57lSnxlzsJ2boUYb4xmVxmtw', sheet_id='1749064367'):
#     file = gd.download_gspread(spreadsheet_id=spreadsheet_id, sheet_id=sheet_id)
#     file = file[['SKU','ASIN', 'Collection','Sub-collection', 'Size Map','Color','Actuality']]
#     file.columns = [x.lower() for x in file.columns]
#     file = file[~file['collection'].isin(excluded_collections)]
#     file = file.drop_duplicates('sku')
#     return file

def get_dictionary(folder_id='1zIHmbWcRRVyCTtuB9Atzam7IhAs8Ymx4',filename='Dictionary.xlsx'):
    file = pd.read_excel(gd.download_file(gd.find_file_id(folder_id, drive_id, filename)))
    file = file[['SKU','ASIN', 'Collection','Sub-collection', 'Size Map','Color','Actuality']]
    file.columns = [x.lower() for x in file.columns]
    file = file[~file['collection'].isin(excluded_collections)]
    file = file.drop_duplicates('sku')
    file['country_code'] = 'US'
    file['sales_channel'] = 'Amazon.com'
    return file

def combine_files(dimensions, dictionary):
    collections = dimensions['collection'].unique()
    dictionary['c'] = dictionary['sub-collection'].copy()
    dictionary.loc[~dictionary['sub-collection'].isin(collections),'c'] = dictionary['collection']
    for c in ['collection','sub-collection']:
        del dictionary[c]
    dictionary = dictionary.rename(columns = {'size map':'size','c':'collection'})
    dictionary = dictionary[['sku', 'asin', 'collection', 'size', 'color', 'actuality','sales_channel']] 
    result = pd.merge(dictionary, dimensions, how = 'right', on = ['collection','size'])
    return result

def get_shipping_weight(df):
    df['dim_weight'] = (df['l'] * df['w'] * df['h']) / 139
    df['shipping_weight'] = df[['dim_weight','individual weight lbs']].max(axis = 1)
    return df

def get_size_tier(df):
    max_side = df[['l','w','h']].max(axis = 1)
    min_side = df[['l','w','h']].min(axis = 1)
    med_side = df[['l','w','h']].median(axis = 1)
    length_girth = (min_side + med_side) * 2
    weight = df['shipping_weight'].copy()

    conditions = [
        (SMALL_STANDARD, (weight <= 1) & (max_side <= 15) & (med_side <= 12) & (min_side <= 0.75)),
        (LARGE_STANDARD, (weight <= 20) & (max_side <= 18) & (med_side <= 14) & (min_side <= 8)),
        (LARGE_BULKY, (weight <= 50) & (max_side <= 59) & (med_side <= 33) & (min_side <= 33) & (length_girth <= 130)),
        (EXTRA_LARGE_0_50, (weight <= 50) | (max_side > 59) | (med_side > 33) | (min_side > 33) | (length_girth > 130)),
        (EXTRA_LARGE_50_70, 50 < (weight <= 70) | (max_side > 59) | (med_side > 33) | (min_side > 33) | (length_girth > 130)),
        (EXTRA_LARGE_70_150, 70 < (weight <= 150) | (max_side > 59) | (med_side > 33) | (min_side > 33) | (length_girth > 130)),
        (EXTRA_LARGE_150_PLUS, (weight > 150) | (max_side > 59) | (med_side > 33) | (min_side > 33) | (length_girth > 130))
        ]
    
    df['size_tier'] = np.select([x[1] for x in conditions],[x[0] for x in conditions],'Unidentified')
    return df

def get_storage_fee(df):
    this_month = today.month
    size_tier = df['size_tier']
    volume = (df['l'] * df['w'] * df['h']) / 1728
    fees_jan_sept = [
        (0.78 * volume, size_tier.isin([SMALL_STANDARD,LARGE_STANDARD])),
        (0.56 * volume, ~size_tier.isin([SMALL_STANDARD,LARGE_STANDARD]))
        ]
    fees_oct_dec = [
        (2.40 * volume, size_tier.isin([SMALL_STANDARD,LARGE_STANDARD])),
        (1.40 * volume, ~size_tier.isin([SMALL_STANDARD,LARGE_STANDARD]))
        ]
    
    fees_jan_sept_181_270 = [
        (1.56 * volume, size_tier.isin([SMALL_STANDARD,LARGE_STANDARD])),
        (1.02 * volume, ~size_tier.isin([SMALL_STANDARD,LARGE_STANDARD]))
        ]

    fees_jan_sept_271_365 = [
        (1.81 * volume, size_tier.isin([SMALL_STANDARD,LARGE_STANDARD])),
        (1.19 * volume, ~size_tier.isin([SMALL_STANDARD,LARGE_STANDARD]))
        ]
    
    fees_oct_dec_181_270 = [
        (3.09 * volume, size_tier.isin([SMALL_STANDARD,LARGE_STANDARD])),
        (1.86 * volume, ~size_tier.isin([SMALL_STANDARD,LARGE_STANDARD]))
        ]

    fees_oct_dec_271_365 = [
        (3.34 * volume, size_tier.isin([SMALL_STANDARD,LARGE_STANDARD])),
        (2.03 * volume, ~size_tier.isin([SMALL_STANDARD,LARGE_STANDARD]))
        ]

    if this_month in range(1,10):
        storage = fees_jan_sept
        storage_181_270 = fees_jan_sept_181_270
        storage_271_365 = fees_jan_sept_271_365

    elif this_month in range(10,13):
        storage = fees_oct_dec
        storage_181_270 = fees_oct_dec_181_270
        storage_271_365 = fees_oct_dec_271_365

        
    df['current_storage_fee'] = np.select([x[1] for x in storage],[x[0] for x in storage],np.nan).astype('float')
    df['storage_jan_sept'] = np.select([x[1] for x in fees_jan_sept],[x[0] for x in fees_jan_sept],np.nan).astype('float')
    df['storage_oct_dec'] = np.select([x[1] for x in fees_oct_dec],[x[0] for x in fees_oct_dec],np.nan).astype('float')
    df['avg_yearly_storage'] = ((df['storage_jan_sept'] * 9) + (df['storage_oct_dec'] * 3)) / 12

    df['current_storage_181-270'] = np.select([x[1] for x in storage_181_270],[x[0] for x in storage_181_270],np.nan).astype('float')
    df['current_storage_271-365'] = np.select([x[1] for x in storage_271_365],[x[0] for x in storage_271_365],np.nan).astype('float')


    df['jan_sept_storage_181-270'] = np.select([x[1] for x in fees_jan_sept_181_270],[x[0] for x in fees_jan_sept_181_270],np.nan).astype('float')
    df['jan_sept_storage_271-365'] = np.select([x[1] for x in fees_jan_sept_271_365],[x[0] for x in fees_jan_sept_271_365],np.nan).astype('float')

    df['oct_dec_storage_181-270'] = np.select([x[1] for x in fees_oct_dec_181_270],[x[0] for x in fees_oct_dec_181_270],np.nan).astype('float')
    df['oct_dec_storage_271-365'] = np.select([x[1] for x in fees_oct_dec_271_365],[x[0] for x in fees_oct_dec_271_365],np.nan).astype('float')
    return df
    
def get_fulfillment_fee(df):
    today_date = today.date()
    df.loc[df['size_tier'].isin([SMALL_STANDARD,EXTRA_LARGE_150_PLUS]),'shipping_weight'] = df['individual weight lbs']
    df['shipping_weight, oz'] = df['shipping_weight'] * 16
    size_tier = df['size_tier']
    weight = df['shipping_weight']
    weight_oz = df['shipping_weight, oz']
    
    fees_before_feb2024 = [ #obsolete
        (3.22, (size_tier == SMALL_STANDARD) & (weight_oz <= 4)),
        (3.40, (size_tier == SMALL_STANDARD) & (weight_oz.between(4, 8, inclusive = 'right'))),
        (3.58, (size_tier == SMALL_STANDARD) & (weight_oz.between(8, 12, inclusive = 'right'))),
        (3.77, (size_tier == SMALL_STANDARD) & (weight_oz.between(12, 16, inclusive = 'right'))),

        (3.86, (size_tier == LARGE_STANDARD) & (weight_oz <= 4)),
        (4.08, (size_tier == LARGE_STANDARD) & (weight_oz.between(4, 8, inclusive = 'right'))),
        (4.24, (size_tier == LARGE_STANDARD) & (weight_oz.between(8, 12, inclusive = 'right'))),
        (4.75, (size_tier == LARGE_STANDARD) & (weight_oz.between(12, 16, inclusive = 'right'))),
        (5.40, (size_tier == LARGE_STANDARD) & (weight.between(1, 1.5, inclusive = 'right'))),
        (5.69, (size_tier == LARGE_STANDARD) & (weight.between(1.5, 2, inclusive = 'right'))),
        (6.10, (size_tier == LARGE_STANDARD) & (weight.between(2, 2.5, inclusive = 'right'))),
        (6.39, (size_tier == LARGE_STANDARD) & (weight.between(2.5, 3, inclusive = 'right'))),
        (7.17 + 0.16*((weight-3)*2), (size_tier == LARGE_STANDARD) & (weight.between(3, 20, inclusive = 'right'))),
        
        (9.73 + 0.42*(weight-1), (size_tier == 'Small oversize') & (weight <= 70)),
        (19.05 + 0.42*(weight-1), (size_tier == 'Medium oversize') & (weight <= 150)),
        (89.98 + 0.83*(weight-90), (size_tier == 'Large oversize') & (weight <= 150)),
        (158.49 + 0.83*(weight-90), (size_tier == 'Special oversize') & (weight > 150))        
        ]
    fees_feb_apr_2024 = [ #obsolete
        (3.22, (size_tier == SMALL_STANDARD) & (weight_oz <= 2)),
        (3.31, (size_tier == SMALL_STANDARD) & (weight_oz.between(2, 4, inclusive = 'right'))),
        (3.40, (size_tier == SMALL_STANDARD) & (weight_oz.between(4, 6, inclusive = 'right'))),
        (3.49, (size_tier == SMALL_STANDARD) & (weight_oz.between(6, 8, inclusive = 'right'))),
        (3.58, (size_tier == SMALL_STANDARD) & (weight_oz.between(8, 10, inclusive = 'right'))),
        (3.68, (size_tier == SMALL_STANDARD) & (weight_oz.between(10, 12, inclusive = 'right'))),
        (3.77, (size_tier == SMALL_STANDARD) & (weight_oz.between(12, 14, inclusive = 'right'))),
        (3.82, (size_tier == SMALL_STANDARD) & (weight_oz.between(14, 16, inclusive = 'right'))),

        (3.86, (size_tier == LARGE_STANDARD) & (weight_oz <= 4)),
        (4.08, (size_tier == LARGE_STANDARD) & (weight_oz.between(4, 8, inclusive = 'right'))),
        (4.32, (size_tier == LARGE_STANDARD) & (weight_oz.between(8, 12, inclusive = 'right'))),
        (4.75, (size_tier == LARGE_STANDARD) & (weight_oz.between(12, 16, inclusive = 'right'))),
        (5.19, (size_tier == LARGE_STANDARD) & (weight.between(1, 1.25, inclusive = 'right'))),
        (5.57, (size_tier == LARGE_STANDARD) & (weight.between(1.25, 1.5, inclusive = 'right'))),
        (5.75, (size_tier == LARGE_STANDARD) & (weight.between(1.5, 1.75, inclusive = 'right'))),
        (6.00, (size_tier == LARGE_STANDARD) & (weight.between(1.75, 2, inclusive = 'right'))),
        (6.10, (size_tier == LARGE_STANDARD) & (weight.between(2, 2.25, inclusive = 'right'))),
        (6.28, (size_tier == LARGE_STANDARD) & (weight.between(2.25, 2.5, inclusive = 'right'))),
        (6.45, (size_tier == LARGE_STANDARD) & (weight.between(2.5, 2.75, inclusive = 'right'))),
        (6.86, (size_tier == LARGE_STANDARD) & (weight.between(2.75, 3, inclusive = 'right'))),
        (7.25 + 0.08*((weight-3)*4), (size_tier == LARGE_STANDARD) & (weight.between(3, 20, inclusive = 'right'))),
        
        (9.73 + 0.42*(weight-1), (size_tier == LARGE_BULKY) & (weight <= 50)),
        (26.33 + 0.38*(weight-1), (size_tier == EXTRA_LARGE_0_50) & (weight <= 50)),
        (40.12 + 0.75*(weight-51), (size_tier == EXTRA_LARGE_50_70) & (weight.between(50, 70, inclusive = 'right'))),
        (54.81 + 0.75*(weight-71), (size_tier == EXTRA_LARGE_70_150) & (weight.between(70, 150, inclusive = 'right'))),
        (194.95 + 0.19*(weight-151), (size_tier == EXTRA_LARGE_70_150) & (weight > 150))        
        ]
    fees_fba_non_peak = [
        (3.06, (size_tier == SMALL_STANDARD) & (weight_oz <= 2)),
        (3.15, (size_tier == SMALL_STANDARD) & (weight_oz.between(2, 4, inclusive = 'right'))),
        (3.24, (size_tier == SMALL_STANDARD) & (weight_oz.between(4, 6, inclusive = 'right'))),
        (3.33, (size_tier == SMALL_STANDARD) & (weight_oz.between(6, 8, inclusive = 'right'))),
        (3.43, (size_tier == SMALL_STANDARD) & (weight_oz.between(8, 10, inclusive = 'right'))),
        (3.53, (size_tier == SMALL_STANDARD) & (weight_oz.between(10, 12, inclusive = 'right'))),
        (3.60, (size_tier == SMALL_STANDARD) & (weight_oz.between(12, 14, inclusive = 'right'))),
        (3.65, (size_tier == SMALL_STANDARD) & (weight_oz.between(14, 16, inclusive = 'right'))),

        (3.68, (size_tier == LARGE_STANDARD) & (weight_oz <= 4)),
        (3.90, (size_tier == LARGE_STANDARD) & (weight_oz.between(4, 8, inclusive = 'right'))),
        (4.15, (size_tier == LARGE_STANDARD) & (weight_oz.between(8, 12, inclusive = 'right'))),
        (4.55, (size_tier == LARGE_STANDARD) & (weight_oz.between(12, 16, inclusive = 'right'))),
        (4.99, (size_tier == LARGE_STANDARD) & (weight.between(1, 1.25, inclusive = 'right'))),
        (5.37, (size_tier == LARGE_STANDARD) & (weight.between(1.25, 1.5, inclusive = 'right'))),
        (5.52, (size_tier == LARGE_STANDARD) & (weight.between(1.5, 1.75, inclusive = 'right'))),
        (5.77, (size_tier == LARGE_STANDARD) & (weight.between(1.75, 2, inclusive = 'right'))),
        (5.87, (size_tier == LARGE_STANDARD) & (weight.between(2, 2.25, inclusive = 'right'))),
        (6.05, (size_tier == LARGE_STANDARD) & (weight.between(2.25, 2.5, inclusive = 'right'))),
        (6.21, (size_tier == LARGE_STANDARD) & (weight.between(2.5, 2.75, inclusive = 'right'))),
        (6.62, (size_tier == LARGE_STANDARD) & (weight.between(2.75, 3, inclusive = 'right'))),
        (6.92 + 0.08*((weight-3)*4), (size_tier == LARGE_STANDARD) & (weight.between(3, 20, inclusive = 'right'))),
        
        (9.61 + 0.38*(weight-1), (size_tier == LARGE_BULKY) & (weight <= 50)),
        (26.33 + 0.38*(weight-1), (size_tier == EXTRA_LARGE_0_50) & (weight <= 50)),
        (40.12 + 0.75*(weight-51), (size_tier == EXTRA_LARGE_50_70) & (weight.between(50, 70, inclusive = 'right'))),
        (54.81 + 0.75*(weight-71), (size_tier == EXTRA_LARGE_70_150) & (weight.between(70, 150, inclusive = 'right'))),
        (194.95 + 0.19*(weight-151), (size_tier == EXTRA_LARGE_150_PLUS) & (weight > 150))        
        ]

    fees_fba_peak = [
        (3.25, (size_tier == SMALL_STANDARD) & (weight_oz <= 2)),
        (3.34, (size_tier == SMALL_STANDARD) & (weight_oz.between(2, 4, inclusive = 'right'))),
        (3.44, (size_tier == SMALL_STANDARD) & (weight_oz.between(4, 6, inclusive = 'right'))),
        (3.53, (size_tier == SMALL_STANDARD) & (weight_oz.between(6, 8, inclusive = 'right'))),
        (3.64, (size_tier == SMALL_STANDARD) & (weight_oz.between(8, 10, inclusive = 'right'))),
        (3.74, (size_tier == SMALL_STANDARD) & (weight_oz.between(10, 12, inclusive = 'right'))),
        (3.82, (size_tier == SMALL_STANDARD) & (weight_oz.between(12, 14, inclusive = 'right'))),
        (3.87, (size_tier == SMALL_STANDARD) & (weight_oz.between(14, 16, inclusive = 'right'))),

        (3.92, (size_tier == LARGE_STANDARD) & (weight_oz <= 4)),
        (4.16, (size_tier == LARGE_STANDARD) & (weight_oz.between(4, 8, inclusive = 'right'))),
        (4.43, (size_tier == LARGE_STANDARD) & (weight_oz.between(8, 12, inclusive = 'right'))),
        (4.84, (size_tier == LARGE_STANDARD) & (weight_oz.between(12, 16, inclusive = 'right'))),
        (5.29, (size_tier == LARGE_STANDARD) & (weight.between(1, 1.25, inclusive = 'right'))),
        (5.68, (size_tier == LARGE_STANDARD) & (weight.between(1.25, 1.5, inclusive = 'right'))),
        (5.84, (size_tier == LARGE_STANDARD) & (weight.between(1.5, 1.75, inclusive = 'right'))),
        (6.10, (size_tier == LARGE_STANDARD) & (weight.between(1.75, 2, inclusive = 'right'))),
        (6.24, (size_tier == LARGE_STANDARD) & (weight.between(2, 2.25, inclusive = 'right'))),
        (6.44, (size_tier == LARGE_STANDARD) & (weight.between(2.25, 2.5, inclusive = 'right'))),
        (6.61, (size_tier == LARGE_STANDARD) & (weight.between(2.5, 2.75, inclusive = 'right'))),
        (7.03, (size_tier == LARGE_STANDARD) & (weight.between(2.75, 3, inclusive = 'right'))),
        (7.46 + 0.08*((weight-3)*4), (size_tier == LARGE_STANDARD) & (weight.between(3, 20, inclusive = 'right'))),
        
        (10.65 + 0.38*(weight-1), (size_tier == LARGE_BULKY) & (weight <= 50)),
        (29.06 + 0.38*(weight-1), (size_tier == EXTRA_LARGE_0_50) & (weight <= 50)),
        (42.93 + 0.75*(weight-51), (size_tier == EXTRA_LARGE_50_70) & (weight.between(50, 70, inclusive = 'right'))),
        (59.23 + 0.75*(weight-71), (size_tier == EXTRA_LARGE_70_150) & (weight.between(70, 150, inclusive = 'right'))),
        (203.46 + 0.19*(weight-151), (size_tier == EXTRA_LARGE_150_PLUS) & (weight > 150))        
        ]

    peak_start, peak_end = (10, 15), (1, 14)
    if peak_start <= (today_date.month, today_date.day) or (today_date.month, today_date.day) <= peak_end:
        fees = fees_fba_peak
    else:
        fees = fees_fba_non_peak
    df['fba_fee'] = np.select([x[1] for x in fees],[x[0] for x in fees],np.nan)
    df['fba_peak_fee'] = np.select([x[1] for x in fees_fba_peak],[x[0] for x in fees_fba_peak],np.nan)
    return df

def get_removal_fee(df):
    size_tier = df['size_tier']
    weight = df['shipping_weight']
    removal_fee = [
        (1.04, (size_tier.isin([SMALL_STANDARD,LARGE_STANDARD])) & (weight.between(0,0.5, inclusive = 'right'))),
        (1.53, (size_tier.isin([SMALL_STANDARD,LARGE_STANDARD])) & (weight.between(0.5,1, inclusive = 'right'))),
        (2.27, (size_tier.isin([SMALL_STANDARD,LARGE_STANDARD])) & (weight.between(1,2, inclusive = 'right'))),
        (2.89 + 1.06 * (weight-2), (size_tier.isin([SMALL_STANDARD,LARGE_STANDARD])) & (weight > 2)),
        
        (3.12, (~size_tier.isin([SMALL_STANDARD,LARGE_STANDARD])) & (weight.between(0,1, inclusive = 'right'))),
        (4.30, (~size_tier.isin([SMALL_STANDARD,LARGE_STANDARD])) & (weight.between(1,2, inclusive = 'right'))),
        (6.36, (~size_tier.isin([SMALL_STANDARD,LARGE_STANDARD])) & (weight.between(2,4, inclusive = 'right'))),
        (10.04, (~size_tier.isin([SMALL_STANDARD,LARGE_STANDARD])) & (weight.between(4,10, inclusive = 'right'))),
        (14.32 + 1.06 * (weight-10), (~size_tier.isin([SMALL_STANDARD,LARGE_STANDARD])) & (weight > 10)),
        
        
        ]
    df['removal_fee'] = np.select([x[1] for x in removal_fee],[x[0] for x in removal_fee],np.nan)
    return df

def sipp_discount(df):
    df.loc[df['size_tier'].isin([SMALL_STANDARD,EXTRA_LARGE_150_PLUS]),'shipping_weight'] = df['individual weight lbs']
    df['shipping_weight, oz'] = df['shipping_weight'] * 16
    size_tier = df['size_tier']
    weight = df['shipping_weight']
    weight_oz = df['shipping_weight, oz']

    sipp_discount = [
        (0.04, (size_tier == SMALL_STANDARD) & (weight_oz <= 2)),
        (0.04, (size_tier == SMALL_STANDARD) & (weight_oz.between(2, 4, inclusive = 'right'))),
        (0.05, (size_tier == SMALL_STANDARD) & (weight_oz.between(4, 6, inclusive = 'right'))),
        (0.05, (size_tier == SMALL_STANDARD) & (weight_oz.between(6, 8, inclusive = 'right'))),
        (0.06, (size_tier == SMALL_STANDARD) & (weight_oz.between(8, 10, inclusive = 'right'))),
        (0.06, (size_tier == SMALL_STANDARD) & (weight_oz.between(10, 12, inclusive = 'right'))),
        (0.07, (size_tier == SMALL_STANDARD) & (weight_oz.between(12, 14, inclusive = 'right'))),
        (0.07, (size_tier == SMALL_STANDARD) & (weight_oz.between(14, 16, inclusive = 'right'))),

        (0.04, (size_tier == LARGE_STANDARD) & (weight_oz <= 4)),
        (0.04, (size_tier == LARGE_STANDARD) & (weight_oz.between(4, 8, inclusive = 'right'))),
        (0.07, (size_tier == LARGE_STANDARD) & (weight_oz.between(8, 12, inclusive = 'right'))),
        (0.08, (size_tier == LARGE_STANDARD) & (weight_oz.between(12, 16, inclusive = 'right'))),
        (0.09, (size_tier == LARGE_STANDARD) & (weight.between(1, 1.25, inclusive = 'right'))),
        (0.09, (size_tier == LARGE_STANDARD) & (weight.between(1.25, 1.5, inclusive = 'right'))),
        (0.10, (size_tier == LARGE_STANDARD) & (weight.between(1.5, 1.75, inclusive = 'right'))),
        (0.11, (size_tier == LARGE_STANDARD) & (weight.between(1.75, 2, inclusive = 'right'))),
        (0.12, (size_tier == LARGE_STANDARD) & (weight.between(2, 2.25, inclusive = 'right'))),
        (0.13, (size_tier == LARGE_STANDARD) & (weight.between(2.25, 2.5, inclusive = 'right'))),
        (0.14, (size_tier == LARGE_STANDARD) & (weight.between(2.5, 2.75, inclusive = 'right'))),
        (0.14, (size_tier == LARGE_STANDARD) & (weight.between(2.75, 3, inclusive = 'right'))),
        (0.23, (size_tier == LARGE_STANDARD) & (weight.between(3, 20, inclusive = 'right'))),
        
        (1.32, (size_tier == LARGE_BULKY) & (weight <= 50)),
        (0, (size_tier == EXTRA_LARGE_0_50) & (weight <= 50)),
        (0, (size_tier == EXTRA_LARGE_50_70) & (weight.between(50, 70, inclusive = 'right'))),
        (0, (size_tier == EXTRA_LARGE_70_150) & (weight.between(70, 150, inclusive = 'right'))),
        (0, (size_tier == EXTRA_LARGE_150_PLUS) & (weight > 150))        
        ]

    df['sipp_discount'] = np.select([x[1] for x in sipp_discount],[x[0] for x in sipp_discount],np.nan)
    return df


def export_to_excel(df):
    try:
        with pd.ExcelWriter(os.path.join(user_folder,'fees.xlsx'), engine = 'xlsxwriter') as writer:
            df.to_excel(writer, sheet_name = 'Fees', index = False)
            mm.format_header(df, writer, 'Fees')
    except PermissionError:
        PopupError("Please close the file first")
        export_to_excel(df)
    mm.open_file_folder(user_folder)

def separate_file():
    file_path = filedialog.askopenfilename(
        title="Select file with dimensions",
        filetypes=[("Excel files","*.xls*"),])
    df = pd.read_excel(file_path)
    df = get_shipping_weight(df)
    df = get_size_tier(df)
    df = get_fulfillment_fee(df)
    df = get_removal_fee(df)    
    df = get_storage_fee(df)
    df = sipp_discount(df)
    export_to_excel(df)


def main(out = True):
    dimensions = get_dims_file()
    dictionary = get_dictionary()
    combined = combine_files(dimensions.copy(), dictionary.copy())
    combined = get_shipping_weight(combined)
    combined = get_size_tier(combined)
    combined = get_fulfillment_fee(combined)
    combined = get_removal_fee(combined)    
    combined = get_storage_fee(combined)
    for nc in num_cols:
        combined[nc] = combined[nc].astype(float, errors = 'ignore')
    if out == False:
        return combined
    export_to_excel(combined)
    

    dims_cloud = dimensions.copy().fillna(0)
    dims_cloud = gc.normalize_columns(dims_cloud)
    gc.push_to_cloud(dims_cloud, destination= 'auxillary_development.dimensions', if_exists = 'replace')
    del dims_cloud
    

if __name__ == '__main__':
    try:
        mode = int(input("Select mode: 1 for Mellanni dimensions, or 2 for file upload"))
        if mode==1:
            main(out = True)
        elif mode==2:
            separate_file()
    except Exception as e:
        PopupError(e)