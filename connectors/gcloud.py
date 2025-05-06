from google.cloud import bigquery
from google.cloud import storage
import os
import pandas as pd
import pandas_gbq
from utils import mellanni_modules as mm

def gcloud_connect():
    key_path = 'credentials/gcloud.json'
    os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = key_path
    
    client = bigquery.Client()
    return client

def cgk_pricing():
    with gcloud_connect() as client:
        query = '''SELECT datetime, asin, brand, full_price, coupon, ld, final_price
                    from `auxillary_development.price_comparison`
                    WHERE asin = "B01M16WBW1"'''
        prices = client.query(query).result().to_dataframe()
    prices['datetime'] = pd.to_datetime(prices['datetime'])
    prices['year'] = prices['datetime'].dt.year
    prices['week'] = (prices['datetime']+pd.DateOffset(days=1)).dt.isocalendar().week
    prices = prices.sort_values('datetime')
    with pd.ExcelWriter(os.path.join(os.path.expanduser('~'),'temp/cgk_pricing.xlsx'), engine = 'xlsxwriter') as writer:
        prices.to_excel(writer, sheet_name = 'cgk', index = False)
        mm.format_header(prices, writer, 'cgk')
    return None

def pull_raw(dataset = 'auxillary_development',
                report = 'dictionary',
                custom_query = None) -> list:
    if not custom_query:
        query = f'SELECT * FROM `{dataset}.{report}` LIMIT 10'
    else:
        query = custom_query
    with gcloud_connect() as client:
        data = client.query(query).result()
    return data

def pull_gcloud(dataset = 'auxillary_development',
                report = 'dictionary',
                custom_query = None) -> pd.DataFrame:
    if not custom_query:
        query = f'SELECT * FROM `{dataset}.{report}` LIMIT 10'
    else:
        query = custom_query
    with gcloud_connect() as client:
        data = client.query(query).result().to_dataframe()
    return data
    
def get_datasets() -> list:
    with gcloud_connect() as client:
        return [x.dataset_id for x in client.list_datasets()]

def get_tables(dataset) -> list:
    with gcloud_connect() as client:
        return [x.table_id for x in client.list_tables(dataset)]

def create_storage_bucket(bucket_name):
    storage_client = storage.Client()
    storage_client.create_bucket(bucket_name)
    return None

def normalize_columns(df):
    import re
    pattern = '^([0-9].)'
    new_cols = [x.strip()
                .replace(' ','_')
                .replace('-','_')
                .replace('?','')
                .replace(',','')
                .replace('.','')
                .replace('/','_')
                .replace('(','')
                .replace(')','')
                .lower()
                for x in df.columns]
    new_cols = [re.sub(pattern, '_'+re.findall(pattern,x)[0], x) if re.findall(pattern,x) else x for x in new_cols]
    df.columns = new_cols
    date_cols = [x for x in df.columns if 'date' in x.lower()]
    if date_cols !=[]:
        df[date_cols] = df[date_cols].astype('str')
        df = df.sort_values(date_cols, ascending = True)
    float_cols = [x for x in df.select_dtypes('float64').columns]
    int_cols = [x for x in df.select_dtypes('int64').columns]
    df[float_cols] = df[float_cols].astype('float32')
    df[int_cols] = df[int_cols].astype('int32')
    return df

def push_to_cloud(df: pd.DataFrame, destination: str, if_exists: str = 'append') -> None:
    client = gcloud_connect()
    df = normalize_columns(df)
    _ = pandas_gbq.to_gbq(df, destination_table=destination, if_exists=if_exists)
    return None
