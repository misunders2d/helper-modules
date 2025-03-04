import os
from typing import List, Literal
import pandas as pd
import re
import sys, subprocess
from connectors import gdrive as gd
import customtkinter
from io import BytesIO
import webbrowser

from ctk_gui.ctk_windows import PopupError, PopupYesNo

def open_file_folder(path: str) -> None:
    try:
        os.startfile(path)
    except AttributeError:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, path])
    except Exception as e:
        print(f'Uncaught exception occurred: {e}')
    return None

def push_dicts(threaded=False, *args, **kwargs):
    from connectors import gcloud as gc
    import threading
    drive_id = '0AMdx9NlXacARUk9PVA'
    
    usa_path = gd.find_file_id(folder_id='1zIHmbWcRRVyCTtuB9Atzam7IhAs8Ymx4', drive_id=drive_id, filename='Dictionary.xlsx')
    ca_path = gd.find_file_id(folder_id='1ZijSZTqY1_5F307uMkdcneqTKIoNSsds', drive_id=drive_id, filename='Dictionary_CA.xlsx')
    eu_path = gd.find_file_id(folder_id='1uye8_FNxI11ZUOKnUYUfko1vqwpJVnMj', drive_id=drive_id, filename='Dictionary_EU.xlsx')
    uk_path = gd.find_file_id(folder_id='1vt8UB2FeQp0RJimnCATI8OQt5N-bysx-', drive_id=drive_id, filename='Dictionary_UK.xlsx')
    
    
    paths = (usa_path,ca_path,eu_path,uk_path)
    dicts = {}
    
    def read_dictionary(dicts, path, name):

        market_dict = pd.read_excel(gd.download_file(path))
        market_dict = gc.normalize_columns(market_dict)
        dicts[name] = market_dict
        return None
    if threaded:
        threads = [
            threading.Thread(target = read_dictionary, args = (dicts, paths[0], 'US')),
            threading.Thread(target = read_dictionary, args = (dicts, paths[1], 'CA')),
            threading.Thread(target = read_dictionary, args = (dicts, paths[2], 'EU')),
            threading.Thread(target = read_dictionary, args = (dicts, paths[3], 'UK'))
            ]
        
        _ = [thread.start() for thread in threads]
        _ = [thread.join() for thread in threads]
    else:
        for path, name in zip(paths,['US','CA','EU','UK']):
            read_dictionary(dicts, path, name)
    
    
    dictionary_us = dicts.get('US')
    dictionary_ca = dicts.get('CA')
    dictionary_eu = dicts.get('EU')
    dictionary_uk = dicts.get('UK')
    
    def push_dictionary(dictionary, table):
        dictionary['pattern'] = dictionary['pattern'].astype('str')
        for col in dictionary.columns:
            dictionary[col] = dictionary[col].astype(str)
            dictionary[col] = dictionary[col].replace({r'[^\x00-\x7F]+':''}, regex=True)
        try:
            gc.push_to_cloud(dictionary, destination=table, if_exists='replace')
           
        except Exception as e:
            if 'window' in kwargs:
                window = kwargs['window']
                window.write_event_value('DICTS_ERROR',e)
            else:
                print(f'The following error occurred:\n{e}')
        return None
    
    if threaded:
        threads = [
            threading.Thread(target = push_dictionary, args = (dictionary_us,'auxillary_development.dictionary')),
            threading.Thread(target = push_dictionary, args = (dictionary_ca,'auxillary_development.dictionary_ca')),
            threading.Thread(target = push_dictionary, args = (dictionary_eu,'auxillary_development.dictionary_eu')),
            threading.Thread(target = push_dictionary, args = (dictionary_uk,'auxillary_development.dictionary_uk'))
            ]

        _ = [thread.start() for thread in threads]
        _ = [thread.join() for thread in threads]
    else:
        push_dictionary(dictionary_us,'auxillary_development.dictionary')
        push_dictionary(dictionary_ca,'auxillary_development.dictionary_ca')
        push_dictionary(dictionary_eu,'auxillary_development.dictionary_eu')
        push_dictionary(dictionary_uk,'auxillary_development.dictionary_uk')
    print('All done')
    return None

def normalize_columns(df):
    # import pandas as pd
    pattern = r'^([0-9].)'
    new_cols = [x.strip()
                .replace(' ','_')
                .replace('-','_')
                .replace('?','')
                .replace(',','')
                .replace('.','')
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

def export_to_excel(
        dfs: List[pd.DataFrame],
        sheet_names: List[str],
        filename: str = 'test.xlsx',
        out_folder: Literal[str,None] = None
        ) -> None:
    from customtkinter import filedialog
    
    if not out_folder:
        out_folder = filedialog.askdirectory(title='Select output folder', initialdir=os.path.expanduser('~'))
    full_output = os.path.join(out_folder,filename)
    try:
        with pd.ExcelWriter(full_output, engine = 'xlsxwriter') as writer:
            for df, sheet_name in list(zip(dfs,sheet_names)):
                df.to_excel(excel_writer = writer, sheet_name = sheet_name, index = False)
                format_header(df, writer, sheet_name)
    except PermissionError:
        print(f'{filename} is open, please close the file first')
        export_to_excel(dfs, sheet_names, filename, out_folder)
    except Exception as e:
        print(e)
        
    return None

def get_comments():
    from openpyxl import load_workbook
    import os
    
    file_path = customtkinter.filedialog.askopenfilename(title='Select the file processing report')
    
    if file_path and file_path != '':
        try:
            wb = load_workbook(file_path, data_only = True)
            ws = wb["Template"] # or whatever sheet name
            ws.delete_rows(0)
            ws.delete_rows(0)
            
            
            comments = []
            for row in ws.rows:
                comments.append(row[2].comment)
            comments = comments[1:]
            
            
            
            file = pd.read_excel(file_path, sheet_name = 'Template', skiprows =2)
            cols = file.columns.tolist()
            cols.insert(4,'comments')
            
            file['comments'] = comments
            file = file[cols]
            
            output = customtkinter.filedialog.askdirecotry(title='Select output folder')
            with pd.ExcelWriter(os.path.join(output,'comments.xlsx')) as writer:
                file.to_excel(writer,sheet_name = 'Comments', index = False)
                format_header(file, writer, 'Comments')
            os.startfile(output)
        except Exception as e:
            PopupError(f'Sorry, error occurred:\n{e}')
    else:
        PopupError('Please select a file')
    return None

def sku_creation():
    '''
    Quick function to rename sizes, capitalize and remove spaces from colors
    for new SKU creation
    Accepts: nothing
    Returns
    -------
    df : TYPE
        DESCRIPTION.

    '''
    colors = input('Colors?\n').split('\n')
    sizes = input('Sizes?\n').split('\n')
    
    size_renaming = {
        'Twin XL':'TXL',
        'Twin':'T',
        'Full':'F',
        'Queen':'Q',
        'Split King':'SK',
        'Cal King':'CK',
        'King':'K'
        }
    
    prefix = input('prefix?')
    df = pd.DataFrame(list(zip(sizes,colors)), columns = ['Size','Color'])
    df['Prefix'] = prefix
    df['SKU_Size'] = df['Size'].replace(size_renaming)
    df['SKU_Color'] = df['Color'].str.strip().str.replace(' ','-').str.upper()
    df['Result'] = df['Prefix']+'-'+df['SKU_Size']+'-'+df['SKU_Color']
    df['len'] = df['Result'].apply(lambda x: len(x))
    df['Warning'] = df['len'].apply(lambda x: 'Warning' if x > 40 else '')
    return df

def kw_ranking_weighted(positions = None, sales = None):
    '''
    Generate a weighted position (by sales) for our KW ranking standings

    Returns
    -------
    (int) Number of our position.

    '''
    import numpy as np
    if all([positions is None, sales is None]):
        positions = input('Input list of positions\n').split('\n')
        sales = input('Input keyword sales\n').split('\n')
    positions = [int(x.replace('>','').replace('#Н/Д','306')) for x in positions]
    sales = [int(x.replace('-','0').replace(',','').replace('#Н/Д','0')) for x in sales]
    if len(positions) == len(sales):
        result = int(np.dot(sales,positions) / sum(sales))
        return result
    else:
        return 0


def bash_quote(dump = False,lang = 'ru'):
    wpath = get_db_path('US')[9]
    link = "http://bash.org/?random" if lang == 'en' else "http://bashorg.org/random"
    search = 'qt' if lang == 'en' else 'quote'
    bash_file = os.path.join(wpath,'bash') if lang == 'en' else os.path.join(wpath,'bash_ru')
    import random
    import pickle
    def get_jokes():
        from bs4 import BeautifulSoup as bs
        import requests
        page = requests.get(link)
        soup = bs(page.content, features = 'lxml')
        jokes = soup.find_all(class_ = search)
        jokes = [joke.get_text('<br/>').replace('<br/>','\n') for joke in jokes] if lang == 'ru' else [joke.get_text() for joke in jokes]
        return jokes
    def read_bash():
        with open(bash_file,'rb') as f:
            check = pickle.load(f)
        return check
    def write_bash(text):
        with open(bash_file,'wb') as f:
            pickle.dump(text,f)
        return None
        
    if os.path.isfile(bash_file):
        if dump == False:
            with open(bash_file,'rb') as f:
                jokes = pickle.load(f)
            joke = jokes[random.randint(0,len(jokes)-1)]
            return joke
        if dump == True:
            jokes = get_jokes()
            check = read_bash()
            for j in jokes:
                if j not in check:
                    check.append(j)
            write_bash(check)
            joke = check[random.randint(0,len(check)-1)]
            return joke
    elif not os.path.isfile(bash_file):
        jokes = get_jokes()
        with open(bash_file,'wb') as f:
            pickle.dump(jokes,f)
        joke = jokes[random.randint(0,len(jokes)-1)]
        return joke

def cancelled_shipments(account = 'US'):
    paths = get_db_path('US')
    path = os.path.join(paths[7],'Cancelled shipments')
    d_path = get_db_path(account)[1]
    d = pd.read_excel(d_path, usecols = ['ASIN','SKU'])
    file = get_file_paths([path])[-1]
    print(f'Latest cancelled shipments file is: {os.path.basename(file)}')
    cancelled = pd.read_excel(file)
    cancelled = pd.merge(cancelled, d, how = 'left', left_on = 'Sku', right_on = 'SKU').dropna(subset = ['ASIN'])
    # if account == 'CA':
    #     cancelled = cancelled[cancelled['Comments'] == 'Canada']
    cancelled = cancelled.pivot_table(
        values = 'Units to Cancel',
        index = 'ASIN',
        aggfunc = 'sum'
        ).reset_index()
    return cancelled

def password_generator(x):
    '''
    Generates a password of 'x' lenght from letters, digits and punctuation marks

    Parameters
    ----------
    x : int
        number of symbols to use.

    Returns
    str
    password
    '''
    import string
    import random
    text_part = x//3*2
    num_part = x-text_part
    text_str = [random.choice(string.ascii_letters) for x in range(text_part)]
    num_str = [random.choice(string.digits*2+string.punctuation) for x in range(num_part)]
    text = text_str+num_str
    random.shuffle(text)
    password = ''.join(text)
    return password

def convert_to_pacific(db,columns):
    import pytz
    pacific = pytz.timezone('US/Pacific')
    db['pacific-date'] = pd.to_datetime(db[columns]).dt.tz_convert(pacific)
    db['pacific-date'] = pd.to_datetime(db['pacific-date']).dt.tz_localize(None)
    return db['pacific-date']

def format_header(df,writer,sheet):
    workbook  = writer.book
    cell_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'center', 'font_size':9})
    worksheet = writer.sheets[sheet]
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, cell_format)
    max_row, max_col = df.shape
    worksheet.autofilter(0, 0, max_row, max_col - 1)
    worksheet.freeze_panes(1,0)
    return None

def format_columns(df,writer,sheet,col_num):
    worksheet = writer.sheets[sheet]
    if not isinstance(col_num,list):
        col_num = [col_num]
    else:
        pass
    for c in col_num:
        width = max(df.iloc[:,c].astype(str).map(len).max(),len(df.iloc[:,c].name))
        worksheet.set_column(c,c,width)
    return None

def encrypt_string(hash_string):
    '''
    Create a hashed string from any input
    '''
    import hashlib
    string_encoded = hashlib.sha256(str(hash_string).encode()).hexdigest()
    return string_encoded    

