# %%
import numpy as np
import pandas as pd
import os
import checklist_utils

link = r'https://example_url/expenses/'
maintenance_path = r'\maintenance\expenses'
excel = 'expenses_export.xlsx'
excel_path = os.path.join(maintenance_path, excel)

# %%
checklist_utils.open_link(link)

checklist_utils.export(excel)

checklist_utils.generate_export(__file__, 'item_success.png')

checklist_utils.save_as_xlsx(excel, maintenance_path)

# %%
print(f'Reading {excel} and performing treatments...')
df_current = pd.read_excel(excel_path)
df_history = pd.read_excel(os.path.join(maintenance_path, 'expenses_06_2025.xlsx'))
df = pd.concat([df_current, df_history], ignore_index=True)

df = df.sort_values(by='Evaluation code', ascending=False)

dict_list = []
for code in df['Evaluation code'].unique():
    df_code = df[df['Evaluation code'] == code].copy() # creates df of unique code
    row_dict = dict(df_code.iloc[0])
    for key in ['Item', 'Response', 'Images', 'Item comment']:
        row_dict.pop(key, None)

    for i in df_code.index:
        item, response = df_code.at[i, 'Item'], df_code.at[i, 'Response']
        row_dict[item] = response

    img_list = '-'.join((df_code['Images'].fillna(''))).split(' ')
    for x, img in enumerate(img_list):
        row_dict[f'Image.{x+1}'] = img.replace('-', '')

    dict_list.append(row_dict)


df_unique = pd.DataFrame(dict_list)
df_unique['What amount was spent ?'] = df_unique['What amount was spent ?'].str.replace('R$', '').str.replace('.', '').str.replace(',', '.').astype(float)

df_unique = df_unique.rename(columns={'End date': 'DateTime completion', 'Start date': 'DateTime start'})
for col in ['start', 'completion']:
    df_unique[[f'Date {col}', f'Time {col}']] = df_unique[f'DateTime {col}'].str.split(' ').apply(pd.Series) # separate date from time
    df_unique[f'Date {col}'] = pd.to_datetime(df_unique[f'Date {col}'], format='%d/%m/%Y')
    df_unique[f'Time {col}'] = pd.to_datetime(df_unique[f'Time {col}'], format='%H:%M:%S').dt.time
    df_unique[f'DateTime {col}'] = pd.to_datetime(df_unique[f'DateTime {col}'], format='%d/%m/%Y %H:%M:%S')


df_unique['Checklist name'] = df_unique['Checklist name'].str.replace('( Expense )', '').str.strip()
df_unique['Unit'], df_unique['Line'] = None, np.nan
unit_dict = {'B': 'Bottling', 'L': 'Labeling', 'W': 'Weighing'}
for i in df_unique.index:
    if ' - ' in df_unique.at[i, 'Checklist name']:
        unit_line = df_unique.at[i, 'Checklist name'].split(' - ')[0]
        if len(unit_line) > 1 and unit_line[1].isdigit():
            df_unique.at[i, 'Unit'] = unit_dict[unit_line[0]]
            df_unique.at[i, 'Line'] = int(unit_line[1])

with pd.ExcelWriter(os.path.join(maintenance_path, 'expenses_python_v2.xlsx')) as writer:
    df_unique.to_excel(writer, sheet_name='items', index=False)

print('- Saved and finished! -\n')
