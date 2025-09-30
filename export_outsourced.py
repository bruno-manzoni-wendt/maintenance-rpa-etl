# %%

import pandas as pd
import os
import checklist_utils

link = r'https://example_url/outsourced/'
maintenance_path = r'maintenance\outsourced'
excel = 'outsourced_export.xlsx'
excel_path = os.path.join(maintenance_path, excel)

# %%
checklist_utils.open_link(link)

checklist_utils.export(excel)

checklist_utils.generate_export(__file__, 'item_success.png')

checklist_utils.save_as_xlsx(excel, maintenance_path)

#%%
print(f'Reading {excel} and performing treatments...')
df = pd.read_excel(excel_path)
df = df.sort_values(by='Evaluation code', ascending=False)

#===================================================
dict_list = []
for code in df['Evaluation code'].unique():
    df_code = df[df['Evaluation code'] == code] # creates df of unique code
    row_dict = dict(df_code.iloc[0])

    for key in ['Item', 'Response', 'Images', 'Item comment']:
        row_dict.pop(key, None)

    for i in df_code.index:
        item, response = df_code.at[i, 'Item'], df_code.at[i, 'Response']
        row_dict[item] = response

    dict_list.append(row_dict)

df_unique = pd.DataFrame(dict_list)
df_unique = df_unique.rename(columns={'End date': 'DateTime completion', 'Start date': 'DateTime start'})
for col in ['start', 'completion']:
    df_unique[[f'Date {col}', f'Time {col}']] = df_unique[f'DateTime {col}'].str.split(' ').apply(pd.Series) # separate date from time
    df_unique[f'Date {col}'] = pd.to_datetime(df_unique[f'Date {col}'], format='%d/%m/%Y')
    df_unique[f'Time {col}'] = pd.to_datetime(df_unique[f'Time {col}'], format='%H:%M:%S').dt.time
    df_unique[f'DateTime {col}'] = pd.to_datetime(df_unique[f'DateTime {col}'], format='%d/%m/%Y %H:%M:%S')

# #===================================================
df_image = df[['Evaluation code', 'Item', 'Response', 'Item comment', 'Images']].copy()
df_image['Images'] = df_image['Images'].str.split(' ')
df_image = df_image.reset_index(drop=True)

for i in df_image.index:
    links = df_image.at[i, 'Images']
    if isinstance(links, list):
        for x, link in enumerate(links):
            df_image.at[i, f'Image.{x+1}'] = link
df_image = df_image.drop(columns=['Images'])

with pd.ExcelWriter(os.path.join(maintenance_path, 'outsourced_python.xlsx')) as writer:
    df_unique.to_excel(writer, sheet_name='items', index=False)
    df_image.to_excel(writer, sheet_name='images', index=False)
