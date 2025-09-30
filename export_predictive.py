# %%
import numpy as np
import pandas as pd
import os
import checklist_utils

link = r'https://example_url/predictive/'
maintenance_path = r'G:\Meu Drive\Bruno\Projetos\Manutenção\predictive'
excel = 'predictive_export.xlsx'
excel_path = os.path.join(maintenance_path, excel)

# %%
checklist_utils.open_link(link)

checklist_utils.export(excel)

checklist_utils.generate_export(__file__, 'item_success.png')

checklist_utils.save_as_xlsx(excel, maintenance_path)

#%%
print(f"Reading '{excel}' and performing treatments...")
df_current = pd.read_excel(excel_path)
df_history = pd.read_excel(os.path.join(maintenance_path, 'predictive_06_2025.xlsx'))
df = pd.concat([df_current, df_history], ignore_index=True)
df = df.sort_values(by='Evaluation code', ascending=False)

# Unit and Line
df['Unit type'] = df['Unit type'].str.replace('Predictive', '').str.strip()

df['Checklist name'] = df['Checklist name'].str.split(' Predictive ').str[1].str.replace(' ( Scheduled )', '').str.replace('( ', '').str.replace(' )', '')
df['Unit'], df['Line'] = None, np.nan
for i, row in df.iterrows():
    if row['Item'] == 'On which line will the check-list be performed ?' and 'Line' in row['Response']: # new predictive
        unit, line = row['Response'].split(' ')[0], int(row['Response'].split(' ')[-1])
        same_codes = df['Evaluation code'] == row['Evaluation code']
        df.loc[same_codes, 'Unit'] = unit
        df.loc[same_codes, 'Line'] = line

    try: # old predictive
        df.at[i, 'Line'] = int(row['Checklist name'][-2:])
        df.at[i, 'Unit'] = row['Checklist name'][:-3]
    except (ValueError, TypeError):
        pass

# unique df
dict_list = []
for code, df_code in df.groupby('Evaluation code'):
    row_dict = dict(df_code.iloc[0])
    for key in ['Item', 'Response', 'Images', 'Item comment']:
        row_dict.pop(key, None)
    dict_list.append(row_dict)

df_unique = pd.DataFrame(dict_list)

# DateTime start and completion
df_unique = df_unique.rename(columns={'End date': 'DateTime completion', 'Start date': 'DateTime start'})
for col in ['start', 'completion']:
    df_unique[[f'Date {col}', f'Time {col}']] = df_unique[f'DateTime {col}'].str.split(' ').apply(pd.Series) # separate date from time
    df_unique[f'Date {col}'] = pd.to_datetime(df_unique[f'Date {col}'], format='%d/%m/%Y')
    df_unique[f'Time {col}'] = pd.to_datetime(df_unique[f'Time {col}'], format='%H:%M:%S').dt.time
    df_unique[f'DateTime {col}'] = pd.to_datetime(df_unique[f'DateTime {col}'], format='%d/%m/%Y %H:%M:%S')

# Images
df_image = df[['Evaluation code', 'Item', 'Response', 'Item comment', 'Images']].copy()
df_image['Images'] = df_image['Images'].str.split(' ')

imgs_dict = df_image['Images'].to_dict()
for i, links in imgs_dict.items():
    if isinstance(links, list):
        for x, link in enumerate(links):
            df_image.loc[i, f'Image.{x+1}'] = link
df_image = df_image.drop(columns=['Images'])


with pd.ExcelWriter(os.path.join(maintenance_path, 'predictive_python.xlsx')) as writer:
    df_unique.to_excel(writer, sheet_name='items', index=False)
    df_image.to_excel(writer, sheet_name='images', index=False)
print('- Saved and finished! -\n')
