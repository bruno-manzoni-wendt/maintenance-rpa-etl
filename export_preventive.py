# %%
import numpy as np
import pandas as pd
import os
import checklist_utils

link = r'https://example_url/preventive/'
maintenance_path = r'maintenance\preventive'
excel = 'preventive_export.xlsx'
excel_path = os.path.join(maintenance_path, excel)

# %%
checklist_utils.open_link(link)

checklist_utils.export(excel)

checklist_utils.generate_export(__file__, 'item_success.png')

checklist_utils.save_as_xlsx(excel, maintenance_path)

# %%
print(f'Reading {excel} and performing treatments...')
df_current = pd.read_excel(excel_path)
df_history = pd.read_excel(os.path.join(maintenance_path, 'preventive_06_2025.xlsx'))
df = pd.concat([df_current, df_history], ignore_index=True)
df = df.sort_values(by='Evaluation code', ascending=False)

df['Checklist name'] = df['Checklist name'].str.split('Preventive ').str[-1] #TODO this might fail when there's air conditioning

df['Unit or Line'] = df['Checklist name'].copy()
df['Unit or Line'] = np.where(df['Unit or Line'].str.contains('Dehumidifier', case=False, na=False), np.nan, df['Unit or Line'])
df['Unit or Line'] = np.where(df['Unit or Line'].str.contains('Filter', case=False, na=False), np.nan, df['Unit or Line'])
df['Unit or Line'] = df['Unit or Line'].str.split(r'(').str[1].str.replace(r')', '').str.strip()

df['Line'] = df['Unit or Line'].str.split(' ').str[1]

df['Unit'] = df['Unit or Line'].str.split(' ').str[0]
df.loc[df['Unit'].str.startswith('Line', na=False), 'Unit'] = np.nan

df = df.drop(columns=['Unit or Line'])

df = df.rename(columns={'End date': 'DateTime completion', 'Start date': 'DateTime start'})
for col in ['start', 'completion']:
    df[[f'Date {col}', f'Time {col}']] = df[f'DateTime {col}'].str.split(' ').apply(pd.Series) # separate date from time
    df[f'Date {col}'] = pd.to_datetime(df[f'Date {col}'], format='%d/%m/%Y')
    df[f'Time {col}'] = pd.to_datetime(df[f'Time {col}'], format='%H:%M:%S').dt.time
    df[f'DateTime {col}'] = pd.to_datetime(df[f'DateTime {col}'], format='%d/%m/%Y %H:%M:%S')

df = df.sort_values(by='DateTime start', ascending=False)

#===================================================

df_unique = df.drop_duplicates('Evaluation code').copy()
df_unique = df_unique[['Evaluation code', 'Unit', 'Line', 'Checklist name', 'Author',
                    'DateTime start', 'Date start', 'Time start',
                    'DateTime completion', 'Date completion', 'Time completion']]
df_unique = df_unique.reset_index(drop=True)


df_image = df[['Evaluation code', 'Item', 'Response', 'Images', 'Item comment']].copy()
df_image['Images'] = df_image['Images'].str.split(' ')

for i in df_image.index:
    links = df_image.at[i, 'Images']
    if isinstance(links, list):
        for x, link in enumerate(links):
            df_image.at[i, f'Image.{x+1}'] = link
df_image = df_image.drop(columns=['Images'])


with pd.ExcelWriter(os.path.join(maintenance_path, 'preventive_python.xlsx')) as writer:
    df_unique.to_excel(writer, sheet_name='items', index=False)
    df_image.to_excel(writer, sheet_name='images', index=False)
print('- Saved and finished! -\n')