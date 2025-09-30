# %%

import pandas as pd
import os
import checklist_utils

link = r'https://example_url/support/'
maintenance_path = r'maintenance\support'
excel = 'support_export.xlsx'
excel_path = os.path.join(maintenance_path, excel)

# %%
checklist_utils.open_link(link)

checklist_utils.export(excel)

checklist_utils.generate_export(__file__, 'item_success.png')

checklist_utils.save_as_xlsx(excel, maintenance_path)

# %%
print(f'Reading {excel} and performing treatments...')
df = pd.read_excel(excel_path)
df = df[['Evaluation code', 'Author', 'Item', 'Response', 'Start date']]
df = df.sort_values(by='Evaluation code', ascending=False)

#===================================================
dict_list = []
for code in df['Evaluation code'].unique():
    df_code = df[df['Evaluation code'] == code] # creates df of unique code
    row_dict = dict(df_code.iloc[0])

    for key in ['Item', 'Response']:
        row_dict.pop(key, None)

    for i in df_code.index:
        item, response = df_code.at[i, 'Item'], df_code.at[i, 'Response']
        row_dict[item] = response

    dict_list.append(row_dict)

df_unique = pd.DataFrame(dict_list)

df_unique.rename(columns={'Service Code.': 'Support Code'}, inplace=True)
df_unique['Support Code'] = df_unique['Support Code'].replace('#', '', regex=True)
df_unique['Support Code'] = pd.to_numeric(df_unique['Support Code'], errors='coerce')

df_unique.dropna(subset=['Support Code'], inplace=True)

df_unique['Support Code'] = df_unique['Support Code'].astype('int64')

with pd.ExcelWriter(os.path.join(maintenance_path, 'support_python.xlsx')) as writer:
    df_unique.to_excel(writer, sheet_name='items', index=False)