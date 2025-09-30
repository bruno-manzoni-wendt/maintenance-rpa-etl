# %%

import pandas as pd
import os
import checklist_utils

link = r'https://example_url/predictive_scheduling/'
maintenance_path = r'maintenance\predictive'
excel = 'predictive_scheduling.xlsx'
excel_path = os.path.join(maintenance_path, excel)

# %%
checklist_utils.open_link(link)

checklist_utils.export(excel)

checklist_utils.generate_export(__file__, 'scheduling_success.png')

checklist_utils.save_as_xlsx(excel, maintenance_path)

# %%
df_main = pd.read_excel(excel_path)
df_history = pd.read_excel(os.path.join(maintenance_path, 'predictive_scheduling_06_2025.xlsx'))
df = pd.concat([df_main, df_history], ignore_index=True)

with pd.ExcelWriter(os.path.join(maintenance_path, 'predictive_scheduling_python.xlsx')) as writer:
    df.to_excel(writer, sheet_name='Worksheet', index=False)

