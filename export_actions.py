# %%
import os
import checklist_utils

import sys
sys.path.append(r'\Python\LIB')
import EFX_lib as efx

link = r'https://example_url/actions/'
maintenance_path = r'\maintenance\actions'
excel = 'export_action.xlsx'
excel_path = os.path.join(maintenance_path, excel)

# %%
checklist_utils.open_link(link)

efx.procurar('filter_action.png', conf=0.95)
efx.sleep(2)
efx.procurar('filter_action.png', True, conf=0.95)
efx.procurar('acoes_maintenance.png', True)
efx.pyg.press('tab', 3)
efx.pyg.press('enter')
efx.sleep(3)
efx.procurar('export_cloud.png', True, True)

checklist_utils.generate_export(__file__, 'sheets_action_success.png')

checklist_utils.save_as_xlsx(excel, maintenance_path)
