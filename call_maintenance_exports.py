# %%
import pyautogui as pyg
import subprocess
import sys
sys.path.append(r'\Python\LIB')
import EFX_lib as efx
import os

efx.PYG_CONFIDENCE = 0.90
efx.dynamic_pyg_img_path = r'\python\pyautogui_img\checklist'
path_short = efx.dynamic_pyg_img_path


def checklist_login():
    efx.procurar('checklist_login.png', True)
    pyg.sleep(0.5)
    pyg.press('tab', 3)
    pyg.sleep(1)
    pyg.press('enter')

    efx.procurar('checklist_login.png')
    pyg.sleep(0.5)
    pyg.press('enter')
    efx.procurar('three_line.png')


def check_if_logged():
    efx.chrome_to_main_monitor(r'www.website_example.com/login')
    print('Verifying login status...\n')

    while True:
        if efx.procurar_uma('checklist_login.png', path_short) is not None:
            print('Not logged in, logging in\n')
            checklist_login()
            break

        if efx.procurar_uma('three_line.png', path_short) is not None:
            print('Is logged')
            break
        pyg.sleep(1)

    pyg.sleep(1)
    pyg.hotkey('ctrl', 'w')


def run_script(script: str):
    pyg.sleep(5)
    maintenance_path = r'\maintenance\python'
    print(f"\n--- STARTING {script} ---")
    subprocess.run(['python', os.path.join(maintenance_path, script)])


def load_data_qlik_sense():
    qlik_path = r'\Python\qlik'
    efx.chrome_to_main_monitor(r'https://qlikcloud.com/dataloadeditor/app/123456789')
    efx.check_if_logged('qlik_logged.png', 'Log in qlik.png', img_path=qlik_path)
    efx.procurar('load_data.png', True, path=qlik_path)
    efx.procurar('keep_editing.png', True, True, path=qlik_path)
    efx.sleep(2)
    pyg.hotkey('ctrl', 'w')


# %%

check_if_logged()

run_script('export_corrective.py')

run_script('export_preventive.py')

run_script('export_predictive.py')

run_script('export_external.py')

run_script('export_expenses.py')

run_script('export_support.py')

run_script('export_outsourced.py')

run_script('export_predictive scheduling.py')

run_script('export_actions.py')

load_data_qlik_sense()