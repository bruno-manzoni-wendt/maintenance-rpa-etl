# %%
import pyautogui as pyg
import os

import sys
sys.path.append(r'\Python\LIB')
import EFX_lib as efx

efx.dynamic_pyg_img_path = r'\pyautogui_img\checklist'
efx.PYG_CONFIDENCE = 0.97
efx.set_pyg_pause(0.45)

# %%
def open_link(link):
    efx.open_link_chrome(link)
    pyg.sleep(1)


def export(excel: str):
    efx.search('filterblack.png', True)
    efx.search('filterblack.png')
    pyg.sleep(2)
    pyg.scroll(-1500)
    pyg.moveTo(1, 1)
    pyg.sleep(0.5)
    efx.search('export_cloud.png', True, conf=0.95)

    if excel != 'predictive_scheduling.xlsx':
        pyg.sleep(1.5)
        pyg.press('tab', presses=2)
        pyg.press('enter')
        efx.search('line_item.png', True)

    efx.search('export_button.png', True, True)
    pyg.sleep(4)


def failure_check(script_path: str):
    if pyg.pixelMatchesColor(364, 417, (176, 0, 32), tolerance=8): # check if pixel is red == failure
        print(f'\n--- ⚠️  Export failed, restarting script {os.path.basename(script_path)} ---\n')
        pyg.sleep(3)
        pyg.hotkey('ctrl', 'w')
        efx.subprocess.run([sys.executable, script_path]) # Restarts a new process
        sys.exit()  # Exit the current process


def generate_export(script_path: str, png_export_success: str):
    pyg.press('F6')
    efx.copy_paste(r'https://site_example.com.br/queues')
    pyg.press('enter')

    efx.search('search.png')
    pyg.moveTo(1, 1)
    pyg.sleep(3)

    count = 0
    while True:
        failure_check(script_path) # restart the script if export fails
        efx.print_same_line(f'Export in progress: {count}s')

        position_success = efx.search_uma(png_export_success, efx.dynamic_pyg_img_path, conf=0.96, region=(0, 0, 1920, 461), grayscale=False)
        if position_success is not None:
            print('\n✅ Export completed\n')
            pyg.moveTo(position_success)
            pyg.moveRel(15, 0, duration=0.25)
            break

        if count % 14 == 0 and count != 0:
            pyg.press('F5')

        pyg.sleep(2)
        count += 2

    efx.search('export_black_cloud.png', True, region=(0, 0, 1920, 461))
    pyg.moveTo(1, 1)
    pyg.sleep(1)


def save_as_xlsx(excel: str, maintenance_path: str):
    print(f'----- DOWNLOAD {excel} -----')
    efx.save_select_file(excel, maintenance_path, True)
    efx.file_last_update(os.path.join(maintenance_path, excel))
    pyg.sleep(1)
    pyg.hotkey('ctrl', 'w')

