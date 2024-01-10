from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys          import Keys
from selenium.webdriver.support.ui           import Select, WebDriverWait
from selenium.webdriver.support              import expected_conditions as EC
from selenium.webdriver.chrome.options       import Options
from selenium.webdriver.common.by            import By
from selenium                                import webdriver

from pathlib                                 import Path
from random                                  import random
from sys                                     import exit

import PySimpleGUI  as sg 
import pandas       as pd

import datetime
import shutil
import json
import time
import os

#Constants
with open(r'C:\Users\Kassa\Documents\Scripts\ReturnRental\scripts\config.json') as config_file:
    CONFIG = json.load(config_file)

WEBDRIVER_LOCATION = r'C:\Users\Kassa\Documents\Scripts\ReturnRental\scripts\chromedriver.exe'
DEBUG_PORT = CONFIG['chrome_port']
ERROR_COUNTER = 0

options = Options()

options.add_experimental_option("debuggerAddress", f"127.0.0.1:{DEBUG_PORT}")
driver = webdriver.Chrome(options=options, executable_path=WEBDRIVER_LOCATION)
actions = ActionChains(driver)
    
def start_task(ERROR_COUNTER):
    try:
        driver.switch_to.default_content()
        actions = ActionChains(driver)
        
        btn_newtask = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="RULE_KEY"]/div[1]/div/div/div[3]/div/div/div/div/span/a')))
        btn_newtask.click()

        txt_retail = WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[11]/ul/li[2]')))
        id = txt_retail.get_attribute('data-childnodesid')
        actions.move_to_element(txt_retail).perform()

        btn_retailtelenet = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="{id}"]/li[2]/a')))
        btn_retailtelenet.click()
        
        ERROR_COUNTER = 0
    except Exception as e:
        ERROR_COUNTER += 1

        if (ERROR_COUNTER >= 3):
            exit(1)

        return start_task(ERROR_COUNTER)

def read_excel_file(dir):
    df_cust_numbers = pd.read_excel(dir)

    #Select only the necessary columns
    df_cust_numbers = df_cust_numbers[['Klant','Klant nummer', 'Store']]

    #Remove NaN values
    df_cust_numbers = df_cust_numbers.dropna()

    #Only keep the numeric values in de dataframe
    df_cust_numbers = df_cust_numbers.astype({'Klant nummer':'string'})
    df_cust_numbers['Klant nummer'] = df_cust_numbers['Klant nummer'].str.extract('(\d+)')

    #Remove duplicate klant nummers
    df_cust_numbers = df_cust_numbers.drop_duplicates(subset=['Klant nummer'])

    #Reset the index
    df_cust_numbers = df_cust_numbers.reset_index(drop=True)

    return df_cust_numbers

def make_shop_files(dir):
    df_big = pd.read_excel(dir)
    
    dfs = []
    dfs.append(df_big[df_big['Store'].isin(['Sint-Denijs-Westrem', 'In&outbound'])])
    dfs.append(df_big[df_big['Store'].isin(['Sint-Niklaas'])])
    dfs.append(df_big[df_big['Store'].isin(['Aartselaar'])])
    dfs.append(df_big[df_big['Store'].isin(['Oostakker'])])
    dfs.append(df_big[df_big['Store'].isin(['Wetteren'])])
    dfs.append(df_big[df_big['Store'].isin(['Lokeren'])])
    dfs.append(df_big[df_big['Store'].isin(['Eeklo'])])

    for index, df in enumerate(dfs):
        # Select only the necessary columns
        df = df[['Klant','Klant nummer', 'Store']]

        # Remove NaN values
        df = df.dropna()

        # Only keep the numeric values in de dataframe
        df = df.astype({'Klant nummer':'string'})
        df['Klant nummer'] = df['Klant nummer'].str.extract('(\d+)')

        # Remove duplicate klant nummers
        df = df.drop_duplicates(subset=['Klant nummer'])

        # Reset the index
        df = df.reset_index(drop=True)
        print(len(df))

        if index == 0:
            save_output(df, "MobileApp_Sint_Denijs", False)
        elif index == 1:
            save_output(df, "MobileApp_Sint_Niklaas", False)
        elif index == 2:
            save_output(df, "MobileApp_Aartselaar", False)
        elif index == 3:
            save_output(df, "MobileApp_Oostakker", False)
        elif index == 4:
            save_output(df, "MobileApp_Wetteren", False)
        elif index == 5:
            save_output(df, "MobileApp_Lokeren", False)
        else:
            save_output(df, "MobileApp_Eeklo", False)

def search_customer(cust_number, ERROR_COUNTER, serial_number):

    try:
        #driver.refresh()

        start_task(ERROR_COUNTER)
        
        time.sleep(1.5 *  2.5)
        
        # Switch to the working frame
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        driver.switch_to.frame(frames[-1])

        # Look for txtField, 2 checkboxes and the search button
        txt_cust = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "$PpyWorkPage$pIntelligentSearchString")))

        # Give all the input to the elements
        txt_cust.send_keys(cust_number)
        txt_cust.send_keys(Keys.ENTER)
        
        time.sleep(1.5 *  8)

        interactions = get_tasks()

        search_interactions(interactions, serial_number)



        ERROR_COUNTER = 0
    except Exception as inst:
        ERROR_COUNTER += 1
        print(inst)

        if (ERROR_COUNTER >= 3):
            return True

        return search_customer(cust_number, ERROR_COUNTER, serial_number)
  
def get_tasks():

    actions = ActionChains(driver)

    # Click on the Tasks button
    btn_tasks = None
    btns = driver.find_elements(By.TAG_NAME, 'h3')

    for btn in btns:
        if btn.get_attribute('class') == "layout-group-item-title" and btn.text.strip() == "Tasks":
            btn_tasks = btn
            break

    btn_tasks.click()

    time.sleep(1.5 *  3)

    # Filter on request
    btn_request = None
    btns = driver.find_elements(By.XPATH, "//tbody/tr/th/div[@class='oflowDiv']")

    for btn in btns:
        try:
            WebDriverWait(btn, 0.05).until(EC.presence_of_element_located((By.XPATH, ".//*[contains(text(),'Request')]")))
            btn_request = WebDriverWait(btn, 0.05).until(EC.presence_of_element_located((By.XPATH, ".//a[@id='pui_filter']")))
            break
        except Exception as _:
            pass
    
    btn_request.click()

    time.sleep(1)

    # Enter the filter for Return Hardware
    input_request = None
    inputs = driver.find_elements(By.TAG_NAME, 'input')

    for input in inputs:
        if '$PpyFilterCriteria_D_TaskTabDetails' in input.get_attribute('name') and input.get_attribute('class') == 'leftJustifyStyle':
            input_request = input
            break
    
    input_request.send_keys('Return Hardware')

    btn_request_apply = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//li/div/button[text()='Apply']")))
    btn_request_apply.click()

    time.sleep(1.5)

    # Filter on Status
    btn_status = None
    btns = driver.find_elements(By.XPATH, "//tbody/tr/th/div[@class='oflowDiv']")

    for btn in btns:
        try:
            WebDriverWait(btn, 0.05).until(EC.presence_of_element_located((By.XPATH, ".//*[contains(text(),'Status')]")))
            btn_status = WebDriverWait(btn, 0.05).until(EC.presence_of_element_located((By.XPATH, ".//a[@id='pui_filter']")))
            break
        except Exception as _:
            pass
    
    btn_status.click()

    time.sleep(1)

    chks = driver.find_elements(By.XPATH, "//tbody/tr/td/label[contains(text(),'Pending-Completion') or contains(text(),'Open')]")

    for chk in chks:
        chk.click()
        time.sleep(0.5)
    
    btn_status_apply = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//li/div/button[text()='Apply']")))
    btn_status_apply.click()

    table = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//table[@pl_prop_class='Telenet-ADT-Work']")))

    # Get all interaction ID's from the table
    interaction_ids = []

    ids = table.find_elements(By.XPATH, ".//a[contains(text(),'S-')]")

    for id in ids:
        interaction_ids.append(id.text.strip())

    # If there are more pages then look at those pages
    search_pages = True

    while search_pages:
        try:
            btn_next_page = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//div/div/div/a[@title='Next Page']")))
            btn_next_page.click()
            
            time.sleep(1)
            
            table = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//table[@pl_prop_class='Telenet-ADT-Work']")))

            # Get all interaction ID's from the table
            ids = table.find_elements(By.XPATH, ".//a[contains(text(),'S-')]")

            for id in ids:
                interaction_ids.append(id.text.strip())

        except Exception as _:
            search_pages = False

    return interaction_ids

def search_interactions(interactions, serial_number):
    for interaction in interactions:
        driver.switch_to.default_content()
        input_interaction = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, '$PpyDisplayHarness$ppySearchText')))

        input_interaction.clear()
        input_interaction.send_keys(interaction)
        input_interaction.send_keys(Keys.ENTER)

        time.sleep(1.5)

        frames = driver.find_elements(By.TAG_NAME, "iframe")
        driver.switch_to.frame(frames[-1])

        time.sleep(1)
        
        # //*[@id="$PpyWorkPage$pReturnDeviceDetails$l1"]/td[1]/span
        # //table[@class='gridTable']/tbody/tr/td/span[@class='expandRowDetails']
        all_hardware = driver.find_elements(By.XPATH, "//table/tbody/tr/td/span[@class='expandRowDetails']")
        print(len(all_hardware))

        for hardware in all_hardware:
            try:
                hardware.click()

                time.sleep(1)

                parent_el = WebDriverWait(hardware, 5).until(EC.presence_of_element_located((By.XPATH, '..')))
                id = f"rowDetail{parent_el.get_attribute('id')}"
                print(id)
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, f"span[@id='{id}' and text()='{serial_number}']")))
                
                btn_edit = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//span/a[contains(@name,'ServiceCaseHeader_pyWorkPage') and text()='Edit']")))
                btn_edit.click()

                time.sleep(1.5)

                btn_confirm = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//button[contains(@name,'ConfirmEditSC_pyWorkPage') and text()='Yes']")))
                btn_confirm.click()

                time.sleep(1.5)

                #input_serial_number = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "")))



                break
            except Exception as _:
                hardware.click()
                pass







def save_output(df, basename, error):

    if error:
        DESTINATION_PATH = Path(f"{CONFIG['base_path']}/ERROR")
    else:
        DESTINATION_PATH = Path(f"{CONFIG['base_path']}/SUCCES")

    # Maak een excel aan die later wordt weggegooid
    writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Processed Cust.')

    writer.sheets['Processed Cust.'].autofit()
    writer.close()

    src_dir = os.getcwd()
    src_file = os.path.join(src_dir, 'output.xlsx')

    # Kijk of het dest path bestaat anders aanmaken van dest path
    if not os.path.isdir(DESTINATION_PATH):
        os.mkdir(DESTINATION_PATH)

    shutil.copy(src_file, DESTINATION_PATH)

    day = datetime.datetime.today()

    dest_file = os.path.join(DESTINATION_PATH, 'output.xlsx')
    new_dest_file = os.path.join(DESTINATION_PATH, basename + '__%02d_%02d_%d' % (day.day, day.month, day.year))
    
    # In case of duplicate file, the file gets a new name
    teller = 1
    if os.path.isfile(new_dest_file + '.xlsx'):
        new_dest_file = next_path(new_dest_file + '_%s.xlsx')
    else:
        new_dest_file = new_dest_file + '.xlsx'

    os.rename(dest_file, new_dest_file)
    os.remove(src_file)

    return new_dest_file

def next_path(path_pattern):
    i = 1

    # First do an exponential search
    while os.path.exists(path_pattern % i):
        i = i * 2

    # Result lies somewhere in the interval (i/2..i]
    # We call this interval (a..b] and narrow it down until a + 1 = b
    a, b = (i // 2, i)
    while a + 1 < b:
        c = (a + b) // 2 # interval midpoint
        a, b = (c, b) if os.path.exists(path_pattern % c) else (a, c)

    return path_pattern % b

def main():

    # Make the GUI using PySimpleGui
    sg.theme('DarkAmber')

    layout_main = [
        [sg.Text('Please choose the Telenet customer file: '), sg.Input(key='-telenetfile-', text_color='black', size=(90,1), disabled=True), sg.FileBrowse(key='-btnbrowse-', file_types=(('Excel Files', '*.xlsx'), ('Excel Files', '*.xls')), disabled=False)],
        [sg.Button('Make store files', key='-btnstores-', disabled=False), sg.Submit(key='-btnsubmit-', disabled=False), sg.Exit(button_color='red')],
        [sg.Text('Currently working on: ', key='-CURRCUSTXT-', visible=False), sg.Text('', size=(0, 1), key='-CURRCUS-', visible=False)],
        [sg.Text('', size=(0, 1), key='-PERCENTAGE-', visible=False)],
        [sg.ProgressBar(max_value=100, orientation='h', size=(65, 15), key='-PBAR-', visible=False)]
    ]

    main_window = sg.Window('Origin Automation', layout_main, icon=r'CCP.ico')

    while True:

        event, values = main_window.read()

        if event in (sg.WIN_CLOSED, 'Exit'):
            main_window.close()
            driver.quit()
            exit(0)

        telenet_file = values['-telenetfile-']

        if event == '-btnstores-':
            # Check if file has the right extension (.xlsx, .xls)
            if '.' in telenet_file:
                if telenet_file.split('.')[1] not in ('xlsx', 'xls'):
                    sg.popup_error('Please enter a supported Excel file (.xlsx or .xls)', title='Unsupported file exception', icon=r'CCP.ico')
                    main_window['-telenetfile-'].set_focus()
                #Check if file exists
                elif not Path(telenet_file).is_file():
                    sg.popup_error('The file was not found, please enter an existing file.', title='No such file exception',icon=r'CCP.ico')
                else:
                    
                    # Disable the buttons while the program is running
                    main_window['-btnstores-'].update(disabled=True)
                    main_window['-btnsubmit-'].update(disabled=True)
                    main_window['-btnbrowse-'].update(disabled=True)

                    make_shop_files(telenet_file)

                    sg.popup(f'Files for the different stores have been made. Please check the output in {CONFIG["base_path"]}\\SUCCES for the excel files.', title='Finished', icon=r'CCP.ico')

                    main_window['-telenetfile-'].update('')
                    main_window['-btnstores-'].update(disabled=False)
                    main_window['-btnsubmit-'].update(disabled=False)
                    main_window['-btnbrowse-'].update(disabled=False)



        if event == '-btnsubmit-':
            start = time.time()

            # Check if file has the right extension (.xlsx, .xls)
            if '.' in telenet_file:

                if telenet_file.split('.')[1] not in ('xlsx', 'xls'):
                    sg.popup_error('Please enter a supported Excel file (.xlsx or .xls)', title='Unsupported file exception', icon=r'CCP.ico')
                    main_window['-telenetfile-'].set_focus()
                #Check if file exists
                elif not Path(telenet_file).is_file():
                    sg.popup_error('The file was not found, please enter an existing file.', title='No such file exception',icon=r'CCP.ico')
                else:
                    sg.popup_auto_close('Application is running, please don\'t close the application.', title='Submitted Telenet file', icon=r'CCP.ico')
                    
                    # Disable the buttons while the program is running
                    main_window['-btnstores-'].update(disabled=True)
                    main_window['-btnsubmit-'].update(disabled=True)
                    main_window['-btnbrowse-'].update(disabled=True)
                    
                    # Making the list of customers
                    df = read_excel_file(telenet_file)
                    save_output(df, 'Return_Rental_Workingfile', False)
                    
                    main_window['-PERCENTAGE-'].update(value='%0.2f procent voltooid.' % (0))
                    main_window['-PBAR-'].update(current_count=int(0))

                    main_window['-CURRCUSTXT-'].update(visible=True)
                    main_window['-PERCENTAGE-'].update(visible=True)
                    main_window['-CURRCUS-'].update(visible=True)
                    main_window['-PBAR-'].update(visible=True)

                    # Doing the business logic
                    df_completed = pd.DataFrame(columns=df.columns)
                    df_error = pd.DataFrame(columns=df.columns)
                    
                    completed = 0

                    for _, value in df.iterrows():
                        try:
                            main_window['-CURRCUS-'].update(value=value['Klant'] + ', met klantnummer: ' + value['Klant nummer'])

                           

                            main_window['-PERCENTAGE-'].update(value='%0.2f procent voltooid.' % ((completed / len(df)) * 100))
                            main_window['-PBAR-'].update(current_count=int((completed / len(df)) * 100))

                        except Exception as inst:
                            sg.PopupError(f"Error while doing tasks for the customer: {value['Klant']} with customer number: {value['Klant nummer']}\nError: {inst}", title='Error', icon=r'CCP.ico')
                    try: 
                        save_output(df_completed, "Origin_Automation_Succes", False)
                        save_output(df_error, "Origin_Automation_Error", True)
                    except Exception as ex:
                        sg.PopupError(f"Error while saving the output to an excel. Exception: {ex}", icon=r'CCP.ico')

                    end = time.time()

                    # Application has finished
                    output = str(round((end - start) / 60)) + ' minutes and ' + str(round((((end - start) / 60.00) - ((end - start) // 60.0)) * 60, 2)) + " seconds."
                    sg.popup(f'Application has finished succesfully in {output}\nThe application handled {completed} return/rentals.\nThe program ran {len(df_completed)}/{completed} succesfull.\nOutput is located at {CONFIG["base_path"]}.', title='Finished', icon=r'CCP.ico')

                    main_window['-telenetfile-'].update('')
                    main_window['-btnstores-'].update(disabled=False)
                    main_window['-btnsubmit-'].update(disabled=False)
                    main_window['-btnbrowse-'].update(disabled=False)

                    main_window['-CURRCUSTXT-'].update(visible=False)
                    main_window['-PERCENTAGE-'].update(visible=False)
                    main_window['-CURRCUS-'].update(visible=False)
                    main_window['-PBAR-'].update(visible=False)


def main_console():
    ERROR_COUNTER = 0
    search_customer(1091168558, ERROR_COUNTER, 226253896477)


if __name__ == '__main__':
    main_console()