from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys          import Keys
from selenium.webdriver.support.ui           import WebDriverWait
from selenium.webdriver.support              import expected_conditions as EC
from selenium.common.exceptions              import TimeoutException, ElementNotInteractableException
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
#with open(r'config.json') as config_file:
with open(r'C:\Users\Kassa\Documents\Scripts\ReturnRental\scripts\config.json') as config_file:
    CONFIG = json.load(config_file)

#WEBDRIVER_LOCATION = r'chromedriver.exe'
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
        
        time.sleep(3)
        
        btn_newtask = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@name,'CPMInteractionPortalHeaderTop') and @class='Interaction_tab_header']")))
        btn_newtask.click()
        
        time.sleep(1.5)

        id = btn_newtask.get_attribute('data-menuid')

        time.sleep(3)

        ul_retail = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, f"//ul[@id='{id}']")))
        txt_retail = WebDriverWait(ul_retail, 10).until(EC.presence_of_element_located((By.XPATH, ".//span[text()='Retail']")))

        actions.move_to_element(txt_retail).perform()

        time.sleep(3)

        btn_retailtelenet = WebDriverWait(txt_retail, 10).until(EC.element_to_be_clickable((By.XPATH, f"./ancestor::li[contains(@id,'{id}')]//span[contains(text(),'Telenet')]")))
        btn_retailtelenet.click()
        
    except Exception as e:
        print("An exception occured: ", e, "Type of exception: ", type(e).__name__)
        ERROR_COUNTER += 1

        if (ERROR_COUNTER >= 2):
            driver.quit()
            exit(1)
        return start_task(ERROR_COUNTER)

def read_excel_file(dir):
    df = pd.read_excel(dir)

    # Select only the necessary columns
    df = df[['Klant','Klant nummer', 'Store', 'Serienummer']]

    # Remove NaN values
    df = df.dropna()

    # Only keep the numeric values in de dataframe
    df = df.astype({'Klant nummer':'string'})
    df['Klant nummer'] = df['Klant nummer'].str.extract('(\d+)')

    # Remove duplicate klant nummers
    df = df.drop_duplicates(subset=['Serienummer'])

    # Reset the index
    df = df.reset_index(drop=True)

    # Make the DataFrame smaller by combining the serial numbers together in an array
    df_working_data = pd.DataFrame(columns=['Customer', 'Customer Number', 'Serial Numbers', 'Store'])
    for index, row in df.iterrows():
        cust_number = row['Klant nummer']
        serial_numbers = df['Serienummer'].loc[df['Klant nummer'] == cust_number].tolist()
        df_working_data.loc[len(df_working_data)] = [row['Klant'], cust_number, serial_numbers, row['Store']]

    # Remove duplicate customer numbers
    df_working_data = df_working_data.drop_duplicates(subset=['Customer Number'])

    return df_working_data

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

def search_customers(df, ERROR_COUNTER):
    # Make DataFrames
    df_working_data = pd.DataFrame(columns=['Customer','Customer Number', 'Delivery Order', 'Serial Number', 'Interaction', 'Store'])
    df_succes = pd.DataFrame(columns=['Customer','Customer Number', 'Serial Number', 'Interaction', 'Store'])
    df_failed = pd.DataFrame(columns=df_succes.columns.append(pd.Index(['Reason'])))

    for index, row in df.iterrows():
        start_task(ERROR_COUNTER)
        print(f"Going to work on cust num: {row['Customer Number']} with serial numbers: {row['Serial Numbers']}")
        df_succes, df_failed = search_customer(row, ERROR_COUNTER, df_working_data, df_succes, df_failed, True)
    
    return (df_succes, df_failed)

def search_customer(row, ERROR_COUNTER, df, df_succes, df_failed, starttask):
    try:
        if starttask:
            driver.refresh()
        
            time.sleep(5)
            ERROR_COUNTER = 0

            start_task(ERROR_COUNTER)

        cust = row['Customer']
        cust_number = row['Customer Number']
        serial_numbers = row['Serial Numbers']
        store = row['Store']
        
        time.sleep(10)
        
        # Switch to the working frame
        frames = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.TAG_NAME, 'iframe')))
        driver.switch_to.frame(frames[-1])

        # Look for txtField and 1 checkbox
        chk_inactive = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//label[contains(text(),'Include inactive customers')]")))
        chk_inactive.click()

        txt_cust = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.NAME, "$PpyWorkPage$pIntelligentSearchString")))

        # Give all the input to the elements
        txt_cust.send_keys(cust_number)
        txt_cust.send_keys(Keys.ENTER)
        
        time.sleep(10)

        interactions = get_tasks()
        
        df_succes, df_failed = search_interactions(row, interactions, df, df_succes, df_failed)

        ERROR_COUNTER = 0

        return (df_succes, df_failed)
    except Exception as inst:
        print("An exception occured: ", inst, "Type of exception: ", type(inst).__name__)
        ERROR_COUNTER += 1

        if (ERROR_COUNTER >= 2):
            for serial_number in serial_numbers:
                df_failed.loc[len(df_failed)] = [str(cust), str(cust_number), str(serial_number), None, store, 'Te veel fouten tijdens uitvoering']
            
            time.sleep(5)
            try:
                print("Close tab after exception")
                close_current_tab()
            except TimeoutException as _:
                pass
            return (df_succes, df_failed)

        return search_customer(row, ERROR_COUNTER, df, df_succes, df_failed, False)
  
def close_current_tab():
    driver.refresh()
    
    time.sleep(5)

    driver.switch_to.default_content()
    selected_tab = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//li[@aria-selected='true']")))
    
    close_btn = WebDriverWait(selected_tab, 5).until(EC.presence_of_element_located((By.XPATH, ".//span[@aria-label='Close Tab' and @id='close']")))
    close_btn.click()

    time.sleep(1.5)

    frames = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.TAG_NAME, 'iframe')))
    driver.switch_to.frame(frames[-1])

    time.sleep(1.5)
  
def get_tasks():

    actions = ActionChains(driver)

    # Click on the Tasks button
    btn_tasks = None
    btns = WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.TAG_NAME, 'h3')))

    for btn in btns:
        if btn.get_attribute('class') == "layout-group-item-title" and btn.text.strip() == "Tasks":
            btn_tasks = btn
            break

    btn_tasks.click()

    time.sleep(5)

    # Filter on request
    btn_request = None
    btns = WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr/th/div[@class='oflowDiv']")))

    for btn in btns:
        try:
            WebDriverWait(btn, 0.05).until(EC.presence_of_element_located((By.XPATH, ".//*[contains(text(),'Request')]")))
            btn_request = WebDriverWait(btn, 0.05).until(EC.element_to_be_clickable((By.XPATH, ".//a[@id='pui_filter']")))
            btn_request.click()
            break
        except Exception as _:
            pass

    time.sleep(3)

    # Enter the filter for Return Hardware
    input_request = None
    inputs = WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.TAG_NAME, 'input')))

    for input in inputs:
        if '$PpyFilterCriteria_D_TaskTabDetails' in input.get_attribute('name') and input.get_attribute('class') == 'leftJustifyStyle':
            input_request = input
            break
    
    input_request.send_keys('Return Hardware')

    btn_request_apply = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//li/div/button[text()='Apply']")))
    btn_request_apply.click()

    time.sleep(3)

    # Filter on Status
    btn_status = None
    btns = WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr/th/div[@class='oflowDiv']")))

    for btn in btns:
        try:
            WebDriverWait(btn, 0.05).until(EC.presence_of_element_located((By.XPATH, ".//*[contains(text(),'Status')]")))
            btn_status = WebDriverWait(btn, 0.05).until(EC.element_to_be_clickable((By.XPATH, ".//a[@id='pui_filter']")))
            btn_status.click()
            break
        except Exception as _:
            pass

    time.sleep(3)

    chks = WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr/td/label[contains(text(),'Pending-Completion') or contains(text(),'Open')]")))
    
    if len(chks) == 0:
        return []
    
    for chk in chks:
        chk.click()
        time.sleep(0.5)
    
    btn_status_apply = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//li/div/button[text()='Apply']")))
    btn_status_apply.click()

    time.sleep(3)

    table = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//table[@pl_prop_class='Telenet-ADT-Work']")))

    # Get all interaction ID's from the table
    interaction_ids = []

    # Searching inside the table element so no need to do webdriverwait
    ids = table.find_elements(By.XPATH, ".//a[contains(text(),'S-') and contains(@name,'TaskTabDetails')]")

    for id in ids:
        interaction_ids.append(id.text.strip())

    # If there are more pages then look at those pages
    search_pages = True

    while search_pages:
        try:
            btn_next_page = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//div/div/div/a[@title='Next Page']")))
            btn_next_page.click()
            
            time.sleep(10)
            
            table = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//table[@pl_prop_class='Telenet-ADT-Work']")))

            # Get all interaction ID's from the table
            ids = table.find_elements(By.XPATH, ".//a[contains(text(),'S-')]")

            for id in ids:
                interaction_ids.append(id.text.strip())

        except Exception as _:
            search_pages = False

    return interaction_ids

def search_interactions(row, interactions, df, df_succes, df_failed):
    cust = row['Customer']
    cust_number = row['Customer Number']
    serial_numbers = row['Serial Numbers']
    store = row['Store']

    for interaction in interactions:
        print(f"Working on interaction: {interaction}")

        driver.switch_to.default_content()
        input_interaction = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.NAME, '$PpyDisplayHarness$ppySearchText')))

        input_interaction.clear()
        input_interaction.send_keys(interaction)
        input_interaction.send_keys(Keys.ENTER)

        time.sleep(3)

        frames = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.TAG_NAME, 'iframe')))
        driver.switch_to.frame(frames[-1])

        table = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@pl_prop_class,'Telenet-FW-ADTFW-Work-OrderMgmt-ReturnDevice')]")))

        all_hardware = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, ".//*[contains(@id,'$PpyWorkPage$pReturnDeviceDetails')]")))
        to_change = 0

        # Iterate through all the delivery orders
        for hardware in all_hardware:
            order = WebDriverWait(hardware, 15).until(EC.presence_of_element_located((By.XPATH, ".//td[@data-attribute-name='Delivery Order']/div/span"))).text
            expand = WebDriverWait(hardware, 15).until(EC.presence_of_element_located((By.XPATH, ".//*[@class='expandRowDetails']")))
            expand.click()

            time.sleep(2)
             
            for index, serial_number in enumerate(serial_numbers):
                try:
                    device_indetifier = WebDriverWait(table, 5).until(EC.presence_of_element_located((By.XPATH, f".//div[@id='rowDetail{hardware.get_attribute('id')}']")))
                    WebDriverWait(device_indetifier, 2).until(EC.presence_of_element_located((By.XPATH, f"//span[contains(text(),'{str(serial_number)}')]")))
                    to_change += 1
                    
                    print(f"Found a serial number in {interaction}, the serial number is {serial_number} and the order is {order}")

                    df.loc[len(df)] = [cust, cust_number, order, serial_number, interaction, store]
                    del serial_numbers[index]
                except Exception as _:
                    pass
            
            if len(serial_numbers) == 0:
                break

        if to_change == 0:
            print("Nothing found to change so close the tab")
            close_current_tab()
            continue

        # Try to edit the interaction
        btn_edit = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span/a[contains(@name,'ServiceCaseHeader_pyWorkPage') and text()='Edit']")))
        btn_edit.click()

        time.sleep(10)
        btn_confirm = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@name,'ConfirmEditSC_pyWorkPage') and text()='Yes']")))
        btn_confirm.click()
        
        WebDriverWait(driver, 60).until(EC.alert_is_present())
        alert = driver.switch_to.alert

        # If the interaction gives an empty assignment key error then add it to the failed list
        if 'empty assignment key' in alert.text.lower():
            rows = df.loc[df['Interaction'] == interaction]
            for index, row in rows.iterrows():
                df_failed.loc[len(df_failed)] = pd.concat([row.drop('Delivery Order'), pd.Series(['Empty assignment key error'], index=['Reason'])])
            alert.accept()
            continue
        
        alert.accept()

        # Try to edit the return rental after accepting the alert
        try:
            driver.switch_to.default_content()

            WebDriverWait(driver, 90).until(EC.visibility_of_element_located((By.XPATH, "//div")))

            time.sleep(30)

            frames = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.TAG_NAME, 'iframe')))
            driver.switch_to.frame(frames[-1])

            time.sleep(5)

            actions = ActionChains(driver)

            # Use the dataframe to get the order id and fill in the serial number
            table = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@pl_prop_class,'Telenet-FW-ADTFW-Work-OrderMgmt-ReturnDevice')]")))

            all_delivery_orders = WebDriverWait(table, 60).until(EC.presence_of_all_elements_located((By.XPATH, ".//tr[contains(@id,'$PpyWorkPage$pReturnDeviceDetails')]")))

            # Check for each row the delivery order to be the same as the order in the all_delivery_orders
            succes_flag = False
            for index, row in df.iterrows():
                for delivery_order in all_delivery_orders:
                    try:
                        print(f"Working on delivery order: {row['Delivery Order']}")
                        
                        id = delivery_order.get_attribute('id')

                        WebDriverWait(delivery_order, 0.5).until(EC.presence_of_element_located((By.XPATH, f"//span[text()='{row['Delivery Order']}']")))
                        input_serial_number = WebDriverWait(delivery_order, 0.5).until(EC.presence_of_element_located((By.XPATH, f".//input[@name='{delivery_order.get_attribute('id')}$pAgentEnteredReturnDeviceSerialNumber']")))
                        input_serial_number.send_keys(row['Serial Number'])
                        input_serial_number.send_keys(Keys.ENTER)
                        
                        try:
                            time.sleep(10)

                            send_serial_number = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f"//tr[@id='{id}']//i/..")))
                            send_serial_number.click()

                            time.sleep(10)

                            parent = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//div[@data-node-id='SubmitReturnDelivery' and @node_name='SubmitReturnDelivery']")))
                            proceed_btn = WebDriverWait(parent, 10).until(EC.element_to_be_clickable((By.XPATH, ".//button[text()='Proceed']")))
                            proceed_btn.click()

                            time.sleep(10)

                            df_succes.loc[len(df_succes)] = row.drop('Delivery Order')
                            succes_flag = True
                            print("Added a new line to the succes table")

                            time.sleep(10)
                            break

                        except Exception as inst:
                            print("An exception occured while inputting the serial number: ", inst, "Type of exception: ", type(inst).__name__)
                            rows = df.loc[df['Interaction'] == interaction]
                            rows.drop('Delivery Order', axis='columns')

                            for index, row in rows.iterrows():
                                df_failed.loc[len(df_failed)] = pd.concat([row, pd.Series(["Error while inputting the serial number"], index=['Reason'])])
                            break
                    except TimeoutException as inst:
                        print("An exception occured while searching for the serial number: ", inst, "Type of exception: ", type(inst).__name__)
                if succes_flag:
                    break
            tabs = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, "//iframe[contains(@name,'PegaGadget')]")))
            print(len(tabs))
            
            if len(tabs) > 2:
                print("Closing tab after either a fail or succes in the editing of the interaction")
                close_current_tab()

            if len(serial_numbers) == 0:
                break
        except ElementNotInteractableException as _:
            print("You were the victim of the white screen bandit")
            
            rows = df.loc[df['Interaction'] == interaction]
            rows.drop('Delivery Order', axis='columns')

            for index, row in rows.iterrows():
                df_failed.loc[len(df_failed)] = pd.concat([row, pd.Series(['White screen fail'], index=['Reason'])])
            print("Close tab after not interactable exception")
            close_current_tab()
            continue
        except  TimeoutException as _:
            print("Timed out while waiting for the element to appear on screen")
            print("Close tab after a timeoutexception")
            close_current_tab()
            continue
    
    # If there are leftover serial numbers then put them in the failed DataFrame
    for serial_number in serial_numbers:
        df_failed.loc[len(df_failed)] = [str(cust), str(cust_number), str(serial_number), None, store, 'Serial number not found inside the interactions']

    wrap_up()

    return (df_succes, df_failed)

def wrap_up():
    # Wrap up the task

    time.sleep(10)

    # Wrap up the task
    btn_wrapup = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@name,'CPMInteractionDriver') and contains(@title,'Wrap Up')]")))
    btn_wrapup.click()

    time.sleep(10)

    # Search for the go to wrap up button
    btn_goto = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Go to Wrap Up') and contains(@name,'ConfirmWrapup')]")))
    btn_goto.click()

    time.sleep(10)

    # Click on the submit button
    btn_submit = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'Submit') and contains(@class,'pzbtn')]")))
    btn_submit.click()

    time.sleep(10)

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

    main_window = sg.Window('Return Rental', layout_main, icon=r'CCP.ico')

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
                    df_working_data = pd.DataFrame(columns=['Customer','Customer Number', 'Delivery Order', 'Serial Number', 'Interaction', 'Store'])
                    df_succes = pd.DataFrame(columns=['Customer','Customer Number', 'Serial Number', 'Interaction', 'Store'])
                    df_failed = pd.DataFrame(columns=df_succes.columns.append(pd.Index(['Reason'])))
                    
                    completed = 0

                    for _, row in df.iterrows():
                        try:
                            main_window['-CURRCUS-'].update(value=row['Customer'] + ', met klantnummer: ' + row['Customer Number'])
                            
                            main_window.perform_long_operation(lambda: search_customer(row, ERROR_COUNTER, df_working_data, df_succes, df_failed, True), '-searchcomplete-')
                            completed += 1
                            
                            event, values = main_window.read()

                            if event in (sg.WIN_CLOSED, 'Exit'):
                                try:
                                    save_output(df_succes, "Return_Rental_Succes", False)
                                    save_output(df_failed, "Return_Rental_Error", True)
                                except Exception as ex:
                                    sg.PopupError(f"Error while saving the output to an excel. Exception: {ex}", icon=r'CCP.ico')

                                end = time.time()
                                output = str(round((end - start) / 60)) + ' minutes and ' + str(round((((end - start) / 60.00) - ((end - start) // 60.0)) * 60, 2)) + " seconds."
                                sg.popup(f'Exiting program early after {output}\nThe application handled {completed} customers.\nThe program ran {len(df_succes)}/{completed} succesfull.\nOutput is located at {CONFIG["base_path"]}.', title='Finished', icon=r'CCP.ico')
                                main_window.close()
                                driver.quit()
                                exit(0)
                            
                            if values['-searchcomplete-']:
                                df_succes, df_failed = values['-searchcomplete-']

                            main_window['-PERCENTAGE-'].update(value='%0.2f procent voltooid.' % ((completed / len(df)) * 100))
                            main_window['-PBAR-'].update(current_count=int((completed / len(df)) * 100))

                        except Exception as inst:
                            print("An exception occured: ", inst, "Type of exception: ", type(inst).__name__)
                            sg.PopupError(f"Error while doing tasks for the customer: {row['Customer']} with customer number: {row['Customer Number']}\nError: {inst}", title='Error', icon=r'CCP.ico')
                    try: 
                        save_output(df_succes, 'Return_Rental_Succes', False)
                        save_output(df_failed, 'Return_Rental_Error', True)
                    except Exception as ex:
                        sg.PopupError(f"Error while saving the output to an excel. Exception: {ex}", icon=r'CCP.ico')

                    end = time.time()

                    # Application has finished
                    output = str(round((end - start) / 60)) + ' minutes and ' + str(round((((end - start) / 60.00) - ((end - start) // 60.0)) * 60, 2)) + " seconds."
                    sg.popup(f'Application has finished succesfully in {output}\nThe application handled {completed} return/rentals.\nThe program ran {len(df_succes)}/{completed} succesfull.\nOutput is located at {CONFIG["base_path"]}.', title='Finished', icon=r'CCP.ico')

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
    

    df = read_excel_file(r'C:\Users\Kassa\Documents\Scripts\ReturnRental\excel\TestReturnRental_Case1.xlsx')

    df_succes, df_failed = search_customers(df, ERROR_COUNTER)
    
    # Empty assignment key error:
    #search_customer(32760435, ERROR_COUNTER, [972352214451, 964385572542], df, df_succes, df_failed, True)

    driver.quit()

    location_succes = save_output(df_succes, 'Return_Rental_Succes', False)
    location_failed = save_output(df_failed, 'Return_Rental_Error', True)

    print(f"Saved the succes file to {location_succes}\nSaved the failed file to {location_failed}")


if __name__ == '__main__':
    main()