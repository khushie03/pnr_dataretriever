from undetected_chromedriver import Chrome, ChromeOptions
import pandas as pd
import time
import os
import keyboard
import pyautogui
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, WebDriverException

def process_pnr(driver, data_dict, pnr_id):
    time.sleep(5)
    print_invoice_clicked = False
    if pnr_id not in data_dict:
        data_dict[pnr_id] = {'buttons_found': 0, 'downloads': 0, 'visited_paths': set(), 'invoice_numbers': [], 'status': "Not processed"}
    
    try:
        pnr_element = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="PNRId"]'))
        )
        pnr_element.clear()
        pnr_element.click()
        pyautogui.write(pnr_id)
        
        retrieve_element = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="GstRetrievePageInteraction"]'))
        )
        retrieve_element.click()
        
        invoice_button_xpaths = [
        '/html/body/div[3]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul/li[1]',
        '/html/body/div[3]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul[2]/li[1]',
        '/html/body/div[3]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul[3]/li[1]',
        '/html/body/div[3]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul[4]/li[1]',
        '/html/body/div[3]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul[5]/li[1]',
        '/html/body/div[3]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul[6]/li[1]',
        '/html/body/div[3]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul[7]/li[1]',
        '/html/body/div[3]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul[8]/li[1]',
        '/html/body/div[3]/div[2]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul[1]/li[1]',
        '/html/body/div[3]/div[2]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul[2]/li[1]',
        '/html/body/div[3]/div[2]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul[3]/li[1]',
        '/html/body/div[3]/div[2]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul[4]/li[1]',
        '/html/body/div[3]/div[2]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul[5]/li[1]',
        '/html/body/div[3]/div[2]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul[6]/li[1]',
        '/html/body/div[3]/div[2]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul[7]/li[1]',
        '/html/body/div[3]/div[2]/div[2]/div[2]/div[2]/table/tbody/tr/td[2]/ul[8]/li[1]'
    ]


        for xpath in invoice_button_xpaths:
            try:
                invoice_button = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                if invoice_button and xpath not in data_dict[pnr_id]['visited_paths']:
                    invoice_number = invoice_button.text  
                    if invoice_number and invoice_number not in data_dict[pnr_id]['invoice_numbers']:
                        data_dict[pnr_id]['visited_paths'].add(xpath)
                        data_dict[pnr_id]['buttons_found'] += 1
                        data_dict[pnr_id]['invoice_numbers'].append(invoice_number)
            except TimeoutException:
                continue

        while data_dict[pnr_id]['downloads'] < data_dict[pnr_id]['buttons_found']:
            try:
                for invoice_number in data_dict[pnr_id]['invoice_numbers']:
                    download_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, f'//a[@id="PrintInvoice" and @invoice-number="{invoice_number}"]'))
                    )
                    if download_button:
                        download_button.click()
                        data_dict[pnr_id]['downloads'] += 1
                        print_invoice_clicked = True
                        time.sleep(40)
                        pyautogui.click(1508, 909)
                        time.sleep(10)
                        if data_dict[pnr_id]['downloads'] == 1:
                            file_name = pnr_id
                        else:
                            file_name = f"{pnr_id}_{data_dict[pnr_id]['downloads'] - 1}"

                        pyautogui.typewrite(file_name, interval=0.1)
                        time.sleep(10)

                        pyautogui.click(702, 672)
                        
                        print_invoice_clicked = False
                        time.sleep(2) 
                    
                    file_path = os.path.join(r"C:\mimt", file_name + ".pdf")
                    if not os.path.exists(file_path):
                        data_dict[pnr_id]['status'] = f"Failed to download {file_name}.pdf"
                    else:
                        data_dict[pnr_id]['status'] = "Successful"

            except TimeoutException:
                data_dict[pnr_id]['status'] = "PrintInvoice button not found or not clickable."
                break
            except Exception as e:
                data_dict[pnr_id]['status'] = f"Error during downloading and renaming file: {e}"
                break

    except TimeoutException as e:
        data_dict[pnr_id]['status'] = f"Timeout exception: {e}"
    except StaleElementReferenceException as e:
        data_dict[pnr_id]['status'] = f"Stale element reference: {e}"
    except WebDriverException as e:
        data_dict[pnr_id]['status'] = f"WebDriver exception: {e}"
    except Exception as e:
        data_dict[pnr_id]['status'] = f"An unexpected error occurred: {e}"

def update_csv(data_dict, csv_file):
    data = pd.read_csv(csv_file)
    data['numberofinvoices'] = data['pnr_id'].apply(lambda x: data_dict[x]['buttons_found'] if x in data_dict else 0)
    data['numberofdownloads'] = data['pnr_id'].apply(lambda x: data_dict[x]['downloads'] if x in data_dict else 0)
    data['status'] = data['pnr_id'].apply(lambda x: data_dict[x]['status'] if x in data_dict else "Not processed")
    data.to_csv(csv_file, index=False)


def run_script():
    driver = None
    try:
        print("Press 'Q' to exit the script.")
        data_dict = {}
        options = ChromeOptions()
        options.add_argument("--disable-blink-features=AutomationControlled")
        prefs = {"download.default_directory": r"C:\PROJECTS\pnr_retriever\downloaded_file"}  
        options.add_experimental_option("prefs", prefs)
        driver = Chrome(options=options)
        driver.get("https://www.goindigo.in/view-gst-invoice.html")

        while True:
            try:
                data = pd.read_csv("data.csv")
                pnr_list = data['pnr_id'].tolist()
                
                for pnr_id in pnr_list:
                    process_pnr(driver, data_dict, pnr_id)
                    update_csv(data_dict, "data.csv")

                if keyboard.is_pressed('q'):
                    print("Script terminated.")
                    break

                time.sleep(0.01)

            except TimeoutException as e:
                print(f"Timeout exception: {e}")
            except StaleElementReferenceException as e:
                print(f"Stale element reference: {e}")
            except WebDriverException as e:
                print(f"WebDriver exception: {e}")
            except Exception as e:
                print(f"An unexpected error occurred: {e}")

    except KeyboardInterrupt:
        print("Script interrupted by user.")
    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        if driver:
            try:
                driver.quit()
            except WebDriverException:
                pass

        update_csv(data_dict, "data.csv")
        print("Final dictionary of the data retrieved:", data_dict)

run_script()
