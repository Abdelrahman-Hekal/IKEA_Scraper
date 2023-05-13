from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import undetected_chromedriver as uc
import time
import os
import re
from datetime import datetime
import pandas as pd
import warnings
import sys
import xlsxwriter
from multiprocessing import freeze_support
import calendar 
import shutil
warnings.filterwarnings('ignore')

def initialize_bot():

    # Setting up chrome driver for the bot
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument('--headless')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # installing the chrome driver
    driver_path = ChromeDriverManager().install()
    chrome_service = ChromeService(driver_path)
    # configuring the driver
    driver = webdriver.Chrome(options=chrome_options, service=chrome_service)
    ver = int(driver.capabilities['chrome']['chromedriverVersion'].split('.')[0])
    driver.quit()
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument("--enable-javascript")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--lang=en")
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument('--headless=new')
    chrome_options.page_load_strategy = 'normal'
    driver = uc.Chrome(version_main = ver, options=chrome_options) 
    driver.set_window_size(1920, 1080)
    driver.maximize_window()
    driver.set_page_load_timeout(20000)

    return driver

def scrape_IKEA(driver, output1, page, cat):

    print('-'*75)
    print(f'Scraping products Links from: {page}')
    stamp = datetime.now().strftime("%d_%m_%Y_%H_%M")

    # getting the products list
    links = []
    driver.get(page)

    while True:
        # scraping products urls 
        try:
            prods = wait(driver, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class*='itemInfo']")))    
        except:
            print('No products are available')
            return

        for prod in prods:
            try:
                link = wait(prod, 2).until(EC.presence_of_element_located((By.TAG_NAME, "a"))).get_attribute('href') 
                if link not in links:
                    links.append(link)
            except:
                pass

        # moving to the next page
        try:
            url = wait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//a[@aria-label='Next']"))).get_attribute('data-sitemap-url') 
            driver.get(url)
        except:
            break

    # scraping Products details
    print('-'*75)
    print('Scraping Products Details...')
    print('-'*75)

    n = len(links)
    data = pd.DataFrame()
    for i, link in enumerate(links):
        try:
            if '/en/' not in link and '/zh/' not in link:
                link = link.replace('/products/', '/en/products/')
            try:
                driver.get(link.replace('/en/', '/zh/'))   
            except:
                print(f'Warning: Failed to load the url: {link}')
                continue
       
            print(f'Scraping the details of product {i+1}\{n}')
            details = {}

            details['Store'] = 'IKEA'

            # Chinese name
            series, name = '', ''  
            try:
                series = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[class='itemName']"))).get_attribute('textContent').strip()
                name = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='itemDetails']"))).get_attribute('textContent').strip()
            except Exception as err:
                pass
            
            details['Product Name (Chinese)'] = f"[{series}] {name}"  
  
            # English name
            try:
                driver.get(link.replace('/zh/', '/en/'))   
            except:
                print(f'Warning: Failed to load the url: {link}')
                continue

            series, name = '', ''  
            try:
                series = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[class='itemName']"))).get_attribute('textContent').strip()
                name = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='itemDetails']"))).get_attribute('textContent').strip()
            except Exception as err:
                pass
            
            details['Product Name (English)'] = f"[{series}] {name}"
                                
            # Product ID
            prod_id = ''             
            try:
                prod_id = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[class='item-code']"))).get_attribute('textContent').strip()
            except:
                continue  

            details['Product ID'] = prod_id 
            details['Link'] = link 
                           
            # product image
            img = ''             
            try:
                a = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[class='slideImg']")))
                img = wait(a, 2).until(EC.presence_of_element_located((By.TAG_NAME, "img"))).get_attribute('src')
            except:
                continue 
                
            details['Image Link'] = img            
            
            # Product outline
            outline = ''             
            try:
                outline = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='product-desc-wrapper']"))).get_attribute('textContent').strip()
            except:
                pass 
            
            details['Product Outline'] = outline

            # product price
            price = ''
            try:
                price = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='itemPrice-wrapper']"))).get_attribute('textContent').replace('$', '').strip().replace(',', '')
            except:
                pass

            details['Price (HKD)'] = price  

            # product description
            des = ''             
            try:
                button = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[data-open*='product-detail-details']")))
                driver.execute_script("arguments[0].click();", button)
                time.sleep(1)
                try:
                    des = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class*='full-length-text-content']"))).get_attribute('textContent').strip()
                except:
                    des = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class*='product-desc-wrapper']"))).get_attribute('textContent').strip()
                htmlelement= wait(driver, 2).until(EC.presence_of_element_located((By.TAG_NAME, "html")))
                htmlelement.send_keys(Keys.ENTER)
                time.sleep(1)
            except:
                pass 
             
            details['Product Description'] = des            
            
            if details['Product Outline'] == '':
                details['Product Outline'] = des

            # other info
            info = ''             
            try:
                button = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[data-open*='measuarements-details']")))
                driver.execute_script("arguments[0].click();", button)
                time.sleep(1)
                div = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class*='measurements-container']")))
                trs = wait(div, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "tr")))
                for tr in trs:
                    try:
                        info += tr.get_attribute('textContent').replace('\n', '').replace(':', ':, ').strip() + '; '
                    except:
                        pass
                info = info[:-2]
                htmlelement= wait(driver, 2).until(EC.presence_of_element_located((By.TAG_NAME, "html")))
                htmlelement.send_keys(Keys.ENTER)
                time.sleep(1)
            except:
                pass 
             
            details['Other Information'] = info
            details['Product Type (Original)'] = cat.strip()

            # product color
            color = ''
            try:
                color = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[id='variation-selected-subtitle']"))).get_attribute('textContent').strip().title()
            except:
                pass

            details['Colour'] = color  

            # product materials
            mat = ''             
            try:
                button = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[data-open*='product-detail-details']")))
                driver.execute_script("arguments[0].click();", button)
                time.sleep(2)
                div = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[id='materials-details']")))
                mat_div = wait(div, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class='mb-3']")))[-1]
                mat = mat_div.get_attribute('innerHTML').replace('\n', '').replace('<br>', '\n').strip()
                #removing other HTML tags from text
                mat = re.compile(r'<[^>]+>').sub('', mat)
            except:
                pass 
             
            details['Materials (IKEA)'] = mat
            details['Extraction Date'] = stamp
            # appending the output to the datafame       
            data = data.append([details.copy()])
           
        except Exception as err:
            print(f'Warning: the below error occurred while scraping the product link: {link}')
            print(str(err))
           
    # output to excel
    if data.shape[0] > 0:
        data['Extraction Date'] = pd.to_datetime(data['Extraction Date'])
        df1 = pd.read_excel(output1)
        df1 = df1.append(data)   
        df1 = df1.drop_duplicates()
        df1.to_excel(output1, index=False)
    else:
        print('-'*75)
        print(f'No valid products Found in: {page}')
        
def get_inputs():
 
    print('-'*75)
    print('Processing The Settings Sheet ...')
    # assuming the inputs to be in the same script directory
    path = os.getcwd()
    if '\\' in path:
        path += '\\IKEA_settings.xlsx'
    else:
        path += '/IKEA_settings.xlsx'

    if not os.path.isfile(path):
        print('Error: Missing the settings file "IKEA_settings.xlsx"')
        input('Press any key to exit')
        sys.exit(1)
    try:
        urls = []
        df = pd.read_excel(path)
        cols  = df.columns
        for col in cols:
            df[col] = df[col].astype(str)

        inds = df.index
        for ind in inds:
            row = df.iloc[ind]
            link, link_type, status = '', '', ''
            for col in cols:
                if row[col] == 'nan': continue
                elif col == 'Category Link':
                    link = row[col]
                elif col == 'Scrape':
                    status = row[col]                
                elif col == 'Type':
                    link_type = row[col]

            if link != '' and status != '' and link_type != '':
                try:
                    status = int(status)
                    urls.append((link, status, link_type))
                except:
                    urls.append((link, 0, link_type))
    except:
        print('Error: Failed to process the settings sheet')
        input('Press any key to exit')
        sys.exit(1)

    # checking the settings dictionary
    #keys = ["Posts Limit"]
    #for key in keys:
    #    if key not in settings.keys():
    #        print(f"Warning: the setting '{key}' is not present in the settings file")
    #        settings[key] = 1
    #    try:
    #        settings[key] = int(float(settings[key]))
    #    except:
    #        input(f"Error: Incorrect value for '{key}', values must be numeric only, press an key to exit.")
    #        sys.exit(1)

    return urls

def initialize_output():

    stamp = datetime.now().strftime("%d_%m_%Y_%H_%M")
    path = os.getcwd() + '\\Scraped_Data\\' + stamp
    if os.path.exists(path):
        shutil.rmtree(path)
    os.makedirs(path)

    file1 = f'IKEA_{stamp}.xlsx'

    # Windws and Linux slashes
    if os.getcwd().find('/') != -1:
        output1 = path.replace('\\', '/') + "/" + file1
    else:
        output1 = path + "\\" + file1  

    # Create an new Excel file and add a worksheet.
    workbook1 = xlsxwriter.Workbook(output1)
    workbook1.add_worksheet()
    workbook1.close()    

    return output1

def main():

    print('Initializing The Bot ...')
    freeze_support()
    start = time.time()
    output1 = initialize_output()
    urls = get_inputs()
    try:
        driver = initialize_bot()
    except Exception as err:
        print('Failed to initialize the Chrome driver due to the following error:\n')
        print(str(err))
        print('-'*75)
        input('Press any key to exit.')
        sys.exit()

    for url in urls:
        if url[1] == 0: continue
        link = url[0]
        cat = url[2]
        try:
            scrape_IKEA(driver, output1, link, cat)
        except Exception as err: 
            print(f'Warning: the below error occurred:\n {err}')
            driver.quit()
            time.sleep(5)
            driver = initialize_bot()

    driver.quit()
    print('-'*75)
    elapsed_time = round(((time.time() - start)/60), 2)
    hrs = round(elapsed_time/60, 2)
    input(f'Process is completed in {elapsed_time} mins ({hrs} hours), Press any key to exit.')

if __name__ == '__main__':

    main()
