import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)
from selenium import webdriver
from urllib.request import urlopen
import csv , time 
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium. webdriver. common. keys import Keys
class Lowes(webdriver.Chrome):
    def __init__(self,driver_path='C:\\SeleniumDrivers\\chromedriver.exe', teardown = False) -> None:
        options = webdriver.ChromeOptions()
        
        options.add_argument("disable-infobars")
        options.add_argument("--disable-extensions")
        options.add_experimental_option("detach", True)
        options.add_experimental_option( 'excludeSwitches' , ['enable-logging'])
        options.page_load_strategy = 'eager'
        self.driver_path = driver_path
        self.teardown = teardown
        super(Lowes , self).__init__(options = options ,executable_path=self.driver_path )

        self.implicitly_wait(2)
        

    def __exit__(self, *args):
        if self.teardown:
            self.quit()
   
    def land_on_page(self , u , sku):
        '''
        function to land on homepage.

        Returns
        ----------
        None
        '''
        self.get(u + f"search?searchTerm={sku}")
    def get_info(self):
        try:
            WebDriverWait(self, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR , 'h1')))
            name = self.find_element(By.CSS_SELECTOR , 'h1').text.rstrip()
        except:
            name = None
        try:
            WebDriverWait(self, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR , 'ul[class="bullets"]')))
            desc = self.find_element(By.CSS_SELECTOR , 'ul[class="bullets"]').text.rstrip()
        except:
            desc = None
        try:    
            WebDriverWait(self, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR , 'ol')))
            cat = self.find_element(By.CSS_SELECTOR , 'ol').text.rstrip()
        except:
            cat = None
        return name, desc, cat
    def click_on_product(self):
        #open in new tab using execute_script
        try:
            link = self.find_element(By.CSS_SELECTOR , 'a[data-clicktype="product_tile_click"]').get_attribute('href')
            self.execute_script(f'''window.open("{link}","_blank");''')
            self.switch_to.window(self.window_handles[1])
            name , desc , cat = self.get_info()
            self.close()
            self.switch_to.window(self.window_handles[0])
        except:
            name , desc , cat = self.get_info()
            link  = self.current_url
        return name, desc, cat , link
st_num = 0
try:
    fname = input("Enter the file name: ")
    sname = input("Enter the sheet name: ")
    st_num = int(input("Enter the starting number: "))

    try:
        df = pd.read_excel(fname , sname)
    except:
        df = pd.read_excel(fname)


    df["Name"] = ""
    df["Description"] = ""
    df["Category"] = ""
    df["Link"] = ""
    l = Lowes()
    for sku in df['SKU'][st_num:]:
        l.land_on_page('https://www.lowes.com/' , sku )
        name, desc , cat , link = l.click_on_product()
        df.loc[df['SKU']==sku, 'Name'] = name
        df.loc[df['SKU']==sku, 'Description'] = desc
        df.loc[df['SKU']==sku, 'Category'] = cat
        df.loc[df['SKU']==sku, 'Link'] = link
        if st_num%20 == 0:
            print("writing to file")
            df.to_excel('filled.xlsx' , index = False)
        print(st_num , "record completed")
        st_num = st_num + 1
        

        

    df.to_excel('filled.xlsx' , index = False)
except Exception as e:
    #save error to txt file
    with open('logfile.txt' , 'w') as f:
        f.write("Error occured at: " + time.ctime() + "\n")
        f.write(f"The last record completed was record number {str(st_num)}\n\n--Error Description-- \n")
        f.write(str(e))
    