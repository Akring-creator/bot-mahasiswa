from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.relative_locator import locate_with
import pandas as pd
import os
import time

url = 'https://pddikti.kemdikbud.go.id/'

# Create Error for data with duplicate
class DataDuplicate(Exception):
    def __init__(self, message) :
        self.message = message
        super().__init__(self.message)

# Create Error for missing data
class MissingData(Exception):
    def __init__(self, message) :
        self.message = message
        super().__init__(self.message)

class Students(webdriver.Chrome):
    def __init__(self, driver_path='D:\coding-lab-fast-track\selenium-bot\chromedriver.exe'):
        self.driver_path = driver_path
        os.environ['PATH'] += self.driver_path
        super(Students, self).__init__()
        self.maximize_window()

    def land_first_page(self):
        self.get(url)
    
    def insert_prompt(self, name):
        """
        This function insert name to search input in PDDIKTI
        """
        selected_element = self.find_element(
            By.XPATH, '//input[@placeholder = "Keyword :  [Nama PT] [Nama Prodi] [Nama MHS] [Nama Dosen] [NIM] [NIDN]."]'
        )
        selected_element.click()

        selected_element = self.find_element(
            By.XPATH, '//input[@pattern="([A-z0-9À-ž\s]){2,}"]'
        )
        selected_element.send_keys(f'{name}')
        # selected_element.send_keys(Keys.ENTER)
        time.sleep(5)
        

    def find_name(self, name, university):
        """
        This function try to find the name and check if the name duplicate or missing
        """
        # cp = self.find_element(
        #     By.XPATH, f'//h2[text()="Mahasiswa"]'
        # )
        # td = cp.find_element(
        #             By.XPATH, f'//following-sibling::div'
        #             )
        dt = self.find_element(By.ID, 'eac-container-search')
        
        check = dt.find_elements(By.XPATH, f'//span[text()="{name}"]')
        correct = []
        for element in check:
            univ_confirmation = element.find_element(By.XPATH, f'//following-sibling::span').text
            print(univ_confirmation)
            if university in univ_confirmation:
                correct.append(element)
        total = len(correct)
        if total > 1:
             raise DataDuplicate(message =f'There are more than 1 data')
        elif total < 1:
             raise MissingData(message ='Data doesnt exist')
        else:
            selected_element = correct[0]
            selected_element.click()
            # time.sleep(2)
            # try:
            #     self.find_element(
            #     By.XPATH, '//table[@class = "table table-bordered"]'
            #     )
            # except:
            #     raise MissingData(message ='Data doesnt exist')
        
        
    def extract_data(self, userid): 
        """
        This function extract student metadata
        """
        selected_table = self.find_element(
                By.XPATH, '//table[@class = "table table-bordered"]'
            )
        data = pd.read_html(selected_table.get_attribute('outerHTML'))[0]
        data=data.drop([0], axis=0)
        column = data[0].unique()
        data.index = column
        data=data.drop(data.columns[[0, 1, 3]], axis=1)
        data = data.transpose()
        data['user_id'] = userid
        return data
    
    def extract_table(self, name, userID):
            
        path_mata_kuliah = r'mata kuliah/'
        path_semester = r'semester/'
        selected_table = self.find_element(
                By.XPATH, '//table[@id = "t01"]'
            )
        data = pd.read_html(selected_table.get_attribute('outerHTML'))[0]
        data = data.drop(['No.'], axis = 1)
        data.to_excel(path_semester + f'{userID}-{name}-semester.xlsx', index = False)

        selected_element = self.find_element(
                By.XPATH, '//a[text()="Riwayat Studi"]'
            )
        selected_element.click()
        time.sleep(1)

        selected_table = self.find_element(
                By.XPATH, '//table[@class = "table table-bordered table-responsive"]'
            )
        data = pd.read_html(selected_table.get_attribute('outerHTML'))[0]
        data = data.drop(['No.'], axis = 1)
        data.to_excel(path_mata_kuliah + f'{userID}-{name}-mata kuliah.xlsx', index = False)

def run():
    student = openFile()
    with Students() as bot:
        for ind in range(0, len(student)):
            bot.land_first_page()
            df = pd.DataFrame()
            error = pd.DataFrame()
            name = student['Name'][ind]
            userid = student['user_id'][ind]
            univ = student['University'][ind]
            time.sleep(2)
            bot.insert_prompt(name)
            try:
                bot.find_name(name, university=univ)
            except DataDuplicate as e:
                data = pd.DataFrame({'user_id':userid, 'name': name, 'reason': e}, index=[0])
                error=error.append(data)
            except MissingData as e:
                data = pd.DataFrame({'user_id':userid,'name': name, 'reason': e}, index=[0])
                error=error.append(data)
            else:
                data = bot.extract_data(userid=userid)
                df = df.append(data, ignore_index = True)
                bot.extract_table(name, userid)
            saveFile('Student Metadata.xlsx', df)
            saveFile('Unsuccesful.xlsx', error)
            print(f'{name} is save')

def openFile():
    path = 'userdata.xlsx'
    df = pd.read_excel(path)
    df = df[['user_id', 'Name', 'University']]
    return df

def saveFile(path, df):
    main = pd.read_excel(path)
    main = main.append(df, ignore_index = True)
    main.to_excel(path, index = False)

def test():
    string = 'Universitas Indonesia'
    sentence = 'Akmal Kuliah di Universitas Indonesia'
    print(string in sentence)
# test()
#run()
# df = openFile()
# print(df.iloc[11])