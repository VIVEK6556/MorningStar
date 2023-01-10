import csv
import os
import shutil
import pandas as pd
import time
import requests
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as Ec
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver as wb
import warnings
warnings.filterwarnings('ignore')
class morning_star:
    def counter(self):
        self.cnt=0
        self.path = input("Enter your Enquery_Sheet Destination Path:")
        self.Download_Excel()
    def Download_Excel(self):
        try:
            shutil.rmtree(''+self.path+'\Dump_Folder\Dump')
        except:
            pass
        code_lst=['LL1019', 'US_NEW_0228', 'LL0612', 'LL0599', 'LL2001', 'LL0274', 'LL0395', 'LL0025', 'LL0082', 'LL0146', 'LL0356', 'AUD_CBA', 'LL0332', 'HK_00992', 'JAP_4755', 'LL0364', 'LL0437', 'LL0339', 'AUD_COH', 'LL1089', 'LL1944', 'LL0346', 'LL0001', 'LL0354', 'LL0028', 'JAP_7751', 'LL1569', 'LL1709', 'LL1210', 'LL1339', 'LL1182', 'HK_00390', 'LL0498', 'LL0195', 'LL0214', 'LL0492', 'LL0517', 'LL0562', 'LL0578', 'LL1700', 'LL0622', 'LL0624', 'LL1059', 'LL0673', 'LL0754', 'LL1088', 'LL0164', 'LL0172', 'LL1830', 'EUR_L_0484', 'EUR_L_0481', 'LL0434', 'LL0440', 'TW_2353', 'LL1812', 'LL1087', 'LL1150', 'LL0446', 'LL0305']
        self.length=len(code_lst)
        symbol_lst=['UNH', 'DASH', 'MSFT', 'INTC', 'COIN', 'TSCO', 'NOKIA', 'SAP', 'ZAL', 'NEM', 'PNDORA', 'CBA', 'HEIA', '00992', '4755', 'BDNNY', 'SXYAY', 'LIGHT', 'COH', 'DMP', 'LOTB', 'DANSKE', 'ADS', 'NOVC', 'VOW3', 'CNN1', 'EBZ', 'DHER', 'FR7', 'HOO', '03988', '00390', 'EDPR', 'BARC', 'EZJ', 'TEF', 'CPR', 'ADSK', 'CTSH', 'LYFT', 'PYPL', 'QCOM', 'BABA', 'AXP', 'DRI', 'DPZ', 'OR', 'RNO', 'C6L.SI', 'INSTAL', 'MIPS', 'NOVN', 'SCMN', '2353', '7974', 'DOL', 'RY', 'CDR', 'VIG']
        exchange_lst=['XNYS', 'XNYS', 'XNAS', 'XNAS', 'XNAS', 'XLON', 'XHEL', 'XETR', 'XETR', 'XETR', 'XCSE', 'XASX', 'XAMS', 'XHKG', 'XTKS', 'PINX', 'PINX', 'XAMS', 'XASX', 'XASX', 'XBRU', 'XCSE', 'XETR', 'XETR', 'XETR', 'XFRA', 'XFRA', 'XFRA', 'XFRA', 'XFRA', 'XHKG', 'XHKG', 'XLIS', 'XLON', 'XLON', 'XMAD', 'XMIL', 'XNAS', 'XNAS', 'XNAS', 'XNAS', 'XNAS', 'XNYS', 'XNYS', 'XNYS', 'XNYS', 'XPAR', 'XPAR', 'XSES', 'XSTO', 'XSTO', 'XSWX', 'XSWX', 'XTAI', 'XTKS', 'XTSE', 'XTSE', 'XWAR', 'XWBO']
        company_name=['UnitedHealth Group Inc', 'DoorDash', 'Microsoft Corp', 'Intel Corp', 'Coinbase', 'Tesco PLC', 'Nokia Oyj', 'SAP SE', 'Zalando SE', 'Nemetschek SE', 'Pandora AS', 'Commonwealth Bank of Aus.', 'Heineken NV', 'Lenovo Group Ltd', 'Rakuten Inc', 'Boliden AB ADR', 'Sika AG ADR', 'Signify NV', 'Cochlear Ltd', "Domino's Pizza Enterprises Ltd", 'Lotus Bakeries NV', 'Danske Bank AS', 'adidas AG', 'Novo Nordisk A/S B', 'Volkswagen AG', 'Canon Inc', 'China Gas', 'Delivery Hero', 'Fast Retailing Co Ltd', 'Herbalife Nutrition', 'Bank Of China Ltd H', 'China Railway Group Ltd H', 'EDP Renovaveis SA', 'Barclays PLC', 'easyJet PLC', 'Telefonica SA', 'Davide Campari-Milano SpA ADR', 'Autodesk Inc', 'Cognizant Technology Solutions Corp A', 'Lyft', 'PayPal Holdings Inc', 'Qualcomm Inc', 'Alibaba Group Holding Ltd ADR', 'American Express Co', 'Darden Restaurants Inc', "Domino's Pizza Inc", "L'Oreal SA", 'Renault SA', 'Singapore Airlines Ltd', 'Instalco AB', 'MIPS AB', 'Novartis AG', 'Swisscom AG', 'Acer Inc', 'Nintendo', 'Dollarama Inc', 'Royal Bank of Canada', 'CD Projekt SA', 'Vienna Insurance Group AG']
        #self.symbol=input("Please enter your symbol:")
        #self.exchange=input("Please enter your exchange:")
        #self.code=input("Please Enter Your code:")
        #self.compony_name=input("Please enter Compony_name:")

        self.symbol =symbol_lst[self.cnt]
        self.exchange = exchange_lst[self.cnt]
        self.code = code_lst[self.cnt]
        self.compony_name = company_name[self.cnt]
        options= wb.ChromeOptions()
        options.add_experimental_option("excludeSwitches", [
            "enable-logging",
            "enable-automation",
            "ignore-certificate-errors"
            "",
            "safebrowsing-disable-download-protection",
            "safebrowsing-disable-auto-update",
            "disable-client-side-phishing-detect"
            "ion"])
        options.add_argument('--no-sandbox')
        options.add_argument('--autoplay-policy=no-user-gesture-required')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument("--disable-blink-features")
        options.add_argument('--start-maximized')
        options.add_argument("--disable-features=ChromeWhatsNewUI")
        options.add_argument("--ignore-certificate-errors")
        options.add_argument('--ignore-ssl-errors=yes')
        options.add_argument("--enable-javascript")
        options.add_argument("--disable-notifications")
        options.add_argument("--enable-popup-blocking")
        options.add_argument("--disable-web-security")
        options.add_argument("--disable-infobars")
        options.add_argument('--no-proxy-server')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument("disable-infobars")
        prefs = {"credentials_enable_service": False,
                 "profile.password_manager_enabled": False}
        options.add_experimental_option("prefs", prefs)
        options.add_experimental_option("prefs", {"download.default_directory": ''+self.path+'\\Dump_Folder\\Dump'})
        response = requests.get('https://www.morningstar.com/stocks/'+self.exchange+'/'+self.symbol+'/financials')
        Web_site_link='https://www.morningstar.com/stocks/'+self.exchange+'/'+self.symbol+'/financials'
        if response.status_code == 200:
            print('Web site exists:',Web_site_link)
        else:
            print('Web site does not exist:',Web_site_link)
            self.Download_Excel()
            self.cnt+=1
        self.driver = wb.Chrome(executable_path="driver/chromedriver.exe", options=options)
        self.driver.maximize_window()
        self.driver.get('https://www.morningstar.com/stocks/'+self.exchange+'/'+self.symbol+'/financials')
        actions = ActionChains(self.driver)
        last_Hieght = self.driver.execute_script('return document.body.scrollHeight')
        while True:
            self.driver.execute_script('window.scrollTo(0,document.body.scrollHeight)')
            time.sleep(5)
            new_Hieght = self.driver.execute_script('return document.body.scrollHeight')
            if new_Hieght == last_Hieght:
                break
            last_Hieght = new_Hieght
        time.sleep(30)
        Income_1 = BeautifulSoup(self.driver.page_source, 'lxml')
        date_currency = Income_1.find_all('div', {'class': 'sal-small-12 sal-columns'})[3].text
        self.date = date_currency[41:47]
        self.currency = date_currency[67:71]
        actions.scroll_to_element(self.driver.find_element(By.CLASS_NAME, 'sal-component-expand')).perform()
        WebDriverWait(self.driver, 60).until(Ec.presence_of_element_located((By.CLASS_NAME, 'sal-component-expand'))).click()
        last_Hieght = self.driver.execute_script('return document.body.scrollHeight')
        while True:
            self.driver.execute_script('window.scrollTo(0,document.body.scrollHeight)')
            time.sleep(5)
            new_Hieght = self.driver.execute_script('return document.body.scrollHeight')
            if new_Hieght == last_Hieght:
                break
            last_Hieght = new_Hieght
        self.driver.execute_script('window.scrollTo(0,300)')
        WebDriverWait(self.driver, 60).until(Ec.presence_of_element_located((By.XPATH,
                                "/html/body/div[2]/div/div/div/div[2]/div[3]/div/main/div/div/div[1]/section/sal-components/div/sal-components-stocks-financials/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div/div/div[2]/div/button"))).click()
        time.sleep(30)
        check=os.path.isfile(''+self.path+'\Dump_Folder\Dump\Income Statement_Annual_As Originally Reported.xls')
        if check == False:
            WebDriverWait(self.driver, 60).until(Ec.presence_of_element_located((By.XPATH,
                                                                            "/html/body/div[2]/div/div/div/div[2]/div[3]/div/main/div/div/div[1]/section/sal-components/div/sal-components-stocks-financials/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div/div/div[2]/div/button"))).click()
        time.sleep(30)
        WebDriverWait(self.driver, 120).until(Ec.presence_of_element_located((By.ID, 'balanceSheet'))).click()
        time.sleep(30)
        self.driver.execute_script('window.scrollTo(0,300)')
        WebDriverWait(self.driver, 60).until(Ec.presence_of_element_located((By.XPATH,"/html/body/div[2]/div/div/div/div[2]/div[3]/div/main/div/div/div[1]/section/sal-components/div/sal-components-stocks-financials/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div/div/div[2]/div/button"))).click()
        time.sleep(30)
        check=os.path.isfile(''+self.path+'\Dump_Folder\Dump\Balance Sheet_Annual_As Originally Reported.xls')
        if check==False:
            WebDriverWait(self.driver, 60).until(Ec.presence_of_element_located((By.XPATH,
                                                                            "/html/body/div[2]/div/div/div/div[2]/div[3]/div/main/div/div/div[1]/section/sal-components/div/sal-components-stocks-financials/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div/div/div[2]/div/button"))).click()
        else:
            pass
        time.sleep(30)
        self.driver.get('https://www.morningstar.com/stocks/'+self.exchange+'/'+self.symbol+'/valuation')
        last_Hieght = self.driver.execute_script('return document.body.scrollHeight')
        while True:
            self.driver.execute_script('window.scrollTo(0,document.body.scrollHeight)')
            time.sleep(5)
            new_Hieght = self.driver.execute_script('return document.body.scrollHeight')
            if new_Hieght == last_Hieght:
                break
            last_Hieght = new_Hieght
        time.sleep(5)
        actions.scroll_to_element(self.driver.find_element(By.ID, 'keyStatsOperatingAndEfficiency')).perform()
        WebDriverWait(self.driver, 60).until(Ec.presence_of_element_located((By.ID, 'keyStatsOperatingAndEfficiency'))).click()
        time.sleep(10)
        WebDriverWait(self.driver, 60).until(Ec.presence_of_element_located((By.XPATH,
                                                                            "/html/body/div[2]/div/div/div/div[2]/div[3]/div/main/div/div/div[1]/section/sal-components/div/sal-components-stocks-valuation/div/div[2]/div/div/div[1]/div[2]/button"))).click()
        check = os.path.isfile(''+self.path+'\Dump_Folder\Dump\OperatingAndEfficiency.xls')
        if check == False:
                WebDriverWait(self.driver, 60).until(Ec.presence_of_element_located((By.XPATH,
                                                                                "/html/body/div[2]/div/div/div/div[2]/div[3]/div/main/div/div/div[1]/section/sal-components/div/sal-components-stocks-valuation/div/div[2]/div/div/div[1]/div[2]/button"))).click()
        else:
            pass
        time.sleep(30)
        actions.scroll_to_element(self.driver.find_element(By.ID, 'keyStatsfinancialHealth')).perform()
        WebDriverWait(self.driver, 60).until(Ec.presence_of_element_located((By.ID, 'keyStatsfinancialHealth'))).click()
        WebDriverWait(self.driver, 60).until(Ec.presence_of_element_located((By.XPATH,
                                                                            "/html/body/div[2]/div/div/div/div[2]/div[3]/div/main/div/div/div[1]/section/sal-components/div/sal-components-stocks-valuation/div/div[2]/div/div/div[1]/div[2]/button"))).click()
        os.path.isfile(''+self.path+'\Dump_Folder\Dump\FinancialHealth.xls')
        if check == False:
                WebDriverWait(self.driver, 60).until(Ec.presence_of_element_located((By.XPATH,
                                                                                "/html/body/div[2]/div/div/div/div[2]/div[3]/div/main/div/div/div[1]/section/sal-components/div/sal-components-stocks-valuation/div/div[2]/div/div/div[1]/div[2]/button"))).click()
        else:
            pass
        time.sleep(30)
        actions.scroll_to_element(self.driver.find_element(By.ID, 'keyStatscashFlow')).perform()
        WebDriverWait(self.driver, 60).until(Ec.presence_of_element_located((By.ID, 'keyStatscashFlow'))).click()
        WebDriverWait(self.driver, 60).until(Ec.presence_of_element_located((By.XPATH,
                                                                            "/html/body/div[2]/div/div/div/div[2]/div[3]/div/main/div/div/div[1]/section/sal-components/div/sal-components-stocks-valuation/div/div[2]/div/div/div[1]/div[2]/button"))).click()
        check = os.path.isfile(''+self.path+'\Dump_Folder\Dump\cashFlow.xls')
        if check == False:
                WebDriverWait(self.driver, 60).until(Ec.presence_of_element_located((By.XPATH,
                                                                                "/html/body/div[2]/div/div/div/div[2]/div[3]/div/main/div/div/div[1]/section/sal-components/div/sal-components-stocks-valuation/div/div[2]/div/div/div[1]/div[2]/button"))).click()
        else:
                pass
        time.sleep(30)
        self.driver.close()
        self.Enquiry_one()
    def Enquiry_one(self):
        if os.path.isfile(''+self.path+'\Dump_Folder\Dump\Income Statement_Annual_As Originally Reported.xls'):
            Enquiry_one_file = ''+self.path+'\Dump_Folder\Dump\Income Statement_Annual_As Originally Reported.xls'
            read_excel=pd.read_excel(Enquiry_one_file)
            length=len(read_excel.columns)
            year=list(read_excel.columns[1:length-1])
            read_excel.to_csv(''+self.path+'\Dump_Folder\Dump\sample.csv')
            read_csv=pd.read_csv(''+self.path+'\Dump_Folder\Dump\sample.csv')
            stats=pd.DataFrame(read_excel)
            status=stats.isin(["    Total Revenue"]).any()
            status_1=stats.isin(["Diluted Net Income Available to Common Stockholders"]).any()
            status_2=stats.isin(["Diluted Weighted Average Shares Outstanding"]).any()
            status_3=stats.isin(["Diluted EPS"]).any()
            df = read_csv.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            if status[0]==True:
                Total_revenue=df[df[''+self.symbol+'_income-statement_Annual_As_Originally_Reported']=="Total Revenue"]
                Total_revenue_1 = Total_revenue.iloc[0][year]
            else:
                Total_revenue_1=[]
            self.data_1, self.data_2, self.data_3, self.data_4, self.data_5 = list(Total_revenue_1) + [None] * (
                    5 - len(Total_revenue_1))
            if status_1[0]==True:
                Duilted_net_income=df[df[''+self.symbol+'_income-statement_Annual_As_Originally_Reported']=="Diluted Net Income Available to Common Stockholders"]
                Duilted_net_income_1 = Duilted_net_income.iloc[0][year]
            else:
                Duilted_net_income_1=[]
            self.Diluted_1, self.Diluted_2, self.Diluted_3, self.Diluted_4, self.Diluted_5 = list(
                    Duilted_net_income_1) + [None] * (5 - len(Duilted_net_income_1))
            if status_2[0]==True:
                Diluted_Weighted=df[df[''+self.symbol+'_income-statement_Annual_As_Originally_Reported']=="Diluted Weighted Average Shares Outstanding"]
                Diluted_Weighted_1 = Diluted_Weighted.iloc[0][year]
            else:
                Diluted_Weighted_1=[]
            self.Weighted_1, self.Weighted_2, self.Weighted_3, self.Weighted_4, self.Weighted_5 = list(
                    Diluted_Weighted_1) + [None] * (5 - len(Diluted_Weighted_1))
            if status_3[0]==True:
                Diluted_EPS=df[df[''+self.symbol+'_income-statement_Annual_As_Originally_Reported']=="Diluted EPS"]
                Diluted_EPS_1=Diluted_EPS.iloc[0][year]
            else:
                Diluted_EPS_1=[]
            self.EPS_1,self.EPS_2, self.EPS_3, self.EPS_4, self.EPS_5=list(Diluted_EPS_1)+[None]*(5-len(Diluted_EPS_1))

            if (self.cnt == 0):
                with open(''+self.path+'\morning_star_Enquiry_1.csv',newline='',mode='w',encoding='utf-8') as f:
                    fields=["My CODE","Exchange","Symbol","Company_Name","Year Ends","Currency","Revenue (Y1)","Revenue (Y2)","Revenue (Y3)","Revenue (Y4)","Revenue (Y5)",
                        "Diluted Net Income(Y1)","Diluted Net Income(Y2)","Diluted Net Income(Y3)","Diluted Net Income(Y4)","Diluted Net Income(Y5)",
                        "Diluted Weighted(Y1)","Diluted Weighted(Y2)","Diluted Weighted(Y3)","Diluted Weighted(Y4)","Diluted Weighted(Y5)",
                        "Diluted EPS(Y1)","Diluted EPS(Y2)","Diluted EPS(Y3)","Diluted EPS(Y4)","Diluted EPS(Y5)"]
                    data = [[self.code,self.exchange,self.symbol, self.compony_name,self.date,self.currency,self.data_1,self.data_2, self.data_3, self.data_4, self.data_5, self.Diluted_1, self.Diluted_2, self.Diluted_3,
                         self.Diluted_4, self.Diluted_5,
                         self.Weighted_1, self.Weighted_2, self.Weighted_3, self.Weighted_4, self.Weighted_5, self.EPS_1,
                         self.EPS_2, self.EPS_3, self.EPS_4, self.EPS_5]]
                    csvwriter=csv.writer(f)
                    csvwriter.writerow(fields)
                    csvwriter.writerows(data)
            else:
                with open(''+self.path+'\morning_star_Enquiry_1.csv', newline='', mode='a', encoding='utf-8') as f:
                    data = [self.code,self.exchange,self.symbol, self.compony_name,self.date,self.currency,self.data_1,self.data_2, self.data_3, self.data_4, self.data_5, self.Diluted_1, self.Diluted_2, self.Diluted_3,self.Diluted_4, self.Diluted_5,self.Weighted_1, self.Weighted_2, self.Weighted_3, self.Weighted_4, self.Weighted_5,
                             self.EPS_1,self.EPS_2, self.EPS_3, self.EPS_4, self.EPS_5]
                    csvwriter = csv.writer(f)
                    csvwriter.writerow(data)
        else:
            with open(''+self.path+'\morning_star_Enquiry_1.csv', newline='', mode='a', encoding='utf-8') as f:
                data = ['E']
                csvwriter = csv.writer(f)
                csvwriter.writerow(data)
        self.Second_Enquiry()
    def Second_Enquiry(self):
        if os.path.isfile(''+self.path+'\Dump_Folder\Dump\Balance Sheet_Annual_As Originally Reported.xls'):
            Enquiry_second=''+self.path+'\Dump_Folder\Dump\Balance Sheet_Annual_As Originally Reported.xls'
            read_excel=pd.read_excel(Enquiry_second)
            length=len(read_excel.columns)
            year = list(read_excel.columns[1:length])
            df=pd.DataFrame(read_excel)
            status=df.isin(["Total Equity"]).any()
            status_1=df.isin(["Total Assets"]).any()
            if status[0]==True:
                Total_Equity = df[df[''+self.symbol+'_balance-sheet_Annual_As_Originally_Reported'] == "Total Equity"]
                Total_Equity_1 = Total_Equity.iloc[0][year]
            else:
                Total_Equity_1=[]
            self.Equity_1, self.Equity_2, self.Equity_3, self.Equity_4, self.Equity_5 = list(Total_Equity_1) + [
                    None] * (5 - len(Total_Equity_1))
            if status_1[0]==True:
                Total_Assets = df[df[''+self.symbol+'_balance-sheet_Annual_As_Originally_Reported'] == "Total Assets"]
                Total_Assets_1 = Total_Assets.iloc[0][year]
            else:
                Total_Assets_1=[]
            self.Assets_1, self.Assets_2, self.Assets_3, self.Assets_4, self.Assets_5 = list(Total_Assets_1) + [
                    None] * (5 - len(Total_Assets_1))
            if (self.cnt == 0):
                with open(''+self.path+'\morning_star_Enquiry_2.csv', newline='', mode='w', encoding='utf-8') as f:
                    fields = ["My CODE", "Exchange", "Symbol", "Company_Name", "Year Ends", "Currency", "Total Assets(Y1)",
                          "Total Assets(Y2)", "Total Assets(Y3)", "Total Assets(Y4)", "Total Assets(Y5)",
                          "Total Equity(Y1)", "Total Equity(Y2)", "Total Equity(Y3)", "Total Equity(Y4)",
                          "Total Equity(Y5)"]
                    data = [
                    [self.code,self.exchange,self.symbol, self.compony_name,self.date,self.currency,self.Assets_1,
                     self.Assets_2, self.Assets_3, self.Assets_4, self.Assets_5,
                     self.Equity_1, self.Equity_2, self.Equity_3, self.Equity_4, self.Equity_5]]
                    csvwriter = csv.writer(f)
                    csvwriter.writerow(fields)
                    csvwriter.writerows(data)
            else:
                with open(''+self.path+'\morning_star_Enquiry_2.csv', newline='', mode='a', encoding='utf-8') as f:
                    data = [self.code,self.exchange,self.symbol, self.compony_name,self.date,self.currency,self.Assets_1, self.Assets_2, self.Assets_3, self.Assets_4, self.Assets_5,
                        self.Equity_1, self.Equity_2, self.Equity_3, self.Equity_4, self.Equity_5]
                    csvwriter = csv.writer(f)
                    csvwriter.writerow(data)
        else:
            with open(''+self.path+'\morning_star_Enquiry_2.csv', newline='', mode='a', encoding='utf-8') as f:
                data = [self.code, self.exchange, self.symbol, self.compony_name,'E', self.currency]
                csvwriter = csv.writer(f)
                csvwriter.writerow(data)
        self.Third_Enquiry()
    def Third_Enquiry(self):
        if os.path.isfile(''+self.path+'\Dump_Folder\Dump\OperatingAndEfficiency.xls'):
            Third_Enquiry=''+self.path+'\Dump_Folder\Dump\OperatingAndEfficiency.xls'
            read_excel=pd.read_excel(Third_Enquiry)
            length=(len(read_excel.columns))
            if length > 13:
                year = list(read_excel.columns[2:length - 2])
            else:
                year = list(read_excel.columns[1:length - 2])
            df=pd.DataFrame(read_excel)
            status=df.isin(['Return on Equity %']).any()
            status_1=df.isin(['Return on Invested Capital %']).any()
            if status[0]==True:
                Return_on_Equity = df[df["Fiscal"] == "Return on Equity %"]
                Return_on_Equity_1 = Return_on_Equity.iloc[0][year]
            else:
                Return_on_Equity_1=[]
            self.Return_1, self.Return_2, self.Return_3, self.Return_4, self.Return_5, self.Return_6, self.Return_7, self.Return_8, self.Return_9, self.Return_10 = list(
                    Return_on_Equity_1) + [None] * (10 - len(Return_on_Equity_1))
            if status_1[0]==True:
                Return_on_Invested_Capital= df[df["Fiscal"] == "Return on Invested Capital %"]
                Return_on_Invested_Capital_1=Return_on_Invested_Capital.iloc[0][year]
                self.Invested_1, self.Invested_2, self.Invested_3, self.Invested_4, self.Invested_5, self.Invested_6, self.Invested_7, self.Invested_8, self.Invested_9, self.Invested_10=list(Return_on_Invested_Capital_1)+[None]*(10-len(Return_on_Invested_Capital_1))
            else:
                Return_on_Invested_Capital_1=[]
            self.Invested_1, self.Invested_2, self.Invested_3, self.Invested_4, self.Invested_5, self.Invested_6, self.Invested_7, self.Invested_8, self.Invested_9, self.Invested_10 = list(
                Return_on_Invested_Capital_1) + [None] * (10 - len(Return_on_Invested_Capital_1))
            if (self.cnt == 0):
                with open(''+self.path+'\morning_star_Enquiry_3.csv', newline='', mode='w', encoding='utf-8') as f:
                    fields=["My CODE","Exchange","Symbol","Company_Name","Year Ends","Currency","Return on Equity %(Y1)","Return on Equity %(Y2)","Return on Equity %(Y3)","Return on Equity %(Y4)","Return on Equity %(Y5)","Return on Equity %(Y6)","Return on Equity %(Y7)","Return on Equity %(Y8)","Return on Equity %(Y9)","Return on Equity %(Y10)",
                        "Return on Invested(Y1)","Return on Invested(Y2)","Return on Invested(Y3)","Return on Invested(Y4)","Return on Invested(Y5)","Return on Invested(Y6)","Return on Invested(Y7)","Return on Invested(Y8)","Return on Invested(Y9)","Return on Invested(Y10)"]
                    data=[[self.code,self.exchange,self.symbol, self.compony_name,self.currency,self.date,self.Return_1, self.Return_2, self.Return_3, self.Return_4, self.Return_5, self.Return_6,self.Return_7, self.Return_8, self.Return_9, self.Return_10,
                       self.Invested_1,self.Invested_2,self.Invested_3,self.Invested_4,self.Invested_5,self.Invested_6,self.Invested_7,self.Invested_8,self.Invested_9,self.Invested_10,]]
                    csvwriter = csv.writer(f)
                    csvwriter.writerow(fields)
                    csvwriter.writerows(data)
            else:
                with open(''+self.path+'\morning_star_Enquiry_3.csv', newline='', mode='a', encoding='utf-8') as f:
                    data = [ self.code,self.exchange,self.symbol, self.compony_name,self.date,self.currency,self.Return_1,self.Return_2, self.Return_3, self.Return_4, self.Return_5, self.Return_6, self.Return_7,self.Return_8, self.Return_9, self.Return_10,
                         self.Invested_1, self.Invested_2, self.Invested_3, self.Invested_4, self.Invested_5,self.Invested_6, self.Invested_7, self.Invested_8, self.Invested_9, self.Invested_10, ]
                    csvwriter = csv.writer(f)
                    csvwriter.writerow(data)
        else:
            with open(''+self.path+'\morning_star_Enquiry_3.csv', newline='', mode='a', encoding='utf-8') as f:
                data = [self.code, self.exchange, self.symbol, self.compony_name, "E", self.currency]
                csvwriter = csv.writer(f)
                csvwriter.writerow(data)
        self.Four_Enquiry()
    def Four_Enquiry(self):
        if os.path.isfile(''+self.path+'\Dump_Folder\Dump\FinancialHealth.xls'):
            Four_Enquiry=''+self.path+'\Dump_Folder\Dump\FinancialHealth.xls'
            read_excel=pd.read_excel(Four_Enquiry)
            length=len(read_excel.columns)
            if length > 12:
                year = list(read_excel.columns[2:length - 1])
            else:
                year = list(read_excel.columns[1:length - 1])
            df=pd.DataFrame(read_excel)
            status=df.isin(["Current Ratio"]).any()
            status_1=df.isin(["Book Value/Share"]).any()
            if status[0]==True:
                Current_Ratio= df[df["Liquidity/Financial Health"] == "Current Ratio"]
                Current_Ratio_1 = Current_Ratio.iloc[0][year]
            else:
                Current_Ratio_1=[]
            self.Ratio_1, self.Ratio_2, self.Ratio_3, self.Ratio_4, self.Ratio_5, self.Ratio_6, self.Ratio_7, self.Ratio_8, self.Ratio_9, self.Ratio_10 = list(
                    Current_Ratio_1) + [None] * (10 - len(Current_Ratio_1))
            if status_1[0]==True:
                Book_Value_Share= df[df["Liquidity/Financial Health"] == "Book Value/Share"]
                Book_Value_Share_1=Book_Value_Share.iloc[0][year]
                self.Value_1, self.Value_2, self.Value_3, self.Value_4, self.Value_5, self.Value_6, self.Value_7, self.Value_8, self.Value_9, self.Value_10 =list(Book_Value_Share_1)+[None]*(10-len(Book_Value_Share_1))
            else:
                Book_Value_Share_1=[]
            self.Value_1, self.Value_2, self.Value_3, self.Value_4, self.Value_5, self.Value_6, self.Value_7, self.Value_8, self.Value_9, self.Value_10 = list(
                Book_Value_Share_1) + [None] * (10 - len(Book_Value_Share_1))
            if (self.cnt == 0):
                with open(''+self.path+'\morning_star_Enquiry_4.csv', newline='', mode='w', encoding='utf-8') as f:
                    fields=["My CODE","Exchange","Symbol","Company_Name","Year Ends","Currency","Book Value/Share(Y1)","Book Value/Share(Y2)","Book Value/Share(Y3)","Book Value/Share(Y4)","Book Value/Share(Y5)","Book Value/Share(Y6)","Book Value/Share(Y7)","Book Value/Share(Y8)","Book Value/Share(Y9)","Book Value/Share(Y10)",
                        "Current Ratio(Y1)","Current Ratio(Y2)","Current Ratio(Y3)","Current Ratio(Y4)","Current Ratio(Y5)","Current Ratio(Y6)","Current Ratio(Y7)","Current Ratio(Y8)","Current Ratio(Y9)","Current Ratio(Y10)"]
                    data=[[self.code,self.exchange,self.symbol, self.compony_name,self.date,self.currency,self.Value_1,self.Value_2,self.Value_3,self.Value_4,self.Value_5,self.Value_6,self.Value_7,self.Value_8,self.Value_9,self.Value_10,
                       self.Ratio_1,self.Ratio_2,self.Ratio_3,self.Ratio_4,self.Ratio_5,self.Ratio_6,self.Ratio_7,self.Ratio_8,self.Ratio_9,self.Ratio_10,
                       ]]
                    csvwriter = csv.writer(f)
                    csvwriter.writerow(fields)
                    csvwriter.writerows(data)
            else:
                with open(''+self.path+'\morning_star_Enquiry_4.csv', newline='', mode='a', encoding='utf-8') as f:
                    data = [ self.code,self.exchange,self.symbol, self.compony_name,self.date,self.currency,self.Value_1,self.Value_2, self.Value_3, self.Value_4, self.Value_5, self.Value_6, self.Value_7, self.Value_8,self.Value_9, self.Value_10,
                         self.Ratio_1, self.Ratio_2, self.Ratio_3, self.Ratio_4, self.Ratio_5, self.Ratio_6, self.Ratio_7,self.Ratio_8, self.Ratio_9, self.Ratio_10]
                    csvwriter = csv.writer(f)
                    csvwriter.writerow(data)
        else:
            with open(''+self.path+'\morning_star_Enquiry_4.csv', newline='', mode='a', encoding='utf-8') as f:
                data = [self.code, self.exchange, self.symbol, self.compony_name, "E", self.currency]
                csvwriter = csv.writer(f)
                csvwriter.writerow(data)
        self.Fifth_Enquiry()
    def Fifth_Enquiry(self):
        if os.path.isfile(''+self.path+'\Dump_Folder\Dump\cashFlow.xls'):
            fifth_Enquiry = ''+self.path+'\Dump_Folder\Dump\cashFlow.xls'
            read_excel = pd.read_excel(fifth_Enquiry)
            length=len(read_excel.columns)
            if length > 12:
                year = list(read_excel.columns[2:length-1])
            else:
                year = list(read_excel.columns[1:length-1])
            df = pd.DataFrame(read_excel)
            status=df.isin(["Cap Ex as a % of Sales"]).any()
            status_1=df.isin(["Free Cash Flow/Share"]).any()
            if status[0]==True:
                Cap_Ex_as_a_of_Sales= df[df["Cash Flow Ratios"] == "Cap Ex as a % of Sales"]
                Cap_Ex_as_a_of_Sales_1 = Cap_Ex_as_a_of_Sales.iloc[0][year]
            else:
                Cap_Ex_as_a_of_Sales_1=[]
            self.Sales_1, self.Sales_2, self.Sales_3, self.Sales_4, self.Sales_5, self.Sales_6, self.Sales_7, self.Sales_8, self.Sales_9, self.Sales_10 = list(
                    Cap_Ex_as_a_of_Sales_1) + [None] * (10 - len(Cap_Ex_as_a_of_Sales_1))
            if status_1[0]==True:
                Free_Cash_Flow_Share = df[df["Cash Flow Ratios"] == "Free Cash Flow/Share"]
                Free_Cash_Flow_Share_1=Free_Cash_Flow_Share.iloc[0][year]
            else:
                Free_Cash_Flow_Share_1=[]
            self.Flow_1, self.Flow_2, self.Flow_3, self.Flow_4, self.Flow_5, self.Flow_6, self.Flow_7, self.Flow_8, self.Flow_9, self.Flow_10 = list(
                Free_Cash_Flow_Share_1) + [None] * (10 - len(Free_Cash_Flow_Share_1))
            if (self.cnt == 0):
                with open(''+self.path+'\morning_star_Enquiry_5.csv', newline='', mode='w', encoding='utf-8') as f:
                    fields = ["My CODE", "Symbol", "Exchange", "Company_Name", "Year Ends", "Currency","Cap Ex as a % (Y1)", "Cap Ex as a % (Y2)", "Cap Ex as a % (Y3)", "Cap Ex as a % (Y4)","Cap Ex as a % (Y5)", "Cap Ex as a % (Y6)", "Cap Ex as a % (Y7)", "Cap Ex as a % (Y8)","Cap Ex as a % (Y9)", "Cap Ex as a % (Y10)","Free Cash Flow/Share(Y1)", "Free Cash Flow/Share(Y2)", "Free Cash Flow/Share(Y3)",
                              "Free Cash Flow/Share(Y4)", "Free Cash Flow/Share(Y5)", "Free Cash Flow/Share(Y6)","Free Cash Flow/Share(Y7)", "Free Cash Flow/Share(Y8)", "Free Cash Flow/Share(Y9)","Free Cash Flow/Share(Y10)"]
                    data=[[self.code,self.exchange,self.symbol, self.compony_name,self.date,self.currency,self.Sales_1,self.Sales_2, self.Sales_3, self.Sales_4, self.Sales_5, self.Sales_6, self.Sales_7, self.Sales_8,self.Sales_9, self.Sales_10,
                     self.Flow_1, self.Flow_2, self.Flow_3, self.Flow_4, self.Flow_5, self.Flow_6, self.Flow_7,self.Flow_8, self.Flow_9, self.Flow_10
                       ]]
                    csvwriter = csv.writer(f)
                    csvwriter.writerow(fields)
                    csvwriter.writerows(data)
            else:
                with open(''+self.path+'morning_star_Enquiry_5.csv', newline='', mode='a', encoding='utf-8') as f:
                    data = [self.code,self.exchange,self.symbol, self.compony_name,self.date,self.currency,self.Sales_1,self.Sales_2, self.Sales_3, self.Sales_4, self.Sales_5, self.Sales_6, self.Sales_7, self.Sales_8,self.Sales_9, self.Sales_10,
                     self.Flow_1, self.Flow_2, self.Flow_3, self.Flow_4, self.Flow_5, self.Flow_6, self.Flow_7,self.Flow_8, self.Flow_9, self.Flow_10]
                    csvwriter = csv.writer(f)
                    csvwriter.writerow(data)
        else:
            with open(''+self.path+'\morning_star_Enquiry_5.csv', newline='', mode='a', encoding='utf-8') as f:
                data = [self.code, self.exchange, self.symbol, self.compony_name, "E", self.currency]
                csvwriter = csv.writer(f)
                csvwriter.writerow(data)
        shutil.rmtree(''+self.path+'\Dump_Folder\Dump')
        self.Dividends()
    def Dividends(self):
        self.driver = wb.Chrome(executable_path="driver/chromedriver.exe")
        self.driver.maximize_window()
        self.driver.get('https://www.morningstar.com/stocks/'+self.exchange+'/'+self.symbol+'/dividends')
        last_Hieght = self.driver.execute_script('return document.body.scrollHeight')
        while True:
            self.driver.execute_script('window.scrollTo(0,document.body.scrollHeight)')
            time.sleep(5)
            new_Hieght = self.driver.execute_script('return document.body.scrollHeight')
            if new_Hieght == last_Hieght:
                break
            last_Hieght = new_Hieght
        # Dividend Per Share
        Dividend= BeautifulSoup(self.driver.page_source, 'lxml')
        try:
            Dividend_Per_Share = Dividend.find_all('tr', {'class': 'mds-tr__sal'})[1]
            data=Dividend_Per_Share.find_all('td', 'mds-td__sal mds-td--right__sal')
            lst = []
            for u in range(0, 10):
                lst.append(data[u].text.strip())
            self.Dividend_1, self.Dividend_2, self.Dividend_3, self.Dividend_4, self.Dividend_5, self.Dividend_6, self.Dividend_7, self.Dividend_8, self.Dividend_9, self.Dividend_10 = lst + [
                    None] * (10 - len(lst))
            if (self.cnt == 0):
                with open(''+self.path+'\morning_star_Enquiry_6.csv', newline='', mode='w', encoding='utf-8') as f:
                    fields = ["My CODE", "Exchange", "Symbol", "Company_Name", "Year Ends", "Currency", "Dividends(Y1)",
                              "Dividends(Y2)", "Dividends(Y3)", "Dividends(Y4)", "Dividends(Y5)", "Dividends(Y6)",
                              "Dividends(Y7)", "Dividends(Y8)", "Dividends(Y9)", "Dividends(Y10)"]
                    data = [[self.code, self.exchange, self.symbol, self.compony_name, self.date, self.currency,self.Dividend_1, self.Dividend_2, self.Dividend_3, self.Dividend_4, self.Dividend_5,self.Dividend_6, self.Dividend_7, self.Dividend_8, self.Dividend_9, self.Dividend_10]]
                    csvwriter = csv.writer(f)
                    csvwriter.writerow(fields)
                    csvwriter.writerows(data)
            else:
                with open(''+self.path+'\morning_star_Enquiry_6.csv', newline='', mode='a', encoding='utf-8') as f:
                    data = [self.code, self.exchange, self.symbol, self.compony_name, self.date, self.currency,
                            self.Dividend_1, self.Dividend_2, self.Dividend_3, self.Dividend_4, self.Dividend_5,
                            self.Dividend_6, self.Dividend_7, self.Dividend_8, self.Dividend_9, self.Dividend_10]
                    csvwriter = csv.writer(f)
                    csvwriter.writerow(data)
        except:
            with open(''+self.path+'\morning_star_Enquiry_6.csv', newline='', mode='a', encoding='utf-8') as f:
                data = [self.code, self.exchange, self.symbol, self.compony_name,'E', self.currency]
                csvwriter = csv.writer(f)
                csvwriter.writerow(data)

        self.driver.close()
        self.cnt += 1
        if self.length < self.cnt:
            self.Download_Excel()
        else:
            print("Task completed")
object.counter()
