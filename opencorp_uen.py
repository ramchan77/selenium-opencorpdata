from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
#from selenium.webdriver.common.proxy import Proxy,ProxyType
import time
import cookielib
import requests
import csv
import xlsxwriter
from xlutils.copy import copy
from xlrd import open_workbook

input_file_name = raw_input("Enter The Input file Name (with csv Extention ): ")
#start_url=input("Enter The start URL Count (just Enter The Integer ) : ")
#temp=start_url%100
#end_url=(start_url-temp)+100
#print 'It Will Run Upto : '+str(end_url)
#end_url=input("Enter The Rabge Limit URL Count (just Enter The Integer ) : ")
output_file_name = raw_input("Enter The file Name (with xls Extention ) : ")
#print output_file_name
workbook = xlsxwriter.Workbook(output_file_name)
worksheet = workbook.add_worksheet()
workbook.close()
book_ro = open_workbook(output_file_name)
book = copy(book_ro)
sheet1 = book.get_sheet(0)
count=0
roww=0
coll=0
#page_content=''
print 'Launching Chrome..'
#prox = Proxy()
#prox.proxy_type = ProxyType.MANUAL
#prox.http_proxy = "127.0.0.1:9667"
#prox.socks_proxy = "127.0.0.1:9667"
#prox.ssl_proxy = "127.0.0.1:9667"
#capabilities = webdriver.DesiredCapabilities.CHROME
#prox.add_to_capabilities(capabilities)
capa = DesiredCapabilities.CHROME
capa["pageLoadStrategy"] = "none"
browser = webdriver.Chrome(executable_path='C:\Users\lenovo\Desktop\python\chromedriver.exe',desired_capabilities=capa)
#print 'Waiting for 2 mins...'
#time.sleep(90)
print 'Entering to Website...'
with open(input_file_name, "r") as f:
    reader=csv.reader(f)
    for row in reader:
        site = row[0]
        checker={'value': 1}
        attempt_count={'value': 1}
        attempt_count1={'value': 1}
        count+=1
        def page_l():
            if attempt_count['value']<3:
                try:
                    browser.get(site)
                    wait = WebDriverWait(browser, 15)
                    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="overview"]/div[2]/div/table/tbody')))
                #time.sleep(6)
                    browser.execute_script("window.stop();")
                except TimeoutException:
                    attempt_count['value']+=1
                    page_l()
            else:
                pass
            #continue
    #time.sleep(2)
        try:
            browser.get(site)
            wait = WebDriverWait(browser, 15)
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="overview"]/div[2]/div/table/tbody')))
            browser.execute_script("window.stop();")
        except TimeoutException:
            page_l()
        el_count={'value': 1}
        el_count1={'value': 1}
        def element_fun():
            try:
                elements=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody')
                checker['value']=0
            #page_content=browser.find_element_by_xpath("/html/body/div/div[2]/section[1]").get_attribute("outerHTML")
            except NoSuchElementException:
                if el_count['value']<2:
                    el_count['value']+=1
                    print '~~~~~~~~Waiting For 10 Seconds~~~~~~~~~~'
                #browser.get(site)
                #time.sleep(6)
                #wait = WebDriverWait(browser, 15)
                #wait.until(EC.visibility_of_element_located((By.XPATH, "/html/body/section[1]/div[2]")))
                #browser.execute_script("window.stop();")
                    page_l()
                    element_fun()
                elif (el_count['value']==2) and (el_count1['value']==1):
                    print '~~~~~~~~Retrying~~~~~~~~~~'
                    el_count['value']=1
                    el_count1['value']+=1
                    try:
                        browser.get(site)
                        wait = WebDriverWait(browser, 15)
                    #time.sleep(6)
                        wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="overview"]/div[2]/div/table/tbody')))
                        browser.execute_script("window.stop();")
                        element_fun()
                    except TimeoutException:
                        pass
            #page_content=browser.find_element_by_xpath("/html/body/div/div[2]/section[1]").get_attribute("outerHTML")
        element_fun()
        if checker['value']==0:
            print str(count)+' '+site
            try:
                uen=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[1]/td[2]').text
            except:
                uen=''
            try:
                company=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[2]/td[2]').text
            except:
                company=''
            try:
                status=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[3]/td[2]').text
            except:
                status=''
            try:
                address=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[4]/td[2]').text
                #addresses=address.splitlines()
                #address_1=addresses[0]
                #address_2=addresses[1].split(' ')
                #country=address_2[0]
                #postal_code=address_2[1]
            except:
                address=''
                #country=''
                #postal_code=''
            try:
                UENIssueDate=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[5]/td[2]').text
            except:
                UENIssueDate=''
            try:
                agency=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[6]/td[2]').text
            except:
                agency=''
            try:
                enitity_type=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[7]/td[2]').text
            except:
                enitity_type=''
            try:
                company_type_description=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[8]/td[2]').text
            except:
                company_type_description=''
            try:
                incorporated=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[9]/td[2]').text
            except:
                incorporated=''
            try:
                account_due_date=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[10]/td[2]').text
            except:
                account_due_date=''
            try:
                anual_return_date=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[11]/td[2]').text
            except:
                anual_return_date=''
            try:
                number_charges=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[12]/td[2]').text
            except:
                number_charges=''
            try:
                primary_ssic=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[13]/td[2]').text
            except:
                primary_ssic=''
            try:
                ssic_p_d=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[14]/td[2]').text
            except:
                ssic_p_d=''
            try:
                ssic_s=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[15]/td[2]').text
            except:
                ssic_s=''
            try:
                ssice_s_d=browser.find_element_by_xpath('//*[@id="overview"]/div[2]/div/table/tbody/tr[16]/td[2]').text
            except:
                ssic_s_d=''
            sheet1.write(roww,coll,uen)
            sheet1.write(roww,coll+1,company)
            sheet1.write(roww,coll+2,status)
            sheet1.write(roww,coll+3,address)
            sheet1.write(roww,coll+4,UENIssueDate)
            sheet1.write(roww,coll+5,agency)
            sheet1.write(roww,coll+6,enitity_type)
            sheet1.write(roww,coll+7,company_type_description)
            sheet1.write(roww,coll+8,incorporated)
            sheet1.write(roww,coll+9,account_due_date)
            sheet1.write(roww,coll+10,anual_return_date)
            sheet1.write(roww,coll+11,number_charges)
            sheet1.write(roww,coll+12,primary_ssic)
            sheet1.write(roww,coll+13,ssic_p_d)
            sheet1.write(roww,coll+14,ssic_s)
            sheet1.write(roww,coll+15,ssice_s_d)
            roww+=1
            book.save(output_file_name)
        else:
            print str(count)+' *** '+site+' *** Element Not Found'
            pass
print 'Closing Chrome..'
browser.close()
