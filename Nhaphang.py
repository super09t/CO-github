#import java.util.concurrent.TimeUnit
from numpy import var
import openpyxl
from asyncio import sleep, wait, wait_for
import time
import timeit
from selenium import webdriver 
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
username = "0101329672"
password = "0101329672"
print("Mời nhấp số dòng hàng")
Sumdonghang = int(input()) - 1
print("Nhập số invoice:")
invoice = input()
print("nhập ngày invoice:")
dateinvoice = input()

# initialize the Chrome driver 
driver = webdriver.Chrome("chromedriver")
#driver = webdriver.Chrome(executable_path=(r"C:\Users\Duong\Desktop\python\chromedriver.exe"))
# head to ecosy login page
driver.get("https://dichvucong.moit.gov.vn/Login.aspx?clientkey=vGh3F8zA&url=https%3a%2f%2fecosys.gov.vn%3a443%2fvalidate.moitid")
# find username/email field and send the username itself to the input field
driver.find_element_by_id("ctl00_cplhContainer_txtLoginName").send_keys(username)
# find password input field and insert password as well
driver.find_element_by_id("ctl00_cplhContainer_txtPassword").send_keys(password)
# click login button
driver.find_element_by_name("ctl00$cplhContainer$btnLogin").click()
driver.find_element_by_id("ctl00_Header1_lbtnLogin").click()
driver.find_element_by_id("timer").click()
driver.find_element_by_xpath("/html/body/form/div[4]/div[2]/div[1]/div[1]/div/ul/li[1]/ul/li[1]/div/a").click()
#click chon CO
CO = driver.find_element_by_css_selector("#ctl00_cplhContainer_cmbFormCO")
CO.click()
time.sleep(0.5)
try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//li[contains(text(),'Form D')]"))
    )
finally:driver.find_element_by_xpath("//li[contains(text(),'Form D')]").click()
#chon nưdriver.find_element_by_xpath("//li[contains(text(),'Form D')]").click()
time.sleep(2) 
try:
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'select')]"))
    )
finally: driver.find_element_by_css_selector("#ctl00_cplhContainer_cmbMarket").click()
time.sleep(0.5)
try:
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//li[contains(text(),'Thailand')]"))
    )    
finally: driver.find_element_by_xpath("//li[contains(text(),'Thailand')]").click() 
time.sleep(1.5)  
#GOODS
driver.find_element_by_xpath("//span[contains(text(),'Goods')]").click()
#time.sleep(1)
#===========
#==============Cho vào vòng lặp===============


read_excel = './file.xlsx'       #khai báo file làm CO
wb = openpyxl.load_workbook(read_excel)
sheet = wb['Sheet8']            #Tên sheet file CO
g = sheet.iter_rows(min_row=1, max_row=30, min_col=1, max_col=8)
print(type(g))
# <class 'generator'>
cells_list=list(g)
#print("dòng đầu là")
#print(cells_list[1][1].value)
#print(cells_list[i+1][j+1].value)
#driver.find_element_by_xpath("/html/body/form/div[7]/div[2]/div[2]/div[7]/div[1]/div[2]/div[3]/div/ul/li[2]")
driver.find_element_by_css_selector("#ctl00_cplhContainer_RadTabStrip1 > div > ul > li:nth-child(2) > a").click()
#nhập tên hàng
#driver.find_element_by_id("ctl00_cplhContainer_txtName_wrapper").click()
driver.find_element_by_name("ctl00$cplhContainer$txtName").send_keys(cells_list[1][2].value)
#nhập đơn vị khối lượng
driver.find_element_by_name("ctl00$cplhContainer$cmbGwUnitId").clear()
driver.find_element_by_name("ctl00$cplhContainer$cmbGwUnitId").send_keys("KILOGRAM")
#nhập đơn vị số lượng
driver.find_element_by_name("ctl00$cplhContainer$cmbUnit").clear()
driver.find_element_by_name("ctl00$cplhContainer$cmbUnit").send_keys("PIECE")
#Nhập package
driver.find_element_by_name("ctl00$cplhContainer$txtBoxValue").send_keys(cells_list[1][6].value)
#nhập đơn vị PKG
driver.find_element_by_name("ctl00$cplhContainer$cmbBoxUnitId").clear()
driver.find_element_by_name("ctl00$cplhContainer$cmbBoxUnitId").send_keys("PACKAGE")
#driver.find_element_by_link_text("PACKAGE").click()
#nhập số lượng
driver.find_element_by_name("ctl00$cplhContainer$txtUnitValue").clear()
driver.find_element_by_name("ctl00$cplhContainer$txtUnitValue").send_keys(repr(cells_list[1][3].value))
#nhập trọng lượng
driver.find_element_by_name("ctl00$cplhContainer$txtGwValue").clear()
driver.find_element_by_name("ctl00$cplhContainer$txtGwValue").send_keys(repr(cells_list[1][4].value))
#nhập invoice
driver.find_element_by_name("ctl00$cplhContainer$txtInvoiceItem").send_keys(invoice)
#nhập ngay invoice
driver.find_element_by_name("ctl00$cplhContainer$radDpkInvoiceItemDate$dateInput").send_keys(dateinvoice)
#nhập Mark

driver.find_element_by_name("ctl00$cplhContainer$txtShippingMark").send_keys("NoMark")
#nhập FOB

driver.find_element_by_name("ctl00$cplhContainer$txtCurrencyValue").send_keys(repr(cells_list[1][7].value))

#Nhập hs code
driver.find_element_by_name("ctl00$cplhContainer$cmbHSCode").clear()
driver.find_element_by_name("ctl00$cplhContainer$cmbHSCode").send_keys(cells_list[1][1].value)
time.sleep(0.5)
#Đợi Hs hiện ra thì click
#time.sleep(2)
try:
    element = WebDriverWait(driver,10).until(
        EC.presence_of_element_located((By.XPATH, "//li[contains(text(),'87084092 - - - - Dùng cho xe thuộc nhóm 87.03')]"))
    )
finally:driver.find_element_by_xpath("//li[contains(text(),'87084092 - - - - Dùng cho xe thuộc nhóm 87.03')]").click()
#RVC
#driver.find_element_by_xpath("/html/body/form/div[5]/div[2]/div[2]/div[7]/div[1]/div[2]/div[4]/div[2]/div[1]/div[2]/div[2]/div/img").click()
#driver.find_element_by_xpath("/html/body/form/div[4]/div/div[2]/table/tbody/tr[7]/td[2]/input").click()
time.sleep(2)
#driver.execute_script('document.querySelector("#ctl00_cplhContainer_rpvGoods > div:nth-child(1) > div.col-right > div:nth-child(2) > div > img").click();')
#driver.execute_script('document.querySelector("#RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria > table > tbody > tr.rwContentRow > td.rwWindowContent.rwExternalContent > iframe").contentWindow.document.getElementById("txtRVC").value = "88%";document.querySelector("#RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria > table > tbody > tr.rwContentRow > td.rwWindowContent.rwExternalContent > iframe").contentWindow.document.getElementById("chkRVC").click();document.querySelector("#ctl00_cplhContainer_radToolBarDefault > div > div > div > ul > li:nth-child(1)").click()')
#driver.execute_script('document.querySelector("#ctl00_cplhContainer_radToolBarDefault > div > div > div > ul > li:nth-child(1)").click()')
#####
driver.find_element_by_css_selector("#ctl00_cplhContainer_rpvGoods > div:nth-child(1) > div.col-right > div:nth-child(2) > div > img").click()
#wait.until(EC.element_to_be_clickable((By.ID, "RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria")))
frame2 = driver.find_element(By.XPATH,'//*[@id="RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria"]/table/tbody/tr[2]/td[2]/iframe')
    
# switch to frame by frame element
driver.switch_to.frame(frame2)
driver.find_element_by_xpath("/html/body/form/div[4]/div/div[2]/table/tbody/tr[7]/td[2]/input").click()
time.sleep(0.5)
driver.find_element_by_xpath("/html/body/form/div[4]/div/div[2]/table/tbody/tr[7]/td[3]/input").clear()
driver.find_element_by_xpath("/html/body/form/div[4]/div/div[2]/table/tbody/tr[7]/td[3]/input").send_keys(cells_list[1][5].value)
o = driver.find_element_by_xpath("/html/body/form/div[4]/div/div[1]/div/div/div/ul/li[1]/a/span/span/span/span")
o.click()
#
driver.switch_to.parent_frame()
#driver.switch_to_default_content()
time.sleep(1)
driver.find_element_by_xpath("/html/body/form/div[5]/div[2]/div[2]/div[7]/div[1]/div[2]/div[4]/div[2]/div[1]/div[2]/div[6]/input").click()
time.sleep(2)



for i in range(Sumdonghang):
    if cells_list[i+2][1].value == cells_list[i+1][1].value :
        driver.find_element_by_css_selector("#ctl00_cplhContainer_RadTabStrip1 > div > ul > li:nth-child(2) > a").click()
        #nhập tên hàng
        #driver.find_element_by_id("ctl00_cplhContainer_txtName_wrapper").click()
        driver.find_element_by_name("ctl00$cplhContainer$txtName").clear
        driver.find_element_by_name("ctl00$cplhContainer$txtName").send_keys(cells_list[i+2][2].value)
        #nhập đơn vị khối lượng
    # driver.find_element_by_name("ctl00$cplhContainer$cmbGwUnitId").clear()
        #driver.find_element_by_name("ctl00$cplhContainer$cmbGwUnitId").send_keys("KILOGRAM")
        #nhập đơn vị số lượng
    # driver.find_element_by_name("ctl00$cplhContainer$cmbUnit").clear()
    # driver.find_element_by_name("ctl00$cplhContainer$cmbUnit").send_keys("PIECE")
        #Nhập package
        driver.find_element_by_name("ctl00$cplhContainer$txtBoxValue").send_keys(repr(cells_list[i+2][6].value))
        #nhập đơn vị PKG
    # driver.find_element_by_name("ctl00$cplhContainer$cmbBoxUnitId").clear()
    # driver.find_element_by_name("ctl00$cplhContainer$cmbBoxUnitId").send_keys("PACKAGE")
        #driver.find_element_by_link_text("PACKAGE").click()
        #nhập số lượng
        driver.find_element_by_name("ctl00$cplhContainer$txtUnitValue").send_keys(repr(cells_list[i+2][3].value))
        #nhập trọng lượng
        driver.find_element_by_name("ctl00$cplhContainer$txtGwValue").send_keys(repr(cells_list[i+2][4].value))
        #nhập invoice
        #driver.find_element_by_name("ctl00$cplhContainer$txtInvoiceItem").send_keys("invoi")
        #nhập ngay invoice
    # driver.find_element_by_name("ctl00$cplhContainer$radDpkInvoiceItemDate$dateInput").send_keys("21022022")
        #nhập Mark
    # driver.find_element_by_name("ctl00$cplhContainer$txtShippingMark").send_keys("NoMark")
        #nhập FOB
        driver.find_element_by_name("ctl00$cplhContainer$txtCurrencyValue").send_keys(repr(cells_list[i+2][7].value))
        #Nhập hs code
        """
        driver.find_element_by_name("ctl00$cplhContainer$cmbHSCode").clear()
        driver.find_element_by_name("ctl00$cplhContainer$cmbHSCode").send_keys(repr(cells_list[i+2][1].value))
        #Đợi Hs hiện ra thì click
        #time.sleep(2)
        try:
            element = WebDriverWait(driver,10).until(
                EC.presence_of_element_located((By.XPATH, "//li[contains(text(),'87084092 - - - - Dùng cho xe thuộc nhóm 87.03')]"))
            )
        finally:driver.find_element_by_xpath("//li[contains(text(),'87084092 - - - - Dùng cho xe thuộc nhóm 87.03')]").click()
        """
        #RVC

        #driver.find_element_by_xpath("/html/body/form/div[5]/div[2]/div[2]/div[7]/div[1]/div[2]/div[4]/div[2]/div[1]/div[2]/div[2]/div/img").click()
        #driver.find_element_by_xpath("/html/body/form/div[4]/div/div[2]/table/tbody/tr[7]/td[2]/input").click()
        #time.sleep(2)
        #driver.execute_script('document.querySelector("#ctl00_cplhContainer_rpvGoods > div:nth-child(1) > div.col-right > div:nth-child(2) > div > img").click();')
        #time.sleep(1)
        #driver.execute_script('document.querySelector("#RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria > table > tbody > tr.rwContentRow > td.rwWindowContent.rwExternalContent > iframe").contentWindow.document.getElementById("txtRVC").value = "88%";document.querySelector("#RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria > table > tbody > tr.rwContentRow > td.rwWindowContent.rwExternalContent > iframe").contentWindow.document.getElementById("chkRVC").click();document.querySelector("#ctl00_cplhContainer_radToolBarDefault > div > div > div > ul > li:nth-child(1)").click()')
        #driver.execute_script('document.querySelector("#ctl00_cplhContainer_radToolBarDefault > div > div > div > ul > li:nth-child(1)").click()')
        #####
        driver.find_element_by_xpath("/html/body/form/div[5]/div[2]/div[2]/div[7]/div[1]/div[2]/div[4]/div[2]/div[1]/div[2]/div[2]/div/img").click()
        #wait.until(EC.element_to_be_clickable((By.ID, "RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria")))
        frame2 = driver.find_element(By.XPATH,'//*[@id="RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria"]/table/tbody/tr[2]/td[2]/iframe')
            
        # switch to frame by frame element
        driver.switch_to.frame(frame2)
        o = driver.find_element_by_xpath("/html/body/form/div[4]/div/div[1]/div/div/div/ul/li[1]/a/span/span/span/span")
        driver.find_element_by_xpath("/html/body/form/div[4]/div/div[2]/table/tbody/tr[7]/td[3]/input").clear()
        driver.find_element_by_xpath("/html/body/form/div[4]/div/div[2]/table/tbody/tr[7]/td[3]/input").send_keys(cells_list[i+2][5].value)
        o.click()
        #
        driver.switch_to.parent_frame()
        #driver.switch_to_default_content()
        time.sleep(1)
        driver.find_element_by_xpath("/html/body/form/div[5]/div[2]/div[2]/div[7]/div[1]/div[2]/div[4]/div[2]/div[1]/div[2]/div[6]/input").click()
        time.sleep(2)
        #84839095 - - - - Dùng cho hàng hóa khác thuộc Chương 87
        #84099949 - - - - - Loại khác
    else :
            driver.find_element_by_css_selector("#ctl00_cplhContainer_RadTabStrip1 > div > ul > li:nth-child(2) > a").click()
            #nhập tên hàng
            #driver.find_element_by_id("ctl00_cplhContainer_txtName_wrapper").click()
            driver.find_element_by_name("ctl00$cplhContainer$txtName").clear
            driver.find_element_by_name("ctl00$cplhContainer$txtName").send_keys(cells_list[i+2][2].value)
            #nhập đơn vị khối lượng
        # driver.find_element_by_name("ctl00$cplhContainer$cmbGwUnitId").clear()
            #driver.find_element_by_name("ctl00$cplhContainer$cmbGwUnitId").send_keys("KILOGRAM")
            #nhập đơn vị số lượng
        # driver.find_element_by_name("ctl00$cplhContainer$cmbUnit").clear()
        # driver.find_element_by_name("ctl00$cplhContainer$cmbUnit").send_keys("PIECE")
            #Nhập package
            driver.find_element_by_name("ctl00$cplhContainer$txtBoxValue").send_keys(repr(cells_list[i+2][6].value))
            #nhập đơn vị PKG
        # driver.find_element_by_name("ctl00$cplhContainer$cmbBoxUnitId").clear()
        # driver.find_element_by_name("ctl00$cplhContainer$cmbBoxUnitId").send_keys("PACKAGE")
            #driver.find_element_by_link_text("PACKAGE").click()
            #nhập số lượng
            driver.find_element_by_name("ctl00$cplhContainer$txtUnitValue").send_keys(repr(cells_list[i+2][3].value))
            #nhập trọng lượng
            driver.find_element_by_name("ctl00$cplhContainer$txtGwValue").send_keys(repr(cells_list[i+2][4].value))
            #nhập invoice
            #driver.find_element_by_name("ctl00$cplhContainer$txtInvoiceItem").send_keys("invoi")
            #nhập ngay invoice
        # driver.find_element_by_name("ctl00$cplhContainer$radDpkInvoiceItemDate$dateInput").send_keys("21022022")
            #nhập Mark
        # driver.find_element_by_name("ctl00$cplhContainer$txtShippingMark").send_keys("NoMark")
            #nhập FOB
            driver.find_element_by_name("ctl00$cplhContainer$txtCurrencyValue").send_keys(repr(cells_list[i+2][7].value))
            #Nhập hs code
            driver.find_element_by_name("ctl00$cplhContainer$cmbHSCode").clear()
            driver.find_element_by_name("ctl00$cplhContainer$cmbHSCode").send_keys(repr(cells_list[i+2][1].value))
            time.sleep(2)
            driver.find_element_by_xpath("/html/body/form/div[1]/div/div[1]/ul/li[1]").click()
            """
            /html/body/form/div[1]/div/div[1]/ul/li[1]
            #Đợi Hs hiện ra thì click
            #time.sleep(2)
            try:
                element = WebDriverWait(driver,10).until(
                    EC.presence_of_element_located((By.XPATH, "//li[contains(text(),'87084092 - - - - Dùng cho xe thuộc nhóm 87.03')]"))
                )
            finally:driver.find_element_by_xpath("//li[contains(text(),'87084092 - - - - Dùng cho xe thuộc nhóm 87.03')]").click()
            """
            #RVC
            #driver.find_element_by_xpath("/html/body/form/div[5]/div[2]/div[2]/div[7]/div[1]/div[2]/div[4]/div[2]/div[1]/div[2]/div[2]/div/img").click()
            #driver.find_element_by_xpath("/html/body/form/div[4]/div/div[2]/table/tbody/tr[7]/td[2]/input").click()
            time.sleep(2)
            #driver.execute_script('document.querySelector("#ctl00_cplhContainer_rpvGoods > div:nth-child(1) > div.col-right > div:nth-child(2) > div > img").click();')
            time.sleep(1)
            #driver.execute_script('document.querySelector("#RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria > table > tbody > tr.rwContentRow > td.rwWindowContent.rwExternalContent > iframe").contentWindow.document.getElementById("txtRVC").value = "88%";document.querySelector("#RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria > table > tbody > tr.rwContentRow > td.rwWindowContent.rwExternalContent > iframe").contentWindow.document.getElementById("chkRVC").click();document.querySelector("#ctl00_cplhContainer_radToolBarDefault > div > div > div > ul > li:nth-child(1)").click()')
            #driver.execute_script('document.querySelector("#ctl00_cplhContainer_radToolBarDefault > div > div > div > ul > li:nth-child(1)").click()')
            #####
            driver.find_element_by_xpath("/html/body/form/div[5]/div[2]/div[2]/div[7]/div[1]/div[2]/div[4]/div[2]/div[1]/div[2]/div[2]/div/img").click()
            #wait.until(EC.element_to_be_clickable((By.ID, "RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria")))
            frame2 = driver.find_element(By.XPATH,'//*[@id="RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria"]/table/tbody/tr[2]/td[2]/iframe')
                
            # switch to frame by frame element
            driver.switch_to.frame(frame2)
            o = driver.find_element_by_xpath("/html/body/form/div[4]/div/div[1]/div/div/div/ul/li[1]/a/span/span/span/span")
            driver.find_element_by_xpath("/html/body/form/div[4]/div/div[2]/table/tbody/tr[7]/td[3]/input").clear()
            driver.find_element_by_xpath("/html/body/form/div[4]/div/div[2]/table/tbody/tr[7]/td[3]/input").send_keys(cells_list[i+2][5].value)
            o.click()
            #
            driver.switch_to.parent_frame()
            #driver.switch_to_default_content()
            time.sleep(1)
            driver.find_element_by_xpath("/html/body/form/div[5]/div[2]/div[2]/div[7]/div[1]/div[2]/div[4]/div[2]/div[1]/div[2]/div[6]/input").click()
            time.sleep(2)
            #84839095 - - - - Dùng cho hàng hóa khác thuộc Chương 87
            #84099949 - - - - - Loại khác