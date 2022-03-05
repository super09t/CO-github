from tkinter import TOP
from selenium import webdriver 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
username = "0101329672"
password = "0101329672"

# initialize the Chrome driver 
driver = webdriver.Chrome("chromedriver") 
# head to github login page
driver.get("https://dichvucong.moit.gov.vn/Login.aspx?clientkey=vGh3F8zA&url=https%3a%2f%2fecosys.gov.vn%3a443%2fvalidate.moitid")
# find username/email field and send the username itself to the input field
driver.find_element_by_id("ctl00_cplhContainer_txtLoginName").send_keys(username)
# find password input field and insert password as well
driver.find_element_by_id("ctl00_cplhContainer_txtPassword").send_keys(password)
# click login button
driver.find_element_by_name("ctl00$cplhContainer$btnLogin").click()
#input("Press Enter to continue...")
#selenium.click("//a[contains(text(),'Đăng nhập')]");
#click vào chữ đăng nhập
driver.find_element_by_link_text("Đăng nhập").click()
driver.get("https://ecosys.gov.vn/Default.aspx")
#Đăng nhập Thành cmn công
driver.find_element_by_link_text("Khai báo C/O").click()
#driver.find_element_by_css_selector("ul .rcbItem:nth-child(2)").click()
#driver.find_element_by_link_text("Goods").click()
#input("Press Enter to continue...")

CO = driver.find_element_by_css_selector("#ctl00_cplhContainer_cmbFormCO")
#CO.click()
#time.sleep(0.5)
driver.find_element_by_css_selector("#ctl00_cplhContainer_cmbFormCO_Input").send_keys("FF")
driver.find_element_by_css_selector("#ctl00_cplhContainer_cmbFormCO_Input").submit()


"""try:
    element = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, "//li[contains(text(),'Form D')]"))
    )
finally: driver.find_element_by_xpath("/html/body/form/div[1]/div/div/ul/li[2]").click()"""
     #driver.find_element_by_xpath("//li[contains(text(),'Form D')]").click()
#time.sleep(0.5)
#driver.find_element_by_xpath("/html/body/form/div[1]/div/div/ul/li[2]").click()
#time.sleep(3)
try:
    element = WebDriverWait(driver, 10).until(
        EC.element_located_selection_state_to_be((By.CSS_SELECTOR, "#ctl00_cplhContainer_cmbMarket"))
    )
finally: 
    driver.find_element_by_css_selector("#ctl00_cplhContainer_cmbMarket_Input").send_keys("T")
time.sleep(1)

driver.find_element_by_css_selector("#ctl00_cplhContainer_cmbMarket_Input").submit()
driver.find_element_by_css_selector("#ctl00_cplhContainer_cmbMarket_Input").send_keys(Keys.ENTER)

        
"""
"""
"""
#time.sleep(2)
driver.implicitly_wait(15)
driver.find_element_by_css_selector("#ctl00_cplhContainer_cmbMarket").click()
driver.find_element_by_xpath("/html/body/form/div[1]/div/div/ul/li[9]").click()
"""

#GOODS
driver.find_element_by_xpath("//span[contains(text(),'Goods')]").click()
time.sleep(1)
driver.find_element_by_xpath("/html/body/form/div[4]/div[2]/div[2]/div[7]/div[1]/div[2]/div[4]/div[2]/div[1]/div[1]/div[1]/div/table/tbody/tr/td[2]/a").click()
driver.find_element_by_css_selector("#ctl00_cplhContainer_cmbHSCode_Input").send_keys("87084092")
#time.sleep(3)
#driver.find_element_by_xpath("//li[contains(text(),'87084092 - - - - Dùng cho xe thuộc nhóm 87.03')]").click()
#87084092 - - - - Dùng cho xe thuộc nhóm 87.03
try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//li[contains(text(),'87084092 - - - - Dùng cho xe thuộc nhóm 87.03')]"))
    )
finally: driver.find_element_by_xpath("//li[contains(text(),'87084092 - - - - Dùng cho xe thuộc nhóm 87.03')]").click()
"""
"""
