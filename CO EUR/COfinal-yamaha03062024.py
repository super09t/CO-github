from numpy import var
import openpyxl
from asyncio import sleep, wait, wait_for
import time
import timeit
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from seleniumbase import SB
start_time = time.time()
username = "0500555909"
password = "0500555909"
read_excel = r'C:\Users\duong.ns\Desktop\CO EUR\filenhap.xlsx'       #khai báo file làm CO
wb = openpyxl.load_workbook(read_excel)
sheet = wb['Sheet1']            #Tên sheet file CO
g = sheet.iter_rows(min_row=1, max_row=99, min_col=1, max_col=10)
print(type(g))
# <class 'generator'>
cells_list=list(g)


#Sumdonghang = int(input()) - 1
Sumdonghang = int(cells_list[4][9].value)-1
print(str(int(cells_list[4][9].value)) + "dong hang")
invoice = cells_list[1][9].value
print("invoice: " + str(cells_list[1][9].value))
print("nhập ngày invoice: " + str(cells_list[3][9].value))
dateinvoice = cells_list[3][9].value
#driver = websb.Chrome(ChromeDriverManager().install())
#current_version = ChromeDriverManager().version
#is_latest = ChromeDriverManager().is_latest_version()
#if not is_latest:
    # Cài đặt trình điều khiển Chrome mới nhất
    #ChromeDriverManager().install()
# initialize the Chrome driver 
#driver = websb.Chrome("chromedriver")
with SB() as sb:
    
    
    # head to ecosy login page
    sb.get("https://dichvucong.moit.gov.vn/Login.aspx?clientkey=vGh3F8zA&url=https%3a%2f%2fecosys.gov.vn%3a443%2fvalidate.moitid")
    # find username/email field and send the username itself to the input field
    sb.find_element("#ctl00_cplhContainer_txtLoginName").send_keys(username)
    # find password input field and insert password as well
    sb.find_element("#ctl00_cplhContainer_txtPassword").send_keys(password)
    #maxacthuc = input()
    #sb.find_element(By.ID,"ctl00_cplhContainer_txtCaptcha").send_keys(maxacthuc)
    # click login button
    #sb.find_element(By.NAME,"ctl00$cplhContainer$btnLogin").click()
    #p = input()
    sb.find_element("#ctl00_Header1_lbtnLogin",timeout=20).click()
    # try:
    #     element = WebDriverWait(driver, 100).until(
    #         EC.presence_of_element_located((By.ID, "ctl00_Header1_lbtnLogin"))
    #     )
    # finally:
    #     sb.find_element(By.ID,"ctl00_Header1_lbtnLogin").click()
    #sb.find_element(By.ID,"ctl00_Header1_lbtnLogin").click()
    sb.wait_for_element_visible("#timer", timeout=30)
    sb.wait_for_element_clickable("#timer", timeout=30)
    sb.click("#timer")
    sb.find_element("#ctl00_Menu1_radMenu > ul > li.rtLI.rtFirst > ul > li:nth-child(1) > div > a").click()
    #click chon CO
    print("Bắt đầu tạo CO")
    print("Chọn CO form D")
    sb.find_element("#ctl00_cplhContainer_cmbFormCO_Input").send_keys("FFFFffffffffffff")
    start_time = time.time()
    sb.find_element("#ctl00_cplhContainer_cmbFormCO_Input").submit()
    print("đợi load nước xk")
    time.sleep(2)
    sb.find_element("#ctl00_cplhContainer_cmbMarket_Input").send_keys("II")
    #sb.find_element("#ctl00_cplhContainer_cmbMarket_DropDown > div > ul > li:nth-child(9)").click()
    
    print("chọn Italy")
    #p = input()
    #sb.find_element(By.CSS_SELECTOR,"#ctl00_cplhContainer_cmbReceiverPlace_Input").send_keys("FF")
    #sb.find_element(By.CSS_SELECTOR,"#ctl00_cplhContainer_cmbReceiverPlace_Input").submit()
    #sb.find_element(By.CSS_SELECTOR,"#ctl00_cplhContainer_plhCustomsNumber0_txtInvoiceNumber").send_keys("304576143600")
    print("Nhập số tk xuất")
    print("Nhập cty xnk")
    sb.find_element("#ctl00_cplhContainer_PersionNameExportEnglish").clear()
    sb.find_element("#ctl00_cplhContainer_PersionNameExportEnglish").send_keys(cells_list[10][9].value)
    print(str(cells_list[5][9].value))
    
    sb.find_element("#ctl00_cplhContainer_plhCustomsNumber0_radDpkInvoiceDate_dateInput").clear()
    sb.find_element("#ctl00_cplhContainer_plhCustomsNumber0_radDpkInvoiceDate_dateInput").send_keys(cells_list[6][9].value)
    sb.find_element("#ctl00_cplhContainer_txtShipName").send_keys(cells_list[7][9].value)
    sb.find_element("#ctl00_cplhContainer_AddressEnglishExport").clear()
    sb.find_element("#ctl00_cplhContainer_AddressEnglishExport").send_keys(cells_list[11][9].value)
    sb.find_element("#ctl00_cplhContainer_AddressEnglishExport2").send_keys(cells_list[12][9].value)
    sb.find_element("#ctl00_cplhContainer_PersionNameImportEnglish").send_keys(cells_list[13][9].value)
    sb.find_element("#ctl00_cplhContainer_AddressEnglishImport").send_keys(cells_list[14][9].value)
    #sb.find_element("#ctl00_cplhContainer_AddressEnglishImport2").send_keys(cells_list[15][9].value)


    sb.find_element("#ctl00_cplhContainer_plhCustomsNumber0_txtInvoiceNumber").submit()
    time.sleep(3)
    sb.find_element("#ctl00_cplhContainer_plhCustomsNumber0_txtInvoiceNumber").send_keys(cells_list[5][9].value)    
    sb.find_element("#ctl00_cplhContainer_RadTabStrip1 > div > ul > li:nth-child(2) > a").click()
    print("DONE==============")
    
    #===========
    #==============Cho vào vòng lặp===============

    
    #nhập tên hàng
    #sb.find_element(By.ID,"ctl00_cplhContainer_txtName_wrapper").click()
    sb.find_element("#ctl00_cplhContainer_txtName").send_keys(cells_list[1][2].value)
    print("HS 1: " + str(cells_list[1][1].value))
    print("Dòng hàng 1: " + cells_list[1][2].value)
    #nhập đơn vị khối lượng
    sb.find_element("#ctl00_cplhContainer_cmbGwUnitId_Input").clear()
    sb.find_element("#ctl00_cplhContainer_cmbGwUnitId_Input").send_keys("KILOGRAM")
    #print("Đơn vị trọng lượng: KGM ")
    #nhập đơn vị số lượng
    sb.find_element("#ctl00_cplhContainer_cmbUnit_Input").clear()
    sb.find_element("#ctl00_cplhContainer_cmbUnit_Input").send_keys("PIECE")
    #print("Đơn vị số lượng: PIECE ")
    #Nhập package
    #sb.find_element("#ctl00_cplhContainer_txtBoxValue").send_keys(cells_list[1][6].value)
    #print("Số kiện 1: " + str(cells_list[1][6].value))
    #nhập đơn vị PKG
    sb.find_element("#ctl00_cplhContainer_cmbBoxUnitId_Input").clear()
    sb.find_element("#ctl00_cplhContainer_cmbBoxUnitId_Input").send_keys("PACKAGE")
    #print("Đơn vị trọng kiện: PACKAGE ")
    #sb.find_element_by_link_text("PACKAGE").click()
    #nhập số lượng
    sb.find_element("#ctl00_cplhContainer_txtUnitValue").clear()
    sb.find_element("#ctl00_cplhContainer_txtUnitValue").send_keys(repr(cells_list[1][3].value))
    print("Số lượng 1: " + str(cells_list[1][3].value))
    #nhập trọng lượng
    #sb.find_element("#ctl00_cplhContainer_txtGwValue").clear()
    #sb.find_element("#ctl00_cplhContainer_txtGwValue").send_keys(repr(cells_list[1][4].value))
    #print("Trọng lượng 1: " + str(cells_list[1][4].value))
    #nhập invoice
    sb.find_element("#ctl00_cplhContainer_txtInvoiceItem").send_keys(invoice)
    #nhập ngay invoice
    sb.find_element("#ctl00_cplhContainer_radDpkInvoiceItemDate_dateInput").send_keys(dateinvoice)
    #nhập Mark

    sb.find_element("#ctl00_cplhContainer_txtShippingMark").send_keys(invoice + cells_list[9][9].value)
    #print("Mark: HAL ALUMINUM (THAILAND) CO.,LTD ")
    #nhập FOB

    sb.find_element("#ctl00_cplhContainer_txtCurrencyValue").send_keys(repr(cells_list[1][7].value))
    print("FOB 1: " + str(cells_list[1][7].value))
    #Nhập hs code
    sb.find_element("#ctl00_cplhContainer_cmbHSCode_Input").clear()
    sb.find_element("#ctl00_cplhContainer_cmbHSCode_Input").send_keys(cells_list[1][1].value)

    time.sleep(0.5)
    #Đợi Hs hiện ra thì click
    #time.sleep(2)
    old_element_hs = sb.find_element("#ctl00_cplhContainer_cmbHSCodeOutsite")
    sb.find_element("#ctl00_cplhContainer_cmbHSCode_DropDown > div.rcbScroll.rcbWidth.rcbNoWrap > ul > li").click()
    xpath = f"//li[contains(text(),'{cells_list[1][1].value}')]"
    try:
        # Đợi tối đa 10 giây cho đến khi phần tử được tìm thấy
        element = sb.wait_for_element_present(xpath, timeout=10)

        # Kiểm tra xem có phần tử nào được tìm thấy hay không
        if element:
            # Click vào phần tử đầu tiên tìm thấy
            sb.click(xpath)
        else:
            print("Không tìm thấy phần tử nào phù hợp.")

        # Thực hiện thao tác click vào phần tử cố định
        
    except Exception as e:
        print(f"Đã xảy ra lỗi: {e}")
    #p = input()
    #RVC
    #sb.find_element(By.XPATH,"/html/body/form/div[5]/div[2]/div[2]/div[7]/div[1]/div[2]/div[4]/div[2]/div[1]/div[2]/div[2]/div/img").click()
    #sb.find_element(By.XPATH,"/html/body/form/div[4]/div/div[2]/table/tbody/tr[7]/td[2]/input").click()
    #valueephs = sb.find_element(By.NAME,"ctl00$cplhContainer$cmbHSCode").get_attribute("value")
    #valueiphs = sb.find_element(By.NAME,"ctl00$cplhContainer$cmbHSCodeOutsite").get_attribute("value")
    #valueephs = sb.find_element(By.CSS_SELECTOR,"#aspnetForm > div.rcbSlide")
    #valueiphs = sb.find_element(By.CSS_SELECTOR,"#aspnetForm > div:nth-child(1)")
    #while (valueiphs != valueephs):
        #valueephs = sb.find_element(By.CSS_SELECTOR,"#aspnetForm > div.rcbSlide")
        #valueiphs = sb.find_element(By.CSS_SELECTOR,"#aspnetForm > div:nth-child(1)")
    # time.sleep(0.3)
    wait = WebDriverWait(sb, 10)

    wait.until(EC.staleness_of(old_element_hs))
    print("Nhập RVC (1): " + str(cells_list[1][5].value))


    #sb.execute_script('document.querySelector("#ctl00_cplhContainer_rpvGoods > div:nth-child(1) > div.col-right > div:nth-child(2) > div > img").click();')
    #sb.execute_script('document.querySelector("#RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria > table > tbody > tr.rwContentRow > td.rwWindowContent.rwExternalContent > iframe").contentWindow.document.getElementById("txtRVC").value = "88%";document.querySelector("#RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria > table > tbody > tr.rwContentRow > td.rwWindowContent.rwExternalContent > iframe").contentWindow.document.getElementById("chkRVC").click();document.querySelector("#ctl00_cplhContainer_radToolBarDefault > div > div > div > ul > li:nth-child(1)").click()')
    #sb.execute_script('document.querySelector("#ctl00_cplhContainer_radToolBarDefault > div > div > div > ul > li:nth-child(1)").click()')
    #####
    sb.find_element(By.CSS_SELECTOR,"#ctl00_cplhContainer_rpvGoods > div:nth-child(1) > div.col-right > div:nth-child(2) > div > img").click()
    #wait.until(EC.element_to_be_clickable((By.ID, "RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria")))
    
    frame2 = sb.find_element("#RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria")
        
    # switch to frame by frame element
    #sb.switch_to_frame("RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria")
    sb.switch_to_frame('iframe[name="PopupSelectOriginCriteria"]')
    input_xpath = "//input[@id='txtRVC']"

    # Kiểm tra xem phần tử có được kích hoạt hay không
    
    sb.find_element("#chkPSR").click()
    

    sb.find_element("#ctl00_cplhContainer_radToolBarDefault > div > div > div > ul > li:nth-child(1) > a > span > span > span").click()
    sb.switch_to_parent_frame()
    
    time.sleep(0.3)
    old_element = sb.find_element("#ctl00_cplhContainer_radGridItemUpgrade_ctl00")
    sb.find_element("#ctl00_cplhContainer_btnAddItem").click()
    #loadnew = sb.find_element(By.NAME,"ctl00$cplhContainer$txtBoxValue").get_attribute("value")
    #while (loadnew != "0"):
        #loadnew = sb.find_element(By.NAME,"ctl00$cplhContainer$txtBoxValue").get_attribute("value")
        #time.sleep(0.3)
    wait = WebDriverWait(sb, 10)

    wait.until(EC.staleness_of(old_element))
    print("=========Continue====")

    for i in range(Sumdonghang):
        if cells_list[i+2][1].value == cells_list[i+1][1].value :
            sb.find_element(By.CSS_SELECTOR,"#ctl00_cplhContainer_RadTabStrip1 > div > ul > li:nth-child(2) > a").click()
            #nhập tên hàng
            print("HS " + str(i+2) +": " + str(cells_list[i+2][1].value))
            #sb.find_element(By.ID,"ctl00_cplhContainer_txtName_wrapper").click()
            sb.find_element("#ctl00_cplhContainer_txtName").clear
            sb.find_element("#ctl00_cplhContainer_txtName").send_keys(cells_list[i+2][2].value)
            print("Dòng hàng " + str(i+2) +": " + str(cells_list[i+2][2].value))
            #nhập đơn vị khối lượng
        # sb.find_element(By.NAME,"ctl00$cplhContainer$cmbGwUnitId").clear()
            #sb.find_element(By.NAME,"ctl00$cplhContainer$cmbGwUnitId").send_keys("KILOGRAM")
            #nhập đơn vị số lượng
        # sb.find_element(By.NAME,"ctl00$cplhContainer$cmbUnit").clear()
        # sb.find_element(By.NAME,"ctl00$cplhContainer$cmbUnit").send_keys("PIECE")
            #Nhập package
            sb.find_element("#ctl00_cplhContainer_txtBoxValue").send_keys(repr(cells_list[i+2][6].value))
            print("Số kiện " + str(i+2) +": " + str(cells_list[i+2][6].value))
            #nhập đơn vị PKG
        # sb.find_element(By.NAME,"ctl00$cplhContainer$cmbBoxUnitId").clear()
        # sb.find_element(By.NAME,"ctl00$cplhContainer$cmbBoxUnitId").send_keys("PACKAGE")
            #sb.find_element_by_link_text("PACKAGE").click()
            #nhập số lượng
            sb.find_element("#ctl00_cplhContainer_txtUnitValue").send_keys(repr(cells_list[i+2][3].value))
            print("Số lượng " + str(i+2) +": " + str(cells_list[i+2][3].value))
            #nhập trọng lượng
            sb.find_element("#ctl00_cplhContainer_txtGwValue").send_keys(repr(cells_list[i+2][4].value))
            print("Trọng lượng " + str(i+2) +": " + str(cells_list[i+2][4].value))
            #nhập invoice
            #sb.find_element(By.NAME,"ctl00$cplhContainer$txtInvoiceItem").send_keys("invoi")
            #nhập ngay invoice
        # sb.find_element(By.NAME,"ctl00$cplhContainer$radDpkInvoiceItemDate$dateInput").send_keys("21022022")
            #nhập Mark
        # sb.find_element(By.NAME,"ctl00$cplhContainer$txtShippingMark").send_keys("NoMark")
            #nhập FOB
            sb.find_element("#ctl00_cplhContainer_txtCurrencyValue").send_keys(repr(cells_list[i+2][7].value))
            print("FOB " + str(i+2) +": " + str(cells_list[i+2][7].value))
            #Nhập hs code
            """
            sb.find_element(By.NAME,"ctl00$cplhContainer$cmbHSCode").clear()
            sb.find_element(By.NAME,"ctl00$cplhContainer$cmbHSCode").send_keys(repr(cells_list[i+2][1].value))
            #Đợi Hs hiện ra thì click
            #time.sleep(2)
            try:
                element = WebDriverWait(driver,10).until(
                    EC.presence_of_element_located((By.XPATH, "//li[contains(text(),'87084092 - - - - Dùng cho xe thuộc nhóm 87.03')]"))
                )
            finally:sb.find_element(By.XPATH,"//li[contains(text(),'87084092 - - - - Dùng cho xe thuộc nhóm 87.03')]").click()
            """
            #RVC

            #sb.find_element(By.XPATH,"/html/body/form/div[5]/div[2]/div[2]/div[7]/div[1]/div[2]/div[4]/div[2]/div[1]/div[2]/div[2]/div/img").click()
            #sb.find_element(By.XPATH,"/html/body/form/div[4]/div/div[2]/table/tbody/tr[7]/td[2]/input").click()
            #time.sleep(2)
            #sb.execute_script('document.querySelector("#ctl00_cplhContainer_rpvGoods > div:nth-child(1) > div.col-right > div:nth-child(2) > div > img").click();')
            #time.sleep(1)
            #sb.execute_script('document.querySelector("#RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria > table > tbody > tr.rwContentRow > td.rwWindowContent.rwExternalContent > iframe").contentWindow.document.getElementById("txtRVC").value = "88%";document.querySelector("#RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria > table > tbody > tr.rwContentRow > td.rwWindowContent.rwExternalContent > iframe").contentWindow.document.getElementById("chkRVC").click();document.querySelector("#ctl00_cplhContainer_radToolBarDefault > div > div > div > ul > li:nth-child(1)").click()')
            #sb.execute_script('document.querySelector("#ctl00_cplhContainer_radToolBarDefault > div > div > div > ul > li:nth-child(1)").click()')
            #####
            
            # switch to frame by frame element
            
            old_element = sb.find_element("#ctl00_cplhContainer_radGridItemUpgrade_ctl00")
            sb.find_element("#ctl00_cplhContainer_btnAddItem").click()
            #loadnew = sb.find_element(By.NAME,"ctl00$cplhContainer$txtBoxValue").get_attribute("value")
            #while (loadnew != "0"):
                #loadnew = sb.find_element(By.NAME,"ctl00$cplhContainer$txtBoxValue").get_attribute("value")
                #time.sleep(0.3)
            wait = WebDriverWait(sb, 10)
            
            wait.until(EC.staleness_of(old_element))
            
            print("=========Continue====")
            #84839095 - - - - Dùng cho hàng hóa khác thuộc Chương 87
            #84099949 - - - - - Loại khác
        else :
                sb.find_element(By.CSS_SELECTOR,"#ctl00_cplhContainer_RadTabStrip1 > div > ul > li:nth-child(2) > a").click()
                #nhập tên hàng
                #sb.find_element(By.ID,"ctl00_cplhContainer_txtName_wrapper").click()
                sb.find_element("#ctl00_cplhContainer_txtName").clear
                sb.find_element("#ctl00_cplhContainer_txtName").send_keys(cells_list[i+2][2].value)
                print("HS " + str(i+2) +": " + str(cells_list[i+2][1].value))
                print("Dòng hàng " + str(i+2) +": " + str(cells_list[i+2][2].value))
                #nhập đơn vị khối lượng
            # sb.find_element(By.NAME,"ctl00$cplhContainer$cmbGwUnitId").clear()
                #sb.find_element(By.NAME,"ctl00$cplhContainer$cmbGwUnitId").send_keys("KILOGRAM")
                #nhập đơn vị số lượng
            # sb.find_element(By.NAME,"ctl00$cplhContainer$cmbUnit").clear()
            # sb.find_element(By.NAME,"ctl00$cplhContainer$cmbUnit").send_keys("PIECE")
                #Nhập package
                sb.find_element("#ctl00_cplhContainer_txtBoxValue").send_keys(repr(cells_list[i+2][6].value))
                print("Số kiện " + str(i+2) +": " + str(cells_list[i+2][6].value))
                #nhập đơn vị PKG
            # sb.find_element(By.NAME,"ctl00$cplhContainer$cmbBoxUnitId").clear()
            # sb.find_element(By.NAME,"ctl00$cplhContainer$cmbBoxUnitId").send_keys("PACKAGE")
                #sb.find_element_by_link_text("PACKAGE").click()
                #nhập số lượng
                sb.find_element("#ctl00_cplhContainer_txtUnitValue").send_keys(repr(cells_list[i+2][3].value))
                print("Số lượng " + str(i+2) +": " + str(cells_list[i+2][3].value))
                #nhập trọng lượng
                sb.find_element("#ctl00_cplhContainer_txtGwValue").send_keys(repr(cells_list[i+2][4].value))
                print("Trọng lượng " + str(i+2) +": " + str(cells_list[i+2][4].value))
                #nhập invoice
                #sb.find_element(By.NAME,"ctl00$cplhContainer$txtInvoiceItem").send_keys("invoi")
                #nhập ngay invoice
            # sb.find_element(By.NAME,"ctl00$cplhContainer$radDpkInvoiceItemDate$dateInput").send_keys("21022022")
                #nhập Mark
            # sb.find_element(By.NAME,"ctl00$cplhContainer$txtShippingMark").send_keys("NoMark")
                #nhập FOB
                
                sb.find_element("#ctl00_cplhContainer_txtCurrencyValue").send_keys(repr(cells_list[i+2][7].value))
                print("FOB " + str(i+2) +": " + str(cells_list[i+2][7].value))
                #Nhập hs code
                #ctl00_cplhContainer_cmbHSCode_Arrow
                sb.find_element("#ctl00_cplhContainer_cmbHSCode_Arrow").click()
                sb.find_element("#ctl00_cplhContainer_cmbHSCode_Input").clear()
                sb.find_element("#ctl00_cplhContainer_cmbHSCode_Input").send_keys(repr(cells_list[i+2][1].value))
                #time.sleep(2)
                old_element_hs = sb.find_element("#ctl00_cplhContainer_cmbHSCodeOutsite")
                
                xpath = f"//li[contains(text(),'{cells_list[i+2][1].value}')]"
                try:
                    # Đợi tối đa 10 giây cho đến khi phần tử được tìm thấy
                    element = sb.wait_for_element_present(xpath, timeout=10)

                    # Kiểm tra xem có phần tử nào được tìm thấy hay không
                    if element:
                        # Click vào phần tử đầu tiên tìm thấy
                        sb.click(xpath)
                    else:
                        print("Không tìm thấy phần tử nào phù hợp.")

                    # Thực hiện thao tác click vào phần tử cố định
                    
                except Exception as e:
                    print(f"Đã xảy ra lỗi: {e}")
                """
                84839095 - - - - Dùng cho hàng hóa khác thuộc Chương 87
                #or EC.presence_of_element_located((By.XPATH, "//li[contains(text(),'84099949 - - - - - Loại khác')]"))
                /html/body/form/div[1]/div/div[1]/ul/li[1]
                #Đợi Hs hiện ra thì click
                #time.sleep(2)
                try:
                    element = WebDriverWait(driver,10).until(
                        EC.presence_of_element_located((By.XPATH, "//li[contains(text(),'87084092 - - - - Dùng cho xe thuộc nhóm 87.03')]"))
                    )
                finally:sb.find_element(By.XPATH,"//li[contains(text(),'87084092 - - - - Dùng cho xe thuộc nhóm 87.03')]").click()
                """
                #RVC
                #sb.find_element(By.XPATH,"/html/body/form/div[5]/div[2]/div[2]/div[7]/div[1]/div[2]/div[4]/div[2]/div[1]/div[2]/div[2]/div/img").click()
                #sb.find_element(By.XPATH,"/html/body/form/div[4]/div/div[2]/table/tbody/tr[7]/td[2]/input").click()
                
                #sb.execute_script('document.querySelector("#ctl00_cplhContainer_rpvGoods > div:nth-child(1) > div.col-right > div:nth-child(2) > div > img").click();')
                #valueephs = sb.find_element(By.NAME,"ctl00$cplhContainer$cmbHSCode").get_attribute("value")
                #valueiphs = sb.find_element(By.NAME,"ctl00$cplhContainer$cmbHSCodeOutsite").get_attribute("value")
                #while (valueiphs != valueephs):
                    #valueephs = sb.find_element(By.NAME,"ctl00$cplhContainer$cmbHSCode").get_attribute("value")
                    #valueiphs = sb.find_element(By.NAME,"ctl00$cplhContainer$cmbHSCodeOutsite").get_attribute("value")
                    #time.sleep(0.3)
                wait = WebDriverWait(sb, 10)
                
                wait.until(EC.staleness_of(old_element_hs))
                #print("Nhập RVC " + str(i+2)+": " + cells_list[i+2][5].value)
                #sb.execute_script('document.querySelector("#RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria > table > tbody > tr.rwContentRow > td.rwWindowContent.rwExternalContent > iframe").contentWindow.document.getElementById("txtRVC").value = "88%";document.querySelector("#RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria > table > tbody > tr.rwContentRow > td.rwWindowContent.rwExternalContent > iframe").contentWindow.document.getElementById("chkRVC").click();document.querySelector("#ctl00_cplhContainer_radToolBarDefault > div > div > div > ul > li:nth-child(1)").click()')
                #sb.execute_script('document.querySelector("#ctl00_cplhContainer_radToolBarDefault > div > div > div > ul > li:nth-child(1)").click()')
                #####
                #sb.find_element(By.CSS_SELECTOR,"#ctl00_cplhContainer_rpvGoods > div:nth-child(1) > div.col-right > div:nth-child(2) > div > img").click()
                #wait.until(EC.element_to_be_clickable((By.ID, "RadWindowWrapper_ctl00_cplhContainer_PopupSelectOriginCriteria")))
                
                
                time.sleep(0.3)
                old_element = sb.find_element("#ctl00_cplhContainer_radGridItemUpgrade_ctl00")
                sb.find_element("#ctl00_cplhContainer_btnAddItem").click()
                #loadnew = sb.find_element(By.NAME,"ctl00$cplhContainer$txtBoxValue").get_attribute("value")
                #while (loadnew != "0"):
                    #loadnew = sb.find_element(By.NAME,"ctl00$cplhContainer$txtBoxValue").get_attribute("value")
                    #time.sleep(0.3)
                wait = WebDriverWait(sb, 10)
                
                wait.until(EC.staleness_of(old_element))
               
                print("=========Continue====")
                #84839095 - - - - Dùng cho hàng hóa khác thuộc Chương 87
                #84099949 - - - - - Loại khác
   
    p = input()
stop_time = time.time() 
elapsed_time = stop_time - start_time
print("Thời gian làm CO là: ")          
print ("elapsed_time:{0}".format(elapsed_time) + "[sec]")
