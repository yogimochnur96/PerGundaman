import time
from datetime import datetime
start_time = datetime.now()

from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.remote.webelement import WebElement

import XLUtils
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

options = Options()
options.headless = True
driver = webdriver.Chrome(executable_path="chromedriver.exe", chrome_options=options)

#driver = webdriver.Chrome(executable_path="D:\cekgundam\chromedriver.exe")

driver.maximize_window()

path = "DataAkun.xlsx"
rows = XLUtils.getRowCount(path, 'Sheet1')

for r in range(2, rows + 1):
    username = XLUtils.readData(path, "Sheet1", r, 1)
    driver.get("http://twitter.com/{}".format(username))
    time.sleep(7)
    try:
        try:
            Active = "(//*[@class='css-18t94o4 css-1dbjc4n r-1niwhzg r-p1n3y5 r-sdzlij r-1phboty r-rs99b7 r-1w2pmg r-1vuscfd r-1dhvaqw r-1fneopy r-o7ynqc r-6416eg r-lrvibr'])"
            driver.find_element_by_xpath(Active)
            print ("========== Active Sayang ==========")
            print (username," Active Say !!!")
            print ("")
            XLUtils.writeData(path,"Sheet1",r,2,"Active")
        except:
            try:
                Locked = "(//*[@class='css-18t94o4 css-1dbjc4n r-1niwhzg r-p1n3y5 r-sdzlij r-1phboty r-rs99b7 r-1w2pmg r-1jgb5lz r-1x0uki6 r-1vuscfd r-1dhvaqw r-1fneopy r-o7ynqc r-6416eg r-lrvibr'])"
                driver.find_element_by_xpath(Locked)
                print ("========== Locked Sayang ==========")
                print (username," Locked Say !!!")
                print ("")
                XLUtils.writeData(path,"Sheet1",r,2,"Locked")
            except:
                try:
                    Suspend = "(//*[@class='css-4rbku5 css-18t94o4 css-901oao css-16my406 r-1n1174f r-1loqt21 r-1qd0xha r-ad9z0x r-bcqeeo r-qvutc0'])"
                    driver.find_element_by_xpath(Suspend)
                    print ("========== Suspend Sayang ==========")
                    print (username, " Suspend Say !!!")
                    print ("")
                    XLUtils.writeData(path, "Sheet1", r, 2, "Suspend")
                except:
                    try:
                        AkunNotFound = "(//*[@class='css-901oao r-1re7ezh r-1qd0xha r-a023e6 r-16dba41 r-ad9z0x r-bcqeeo r-q4m81j r-ey96lj r-qvutc0'])"
                        driver.find_element_by_xpath(AkunNotFound)
                        print ("========== Not Found Sayang ==========")
                        print (username, " Not Found Say !!!")
                        print ("")
                        XLUtils.writeData(path, "Sheet1", r, 2, "Not Found")
                    except:
                        raise Exception(" ")

    except Exception,e:
        print str(e)

end_time = datetime.now()
print('Waktu Eksekusi : {}'.format(end_time - start_time))