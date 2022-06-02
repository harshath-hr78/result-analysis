from selenium import webdriver
import time

option = webdriver.ChromeOptions()
option.add_argument("-incognito")
option.add_experimental_option("excludeSwitches", ['enable-automation']);

browser = webdriver.Chrome(executable_path='C:\Program Files (x86)\chromedriver.exe', options=option)

browser.get("https://www.vtu4u.com/")


textboxes = browser.find_elements_by_id("usn")
cbcsclick = browser.find_elements_by_id("syl3")
submitclick = browser.find_element_by_class_name("btn-home-search")



time.sleep(2)
textboxes[0].send_keys("1AM18CS067")
cbcsclick[0].click()
time.sleep(1)
submitclick.click()





