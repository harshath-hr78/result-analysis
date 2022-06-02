from selenium import webdriver
import time

option = webdriver.ChromeOptions()
option.add_argument("-incognito")
option.add_experimental_option("excludeSwitches", ['enable-automation']);

browser = webdriver.Chrome(executable_path='C:\Program Files (x86)\chromedriver.exe', options=option)

browser.get("https://results.vtu.ac.in/FMEcbcs22/resultpage.php")


textboxes = browser.find_elements_by_class_name("form-control")

otherboxes = browser.find_element_by_xpath("/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[1]/div/input")
time.sleep(2)
textboxes[0].send_keys("1AM18CS067")




