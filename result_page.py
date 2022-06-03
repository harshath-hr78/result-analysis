from distutils.log import error
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import pyautogui
import cv2 as cv
import pytesseract
from bs4 import BeautifulSoup
import pandas as pd
import requests

# install all these as pip install filename, and pip install opencv-python.

option = webdriver.ChromeOptions()
option.add_argument("-incognito")
option.add_experimental_option("excludeSwitches", ['enable-automation'])
option.add_experimental_option("detach",True)

#add your chrome driver installation path
browser = webdriver.Chrome(executable_path=r'C:\Program Files (x86)\chromedriver.exe', options=option)

def fillLoginpage():

    browser.get("https://results.vtu.ac.in/FMEcbcs22/resultpage.php")

    #getting hold of usn and captcha input fields.

    testbox = browser.find_element_by_xpath("/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[1]/div/input")
    captchabox = browser.find_element_by_xpath("/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[2]/div[1]/input")

    #start with the image capta recognition procedure
    time.sleep(2)
    myScreenshot = pyautogui.screenshot(region=(670,510,230,110)) #region=(horizontal pos, vertical pos, vertical ratio, horizontal ratio)
    myScreenshot.save(r'C:\Users\harsh\Desktop\result_analysis\pics\screenshot.png') #change according to your dir.

    os.chdir(r"C:\Users\harsh\Desktop\result_analysis\pics")
    img = cv.imread('screenshot.png',0)
    ret,thresh = cv.threshold(img,103,150,cv.THRESH_TOZERO_INV)
    cv.imshow('Binary Threshold', thresh)
    # Using cv2.imwrite() method
    # Saving the image
    os.chdir(r'C:\Users\harsh\Desktop\result_analysis\pics')
    cv.imwrite("thresh_img.png", thresh)

    time.sleep(1)
    #os.system('"wsl tesseract thresh_img.jpg result"') #tesseract is ocr function for image to text
    img2 = cv.imread('thresh_img.png',0)
    #install tesseract from https://github.com/UB-Mannheim/tesseract/wiki choose 64-bit 
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' 
    #depends on your tesseract installation path
    custom_config = r'--oem 3 --psm 6'
    captcha = pytesseract.image_to_string(img2, config=custom_config)
    
    captcha.replace(" ", "").strip()

    print("Captcha printing " +captcha)
    print(len(captcha)-1)
    if(len(captcha)-1 != 6 ):
        fillLoginpage()

    #finally input the result pages with required info.
    time.sleep(1)
    try:
        testbox.send_keys("1AM18CS077")
        captchabox.send_keys(captcha) 
    except:
        error
    try:
        print(browser.current_url)
    except:
        fillLoginpage()
    
    #copy the full XPATH for the required cell and add the below code to get the data
    sub_codes = ["18ME751", "18CS71", "18CS72","18CS744","18CS734","18CSL76","18CSP77"]
    rows = []

    for sub_code in sub_codes:
        subject = browser.find_element_by_xpath("//*[@id='dataPrint']//*[contains(text(),'"+sub_code+"')]//following::div[1]").text
        internal_marks = browser.find_element_by_xpath("//*[@id='dataPrint']//*[contains(text(),'"+sub_code+"')]//following::div[2]").text
        external_marks = browser.find_element_by_xpath("//*[@id='dataPrint']//*[contains(text(),'"+sub_code+"')]//following::div[3]").text
        total_marks = browser.find_element_by_xpath("//*[@id='dataPrint']//*[contains(text(),'"+sub_code+"')]//following::div[4]").text
        remarks = browser.find_element_by_xpath("//*[@id='dataPrint']//*[contains(text(),'"+sub_code+"')]//following::div[5]").text

        present_row_data={'Subject Code': sub_code,
                   'Subject Name': subject,
                   'Internal Marks': internal_marks,
                   'External Marks': external_marks,
                   'Total': total_marks,
                   'Remarks': remarks }
        rows.append(present_row_data)
    
    final_result_data = pd.DataFrame(rows)                              #import pandas as pd
    final_result_data.to_excel('vtu_result.xlsx',index=False)
        

    time.sleep(100)

fillLoginpage()



    
    






