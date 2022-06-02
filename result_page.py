from selenium import webdriver
import time
import os
import pyautogui
import cv2 as cv
import pytesseract
# install all these as pip install filename, and pip install opencv-python.

option = webdriver.ChromeOptions()
option.add_argument("-incognito")
option.add_experimental_option("excludeSwitches", ['enable-automation']);

#add your chrome driver installation path
browser = webdriver.Chrome(executable_path=r'C:\Program Files (x86)\chromedriver.exe', options=option)
browser.get("https://results.vtu.ac.in/FMEcbcs22/resultpage.php")

#getting hold of usn and captcha input fields.

testbox = browser.find_element_by_xpath("/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[1]/div/input")
captchabox = browser.find_element_by_xpath("/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[2]/div[1]/input")
submitclick = browser.find_element_by_xpath("/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[3]/div[1]/input")


#start with the image capta recognition procedure
time.sleep(4)
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



time.sleep(2)
#os.system('"wsl tesseract thresh_img.jpg result"') #tesseract is ocr function for image to text
img2 = cv.imread('thresh_img.png',0)
#install tesseract from https://github.com/UB-Mannheim/tesseract/wiki choose 64-bit 
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' 
#depends on your tesseract installation path
custom_config = r'--oem 3 --psm 6'
captcha = pytesseract.image_to_string(img2, config=custom_config)
print(captcha)

#finally input the result pages with required info.
time.sleep(2)
testbox.send_keys("1AM18CS092")
captchabox.send_keys(captcha) 

submitclick.click()
