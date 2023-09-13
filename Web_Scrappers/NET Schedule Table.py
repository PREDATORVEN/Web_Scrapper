from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
import urllib.request
from selenium.webdriver.common.by import By
import base64
import os, argparse
import pytesseract
from pytesseract import pytesseract  # library for pytesseract
from PIL import Image
import cv2
import easyocr
import io, base64
from PIL import Image
import numpy
import time
import datetime
from datetime import date
import glob
import win32com.client as win32
path_to_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # path for .exe file 
delta = 0
#-------- function 

#--------
def date_setting(today):
    print(today)
    target_date = str(today + datetime.timedelta(days=delta))
    today = str(today)
    driver.find_element(By.ID, "scheduleDate").click()
    targetlst = target_date.split("-")
    todaylst = today.split("-")
    today = todaylst[2] + "-" + todaylst[1] + "-" + todaylst[0]
    target_date = targetlst[2] + "-" + targetlst[1] + "-" + targetlst[0]

    if (today.split("-")[1] != target_date.split("-")[1]):
        driver.find_element(By.XPATH, "/html/body/div[3]/div/a[2]").click()

    tble = driver.find_element(By.XPATH, '/html/body/div[3]/table/tbody')
    tareekh = tble.find_elements(By.TAG_NAME, 'td')

    no_date = 0
    for tr in tareekh:
        if tr.text not in [str(item) for item in range(1, 10)]:
            if (tr.text == target_date[0:2]):
                no_date += 1
        else:
            if ('0' + tr.text == target_date[0:2]):
                no_date += 1

    for tr in tareekh:
        if tr.text not in [str(item) for item in range(1, 10)]:
            if (tr.text == target_date[0:2]) and no_date == 1:
                tr.click()
                break
            elif (tr.text == target_date[0:2]) and no_date == 2:
                no_date = 1
        else:
            if ('0' + tr.text == target_date[0:2]):
                tr.click()
                break


def launch():
    try:
        # Clicking on Details Button
        driver.find_element(By.ID, 'byDetails').click()
        time.sleep(1)

        #setting date
        today = date.today()
        date_setting(today)
        time.sleep(2)

        # Region DropDown
        region1 = driver.find_element(By.ID, 'ddlRegion')
        drop_down1 = Select(region1)
        drop_down1.select_by_visible_text('NORTH')
        time.sleep(2)

        # Revision DropDown
        # region2 = driver.find_element(By.ID, 'ddlRevision')
        # drop_down2 = Select(region2)
        # drop_down2.select_by_value('151')
        # time.sleep(2)

        # Buyer DropDown
        region3 = driver.find_element(By.ID, 'ddl_Seller')
        drop_down3 = Select(region3)
        drop_down3.select_by_visible_text('ALL')
        time.sleep(5)

        # Show Data Button
        driver.find_element(By.ID, 'SubmitBtn').click()
        time.sleep(5)

        # Selecting image and taking screenshot
        img_ele = driver.find_element(By.XPATH, '/html/body/div[5]/div[2]/div/div/div/img')
        time.sleep(2)
        ss = img_ele.screenshot_as_base64
        time.sleep(2)
        print("screen shot taken for captcha 1!!")
# -------------------------------------------------working till here------------------
# -------------------------------------- changing this block to save and read the screenshot--
        # img = Image.open(io.BytesIO(base64.decodebytes(bytes(ss, "utf-8"))))
        # im = numpy.array(img)
        # reader = easyocr.Reader(['en'], gpu=False)
        # results = reader.readtext(im)
        # print(results[0][1])
        
        png_screenshot = base64.b64decode(ss)
        with open('D:\\YASH STUFF\\Python\CSV_LOADER\\captcha_one.jpg', 'wb') as f:
            f.write(png_screenshot)
        print("captcha 1 SAVED !!")
        image_path = r"D:\YASH STUFF\Python\CSV_LOADER\captcha_one.jpg"  
        img = Image.open(image_path)
        pytesseract.tesseract_cmd = path_to_tesseract
        text_one = pytesseract.image_to_string(img)
        print("captcha 1 :",text_one)
        captcha_field = driver.find_element(By.ID, 'txtCaptcha') 
        captcha_field.send_keys("o"+text_one)
        time.sleep(2)
        print("Captcha_one entered !!")

#------------------------------------------------------------------------------------------
        # Filling the captcha
        # driver.find_element(By.ID, 'txtCaptcha').send_keys(results[0][1])
        # time.sleep(2)

        # Clicking the submit button
        driver.find_element(By.XPATH, '/html/body/div[5]/div[3]/div/button/span').click()
        time.sleep(8)

        # Clicking the download button
        driver.find_element(By.XPATH,
                            '/html/body/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[1]/div[2]/a/i').click()
        time.sleep(2)

        # Clicking the excel button
        # driver.find_element(By.ID, 'XlsExport').click()
        driver.find_element(By.ID, 'CsvExport').click()
        time.sleep(2)
#-----------------------------------------changing code for captcha 2 -----
        # Second Image
        img_ele = driver.find_element(By.XPATH, '/html/body/div[6]/div[2]/div/div/div/img')
        time.sleep(5)
        ss = img_ele.screenshot_as_base64
        time.sleep(2)
        print("screen shot taken for captcha 2")
    #----------------------------------changed block----------------
        # img = Image.open(io.BytesIO(base64.decodebytes(bytes(ss, "utf-8"))))
        # im = numpy.array(img)
        # reader = easyocr.Reader(['en'], gpu=False)
        # results = reader.readtext(im)

        # print(results[0][1])
        png_screenshot = base64.b64decode(ss)
        with open('D:\\YASH STUFF\\Python\\CSV_LOADER\\captcha_two.jpg', 'wb') as f:
            f.write(png_screenshot)
        print("captcha 2 SAVED !!")
        image_path = r"D:\YASH STUFF\Python\CSV_LOADER\captcha_two.jpg"  
        img = Image.open(image_path)
        pytesseract.tesseract_cmd = path_to_tesseract
        text_two = pytesseract.image_to_string(img)
        print("captcha 2 :",text_two)
        captcha_field = driver.find_element(By.ID, 'txtCaptcha') 
        captcha_field.send_keys(text_two)
        time.sleep(2)
        print("Captcha_two entered !!")

    #-----------------------------------------------------------------
        # Filling the captcha
        # driver.find_element(By.ID, 'txtCaptcha').send_keys(results[0][1])
        # time.sleep(2)
    #-----------------------------------------------------------------

        # Clicking the submit button
        driver.find_element(By.XPATH, '/html/body/div[6]/div[3]/div/button/span').click()
        time.sleep(60)
        downloads_folder = r"C:\Users\Yash Sharma\Downloads"
        if not os.path.isdir(downloads_folder):
            raise ValueError("The provided path is not a valid directory.")
        csv_files = glob.glob(os.path.join(downloads_folder, '*.csv'))

        # Check if any CSV files were found
        if not csv_files:
            print(" none")  
        # Sort CSV files by modification time in descending order
        csv_files.sort(key=os.path.getmtime, reverse=True)
        latest_csv_file =csv_files[0]
        if latest_csv_file:
            print(f"The latest downloaded CSV file is: {latest_csv_file}")
        else:
            print("No CSV files found in the downloads folder.")



        # Replace these paths with the actual input and output paths
        original_csv_path = latest_csv_file
        file_name=original_csv_path.split('\\')
        file_name=(file_name[-1].split('.'))[0]

        new_csv_path = f"D:\\YASH STUFF\\NET_SCHEDULE_FILE\\{file_name}.csv"

        excel = win32.Dispatch('Excel.Application')
        excel.Visible = True  # You can set this to False if you don't want Excel to be visible

        # Open the CSV file
        workbook = excel.Workbooks.Open(original_csv_path)

        # Perform any additional actions or modifications in Excel if needed

        # Save the file as a new CSV with the specified format and encoding
        workbook.SaveAs(new_csv_path, FileFormat=6)  # FileFormat 6 corresponds to 'CSV (Comma delimited)'
        workbook.Close(SaveChanges=True)

        print(f"CSV file saved as '{new_csv_path}' in Excel with format 'CSV (Comma delimited)'.")

        # Close and quit Excel
        excel.Quit()
        print("Conversion completed.")

        print("File downloaded at", f"D:\\YASH STUFF\\NET_SCHEDULE_FILE\\{file_name}.csv")
        print("RUN AGAIN")
    except:
        launch()

print("Accessing driver")
# driver = webdriver.Chrome(ChromeDriverManager().install())
driver = webdriver.Chrome()
print("trying to open")
driver.get("https://wbes.nrldc.in/ReportNetSchedule/GetNetScheduleIndex")
print("opened")
time.sleep(5)
launch()
