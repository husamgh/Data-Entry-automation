#
#  This script is called whenever I want to fill the 
#   ------Google spread sheet------------- 
#
#&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
#https://techwithtim.net/tutorials/google-sheets-python-api-tutorial/
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import time
import sys
import fnmatch,os
import PyPDF2
import re
import textwrap3
from alphabet_detector import AlphabetDetector # to detect the language of the charecter
import arabic_reshaper #to process the arabic languagehttps://github.com/mpcabd/python-arabic-reshaper
from bidi.algorithm import get_display
from textwrap3 import wrap # to wrap the long ext in image
from PIL import Image # for image processing
from PIL import Image, ImageDraw, ImageFont #same for image processing 
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
#-------------------------------------------------------------------------------------------------
#------------------------for the google spread sheets---------------------------------------------
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)
sheet = client.open("dar-lusail").get_worksheet(0)  # Open the spreadhseet, first page 
#-------------------------------------

#print(Name_cells)
Pdf_path=r'C:\\Users\\user\\Desktop\\darlusail\\pdf\\'
browser=webdriver.Chrome()
#------------Update the book pages on the google sheet---------------------------------
def Pages_Nb(x):
 print("setting the page number of row:",x," on the google sheet....")
 Name_cells=sheet.cell(x,19).value
 pdf_file = open(Pdf_path+Name_cells,'rb')
 read_pdf = PyPDF2.PdfFileReader(pdf_file)
 book_pages=read_pdf.getNumPages()
 if(book_pages % 2!=0):
     book_pages=book_pages+1
 else:
      pass   
  
 sheet.update_cell(x,11,book_pages)
#---------------------------------------------------------------
def Book_Name(x): # in this function the browser opens the URl and gets the book name
    print("setting the name of row:",x," on the google sheet....")
    Url_cell=sheet.cell(x,3).value
    browser.get(Url_cell)
    #time.sleep(8)
    web_Name=browser.find_element_by_xpath("/html/body/div[3]/div/header/div/div/div/h1").text
    name=web_Name.replace('كتاب ', '')
    sheet.update_cell(x,4,name)
    print(name)
   
#-----------------------------------------------------------------------------------
#def Book_Year(x):
 #print("setting the year of row:",x," on the google sheet....")
 #Url_cell=sheet.cell(x,3).value
 #browser.get(Url_cell)
 #yr=browser.find_element_by_xpath('//*[@id="bookinfo"]/div[2]/div[1]/div[2]/big/b/br[12]').text
 #yr=browser.find_element_by_xpath('//*[@id="bookinfo"]/div[2]/div[1]/div[2]/big/big[2]').text
#=========================================================================================================================
#-----------------------------------------------------------------------------------------------------------
#----to make a thumbnail picture-----------------------------------------------------------\
def Thumb_Book(x):
 #it uses an existing picture in the Darlusail file and edits it then save the update in the thumbnail file
 #check https://haptik.ai/tech/putting-text-on-image-using-python/
 #https://haptik.ai/tech/putting-text-on-images-using-python-part2/ to adjust the text in image
 im1 = Image.open(r'C:\\Users\\user\\Desktop\\darlusail\\book.jpg')  
 Thumb_cell=sheet.cell(x,16).value
 Title_cell=str(sheet.cell(x,4).value)
 draw = ImageDraw.Draw(im1)
 font= ImageFont.truetype('arial.ttf', size=10) # desired size
 txt = arabic_reshaper.reshape(Title_cell)# starting position of the message
 message=get_display(txt)
 color = 'rgb(0, 0, 0)' # black color
 wrapper = textwrap3.TextWrapper(width=20) # draw the message on the background with text wrapper
 word_list = wrapper.wrap(text=message) 
 caption_new = ''
 ad = AlphabetDetector()
 if (ad.is_arabic(txt)==True or ad.is_latin(txt)!=True):#checks if the text whatever it is has Arabic words
     print(" the name contains arabic words")
     for ii in reversed(word_list):
         caption_new = caption_new + ii + '\n'
 if(ad.is_latin(txt)==True):
     print(" the name doesn't contains arabic words")
     for ii in word_list:
         caption_new = caption_new + ii + '\n'
 print(caption_new)
 font= ImageFont.truetype('arial.ttf', size=10)#define the font for the text
 color = 'rgb(0, 0, 0)'
 #w,h = draw.textsize(message, font=font)
 w,h = draw.textsize(caption_new, font=font)
 W,H = im1.size
 #x,y = 0.5*(W-w),0.90*H-h
 x,y=0.5*(W-w),0.5*H
 draw.text((x,y), caption_new, fill=color,font=font)
 size=392,558 #define the dimensions to crop in the next step
 im1=im1.resize(size,Image.ANTIALIAS)
 im2=im1.save(r"C:"+Thumb_cell) #save the image with the new configurations
#----------------------------------------------------------------------------------------------------
#----------------the same as the Thumbnail_book but for the picture part------------------------------------------------------------------------------------- 
def Pic_Book(x):
 im1 = Image.open(r'C:\\Users\\user\\Desktop\\darlusail\\book.jpg')  
 Pic_cell=sheet.cell(x,17).value
 Title_cell=str(sheet.cell(x,4).value)
 draw = ImageDraw.Draw(im1)
 font= ImageFont.truetype('arial.ttf', size=10) # desired size
 txt = arabic_reshaper.reshape(Title_cell)# starting position of the message
 message=get_display(txt)
 color = 'rgb(0, 0, 0)' # black color
 wrapper = textwrap3.TextWrapper(width=20) # draw the message on the background with text wrapper
 word_list = wrapper.wrap(text=message) 
 caption_new = ''
 ad = AlphabetDetector()
 if (ad.is_arabic(txt)==True or ad.is_latin(txt)!=True):#checks if the text whatever it is has Arabic words
     print(" the name contains arabic words")
     for ii in reversed(word_list):
         caption_new = caption_new + ii + '\n'
 if(ad.is_latin(txt)==True):
     print(" the name doesn't contains arabic words")
     for ii in word_list:
         caption_new = caption_new + ii + '\n'
 print(caption_new)
 font= ImageFont.truetype('arial.ttf', size=10)#define the font for the text
 color = 'rgb(0, 0, 0)'
 #w,h = draw.textsize(message, font=font)
 w,h = draw.textsize(caption_new, font=font)
 W,H = im1.size
 #x,y = 0.5*(W-w),0.90*H-h
 x,y=0.5*(W-w),0.5*H
 draw.text((x,y), caption_new, fill=color,font=font)
 size=563,285 #define the dimensions to crop in the next step
 im1=im1.resize(size,Image.ANTIALIAS)
 im2=im1.save(r"C:"+Pic_cell) #save the image with the new configurations
  
#Thumb_Book(26)
Pic_Book(26)