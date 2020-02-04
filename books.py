#https://github.com/vdasu/Auto-Login-Script
 # #https://techwithtim.net/tutorials/google-sheets-python-api-tutorial/
#https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html
# this script is for the authors part, to search and add to database 
#http://allselenium.info/handle-stale-element-reference-exception-python-selenium/
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import time
import sys
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

#---------------------------------------------

browser=webdriver.Chrome()
#to order the script to open the Authers party directly.
browser.get("*************************************") # put the URL you wish to access
#-------------------------------------------
# define the username and password text feilds and parameters
username=browser.find_element_by_name("ctl00$LogInContainer$txtUserName")
password=browser.find_element_by_name("ctl00$LogInContainer$txtPassword")
signInButton = browser.find_element_by_name("ctl00$LogInContainer$BtnLogin")
username.send_keys("******") # add the username
password.send_keys("******") #add the password
signInButton.click()
#-----------------------------------------------
#this part to access the google sheet for various tasks add, store, get.....
scope = ["************************","************************************"] # in the scope you define the API of the Google Drive
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope) # the cred.json is the json file which you get the access to the Google Drive 
client = gspread.authorize(creds)
sheet = client.open("*********").get_worksheet(0)  # Open the spreadhseet, first page 
#--------------------------------------------------------------------------------------
#------wAITING VARIABLES
wait = WebDriverWait(browser, 25);
#_____________________________Functions Declarations__________________________________
def Book_title(x): #book title
 tit=browser.find_element_by_name("ctl00$Cufex_MainContent$txtTitle")
 tit_cells=sheet.cell(x,4).value
 try:
    tit.send_keys(tit_cells)
 except StaleElementReferenceException as Exception:
    print('StaleElementReferenceException while trying to write the book title, trying to find element again')                 
    tit=browser.find_element_by_name('ctl00$Cufex_MainContent$txtTitle')
    tit.send_keys(tit_cells)
#------------------------------------------------------------------
def Book_Lang(x): #book language
 lang_book=Select(browser.find_element_by_name("ctl00$Cufex_MainContent$cmbBooksLanguages"))
 lang_cells=sheet.cell(x,12).value
 try:
    lang_book.select_by_visible_text(lang_cells)
 except StaleElementReferenceException as Exception:
    print('StaleElementReferenceException while trying to select the language, trying to find element again') 
    lang_book=Select(browser.find_element_by_name("ctl00$Cufex_MainContent$cmbBooksLanguages"))
    lang_book.select_by_visible_text(lang_cells)
#-----------------------------------------------------------------------------------------------
def Book_Pub(x): #book publisher
 pub_book=browser.find_element_by_name("ctl00$Cufex_MainContent$txtPublisher")
 pub_cells=sheet.cell(x,8).value

 try:
    pub_book.send_keys(pub_cells)
 except StaleElementReferenceException as Exception:
    print('StaleElementReferenceException while trying to write the publisher, trying to find element again') 
    pub_book=browser.find_element_by_name("ctl00$Cufex_MainContent$txtPublisher")
    pub_book.send_keys(pub_cells)
#---------------------------------------------------------------------------------------------
def Book_Year(x): #Book Year
 book_year=browser.find_element_by_name("ctl00$Cufex_MainContent$txtBookYear")
 year_cells=sheet.cell(x,7).value

 try:
    book_year.send_keys(year_cells)
 except StaleElementReferenceException as Exception:
    print('StaleElementReferenceException while trying to set the book year, trying to find element again') 
    book_year=browser.find_element_by_name("ctl00$Cufex_MainContent$txtBookYear")
    book_year.send_keys(year_cells)
#-------------------------------------------------------------------------------------------------------
def Book_Contr(x): #book Contractor
 book_contractor=Select(browser.find_element_by_name("ctl00$Cufex_MainContent$cmbContractor"))
 cont_cells=sheet.cell(x,6).value

 try:   
    book_contractor.select_by_visible_text(cont_cells)
 except StaleElementReferenceException as Exception:
    print('StaleElementReferenceException while trying to select the contractor, trying to find element again')
    book_contractor=Select(browser.find_element_by_name("ctl00$Cufex_MainContent$cmbContractor"))      
    book_contractor.select_by_visible_text(cont_cells)
#-------------------------------------------------------------------------------------------------------
def Book_Auth(x): #book Author
 book_auth=Select(browser.find_element_by_name("ctl00$Cufex_MainContent$cmbAuthor"))
 auth_cells=sheet.cell(x,5).value 
 try:
    book_auth.select_by_visible_text(auth_cells)
 except StaleElementReferenceException as Exception:
    print('StaleElementReferenceException while trying to select the author, trying to find element again')
    book_auth=Select(browser.find_element_by_name("ctl00$Cufex_MainContent$cmbAuthor"))    
    book_auth.select_by_visible_text(auth_cells)
#-------------------------------------------------------------------------------------------------------------
def Small_describ(x): #small describtion
 book_sm=browser.find_element_by_name("ctl00$Cufex_MainContent$txtSmallDescription")
 sm_cells=sheet.cell(x,9).value
 try:
    book_sm.send_keys(sm_cells)
 except StaleElementReferenceException as Exception:
    print('StaleElementReferenceException while trying to write the small description, trying to find element again')
    book_sm=browser.find_element_by_name("ctl00$Cufex_MainContent$txtSmallDescription")
    book_sm.send_keys(sm_cells)
#----------------------------------------------------------------------------------------------------------------------
def Book_Desc(x): #book description
 book_desc=browser.find_element_by_name("ctl00$Cufex_MainContent$txtDescription")
 desc_cells=sheet.cell(x,10).value
 try:
    book_desc.send_keys(desc_cells)
 except StaleElementReferenceException as Exception:
    print('StaleElementReferenceException while trying to write the full description, trying to find element again')
    book_desc=browser.find_element_by_name("ctl00$Cufex_MainContent$txtDescription")
    book_desc.send_keys(desc_cells)
#----------------------------------------------------------------------------------------------------------__________
def Book_Parent(x): #books  Category
 parent_book=Select(browser.find_element_by_name("ctl00$Cufex_MainContent$cmbParentSearch"))
 sub_butt=browser.find_element_by_name("ctl00$Cufex_MainContent$LstCategories$MoveAllRight")
 parent_cells=sheet.cell(x,13).value
 try:
    parent_book.select_by_visible_text(parent_cells)
 except StaleElementReferenceException as Exception:
    print('StaleElementReferenceException while trying to select the parent Category, trying to find element again')
    parent_book=Select(browser.find_element_by_name("ctl00$Cufex_MainContent$cmbParentSearch"))  
    parent_book.select_by_visible_text(parent_cells)

 time.sleep(19)#then press the selecet all button
 try:
    sub_butt.click()
 except StaleElementReferenceException as Exception:
    print('StaleElementReferenceException while trying to select available books Categories, trying to find element again')
    sub_butt=browser.find_element_by_name("ctl00$Cufex_MainContent$LstCategories$MoveAllRight")
    sub_butt.click()
 
 #time.sleep(17)
#-------------------------------------------------------------------------------------------------------------------------
def Ebook(x): # the Ebook part in the web page, and this function will call the upload function.
 #time.sleep(10)
 check=browser.find_element_by_xpath('//*[@id="ctl00_Cufex_MainContent_UpdateCufex"]/div[2]/div[5]/div[2]/div[1]/div[2]/div/label/span')
 #check=browser.find_element_by_name('chkEBook')
 Ebook_price=browser.find_element_by_name("ctl00$Cufex_MainContent$txtEBookPriceInUSD")
 Ebook_commission=browser.find_element_by_name("ctl00$Cufex_MainContent$txtEBookCommission")
 Ebook_up=browser.find_element_by_name("ctl00$Cufex_MainContent$BookNamePdf")
 up_cells=sheet.cell(x,15).value
 browser.implicitly_wait(10)
 try: 
    
    check.click()
 except StaleElementReferenceException as Exception:
    print('StaleElementReferenceException while trying to check the Ebook check, trying to find element again')
    check=browser.find_element_by_xpath('//*[@id="ctl00_Cufex_MainContent_UpdateCufex"]/div[2]/div[5]/div[2]/div[1]/div[2]/div/label/span')
    check.click()
 #put 0$ price
 try: 
    Ebook_price.send_keys('0')
 except StaleElementReferenceException as Exception:
    print('StaleElementReferenceException while trying to put the price, trying to find element again')
    Ebook_price=browser.find_element_by_name("ctl00$Cufex_MainContent$txtEBookPriceInUSD")
    Ebook_price.send_keys('0')
 #put 0$ commission
 try: 
    Ebook_commission.send_keys('0')
 except StaleElementReferenceException as Exception:
    print('StaleElementReferenceException while trying to write the commission, trying to find element again')
    Ebook_commission=browser.find_element_by_name("ctl00$Cufex_MainContent$txtEBookCommission")
    Ebook_commission.send_keys('0')
 try:
    Ebook_up.send_keys(up_cells)
 except StaleElementReferenceException as Exception:
    print('StaleElementReferenceException while trying uploading the PDF copy, trying to find element again')
    Ebook_up=browser.find_element_by_name("ctl00$Cufex_MainContent$BookNamePdf")
    Ebook_up.send_keys(up_cells)
#-----------------------------------------------------------------------------------------------------------
def Book_pages(x):
 pages_book=browser.find_element_by_name("ctl00$Cufex_MainContent$txtPageNumber")
 pages_cells=sheet.cell(x,11).value
 try:
   pages_book.send_keys(pages_cells)
 except StaleElementReferenceException as Exception:
    print('StaleElementReferenceException while trying write the book pages, trying to find element again')
    pages_book=browser.find_element_by_name("ctl00$Cufex_MainContent$txtPageNumber")
    pages_cells=sheet.cell(x,11).value
#-------------------------------------------------------------------------------------------------------------------
def Book_Thumb(x): #thumbnail upload
 #thum_open=browser.find_element_by_name("ctl00$Cufex_MainContent$txtThumbnail$txtPicture")
 thum_open=browser.find_element_by_xpath('//*[@id="ctl00_Cufex_MainContent_txtThumbnail_txtPicture"]')
 browser.implicitly_wait(20)
 thum_choose=browser.find_element_by_name("ctl00$Cufex_MainContent$txtThumbnail$fileImages")
 thum_save=browser.find_element_by_css_selector("div[class='BButtonList BtnAttach AttachPicture']")
 thum_cell=sheet.cell(x,16).value
 try: #open the upload space
     thum_open.click()
 except StaleElementReferenceException as Exception:
     print('StaleElementReferenceException while trying to open the upload button, trying to find element again')
     thum_open=browser.find_element_by_xpath('//*[@id="ctl00_Cufex_MainContent_txtThumbnail_txtPicture"]')
     #thum_open=browser.find_element_by_name("ctl00$Cufex_MainContent$txtThumbnail$txtPicture")
     thum_open.click() 
 time.sleep(3)
 try: # choose the thumbnail
     thum_choose.send_keys(thum_cell)
 except StaleElementReferenceException as Exception:
     print('StaleElementReferenceException while trying to choose the thumbnail to upload, trying to find element again')
     thum_choose=browser.find_element_by_name("ctl00$Cufex_MainContent$txtThumbnail$fileImages")
     thum_choose.send_keys(thum_cell)
 time.sleep(3)    
 try: # press the Save and Crop button
     thum_save.click()
 except StaleElementReferenceException as Exception:
     print('StaleElementReferenceException while trying to save the thumbnail picture, trying to find element again')
     thum_save=browser.find_element_by_css_selector("div[class='BButtonList BtnAttach AttachPicture']")
     thum_save.click()
#-----------------------------------------------------------------------------------------------------------------------------------
def Book_Pic(x): #Picture upload
 
 #pic_open=browser.find_element_by_name("ctl00$Cufex_MainContent$txtPicture$txtPicture")
 pic_open=browser.find_element_by_xpath('//*[@id="ctl00_Cufex_MainContent_txtPicture_txtPicture"]')
 pic_choose=browser.find_element_by_name("ctl00$Cufex_MainContent$txtPicture$fileImages")
 pic_save=browser.find_element_by_xpath('//*[@id="ctl00_Cufex_MainContent_txtPicture_tbAttachment"]/div[4]/div[1]')
 pic_cell=sheet.cell(x,17).value
 try: #open the upload space
     pic_open.click()
 except StaleElementReferenceException as Exception:
     print('StaleElementReferenceException while trying to open the upload button, trying to find element again')
     #pic_open=browser.find_element_by_css_selector("div[class='txtPicture Globaltextbox PicturesStyle AnimateMe']")
     pic_open=browser.find_element_by_xpath('//*[@id="ctl00_Cufex_MainContent_txtPicture_txtPicture"]')
     #pic_open=browser.find_element_by_name("ctl00$Cufex_MainContent$txtPicture$txtPicture")
     pic_open.click() 
 time.sleep(3)
 try: # choose the thumbnail
     pic_choose.send_keys(pic_cell)
 except StaleElementReferenceException as Exception:
     print('StaleElementReferenceException while trying to choose the picture to upload, trying to find element again')
     pic_choose=browser.find_element_by_name("ctl00$Cufex_MainContent$txtPicture$fileImages")
     pic_choose.send_keys(pic_cell)
 time.sleep(2)    
 try: # press the Save and Crop button
     pic_save.click()
 except StaleElementReferenceException as Exception:
     print('StaleElementReferenceException while trying to save the picture, trying to find element again')
     pic_save=browser.find_element_by_xpath('//*[@id="ctl00_Cufex_MainContent_txtPicture_tbAttachment"]/div[4]/div[1]')
     pic_save.click()
 time.sleep(15)
#_____________________________________________________________________________________________________________________________________
#######################################################################################################################################
def Data_write(x): 
 print("setting thumbnail")  
 Book_Thumb(x)
 time.sleep(18)
 print("setting picture")
 Book_Pic(x)
 time.sleep(8)
 print("setting contractor")
 Book_Contr(x)
 time.sleep(20)
 print("setting language, both describtions, publisher,and Author")
 Book_Lang(x)
 Book_Desc(x)
 Small_describ(x)
 Book_Pub(x)
 Book_Auth(x)
 print("setting Category")
 Book_Parent(x)
 time.sleep(25)
 print("Setting year, number of pages, and Title")
 Book_Year(x)
 Book_pages(x)
 Book_title(x)
 print("setting the pdf copy")
 Ebook(x)
 time.sleep(10)
#----------------------------------------------------------------------------------------------------------      
def Data_Upload(x): #this function handles the uploading process, then clearing the fields 
        pop_up=browser.find_element_by_class_name("confirm")
        submit=browser.find_element_by_name("ctl00$Cufex_MainContent$btnSave")
        reset_butt=browser.find_element_by_name("ctl00$Cufex_MainContent$btnAdd")
        Data_write(x)
        try: #press on the submit button 
            submit.click()
        except StaleElementReferenceException as Exception:
            print('StaleElementReferenceException while trying to press final submit button, trying to find element again')
            submit=browser.find_element_by_name("ctl00$Cufex_MainContent$btnSave")
            submit.click()  
        print("updating the status on the spread sheet of book",x)
        sheet.update_cell(x,18,"Done")
        print(" wait 60 seconds,untill the data is being uploaded.....")
        time.sleep(60)
        #browser.implicitly_wait(10)
        try:
            pop_up.click() #click ok on the "book is saved" message
        except StaleElementReferenceException as Exception:
            print('StaleElementReferenceException while trying to press the confirm of the popup, trying to find element again')
            pop_up=browser.find_element_by_class_name("confirm")
            pop_up.click()
        try:
            reset_butt.click()
        except StaleElementReferenceException as Exception:
            print('StaleElementReferenceException while trying to press clear button, trying to find element again')
            reset_butt=browser.find_element_by_name("ctl00$Cufex_MainContent$btnAdd")
            reset_butt.click()
        time.sleep(6)
#-------------------------------------------------------------------------------
def Book_Search(x):
 search_book=browser.find_element_by_name("ctl00$Cufex_MainContent$txtSearch")
 search_butt=browser.find_element_by_name("ctl00$Cufex_MainContent$btnSearch")
 Name_cell=sheet.cell(x,4).value
 search_book.clear()
 try:  # now write the book name in the search bar
   search_book.send_keys(Name_cell)
 except StaleElementReferenceException as Exception:
   print('StaleElementReferenceException while trying to search for the book, trying to find element again')
   search_book=browser.find_element_by_name("ctl00$Cufex_MainContent$txtSearch")
   search_book.send_keys(Name_cell)
 try: #now press the search button
   search_butt.click()
 except StaleElementReferenceException as Exception:
   print('StaleElementReferenceException while trying to press the search button, trying to find element again')
   search_butt=browser.find_element_by_name("ctl00$Cufex_MainContent$btnSearch")
   search_butt.click()
 time.sleep(19)
#____________________________________________________________________________________________________
#=========================================The Bot logic begins here========================================================
#____________________________________________________________________________________________________
#\\=================// change the range in the loop for the process in the spread sheet
# instead of the for loop- declare it as function with input x, 
# and call the whole function from other script inside a for loop
for x in range(17,20):
 Status_cell=sheet.cell(x,18).value
 pop_up=browser.find_element_by_class_name("confirm")
 if(Status_cell !='Done'):
          if(sheet.cell(x,4).value==''):
                 pass
          else:  
             print("working on book:",x,"in the current loop")
             Book_Search(x)
             #if(browser.find_element_by_name("confirm").is_displayed()== True):
             if(pop_up.is_displayed()== True):
                    print("the book is not in the CMS, starting the upload process...")
                    pop_up.click() #press on the ok button of the pop up message
                    Data_Upload(x)
             else:
                    print("press 1 to upload the current book")
                    print("or press 2 to move to the next book")
                    choice=input("what do you want to do next?")
                    if (choice=='1'):
                           Data_Upload(x)
                           print("ok then,....uploading it right now...")
                    if (choice=='2'):
                           sheet.update_cell(x,18,"Done")
                           print("passing to the next book in the sheet")
 if(Status_cell =='Done'):
        pass
print("The job is done,the script will close.......")  
                   

                        
