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
#from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
#---------------------------------------------
browser=webdriver.Chrome()
#to order the script to open the Authers party directly.
browser.get("https://www.darlusail.com/cufex/Definition/Authors")
#-------------------------------------------
# define the username and password text feilds and parameters
username=browser.find_element_by_name("ctl00$LogInContainer$txtUserName")
password=browser.find_element_by_name("ctl00$LogInContainer$txtPassword")
signInButton = browser.find_element_by_name("ctl00$LogInContainer$BtnLogin")
username.send_keys("husam")
password.send_keys("&^847HYR356#$")
signInButton.click()
#-----------------------------------------------
#this part to access the google sheet for various tasks add, store, get.....
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)
sheet = client.open("dar-lusail").get_worksheet(1)  # Open the spreadhseet, second page 
#-------------the diioferent fields, texts, buttons...------------------------
#==============================================================================
#____________________Noe Functions Declarations__________________________________
def author_Name(x):   # the Name Field
     Name=browser.find_element_by_name("ctl00$Cufex_MainContent$txtTitle")
     Flcell=sheet.cell(x,2).value #Flcell if for the full name
     try:      
         Name.send_keys(Flcell)
     except StaleElementReferenceException as Exception:
         print('StaleElementReferenceException while trying to type the author name in the text field, trying to find element again')
         Name=browser.find_element_by_name("ctl00$Cufex_MainContent$txtTitle")
         Name.send_keys(Flcell)           
#-------------------------------------------------------------------------
def author_thumbData(x):
      #---------upload thumbnail picture part--------   
 th_cells=sheet.cell(x,7).value #thumbnail picture cells 
 thum_open=browser.find_element_by_name("ctl00$Cufex_MainContent$txtThumbnail$txtPicture")
 thum_choose=browser.find_element_by_name("ctl00$Cufex_MainContent$txtThumbnail$fileImages")  
 try:  
         thum_open.click()
 except StaleElementReferenceException as Exception:
         print('StaleElementReferenceException while trying to press the thumbnail button, trying to find element again') 
         thum_open=browser.find_element_by_name("ctl00$Cufex_MainContent$txtThumbnail$txtPicture")
         thum_open.click() 
 browser.implicitly_wait(6)
 try: 
         thum_choose.send_keys(th_cells)
        
 except StaleElementReferenceException as Exception:
         print('StaleElementReferenceException while trying to attach the thumbnail, trying to find element again') 
         thum_choose=browser.find_element_by_name("ctl00$Cufex_MainContent$txtThumbnail$fileImages")
         thum_choose.send_keys(th_cells)
                 #locate the "Crop & Save" button
 thum_upload=browser.find_element_by_css_selector("div[class='BButtonList BtnAttach AttachPicture']")
 thum_upload.click()  
#-----------------------------------------------------------------------------------------------------
def author_Hon(x): #honorific part
  Hon_cell=sheet.cell(x,1).value #Honorific cells
  honorofic=Select(browser.find_element_by_name("ctl00$Cufex_MainContent$CmbHonorifics"))
  if(Hon_cell !="Empty"):
    
         try: #honorific content
                 honorofic.select_by_visible_text(Hon_cell)
         except StaleElementReferenceException as Exception:
                 print('StaleElementReferenceException while trying to select the Honorific, trying to find element again')
                 honorofic=Select(browser.find_element_by_name("ctl00$Cufex_MainContent$CmbHonorifics"))
                 honorofic.select_by_visible_text(Hon_cell)
  else:
                pass 
  time.sleep(10)
#----------------------------------------------------------------------------------------------------------
def author_Contr(x): # Contractor part
 con_cell=sheet.cell(x,3).value # contractor cells
 Contractor=Select(browser.find_element_by_name("ctl00$Cufex_MainContent$cmbMember"))
 try:     #contractor selection
         Contractor.select_by_visible_text(con_cell)
 except StaleElementReferenceException as Exception:
         print('StaleElementReferenceException while trying to select the contractor, trying to find element again')
         Contractor=Select(browser.find_element_by_name("ctl00$Cufex_MainContent$cmbMember"))
         Contractor.select_by_visible_text(con_cell)
#--------------------------------------------------------------------------------------------------------------------
def author_About(x): # the About field
 about_Breif=browser.find_element_by_name("ctl00$Cufex_MainContent$txtSubTitle")
 ab_cell=sheet.cell(x,4).value # about Breif cells
 try:       
          about_Breif.send_keys(ab_cell)
 except StaleElementReferenceException as Exception:
         print('StaleElementReferenceException while trying to type "about breif" in the text field, trying to find element again')
         about_Breif=browser.find_element_by_name("ctl00$Cufex_MainContent$txtSubTitle")
         about_Breif.send_keys(ab_cell)  
#------------------------------------------------------------------------------------------------------------------------------
def author_picture(x): # the picture part
 pic_press=browser.find_element_by_name("ctl00$Cufex_MainContent$txtPicture$txtPicture")
 pic_choose=browser.find_element_by_name("ctl00$Cufex_MainContent$txtPicture$fileImages") 
 pc_cells=sheet.cell(x,8).value #picture cells
 pic_upload=browser.find_element_by_xpath('//*[@id="ctl00_Cufex_MainContent_txtPicture_tbAttachment"]/div[4]/div[1]')
 try:
         pic_press.click()
 except StaleElementReferenceException as Exception:
         print('StaleElementReferenceException while trying to press the picture upload button, trying to find element again')                 
         pic_press=browser.find_element_by_xpath('//*[@id="ctl00_Cufex_MainContent_txtPicture_txtPicture"]')
         #pic_press=browser.find_element_by_name("ctl00$Cufex_MainContent$txtPicture$txtPicture")
         pic_press.click()
 try:
         pic_choose.send_keys(pc_cells) 
 except StaleElementReferenceException as Exception:
         print('StaleElementReferenceException while trying to locate the picture choose button, trying to find element again')
         pic_choose=browser.find_element_by_name("ctl00$Cufex_MainContent$txtPicture$fileImages")
         pic_choose.send_keys(pc_cells) 

 pic_upload.click()
#-----------------------------------------------------------------------------------------------------------------------------------------
def author_HomeDesc(x): # the Home Description
 home_Description=browser.find_element_by_name("ctl00$Cufex_MainContent$txtSmallDescription")
 Ho_cell=sheet.cell(x,5).value #Home description cells
 try:    #Home Description
         home_Description.send_keys(Ho_cell)
 except  StaleElementReferenceException as Exception:
         print('StaleElementReferenceException while trying to type home description in the text field, trying to find element again')
         home_Description=browser.find_element_by_name("ctl00$Cufex_MainContent$txtSmallDescription")
         home_Description.send_keys(Ho_cell)
#------------------------------------------------------------------------------------------------------ 
def author_ShDescribtion(x): # the short Describtion
 short_Description=browser.find_element_by_name("ctl00$Cufex_MainContent$txtDescription")
 sh_cells=sheet.cell(x,4).value #short description 
 try:     
         short_Description.send_keys(sh_cells)
 except  StaleElementReferenceException as Exception:
         print('StaleElementReferenceException while trying to type short description in the text field, trying to find element again')
         short_Description=browser.find_element_by_name("ctl00$Cufex_MainContent$txtDescription")
         short_Description.send_keys(sh_cells)
#-----------------------------------------------------------------------------------------------------------------------------------------
def data_Write(x): 
 print("setting author contractor....") 
 author_Contr(x)
 time.sleep(15)
 print("setting author name...")
 author_Name(x)
 print("setting the author picture")
 author_picture(x)
 time.sleep(15)
 print("setting the author thumbnbail...")
 author_thumbData(x)
 time.sleep(15)
 print("setting the Honorific,if it exists.... ")
 author_Hon(x)
 print("setting the short describtion, home describtion, and the about field")
 author_ShDescribtion(x)
 author_HomeDesc(x)
 author_About(x)
 print("All the fields are filled")
#-----------------------------Clear the page after the upload in process--------------------------------------------------------------------------------------
def clear_Page():
    clear_butn=browser.find_element_by_name("ctl00$Cufex_MainContent$btnAdd")
    popMess=browser.find_element_by_class_name("confirm")
    print("clearing all the fields...")
    try:
                  clear_butn.click()
    except StaleElementReferenceException as Exception:
                  print('StaleElementReferenceException while trying press clear, trying to find element again')
                  clear_butn=browser.find_element_by_name("ctl00$Cufex_MainContent$btnAdd")
                  clear_butn.click()          
         # now confirm the clear button
    try:
                  popMess.click()
    except StaleElementReferenceException as Exception:
                  print('StaleElementReferenceException while trying to Confirm the Clearing, trying to find element again')
                  popMess=browser.find_element_by_class_name("confirm")
                  popMess.click()
#----------------------------------------------------------------------------------
def Data_Upload(x): # the upload function
 data_Write(x)
 sub_final=browser.find_element_by_name("ctl00$Cufex_MainContent$btnSave")
 popMess=browser.find_element_by_class_name("confirm") 
 try: #press submit
         sub_final.click()
 except StaleElementReferenceException as Exception:
         print('StaleElementReferenceException while trying to press the submit, trying to find element again') 
         sub_final=browser.find_element_by_name("ctl00$Cufex_MainContent$btnSave") 
         sub_final.click()
  #-----update the status row on the spread sheet
 sheet.update_cell(x,9,"Done") 
 time.sleep(18)#-------wait till the upload ends and the message is displayed
 
 try:
         popMess.click()       
 except StaleElementReferenceException as Exception:
         print('StaleElementReferenceException while trying to press the OK button, trying to find element again')
         popMess=browser.find_element_by_class_name("confirm")
         popMess.click()
 time.sleep(8)
 clear_Page()
#----------------search for the author if exists in the CMS-----------------------
def author_Search(x):
 search_name= browser.find_element_by_name("ctl00$Cufex_MainContent$txtSearch")  
 buttonSearch=browser.find_element_by_name("ctl00$Cufex_MainContent$btnSearch") 
 browser.find_element_by_name("ctl00$Cufex_MainContent$txtSearch").clear()
 Flcell=sheet.cell(x,2).value #Flcell if for the full name
 try:
    search_name.send_keys(Flcell)#next step search for the names from tyhe arrray
 except StaleElementReferenceException as Exception:
     print('StaleElementReferenceException while trying to type the author name, trying to find element again')
     search_name= browser.find_element_by_name("ctl00$Cufex_MainContent$txtSearch")
     search_name.send_keys(Flcell)
 try: 
    buttonSearch.click()
 except StaleElementReferenceException as Exception:
     print('StaleElementReferenceException while trying to press the search button, trying to find element again')
     buttonSearch=browser.find_element_by_name("ctl00$Cufex_MainContent$btnSearch")
     buttonSearch.click()  
 time.sleep(10) #wait till the search ends

#_____________________End of Functions declarations______________________________________________________________________________________________
#================================================================================================================================================
#----------------------Script logic begins here------------------------------------------------------------------------------------
# instead of the for loop- declare it as function with input x, 
# and call the whole function from other script inside a for loop
for x in range(2, 20):
    
 clear_butn=browser.find_element_by_name("ctl00$Cufex_MainContent$btnAdd")
        # Define the Range for the script to check from  Row 1 to row 9 without 10
 print("working on step",x)
 Flcell=sheet.cell(x,2).value #Flcell if for the full name
 St_cell=sheet.cell(x,9).value #to get the 'Done'status cells
 #if (sheet.cell(x,2).value!='Full Name' and sheet.cell(x,9).value !='Done' and sheet.cell(x,9).value !='Status'):
 if (Flcell !='Full Name' and St_cell !='Done' and St_cell!='Status'):
  if(sheet.cell(x,2).value==""):
   pass
  else:
   print("working on ",Flcell)
   author_Search(x)
   if (browser.find_element_by_class_name("confirm").is_displayed()==True): #if the author is not in DB, the script adds the data
         browser.find_element_by_class_name("confirm").click()
         Data_Upload(x)
   else:   
           print("press 1 to upload the current auther")
           print("and '2' to pass to the next choice on the spread sheet.......")
           choice=input("what is the next step?")
           if(choice=='1'):
                   print("passing the data to the CMS........")
                   Data_Upload(X)
           if(choice=='2'):
                   print("going to the next author on the spread sheet.......")
                   sheet.update_cell(x,9,'Done')

print("The job is done,the script will close.......")        
          