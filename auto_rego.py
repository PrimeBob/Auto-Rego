#!/usr/bin/env python
# coding: utf-8

# In[19]:


from selenium import webdriver
import time
import os
from gsheets import Sheets
import pandas as pd
import operator
import mysql.connector
import shutil
from shutil import copyfile
import os.path
import csv
from gspread_pandas import Spread, Client 
from oauth2client.service_account import ServiceAccountCredentials
import pprint
import gspread
import math
from selenium.webdriver.chrome.options import Options
import git
from git import Repo


# In[39]:


print("updating program...")

#removing local version
try:
    os.remove('/Users/innovatus6/Desktop/auto-rego/bin/auto_rego.py')
    
except:    
    print("no bin to deleted")
    
    
    
time.sleep(5)

#cloning from the latest version
path='/Users/innovatus6/Desktop/auto-rego/bin'
clone="https://github.com/PrimeBob/Auto-rego"


Repo.clone_from(clone, path)


print("updating completed")


# In[146]:


##removing the annoying tab that toggles and says something like: chrome is monitoring automated something something 
options = Options()
options.add_argument("start-maximized")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)


# In[147]:


#reading credientials 
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

#wd for authentiction file
os.chdir('/Users/innovatus6')
creds = ServiceAccountCredentials.from_json_keyfile_name('linkupdater-631ef6e77556.json', scope)


# In[148]:


#getting permission 
client = gspread.authorize(creds)


# In[149]:


#opening production sheet
sheet= client.open("Master Data Builds").worksheet("test1")


# In[150]:


df = pd.DataFrame(sheet.get_all_records())


# In[151]:


harvestfield=df[df.status == 'need'].reset_index(drop=True)


# In[152]:


harvestfield


# In[153]:


#country dictionary
countrydictionary={
    'USA':'1',
    'Australia':'2',
    'Narnia':'3'    
}


# In[154]:


web = webdriver.Chrome(chrome_options=options, executable_path='/Users/innovatus6/chromedriver')

##the below code fills in the registration page and submits it
for i in range(len(harvestfield.email)):        
    print(str(i)+" start check")
    
    try:
        web.get("https://www.bigmarker.com/innovatus-digital/Jason-He-VS-Cassius-clay")
        time.sleep(10)            

        #name
        first=harvestfield.name[i]
        firstslot=web.find_element_by_xpath('//*[@id="new_member_full_name"]')
        time.sleep(1)
        firstslot.send_keys(first)
        time.sleep(1)

        #email
        email=harvestfield.email[i]
        emailslot=web.find_element_by_xpath('//*[@id="new_member_email"]')
        time.sleep(1)
        emailslot.send_keys(email)
        time.sleep(1)

        #number
        contactnumber=harvestfield.number[i]
        contactnumberslot=web.find_element_by_xpath('//*[@id="conference_registration_pre_conference_responses_attributes_0_response"]')
        time.sleep(1)
        contactnumberslot.send_keys(str(contactnumber))
        time.sleep(1)

        #organisation
        organisation=harvestfield.organisation[i]
        organisationslot=web.find_element_by_xpath('//*[@id="conference_registration_pre_conference_responses_attributes_1_response"]')
        time.sleep(1)
        organisationslot.send_keys(organisation)
        time.sleep(1)

        #job title    
        jobtitle=harvestfield.title[i]
        jobtitleslot=web.find_element_by_xpath('//*[@id="conference_registration_pre_conference_responses_attributes_2_response"]')
        time.sleep(1)
        jobtitleslot.send_keys(jobtitle)
        time.sleep(1)


        #country
        country=int(countrydictionary[harvestfield.country[i]])
        countrybutton=web.find_element_by_xpath('//*[@id="conference_registration_pre_conference_responses_attributes_3_response"]/option['+str(country+1)+']')

        time.sleep(1)
        countrybutton.click()

        #submit 
        submitbutton=web.find_element_by_xpath('//*[@id="register_member"]/div[2]/div/div[8]/div[2]/div/span')
        submitbutton.click()
        time.sleep(2)

        #rego button
        regobutton=web.find_element_by_xpath('//*[@id="register_with_pre_response_template"]')
        regobutton.click()
        time.sleep(15)
        
    except:
        #problem case (when country input has a problem)
        sheet.update_cell(int(df[df.iloc[:,2]==harvestfield.email[i]].index.to_numpy())+2,1,"problem")
        
        continue

    try: 
        time.sleep(2)
        web.find_element_by_xpath('/html/body/div[1]/div[1]/div/div/div[1]/div[2]')
        print(harvestfield.email[i]+ ' has been successfully regoed')
        
        #clear cache because the browser remembers previous regos
        web.delete_all_cookies()
        
        
        #updating pm sheet        
        sheet.update_cell(int(df[df.iloc[:,2]==harvestfield.email[i]].index.to_numpy())+2,1,"yes - regoed")
        
        continue 
        
        
    except:
                  
        #the if clause checks the existence of a string unique to the landing page of an regoed person                  
        if web.find_element_by_xpath('//*[@id="webinar-template-10-li"]/div[1]/div[2]/a').text !="":

            print(harvestfield.email[i]+ ' is already regoed')

            #updating pm sheet        
            sheet.update_cell(int(df[df.iloc[:,2]==harvestfield.email[i]].index.to_numpy())+2,1,"yes - regoed")

            #clear cache because the browser remembers previous regos
            web.delete_all_cookies()

            continue
                
            
        else: 
            
            sheet.update_cell(int(df[df.iloc[:,2]==harvestfield.email[i]].index.to_numpy())+2,1,"problem")
            continue
    


# In[ ]:


web.close()

