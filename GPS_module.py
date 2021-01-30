#import speech_recognition as sr
import os
from math import cos, asin, sqrt, pi
import pandas as pd

vsr_path = '/home/amol/Reports/VSR/'
vs_path = '/home/amol/Reports/VS/'



def find_latest_trip_advance_booking_report():
    #import os


    report_list=[]
    for i in os.listdir('/home/amol/Downloads/'):
        if "TripAdvanceBookingReport_Ver1" in i:
            report_list.append(i)

    newdict={}
    for i in report_list:
        if i=='TripAdvanceBookingReport_Ver1.xlsx':
            newdict.update({0:i})
        else:
            if i[31:32].isdigit():
                number=int(i[31:32])

            if i[31:33].isdigit():
                number=int(i[31:33])

            if i[31:34].isdigit():
                number=int(i[31:34])

            newdict.update({number:i})

    updated_excel_file=newdict[max(newdict.keys())]

    print("I am using {} for report creation".format(updated_excel_file))
    return '/home/amol/Downloads/'+updated_excel_file

def find_latest_current_status_report():
    import os
    import pandas as pd
    import ayansh as gps
    import shutil
    from datetime import datetime
    path= '/home/amol/Downloads/'
    new_path = '/home/amol/Reports/Current Status Report/'
    all_csr_files = gps.getListOfFiles(new_path)

    all_files=os.listdir(path)

    csr_files=[]

    for i in all_files:
        try:
            if i.startswith('Current_Status_Report'):
                csr_files.append(i)
        except:
            pass
      
    for i in csr_files:
        df = pd.read_excel(path+i,sheet_name='REPORT INFO')
        file_date = df.iloc[6,1]
        file_date =datetime.strptime(file_date,'%d/%m/%Y %H:%M:%S')
        file_date=file_date.strftime("%Y-%m-%d %X")
        os.rename(path+i,new_path+"Current_Status_Report "+file_date+".xls")
        #shutil.move(path+file_name,new_path+file_name.split('/')[-1]) Ex.
        print('For loop ran.....')
        
    print("Renaming done ........")

    print("No of files befor moving = ",len(all_csr_files))
    all_csr_files = gps.getListOfFiles(new_path)
    print("No of files after moving = ",len(all_csr_files))
    all_csr_files.sort()
    print("Hi Amol I have selected {} as a latest updated Current Status Report".format(all_csr_files[-1]))
    return all_csr_files[-1]



def find_latest_current_status_report_old():
    import os

    report_list=[]
    for i in os.listdir('/home/amol/Downloads/'):
        if "Current_Status_Report" in i:
            report_list.append(i)

    newdict={}
    for i in report_list:
        if i=='Current_Status_Report.xls':
            newdict.update({0:i})
        else:
            if i[23:24].isdigit():
                number=int(i[23:24])

            if i[23:25].isdigit():
                number=int(i[23:25])

            if i[23:26].isdigit():
                number=int(i[23:26])

            newdict.update({number:i})


    updated_current_status_report=newdict[max(newdict.keys())]

    print('Hi Amol, I have selected "{}" file as a latest updated file '.format(updated_current_status_report))
    return os.path.join('/home/amol/Downloads/',updated_current_status_report)

def find_latest_real_time_report():
    import os

    report_list=[]
    for i in os.listdir('/home/amol/Downloads/'):
        if "Real_Time_Report" in i:
            report_list.append(i)

    newdict={}
    for i in report_list:
        if i=='Real_Time_Report.xls':
            newdict.update({0:i})
        else:
            if i[18:19].isdigit():
                number=int(i[18:19])

            if i[18:20].isdigit():
                number=int(i[18:20])

            if i[18:21].isdigit():
                number=int(i[18:21])
            try:
                newdict.update({number:i})
            except:
                pass


    updated_real_time_report=newdict[max(newdict.keys())]

    print('Hi Amol, I have selected "{}" file as a latest updated file '.format(updated_real_time_report))
    return os.path.join('/home/amol/Downloads/',updated_real_time_report)

def df_to_excel(template,worksheet,row,col,df,output):
    from openpyxl import load_workbook
    r=row
    c=col
    #version openpyxl 3.0.0
    wb = load_workbook(template)
    print("File opened ")

    sheet = wb[worksheet]

    for i in df.index:
        c=col
        for j in df.loc[i]:
            sheet.cell(row=r,column=c).value=j
            c+=1
        r+=1
    wb.save(output)
    print("Finised..............")

def GPS_Email_Report():

    import os

    os.chdir('/home/amol/Downloads')

    file_name=list(os.listdir())
    report_list=[]
    for i in os.listdir():
        if "Current_Status_Report" in i:
            report_list.append(i)

    newdict={}
    for i in report_list:
        if i=='Current_Status_Report.xls':
            newdict.update({0:i})
        else:
            if i[23:24].isdigit():
                number=int(i[23:24])

            if i[23:25].isdigit():
                number=int(i[23:25])

            if i[23:26].isdigit():
                number=int(i[23:26])

            newdict.update({number:i})



    updated_current_status_report=newdict[max(newdict.keys())]

    print('Hi Amol, I am using "{}" file for creating report'.format(updated_current_status_report))



    import pandas as pd
    df=pd.read_excel(str('/home/amol/Downloads/'+updated_current_status_report),sheet_name="Current Status Report")
    df.to_excel("temp.xlsx",index=None)


    #from openpyxl import Workbook
    from openpyxl import load_workbook

    gps = load_workbook('/home/amol/Documents/Excel Files/Email GPS Ver 5.xlsx')
    gps_ws=gps['Current Status Report']
    novire = load_workbook("/home/amol/Downloads/temp.xlsx")
    novire_ws=novire['Sheet1']
    #vhiof = load_workbook('/home/amol/Desktop/VHIOF/vehicle hold in other fleet.xlsx')
    #vhiof_ws=vhiof.active


    gps_ws.delete_cols(1,20) # to delete previous data of the first 1 to 20 columns

    for i in range(1,novire_ws.max_row+1): # novire_ws.max_column+1
        for j in range(1,novire_ws.max_column+1):
            gps_ws.cell(row=i,column=j).value=novire_ws.cell(row=i,column=j).value

    #gps_ws_pastereport=gps['PasteReport']

    #for i in range(1,vhiof_ws.max_row+1): # novire_ws.max_column+1
    #    for j in range(1,vhiof_ws.max_column+1):
    #        gps_ws_pastereport.cell(row=i,column=j).value=vhiof_ws.cell(row=i,column=j).value


    import datetime

    date=datetime.datetime.now()
    date=date.strftime(" %d %b %Y %X")
    date=str(date)

    file_name="/home/amol/Desktop/GPS Email"+date+".xlsx"

    gps.save(file_name)

    print('Your report has been created successfully\nOutput file is located on Desktop, file name is "{}"\nFull path is {}'.format(file_name.split("/")[-1],file_name))

def find_latest_vsr():
    all_vsr_files = getListOfFiles(vsr_path)
    all_vsr_files.sort()
    print("Hi Amol, I am using ",all_vsr_files[-1])
    return all_vsr_files[-1]

def find_changes_in_controling_branch():
    import ayansh as gps
    import pandas as pd
    import numpy as np
    from datetime import timedelta
    import subprocess

    path_last_vsr = gps.find_latest_vsr()

    Updated_Current_Status_Report=gps.find_latest_current_status_report()
    Basic_Info_from_CSR = pd.read_excel(Updated_Current_Status_Report,sheet_name="Current Status Report",usecols=('Vehicle','Date/Time','Location','Speed'))
    Basic_Info_from_CSR = pd.read_excel(Updated_Current_Status_Report,sheet_name="Current Status Report",usecols=('Vehicle','Controlling Branch'))
    Basic_Info_from_CSR.rename(columns= {'Controlling Branch' : 'On Novire'},inplace=True)

    vsr= pd.read_excel(path_last_vsr,sheet_name="Vehicle_Status_Register_Quick",usecols=('Vehicle No','Vehicle control Location'))
    vsr.rename(columns={'Vehicle No' :'Vehicle', 'Vehicle control Location' : 'in Varuna' },inplace=True)
    vsr=vsr[['Vehicle','in Varuna']]
    final = Basic_Info_from_CSR.merge(vsr,on='Vehicle')
    filt = final['On Novire'] == final['in Varuna']
    mismatched = final[-filt]
    mismatched.rename(columns={'On Novire': 'Old Controlling Branch', 'in Varuna' : 'Current Controlling Branch'},inplace=True)
    mismatched['SrNo']= range(1,mismatched['Vehicle'].count()+1,1)
    mismatched=mismatched[['SrNo','Vehicle','Old Controlling Branch','Current Controlling Branch']]
    gps.df_to_excel('/home/amol/Documents/Excel Files/Update Controlling Branch Template.xlsx','Sheet1',row=2,col=1,df=mismatched,output='/home/amol/Desktop/Update Controlling Branch.xlsx')
    subprocess.call(["et",'/home/amol/Desktop/Update Controlling Branch.xlsx'],shell=False)
    print("I used ",path_last_vsr)
    print("End ..........")

def TakeCommand():
    import speech_recognition as sr
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print('Say something')
        audio = r.listen(source,timeout=1,phrase_time_limit=5)

    #use Google's speech recognation
    data = ''
    try:
        data = r.recognize_google(audio)
        print('You said: '+ data)
    except sr.UnkonwnValueError:
        print('Google Speech Recognition could not understand that audio, unknown error')
    except sr.RequestError as e:
        print('Request restults from Google Speech Recognitoin service error')
    return data

def getListOfFiles(dirName):
    # create a list of file and sub directories
    # names in the given directory
    listOfFile = os.listdir(dirName)
    allFiles = list()
    # Iterate over all the entries
    for entry in listOfFile:
        # Create full path
        fullPath = os.path.join(dirName, entry)
        # If entry is a directory then get the list of files in this directory
        if os.path.isdir(fullPath):
            allFiles = allFiles + getListOfFiles(fullPath)
        else:
            allFiles.append(fullPath)

    return allFiles

def update_driver_details_db():
    import ayansh as gps
    import pandas as pd

    updated_file = gps.find_latest_trip_advance_booking_report()
    df=pd.read_excel(updated_file,usecols=["Vehicle No","Manual Driver code","Driver Mobile","Mobile No","Driver Name","Branch","Zone","Request By Mobile","From City","To City","Dated"],header=2)

    #df.sort_index(ascending=False,inplace=True)
    #df.drop_duplicates(inplace=True)
    #df.drop_duplicates(subset="Vehicle No",inplace=True)
    df['Vehicle No'] = df['Vehicle No'].str.upper()
    df.drop_duplicates(subset="Vehicle No",keep='last',inplace=True)
    df_csv = pd.read_csv("/home/amol/Documents/Excel Files/Templates/Create CSV Contacts.csv")
    def change_vehice_number(veh):
        return veh[-4:]+"-"+veh

    df_csv['Name']= df['Vehicle No'].map(change_vehice_number)
    df_csv['Yomi Name'] = df['Driver Name']
    df_csv['Phone 1 - Type'] = "Mobile"
    df_csv['Phone 1 - Value'] = df['Driver Mobile']
    df_csv['Phone 2 - Type'] = "Mobile"
    df_csv['Phone 2 - Value'] = df['Mobile No']
    df_csv['Notes'] = df['Branch'] + "---" + df['Zone'] + " | " + "FROM " + df['From City'].str.title() + " TO " + df['To City'].str.title()

    df_csv.to_csv('/home/amol/Desktop/Driver Contacts.csv',index=None)
    print("CSV created ..... ")
    #Createing file /home/amol/Documents/Excel Files/Driver_Mobile_Number.xlsx
    df_db_excel = df[['Vehicle No','Driver Mobile', 'Mobile No','Request By Mobile','Manual Driver code', 'Driver Name', 'Branch', 'Zone','From City', 'To City']]
    df_db_excel.to_excel("/home/amol/Documents/Excel Files/Driver_Mobile_Number.xlsx",index=None)
    print("Excel file: /home/amol/Documents/Excel Files/Driver_Mobile_Number.xlsx\nCreated ............")

def gps_stop_enroute_report():
    import os
    import ayansh as my_gps
    import pandas as pd
    from openpyxl import load_workbook
    #from datetime import *
    import datetime

    latest_csr_file = my_gps.find_latest_current_status_report()

    df=pd.read_excel(latest_csr_file,sheet_name="Current Status Report")
    #For filling NA with NA string
    df[df['Location'].isna()] = df[df['Location'].isna()].fillna("NA")
    print("Maximum GPS Time is = {}".format(df.loc[1,'Date/Time']))

    #calculating datetime of currnet time - latest csr  
    maximum_gps_time = df.loc[1,'Date/Time']
    maximum_gps_time = datetime.datetime.strptime(maximum_gps_time,"%Y/%m/%d %H:%M")
    print("Report delay time " ,datetime.datetime.now() - maximum_gps_time)


    df.to_excel("/home/amol/Downloads/temp.xlsx",index=None)

    gps = load_workbook('/home/amol/Documents/Excel Files/Email GPS Ver 6.xlsx')
    gps_ws=gps['Current Status Report']
    novire = load_workbook("/home/amol/Downloads/temp.xlsx")
    novire_ws=novire['Sheet1']
    #vhiof = load_workbook('/home/amol/Desktop/VHIOF/vehicle hold in other fleet.xlsx')
    #vhiof_ws=vhiof.active


    gps_ws.delete_cols(1,20) # to delete previous data of the first 1 to 20 columns 

    for i in range(1,novire_ws.max_row+1): # novire_ws.max_column+1
        for j in range(1,novire_ws.max_column+1):
            gps_ws.cell(row=i,column=j).value=novire_ws.cell(row=i,column=j).value

    #gps_ws_pastereport=gps['PasteReport']

    #for i in range(1,vhiof_ws.max_row+1): # novire_ws.max_column+1
    #    for j in range(1,vhiof_ws.max_column+1):
    #        gps_ws_pastereport.cell(row=i,column=j).value=vhiof_ws.cell(row=i,column=j).value

    #New Code
    gps_report_info = gps['REPORT INFO']

    #gps_report_info.delete_cols(6,2)
    #gps_report_info.insert_cols(6,2)

    try:
        last_vsr_file = my_gps.find_latest_vsr()
    except:
        memory_file = open('/home/amol/Documents/Excel Files/Templates/program_memory_files/last_used_vsr_file.mem','r')
        last_vsr_file = memory_file.readline()
        print("Hi I am using VSR file : ",last_vsr_file)



    vsr = pd.read_excel(last_vsr_file,sheet_name='Vehicle_Status_Register_Quick',usecols=['Vehicle No','Vehicle control Location'])
    vsr = vsr[['Vehicle No','Vehicle control Location']]


    driver_number = pd.read_excel('/home/amol/Documents/Excel Files/Driver_Mobile_Number.xlsx',sheet_name='Sheet1')
    vsr = vsr.merge(driver_number,on='Vehicle No',how='outer')


    r=1
    for i in vsr.index:
        gps_report_info.cell(row=r,column=6).value=vsr.iloc[r-1,0]
        gps_report_info.cell(row=r,column=7).value=vsr.iloc[r-1,1]
        gps_report_info.cell(row=r,column=10).value=vsr.iloc[r-1,2] #for driver mobile number add in excel
        r+=1

    #New Code

    date=datetime.datetime.now()     
    date=date.strftime(" %d %b %Y %X")
    date=str(date)

    file_name="/home/amol/Desktop/GPS Email"+date+".xlsx"
    file_name = file_name.replace(" ","_")

    gps.save(file_name)

    print('Your report has been created successfully\nOutput file is located on Desktop, file name is "{}"\nFull path is {}'.format(file_name.split("/")[-1],file_name))
    #memory_file = open('/home/amol/Documents/Excel Files/Templates/program_memory_files/last_used_vsr_file.mem','w')
    #memory_file.write(last_vsr_file)
    #memory_file.close()
    with open('/home/amol/Documents/Excel Files/Templates/program_memory_files/last_used_vsr_file.mem','w',encoding = 'utf-8') as f:
        f.write(last_vsr_file)
        f.close()
    return file_name

def rename_and_move_csr_file():
    import os
    import pandas as pd
    import ayansh as gps
    #import subprocess
    from datetime import datetime
    #path = '/home/amol/Desktop/csr/'
    path= '/home/amol/Downloads/'

    all_files=os.listdir(path)

    csr_files=[]

    for i in all_files:
        try:
            if i[:21]=='Current_Status_Report':
                csr_files.append(i)
        except:
            pass


    for i in csr_files:
        df = pd.read_excel(path+i,sheet_name='REPORT INFO')
        file_date = df.iloc[6,1]
        file_date =datetime.strptime(file_date,'%d/%m/%Y %H:%M:%S')
        file_date=file_date.strftime("%Y-%m-%d %X")
        os.rename(path+i,path+"Current_Status_Report "+file_date+".xls")

    print("Renaming done ........")

    import shutil

    all_files=os.listdir(path)

    csr_files=[]

    for i in all_files:
        try:
            if i[:21]=='Current_Status_Report':
                csr_files.append(i)
        except:
            pass


    new_path = '/home/amol/Reports/Current Status Report/'
    total_files = gps.getListOfFiles(new_path)
    print("No of files befor moving = ",len(total_files))
    for file_name in csr_files:
        shutil.move(path+file_name,new_path+file_name.split('/')[-1])
    print("Filed moved in "+new_path)
    total_files = gps.getListOfFiles(new_path)
    print("No of files after moving = ",len(total_files))

def download_current_status_report_gui():
    import pyautogui
    import time
    time.sleep(1)
    import sys
    import ayansh as gps
    import os
    png_path = '/home/amol/anaconda3/lib/python3.8/site-packages/ayansh/png_for_pyautogui/csr/'
    #os.chdir('/home/amol/Documents/01 - py/CurrentStatusReport/')
    os.chdir(png_path)
    while pyautogui.locateCenterOnScreen('Google_Chrome_Icon.png',confidence=0.80) is None:
            pass
    pyautogui.moveTo(pyautogui.locateCenterOnScreen('Google_Chrome_Icon.png',confidence=0.9),duration=0)
    pyautogui.click()
    time.sleep(1)

    while pyautogui.locateCenterOnScreen('google_chorme_detector_app_icon.png',confidence=0.99) is None:
            pass
    #pyautogui.moveTo(pyautogui.locateCenterOnScreen('google_chorme_detector_app_icon.png',confidence=0.9),duration=0)
    time.sleep(3)

    pyautogui.hotkey('ctrl','t')
    time.sleep(1)
    pyautogui.typewrite('http://www.ivts.noviretechnologies.com/IVTS/')
    #pyautogui.typewrite('https://ivts.noviretechnologies.com/IVTS/logout.do')
    
    pyautogui.press('enter')


    while pyautogui.locateCenterOnScreen('username.png',confidence=0.80) is None:
            pass
    pyautogui.moveTo(pyautogui.locateCenterOnScreen('username.png',confidence=0.80),duration=0)
    pyautogui.click()
    pyautogui.typewrite('varuna')
    pyautogui.typewrite(['tab'])
    #while pyautogui.locateCenterOnScreen('password.png',confidence=0.99) is None:
    #        pass
    #pyautogui.moveTo(pyautogui.locateCenterOnScreen('password.png',confidence=0.9),duration=0)
    #pyautogui.click()
    pyautogui.typewrite('vil2020')
    pyautogui.typewrite(['enter'])

    while pyautogui.locateCenterOnScreen('ReportSelectionTool.png',confidence=0.99) is None:
            pass
    pyautogui.moveTo(pyautogui.locateCenterOnScreen('ReportSelectionTool.png',confidence=0.9),duration=0)
    pyautogui.click()


    while pyautogui.locateCenterOnScreen('SelectReport.png',confidence=0.99) is None:
