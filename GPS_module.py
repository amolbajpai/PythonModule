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
