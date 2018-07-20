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
