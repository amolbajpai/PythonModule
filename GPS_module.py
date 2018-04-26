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
