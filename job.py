# -*- coding: utf-8 -*-
"""
Created on Tue Jul 20 19:52:24 2021

@author: molha
"""

# Import libraries
import pandas as pd
import datetime
from datetime import date
import matplotlib as mpl
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from pathlib import Path # to open excel file using python
import math
import os

# create the path
path = os.path.join('D:/Jobs', 'Applied Jobs1.xlsx')

# import the data
df = pd.read_excel(path)
#df = df.iloc[:,0:-4]

# provide reports
class Jobs:
    def __init__(self):
        self.dates = df['Application Date'].tolist()
        
    def weekly_effeciency(self):
        weekno = []
        effeciency = []
        self.weekstart = 0
        for i in range(int(math.ceil((max(self.dates) - min(self.dates)).days/7))):
            if self.weekstart == 0:
                self.weekstart = min(self.dates)
                if (max(self.dates) - self.weekstart).days >= 6:
                    self.weekend = self.weekstart + datetime.timedelta(days=7)
                else:
                    self.weekend = pd.to_datetime(date.today())
            else:
                self.weekstart = self.weekend
                if (max(self.dates) - self.weekstart).days >= 6:
                    self.weekend = self.weekstart + datetime.timedelta(days=7)
                else:
                    self.weekend = pd.to_datetime(date.today())
            self.AppsVolume = len(df[(df['Application Date'] >= self.weekstart)
                                 & (df['Application Date'] < self.weekend + 
                                    datetime.timedelta(days=1))]
                                  ['Application Date'])
            
            print('Week Start Day: ', str(self.weekstart).split(' ')[0])
            print('Week End Day: ', str(self.weekend).split(' ')[0])
            print('Number of Applications: ', self.AppsVolume)

            if (self.weekend - self.weekstart).days < 7:
                WeekProjection = len(df[(df['Application Date'] >= self.weekstart)
                                 & (df['Application Date'] <= self.weekend)]
                                  ['Application Date']) / (
                                      self.weekend - self.weekstart).days * 7
                print('Expected number of applications this week: ', 
                      round(WeekProjection))
            print()
            print()
            weekno.append(self.weekend.strftime("%d/%b/%Y"))
            if (self.weekend - self.weekstart).days < 7:
                effeciency.append(round(WeekProjection))
            else:
                effeciency.append(self.AppsVolume)
                
        tab1 = pd.DataFrame(weekno)
        tab1.columns = ['Dates']
        tab1['Number of Applications'] = effeciency
        print(tab1)
        charts_path = os.path.join('D:/pys','jobs')
        plt.bar(tab1.Dates,tab1['Number of Applications'], color = 'green', ec = 'black')
        plt.xticks(rotation = 30, fontsize = 8)
        plt.title('Weekly Applications Trend', fontsize = 8)
        plt.savefig(charts_path + '\\main trend.png', dpi=300)
        
        # Python program to read
        # image using PIL module
        # importing PIL
        from PIL import Image
  
        cpath = os.path.join('D:/pys','jobs', 'main trend.png')
        # Read image
        img = Image.open(cpath)
  
        # Output Images
        img.show()
        
    def NewJob(self):
        Sr = df['Sr.'][df. index[-1]] + 1
        print('Job serial is ', Sr)
        Date = datetime.datetime.now()
        print('Job data is ', Date)
        Title = input('Enter the job title: ')
        CompanyName = input('Enter the company name: ')
        if CompanyName in(df.Company.tolist()):
            print('You have applied befor in theis company as follows:')
            print(df[df.Company == CompanyName])
        JobLink = input('Enter the job link: ')
        Location = input ('Enter the job location: ')
        
        print('============================================================')
        # Platform
        print()
        print()
        PlatformCondition = False
        while not PlatformCondition:
            print('Enter 1 for Indeed')
            print('Enter 2 for LinkedIn')
            print('Enter 3 for Facebook')
            print('Enter 4 for others (manual input)')
            choice1 = int(input('Enter your choice: '))
            if choice1 == 1:
                Platform = 'Indeed'
                print()
                print('Selected platform is',Platform)
                PlatformCondition = True
            elif choice1 == 2:
                Platform = 'LinkedIn'
                print()
                print('Selected platform is',Platform)
                PlatformCondition = True
            elif choice1 == 3:
                Platform = 'Facebook'
                print()
                print('Selected platform is',Platform)
                PlatformCondition = True
            elif choice1 == 4:
                Platform = input('Enter the platform name: ')
                print()
                print('Selected platform is',Platform)
                PlatformCondition = True
            else: print('Invalid input')
            
        JobStatus = ''
        
        print('============================================================')
        # Applied through what
        print()
        print()
        SubmitPlaceCondition = False
        while not SubmitPlaceCondition:
            print('Enter 1 for "Applied through ', Platform, '"', sep = '')
            print('Enter 2 for "Applied through thier website"')
            print('Enter 3 for manual input')
            choice2 = int(input('Enter your choice: '))
            if choice2 == 1:
                SubmitPlace = 'Applied through ' + Platform
                print()
                print('Application ', SubmitPlace)
                SubmitPlaceCondition = True
            elif choice2 == 2:
                SubmitPlace = 'Applied through their website'
                print()
                print('Application', SubmitPlace)
                SubmitPlaceCondition = True
            elif choice2 == 3:
                SubmitPlace = input('Enter the platform name: ')
                print()
                print('Application', SubmitPlace)
                SubmitPlaceCondition = True
            else: print('Invalid input')

        print('============================================================')
        print()
        print()
        Salary = input('Enter the Salary range: ')
        Notes = input('Enter your notes: ')
        Progress = ''
        
        # Update the exel file
        df2 = pd.DataFrame([Sr, Date, Title, CompanyName, JobLink, Location, 
                        Platform,JobStatus, SubmitPlace, Salary, Notes, 
                        Progress])
        df2 = df2.T
        df2.columns = df.columns.values
        df1 = df.append(df2, ignore_index = True)
        
   
        excelpath = os.path.join('D:/Jobs', 'Applied Jobs1.xlsx')
        writer = pd.ExcelWriter(excelpath, engine='xlsxwriter')
        df1.to_excel(writer, sheet_name='Sheet1', startrow=0, startcol=0, header=True, index=False)
        writer.save()
        
        # Open the new file
        absolutePath = Path(excelpath).resolve()
        os.system(f'start excel.exe "{absolutePath}"')
    
        ##
        #this function 
        #
    def SearchCompanyName(self):
        CompanyName = input('Enter the company name: ')
        print()
        if CompanyName in(df.Company.tolist()):
            print('You applied in "', CompanyName, '" before', sep ="")
            print()
            print("Previous applications' details:")
            print(print(
                df[df.Company == CompanyName]
                [['Application Date','Job Title','Location',
                  'Platform','Application Place','Feedback']])
                )
        else:
            print('No previous applications for "',CompanyName,'"', sep = "")
            
        
jobs = Jobs()

# Main Enterence
print()
SelectionCondition = False
while not SelectionCondition:
    print('Enter 1 for getting the applied jobs analysis')
    print('Enter 2 for adding new job')
    print('Enter 3 for search company name')
    choice3 = int(input('Select your choice: '))
    if choice3 == 1:
        jobs.weekly_effeciency()
        SelectionCondition = True
    elif choice3 == 2:
        jobs.NewJob()
        SelectionCondition = True
    elif choice3 == 3:
        jobs.SearchCompanyName()
        SelectionCondition = True
    else:
        print('Invalid input')
        print()
        print()
