# -*- coding: utf-8 -*-
"""
Created on Wed Jul 11 10:38:13 2018

pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org pebble

@author: zmohamadazri
"""
import os
import shutil
from os.path import join, dirname, abspath, exists
import pandas as pd
from requests.exceptions import HTTPError, ConnectionError, ReadTimeout
from selenium.common.exceptions import TimeoutException, WebDriverException
import urllib3
urllib3.disable_warnings()
import datetime

import sourcecode
import packagetools.urlaccess as urlaccess
import packagetools.convertfiles as convert
from packagetools.sendemail import SendEmail
from packagetools.datachanges import CheckingResult, ExcelChanges, consolidate

from pebble import ProcessPool as Pool
from concurrent.futures import TimeoutError

import signal
import subprocess


def rc_init(masterfile, masterfolder, mdbfolder, RCmdbfolder):
    global starttime
    starttime = datetime.datetime.now()
    print(starttime.strftime("%d %b %Y  %I:%M:%S %p"))
    
    #define all related files
    global excel1; global excel2; global excel3
    global errorfile; global mdbfile; global htmlfile
    
    excel1 = join(masterfolder, masterfile)
    excel2 = join(masterfolder, "Report.xlsx")
    excel3 = join(masterfolder, "Email.xlsx")
    htmlfile = join(masterfolder, "Email.html")
    errorfile = join(masterfolder, "errorlog.txt")
    mdbfile = join(masterfolder, "Automation.mdb")

    #convert masterfile to dataframe
    global df1; global df12
    df1 = pd.read_excel(excel1, sheet_name='Sheet1')
    df12 = pd.read_excel(excel1, sheet_name = 'Sheet2', index_col=0, header=None)
    
    #create dataframe for use in report attachment
    global df2; global df2m
    df2 = pd.DataFrame(columns=['Source','STP_Name', 'New_Timepoint','Last_Timepoint', 'Changes_Type', 'Key_Status', 
                                'Frequency', 'Edge_Level', 'System_ID', 'Update_Method', 'Remarks', 'Release_Time'])
    df2m = pd.DataFrame(columns=['Source','STP_Name', 'New_Timepoint','Last_Timepoint', 'Changes_Type', 'Key_Status', 
                                'Frequency', 'Edge_Level', 'System_ID', 'Update_Method', 'Remarks', 'Release_Time', 'No of Fail'])
    
    #clear columns value on the masterfile
    df1['TimePoint Source'] = None
    df1['Changes Type'] = None
    df1['Status'] = None
    df1['Last Timepoint'] = None
    
    global fromaddress; global toaddress; global ccaddress; global countrycode; global consolfolder
    fromaddress = str(df12.loc['From'].get(1)).strip()
    toaddress = str(df12.loc['To'].get(1)).strip()
    ccaddress = str(df12.loc['CC'].get(1)).strip()
    countrycode = str(df12.loc['Country Code'].get(1)).strip()
    
    consolfolder = join(masterfolder, 'Consolidated Report')
    if not exists(consolfolder):
        os.makedirs(consolfolder)
    
    global excel4
    excel4 = join(consolfolder, countrycode+' Consolidated Report.xlsx')
    
    global mdbRCfile
    mdbRCfile = join(masterfolder, countrycode+'ConsolidatedReport.mdb')
    
def rc_init_simplified(masterfolder, masterfile):
    #Simplified intialization to restart checking
    print('Loading...')
    
    dlfolder = join(masterfolder, 'Download')
    if not exists(dlfolder):
        os.makedirs(dlfolder)
        
    global df_rc_readonly; df_rc_readonly = pd.read_excel(join(masterfolder, masterfile), sheet_name='Sheet1')
    global driver; driver = urlaccess.openwebdriver(dlfolder)
    global save_path; save_path = join(masterfolder, 'file')

def rc_process():
    with Pool(max_workers=1, initializer=rc_init_simplified, initargs=(masterfolder, masterfile,)) as p:        
        for i in range(0, 5):
#        for i in df1.index:
            c = CheckingResult(i, df1, df2, df2m, errorfile)
            chromepid = p.schedule(return_chromepid).result()
            future = p.map(run_url_checking, (i,), timeout=120)
            
            new_tp_validation = True
            try:
                fr = future.result()
                last_update = [fr1 for fr1 in fr][0]
                df1.loc[i, 'TimePoint Source'] = str(last_update)
            #error handler
            except (ConnectionError, TimeoutException):
                c.failed('Fail - Connection unstable')
                new_tp_validation = False
            except TimeoutError:
                killallchromeprocess(chromepid)
                new_tp_validation = False
                c.failed('Fail - Website Layout Change/Server Down2')
            except Exception as e:
                print(str(e))
                c.failed('Fail - Website Layout Change/Server Down')
                new_tp_validation = False
            finally:                       
                #check whether there are a new updates or failed and make changes in the dataframe
                if new_tp_validation:
                    if (df1.loc[i, 'TimePoint Source'] == '' or df1.loc[i, 'TimePoint Source'] == None):
                        c.failed('Fail - Website Layout Change/Server Down')
                    elif df1.loc[i, 'Current TimePoint'] != df1.loc[i, 'TimePoint Source']:
                        c.updatedetected()
                        if 'Macro' in str(df1.loc[i, 'Update Method']): 
                            c.updatemdb(mdbfile, countrycode)
                    else:
                        c.uptodate()
                c.updatemdbRC(mdbRCfile)                
                print(str(i+1)+' '+str(df1.loc[i,'STP Name'])+'\n'+str(df1.loc[i,'Changes Type'])+'\n')
        
        p.schedule(return_quitdriver)
        
def return_chromepid():
#    https://stackoverflow.com/questions/10752512/get-pid-of-browser-launched-by-selenium
    return driver.service.process.pid 

def return_quitdriver():
    driver.quit()
    
def killallchromeprocess(cpid):
    print('Terminating process...')
    cmd = 'wmic process get processid,parentprocessid,executablepath | find "chrome.exe" |find "' + str(cpid) + '"'
    p = subprocess.Popen(cmd, stdout=subprocess.PIPE, shell=True)
    (output, err) = p.communicate()
    cpid2 = str(output).split(str(cpid))[-1].replace('\\r', '').replace('\\n', '').replace("'", '').strip() 
    os.kill(int(cpid), signal.SIGINT)
    os.kill(int(cpid2), signal.SIGINT)

def rc_final():   
    #count number of url, new releases, and failed
    newreleases = len(df1[df1['Changes Type'] == 'New Detected'])
    failedreleases = len(df1[df1['Changes Type'] != 'Up to date']) - len(df1[df1['Changes Type'] == 'New Detected'])
    manualreleases = len(df2m)
    allurl = len(df1)
    
    #convert dataframe back to excel masterfile
    convert.dftomasterfile(excel1, df1)
    
    #write and convert email excel body to html if new releases or failed and send email
    serverhost = 'ceicdata-com.mail.protection.outlook.com'
    email = SendEmail(serverhost, fromaddress, toaddress, ccaddress, 10)
    
    if newreleases != 0 or failedreleases != 0:
        convert.dftoreport(excel2, df2, df2m)#convert dataframe to excel report attachment
        ec = ExcelChanges(excel2, excel3, df1)
        ec.reporttoemail()#write email body in excel
        convert.exceltohtml(excel3, htmlfile)#convert excel email body to html
        
        if newreleases != 0:
            email.sendmail(excel2, htmlfile, newreleases, failedreleases, manualreleases, allurl, 'Alert! | '+countrycode+'_Release Detected_')
            consolidate(df2, df2m, excel4)
        elif failedreleases > manualreleases:
            email.sendmail(excel2, htmlfile, newreleases, failedreleases, manualreleases, allurl, 'Failed | '+countrycode+'_No Release Detected_')
        else:
            consolidate(df2, df2m, excel4)
            email.sendmail(excel2, None, newreleases, failedreleases, manualreleases, allurl, 'Failed | '+countrycode+'_No Release Detected_')
        
    else:
        email.sendmail(excel2, None, newreleases, failedreleases, manualreleases, allurl, 'All Up To Date | '+countrycode+'_No Release Detected_')
    
    #remove report attachment, excel and html email body
    urlaccess.deletefile(excel2)
    urlaccess.deletefile(excel3)
    urlaccess.deletefile(htmlfile)
    
    endtime = datetime.datetime.now()
    print(endtime.strftime("%d %b %Y  %I:%M:%S %p"))
    
    print('\nTotal Running Time: ' + str(endtime-starttime))
    
def run_url_checking(i):    
    #iterate row by row through the dataframe
    url = df_rc_readonly.loc[i, 'URL']
    indicator_name = df_rc_readonly.loc[i, 'Indicator']
    stpname = df_rc_readonly.loc[i, 'STP Name']
    ref = df_rc_readonly.loc[i, 'Ref']
    timepoint1 = df_rc_readonly.loc[i, 'Current TimePoint']
    
    return sourcecode.checkupdate(url, indicator_name, stpname, save_path, driver, timepoint1, ref)

#"C:\Users\ytee\AppData\Local\Continuum\anaconda3\python.exe" "C:\Users\ytee\OneDrive - Internet Securities, LLC\Desktop\Quickhelp\C4C\script\Python (ver7)\URY\mainfile.py"
if __name__ == "__main__":
    masterfile = "URL Checking.xlsx"
    masterfolder = dirname(abspath(__file__))
    mdbfolder = dirname(abspath(__file__))
    RCmdbfolder = dirname(abspath(__file__))    
    
    rc_init(masterfile, masterfolder, mdbfolder, RCmdbfolder)
    rc_process()
    rc_final()
        
    user = os.getlogin()
    tempfolder = 'C:\\Users\\'+ user+'\\AppData\\Local\\Temp'
    
    for allfiles in os.listdir(tempfolder):
        if allfiles.startswith('scoped_dir'):
            file_path = os.path.join(tempfolder, allfiles)
            try:
                shutil.rmtree(file_path)
            except Exception as e:
                pass
