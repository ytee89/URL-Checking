# -*- coding: utf-8 -*-
"""
Created on Wed Jul 11 10:38:13 2018

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

def run_url_checking(masterfile):
    starttime = datetime.datetime.now()
    print(starttime.strftime("%d %b %Y  %I:%M:%S %p"))
    
    #define all related files
    masterfolder = dirname(abspath(__file__))
    
    excel1 = join(masterfolder, masterfile)
    excel2 = join(masterfolder, "Report.xlsx")
    excel3 = join(masterfolder, "Email.xlsx")
    htmlfile = join(masterfolder, "Email.html")
    mdbfile = join(masterfolder, "Automation.mdb")

    #convert masterfile to dataframe
    df1 = pd.read_excel(excel1, sheet_name='Sheet1')
    df12 = pd.read_excel(excel1, sheet_name = 'Sheet2', index_col=0, header=None)
    
    #create dataframe for use in report attachment
    df2 = pd.DataFrame(columns=['Source','STP Name', 'New Timepoint','Previous Timepoint', 'Changes Type', 
                                'Key', 'Frequency', 'Level', 'System ID', 'Method', 'Remark', 'Requested Time'])
    df2m = pd.DataFrame(columns=['Source','STP Name', 'Changes Type', 'Key', 'Frequency', 'Level', 'System ID', 'Method', 'Remark', 'Requested Time'])
    
    #clear columns value on the masterfile
    df1['TimePoint Source'] = None
    df1['Changes Type'] = None
    df1['Status'] = None
    df1['Last Timepoint'] = None
    
    fromaddress = str(df12.loc['From'].get(1)).strip()
    toaddress = str(df12.loc['To'].get(1)).strip()
    ccaddress = str(df12.loc['CC'].get(1)).strip()
    countrycode = str(df12.loc['Country Code'].get(1)).strip()
    
    consolfolder = join(masterfolder, 'Consolidated Report')
    if not exists(consolfolder):
        os.makedirs(consolfolder)
    else:
        pass
    excel4 = join(consolfolder, countrycode+' Consolidated Report.xlsx')
    mdbRCfile = join(masterfolder, countrycode+'ConsolidatedReport.mdb')
    
    #open selenium webdriver
    driver = urlaccess.openwebdriver()
    
    #iterate row by row through the dataframe
    for i in df1.index:
        url = df1.loc[i, 'URL']
        indicator_name = df1.loc[i, 'Indicator']
        stpname = df1.loc[i, 'STP Name']
        ref = df1.loc[i, 'Ref']
        timepoint1 = df1.loc[i, 'Current TimePoint']
        save_path = join(masterfolder, 'file')
        
        c = CheckingResult(i, df1, df2, df2m)
        
        try:
            #check the latest timepoint
            last_update = sourcecode.checkupdate(url, indicator_name, stpname, save_path, driver, timepoint1, ref)
            
            df1.loc[i, 'TimePoint Source'] = str(last_update)
            
            #check whether there are a new updates or failed and make changes in the dataframe
            if df1.loc[i, 'TimePoint Source'] == '' or df1.loc[i, 'TimePoint Source'] == None:
                c.failed('Fail - Website Layout Change/Server Down')
            elif df1.loc[i, 'Current TimePoint'] != df1.loc[i, 'TimePoint Source']:
                c.updatedetected()
                c.updatemdbRC(mdbRCfile)
                if 'Macro' in str(df1.loc[i, 'Update Method']): 
                    c.updatemdb(mdbfile, countrycode)
            else:
                c.uptodate()
               
        #error handler
        except AttributeError:
            c.failed('Fail - Website Layout Change/Server Down')
        except NameError:
            c.failed('Fail - Website Layout Change/Server Down')
        except HTTPError:
            c.failed('Fail - Website Layout Change/Server Down')
        except ConnectionError:
            c.failed('Fail - Connection unstable')
        except ReadTimeout:
            c.failed('Fail - Website Layout Change/Server Down')
        except TimeoutException:
            c.failed('Fail - Connection unstable')
        except WebDriverException:
            c.failed('Fail - Website Layout Change/Server Down')
        except Exception:
            c.failed('Fail - Website Layout Change/Server Down')
            
        print(str(i+1)+' '+str(df1.loc[i,'STP Name'])+'\n'+str(df1.loc[i,'Changes Type'])+'\n')
    
    #close selenium webdriver
    driver.quit()
    
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
        convert.dftoreport(excel2, df2.drop(['Requested Time'], axis=1), df2m.drop(['Requested Time'], axis=1))#convert dataframe to excel report attachment
        ec = ExcelChanges(excel2, excel3, df1)
        ec.reporttoemail()#write email body in excel
        convert.exceltohtml(excel3, htmlfile)#convert excel email body to html
        
        if newreleases != 0:
            email.sendmail(excel2, htmlfile, newreleases, failedreleases, manualreleases, allurl, 'Alert! | '+countrycode+'_Release Detected_')
            consolidate(df2, df2m, excel4)
        elif failedreleases > manualreleases:
            email.sendmail(excel2, htmlfile, newreleases, failedreleases, manualreleases, allurl, 'Failed | '+countrycode+'_No Release Detected_')            
        else:
            htmlfile = None
            consolidate(df2, df2m, excel4)
            email.sendmail(excel2, htmlfile, newreleases, failedreleases, manualreleases, allurl, 'Failed | '+countrycode+'_No Release Detected_')
        
    else:
        htmlfile = None
        email.sendmail(excel2, htmlfile, newreleases, failedreleases, manualreleases, allurl, 'All Up To Date | '+countrycode+'_No Release Detected_')
    
    #remove report attachment, excel and html email body
    urlaccess.deletefile(excel2)
    urlaccess.deletefile(excel3)
    urlaccess.deletefile(htmlfile)
    
    endtime = datetime.datetime.now()
    print(endtime.strftime("%d %b %Y  %I:%M:%S %p"))
    
    print('\nTotal Running Time: ' + str(endtime-starttime))
    
if __name__ == "__main__":
    masterfile = "URL Checking.xlsx"
    run_url_checking(masterfile)
    
    user = os.getlogin()
    tempfolder = 'C:\\Users\\'+ user+'\\AppData\\Local\\Temp'
    
    for allfiles in os.listdir(tempfolder):
        if allfiles.startswith('scoped_dir'):
            file_path = os.path.join(tempfolder, allfiles)
            try:
                shutil.rmtree(file_path)
            except Exception as e:
                pass
