# -*- coding: utf-8 -*-
"""
Created on Tue Jul  3 11:45:04 2018
@author: zmohamadazri
"""
import os
import openpyxl
import xlrd
import time
from packagetools.urlaccess import URL, deletefile

def checkupdate(url, indicator_name, stpname, save_path, driver, timepoint1, ref):
    u = URL(url)
    
    ## Do your things here ##
    if indicator_name == "Source1":
        last_update = u.urlrequests()
        
    elif indicator_name == "Source2":
        last_update = u.getdriver(driver)
    
    elif indicator_name == "Source3":
        last_update = u.urlgetjson()
    ## Do your things here ##
    
    return last_update
