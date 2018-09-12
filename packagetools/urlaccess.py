# -*- coding: utf-8 -*-
"""
Created on Tue Jul  3 10:41:57 2018

@author: zmohamadazri
"""

import os
import requests
from selenium import webdriver
import json
from bs4 import BeautifulSoup as soup
import PyPDF2

def openwebdriver():
    options = webdriver.ChromeOptions()
    options.add_argument('headless')
    driver = webdriver.Chrome(chrome_options=options)
    return driver

def deletefile(file):
    try:
        os.remove(file)
    except:
       pass
    
class URL(object):
    def __init__(self, url):
        self.url = url
        
    def urlrequests(self):
        #request url and read html using bs4
        headers={'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36'}
        session = requests.Session()
        html = session.get(self.url, verify=False, timeout=30, headers=headers)
        page = soup(html.content, "lxml")
        return page
    
    def pdfmoddate(self, save_path):
        #read last modified date of pdf file    
        pdfFileObj = open(save_path, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        
        try:
            last_update1 = pdfReader.getDocumentInfo()
            last_update = str(last_update1['/ModDate'])
        except PyPDF2.utils.PyPdfError as pdferror:
            last_update = ""
            
        pdfFileObj.close()
        
        return last_update

    def urlgetjson(self):
        #request url by get method and read json
        session = requests.Session()
        html = session.get(self.url, verify=False, timeout=30)
        page = json.loads(html.content.decode('utf-8'))
        return page
    
    def urlpostjson(self,formdata):
        #request url by post method and read json
        session = requests.Session()
        html = session.post(self.url, verify=False, timeout=30, data=formdata)
        page = json.loads(html.content.decode('utf-8'))
        return page
    
    def dlfile(self, save_path):
        #download file from direct url using request
        r = requests.get(self.url,verify=False, timeout=30)
        with open(save_path,'wb') as f:   
            f.write(r.content)
    
    def getdriver(self,dr):
        #open url using selenium
        dr.set_page_load_timeout(30)
        dr.get(self.url)
        html = dr.page_source
        page = soup(html, "lxml")
        return page