# -*- coding: utf-8 -*-
"""
Created on Mon Jul  2 16:07:42 2018

@author: zmohamadazri
"""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import email.mime.application
import datetime

class SendEmail(object):
    def __init__(self, serverhost, fromaddr, toaddr, ccaddr, mnts):
        self.serverhost = serverhost
        self.fromaddr = fromaddr
        self.toaddr = toaddr
        self.ccaddr = ccaddr
        self.mnts = mnts
        
    def sendmail(self, attachments, htmlbody, newreleases, failedreleases, allurl, subj):
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subj + str((datetime.datetime.now() + datetime.timedelta(hours=8)).strftime("%d-%m-%Y"))
        msg['From'] = 'RCTCore <RCTCore@isimarkets.com>'
        msg['To'] = self.toaddr
        msg['CC'] = self.ccaddr
        toaddrs = [self.toaddr] + self.ccaddr.split(',')
        
        if htmlbody != None:
            htmltable = open(htmlbody).read()
            htmltable = htmltable.replace('align=center' , 'align=left')
        else:
            htmltable = ''
        
        HTMLBody = "<font face = 'calibri'><font size = '3'>Next Schedule: " + str((datetime.datetime.now() + datetime.timedelta(hours=8,minutes=self.mnts)).strftime("%d %b %Y  %I:%M %p")) + "<br>" \
                + "Success: " + str(allurl-failedreleases) + "/" + str(allurl) + "<br>" \
                + "Failed: " + str(failedreleases) + "/" + str(allurl) + "<br>" \
                + "New: "  + str(newreleases) \
                + "</font><br>" \
                +  htmltable \
                + "<br>" \
                + "<font face = 'calibri'><font size = '3'>Matrix Manager's checking results are a benchmark for your further action." + "<br>" \
                + "Please contact us at (RCTCore@isimarkets.com) if:" + "<br>" \
                + "1) No email received on the time stated in 'Next Schedule' as above" + "<br>" \
                + "2) Enhancement requests for remarks stated 'Fail - Website Layout Change/Server Down'" + "<br>" \
                + "3) Any other enquires</font>"
            
        msg.attach(MIMEText(HTMLBody, 'html'))
        
        if newreleases != 0 or failedreleases != 0:
            filename=attachments
            fp=open(filename,'rb')
            att = email.mime.application.MIMEApplication(fp.read(),_subtype="xlsx")
            fp.close()
            att.add_header('Content-Disposition','attachment',filename=str((datetime.datetime.now() + 
                                                                            datetime.timedelta(hours=8)).strftime("%d-%m-%Y %H:%M")) + '.xlsx')
            msg.attach(att)
        
        server = smtplib.SMTP(self.serverhost)
        server.sendmail(self.fromaddr, toaddrs, msg.as_string())
        server.quit()
