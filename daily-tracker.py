#!/usr/bin/env python3

from urllib.parse import quote
import webbrowser
from datetime import date
import xlrd               
            
# setup
today = date.today()
date_string = today.strftime("%Y-%m-%d")


# -------Deviations
# --edit this section if there is any issues or changes
filePath = "I:\\Quality\\Deviations\\DR Logs\\"
workbook = xlrd.open_workbook(filePath+"Deviation Log.xlsm")
worksheet = workbook.sheet_by_name('Master')
emails_string ="ECO@izimed.com;greg.groenke@izimed.com"


summary = ""
recipients = []
expired_summary = ""
draft_summary = ""



# reading from log  
devTotal = worksheet.nrows 
for i in range(2,devTotal):
    try:
        days = int(worksheet.cell(i, 10).value)
    except:
        days = 60
    if days <= 7:

        devNumber = str(worksheet.cell(i, 0).value)
        location = str(worksheet.cell(i, 2).value)
        excel_date = worksheet.cell(i, 7).value
        try:
            y, m, d, h, i, s = xlrd.xldate_as_tuple(excel_date, workbook.datemode)
            date = str("{0}-{1}-{2}".format(y, m, d))
        except:
            date = excel_date
        status  = "Deviation "+devNumber + " expires on " + date + ", Owner: " + location  
        summary = summary + "\n" + status   
    
    dev_status = str(worksheet.cell(i, 1).value)
    if dev_status == "Expired":
        devNumber = str(worksheet.cell(i, 0).value)
        location = str(worksheet.cell(i, 2).value)
        excel_date = worksheet.cell(i, 7).value  
        try:
            y, m, d, h, i, s = xlrd.xldate_as_tuple(excel_date, workbook.datemode)
            date = str("{0}-{1}-{2}".format(y, m, d))
        except:
            date = excel_date
        expired_status  = "Deviation "+devNumber + " expired on " +date + ", Owner: " + location  
        expired_summary = expired_summary + "\n" + expired_status
    elif dev_status =="Drafting" :
        try:
            days = int(worksheet.cell(i, 5).value)
        except:
            days = 60
        if days >= 14: 
            devNumber = str(worksheet.cell(i, 0).value)
            location = str(worksheet.cell(i, 2).value)
            excel_date = worksheet.cell(i, 4).value  
            try:
                y, m, d, h, i, s = xlrd.xldate_as_tuple(excel_date, workbook.datemode)
                date = str("{0}-{1}-{2}".format(y, m, d))
            except:
                date = excel_date
            draft_status  = "Deviation "+devNumber + " drafting since " + date + ", Owner: " + location  
            draft_summary = draft_summary + "\n" + draft_status

if summary == "":
        summary = "No upcoming expiring deviations."
else:
    summary = "Expiring Deviations:\n" +summary  

if expired_summary == "":
        expired_summary = "\n\nNo expired deviations."
else:    
    expired_summary = "\n\nExpired Deviations:\n" + expired_summary

if draft_summary == "":
        draft_summary = "\n\nNo deviations drafting longer than 2 weeks."
else:
    draft_summary = "\n\nDeviations Drafting for more than 2 weeks:\n" + draft_summary


summary = summary  + draft_summary + expired_summary

#generate email
def mailto(recipients, subject, body):
    webbrowser.open("mailto:%s?subject=%s&body=%s" %
        (recipients, quote(subject), quote(body)))
subject =  "Deviation Update: " + date_string

def gen(emails, body):
    mailto(emails,subject, body)

    
print (summary)
gen(emails_string, summary)


# ------------------------------------------ECO-------------------------
# edit this section for path changes
filePath = "I:\\Quality\\ECO\\"
workbook = xlrd.open_workbook(filePath+"ECO Log.xlsm")
worksheet = workbook.sheet_by_name('Master')

# setup
summary = ""
recipients = []


# reading from log  
ecoTotal = worksheet.nrows 
for i in range(2,ecoTotal):
    try:
        days = int(worksheet.cell(i, 17).value)
    except:
        days = 0
    if days >= 7:
        location = str(worksheet.cell(i, 15).value)
        ecoNumber = int(worksheet.cell(i, 0).value)
        originator = str(worksheet.cell(i, 1).value)
        status  = "ECO"+str(ecoNumber) + " by " + originator + " has been with " + location + " for " +str(days) + " days."       
        summary = summary + "\n" + status

if summary == "":
        summary = "No ECOs lagging more than 2 weeks as of " + date_string+ "."
else:
    summary = "ECOs without movement in a week - " + date_string+ ":\n"+ summary + "\n\nContact Doc Control if these statuses are not correct and please update Doc Control when ECOs move."      
             

#generate email
print (summary)
subject = "ECO Update: " + date_string
gen(emails_string, summary)


# -------NCMRs----------------------------------------------------
# edit this section only
filePath = "I:\\Quality\\Non-Conformances\\"
workbook = xlrd.open_workbook(filePath+"NCMR Log.xlsm")
worksheet = workbook.sheet_by_name('Master')


drafting_summary = ""
open_summary = ""
recipients = []

# reading from log  
itemTotal = worksheet.nrows 
for i in range(2,itemTotal):
    try:
        days = int(worksheet.cell(i, 5).value)
    except:
        days = 0
    if days >= 60:
        ncmr = str(worksheet.cell(i, 0).value)
        originator = str(worksheet.cell(i, 2).value)

        drafting_status  = "NCMR "+ ncmr + " has been drafting for " +str(days) + " days, Owner: " + originator
        drafting_summary = drafting_summary + "\n" + drafting_status   
    
    ncmr_status = str(worksheet.cell(i, 1).value)
    if ncmr_status == "Executing":
        try:
            days = int(worksheet.cell(i, 9).value)
        except:
            days = 0
        if days >= 60:
            ncmr = str(worksheet.cell(i, 0).value)
            originator = str(worksheet.cell(i, 2).value)
            excel_date = worksheet.cell(i, 6).value  
            try:
                y, m, d, h, i, s = xlrd.xldate_as_tuple(excel_date, workbook.datemode)
                date = str("{0}-{1}-{2}".format(y, m, d))
            except:
                date = excel_date
            open_status  = "NCMR "+ ncmr + " pending close out since " +date + ", Owner: " + originator  
            open_summary = open_summary + "\n" + open_status


if drafting_summary == "":
        drafting_summary = "No NCMRs drafting for more than 60 days."
else:
    drafting_summary = "NCMRs Drafting for more than 60 days:\n" + drafting_summary
    
if open_summary == "":
        open_summary = "\nNo NCMRs open for more than 60 days."
else:    
    open_summary = "\n\nNCMRs Open for more than 60 days:\n"  + open_summary

          
subject =  "NCMR Update: " + date_string
print (drafting_summary + open_summary )
gen(emails_string, drafting_summary + open_summary + "\n\nPlease contact Quality if the statuses are not correct.")










