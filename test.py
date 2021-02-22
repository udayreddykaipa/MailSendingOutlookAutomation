import pywintypes
from win32com.client import Dispatch
import win32com
import time
import datetime
import random
import openpyxl
# from openpyxl import load+_workbook
import pandas as pd


def mailprepare():
    data = pd.ExcelFile('t.xlsx')  # give exact file path here
    sheet = data.parse('Sheet1') # sheet name case sensitive 
    # print(len(sheet['Name'])) # no of rows 
    for kk in range(len(sheet['Name'])):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            for accou in outlook.Session.Accounts:
                print(accou, accou.SmtpAddress)
                if accou.SmtpAddress == 'udayreddy.kaipa@gmail.com':
                    useacc = accou
                    break
            mail = outlook.CreateItem(0)
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, useacc))
            name = sheet['Name'][kk]
            mailto = sheet['Email'][kk]
            subject = sheet['subjectlines'][kk]
            print(mail,mailto,subject)
            template1 = '<table>\
                        <tr><td><b>Dear Dr. {auname},</b><br/><br/></td></tr>\
                        <tr><td>Greetin{atitle}gs</td></tr>\
                        <tr><td><b>Regards,</b><br/></td></tr>\
                        </table>'.format(auname=name)

            templates = [template1]
            eml = random.choice(templates)
            mail.To = "udayreddy.kaipa@gmail.com"
            mail.Subject = subject
            mail.HTMLBody = eml
            
            timelist = [10, 11, 12]
            sltime = random.choice(timelist)
            time.sleep(sltime)
            mail.display(True)
            mail.send
        except (KeyError, pywintypes.com_error, TypeError, AttributeError):
            print("Exception")
            continue



mailprepare()

# mail.send

# outlook
