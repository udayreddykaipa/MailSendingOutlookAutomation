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
    data = pd.ExcelFile(r'mdpi_pharma.xlsx')  # give exact file path here
    sheet = data.parse('Sheet1') # sheet name case sensitive 
    # print(len(sheet['Name'])) # no of rows 
    for kk in range(len(sheet['Name'])):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            for accou in outlook.Session.Accounts:
                print(accou);
                if accou == 'udayreddy.kaipa@gmail.com': # select that account as default profile, change this to what is printed.
                    useacc = accou;
                    break
            mail = outlook.CreateItem(0)
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, accou))
            name = sheet['Name'][kk]
            mailto = sheet['Email'][kk]
            subject = sheet['subjectlines'][kk]
            title = sheet['Title'][kk]
            template1 = '<table>\
                        <tr><td><b>Dear Dr. {auname},</b><br/><br/></td></tr>\
                        <tr><td>Greetings from <a href="http://www.opastonline.com/journal/journal-of-pharmaceutical-research">Journal of Pharmaceutical Research</a>!<br/><br/></td></tr>\
                        <tr><td>We have gone through your recent articles {atitle} perceived outstanding knowledge and valuable information useful to the scientific community.<br/><br/></td></tr>\
                        <tr><td>Please contribute your articles on or before <b>February 28, 2021 </b> at: ruth(at)e-openaccess.info or via <a href="https://www.opastonline.com/submit-manuscript/">online.<br/></td></tr>\
                        <tr><td>Looking forward to work with you.<br/><br/></td></tr>\
                        <tr><td><b>Regards,</b><br/>Ruth<br/>Editorial Manager</td></tr>\
                        </table>'.format(auname=name,atitle=title)
            templates = [template1]
            eml = random.choice(templates)
            mail.To = mailto
            mail.Subject = subject
            mail.HTMLBody = eml
            
            timelist = [2, 3, 4]
            sltime = random.choice(timelist)
            time.sleep(sltime)
            #mail.display(True)
            mail.send
            #print("sent to ",mailto)
        except (KeyError, pywintypes.com_error, TypeError, AttributeError):
            print("Exception")
            continue



mailprepare()

# mail.send

# outlook
