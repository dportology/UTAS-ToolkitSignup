#-------------------------------------------------------------------------------
# Name:        NPI TOOLKIT AUTOSIGNUP TOOL
#
# Purpose: In anticipation of several users for the NPI Toolkit, the purpose
# of this program will be to create database users completely automatically.
# The program will read through an outlook email account (now set to mine
# alex.davenport@utas.utc.com) and pull out any emails with the subject
# 'toolkit signup', or any slight variations, connect to the sql server hosting
# the toolkit, add the user, and email them their account information. This is
# made in response to other UTAS systems which seems to take a minimum of a week
# to get access to. Now users can begin work immedietly.
#
# Author:      Alex Davenport
# Last Edited:     6/18/2015
#-------------------------------------------------------------------------------

from tkFileDialog import askopenfile
import adodbapi
import random
import string
import win32com.client as win32
from collections import OrderedDict
import win32com.client
import time
import pywintypes
import codecs
import tempfile
import win32api
from win32com.mapi.mapitags import *
from win32com.mapi import mapi
from win32com.mapi import mapiutil
import ctypes
import win32con
import win32gui
import win32process
import thread
import threading
from PyQt4.QtCore import *
from PyQt4.QtGui import *

class StoppableThread(threading.Thread):

    def __init__(self):
        super(StoppableThread, self).__init__()
        self._stop = threading.Event()

    def threadedFunction():

        import re
        import win32gui

        windows = []
        classes = {}

        win32gui.EnumWindows(win_enum_handler, classes)

        for key, value in classes.iteritems():
            if value.endswith('Microsoft Outlook'):
                print key, value
                win32gui.SetForegroundWindow(key)

        print 'ending thread'
        self._stop.set()

    def stop(self):
        self._stop.set()

    def stopped(self):
        return self._stop.isSet()


def win_enum_handler(hwnd, titles):
    titles[hwnd] = win32gui.GetWindowText(hwnd)


def threadedFunction():

    import re
    import win32gui

    windows = []
    classes = {}

    win32gui.EnumWindows(win_enum_handler, classes)

    for key, value in classes.iteritems():
        if value.endswith('Microsoft Outlook'):
            print key, value
            win32gui.SetForegroundWindow(key)
            print '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys("{ENTER}") # Delete selected text?  Depends on context. :P


def computerIsLocked():

    user32 = ctypes.windll.User32
    OpenDesktop = user32.OpenDesktopA
    SwitchDesktop = user32.SwitchDesktop
    DESKTOP_SWITCHDESKTOP = 0x0100
    hDesktop = OpenDesktop ("default", 0, False, DESKTOP_SWITCHDESKTOP)
    result = SwitchDesktop (hDesktop)
    if result:
        return False
    else:
        return True


# emails from exchange are sent with name only. need to look up
# smtp address for return email using the address book
# pass list of names from mailChecker function, return their coresponding
# smtp addresses as a list
def getAddressBook(names):

    profileName = "alex.davenport@utas.utc.com"
    session = mapi.MAPIInitialize(None)
    session =mapi.MAPILogonEx(0,profileName,None, mapi.MAPI_EXTENDED |
    mapi.MAPI_LOGON_UI |\
                                 mapi.MAPI_NO_MAIL |mapi.MAPI_USE_DEFAULT)
    hr=session.OpenAddressBook(0,None,mapi.AB_NO_DIALOG)

    root=hr.OpenEntry(None,None,mapi.MAPI_BEST_ACCESS)

    root_htab=root.GetHierarchyTable(0)

    DT_GLOBAL=131072
    restriction = (mapi.RES_PROPERTY,
                           (1,
                            PR_DISPLAY_TYPE,
                            (PR_DISPLAY_TYPE, DT_GLOBAL)))

    gal_id = mapi.HrQueryAllRows(root_htab,
                                       (PR_ENTRYID),
                                       restriction,
                                       None,
                                       0)

    gal_id = gal_id[0][0][1]

    gal=hr.OpenEntry(gal_id,None,mapi.MAPI_BEST_ACCESS)

    gal_list=gal.GetContentsTable(0)

    PR_SMTP=972947486

    rows = mapi.HrQueryAllRows(gal_list,
                                       (PR_ENTRYID,
    PR_DISPLAY_NAME_A,PR_ACCOUNT,PR_SMTP),
                                       None,
                                       None,
                                       0)

    entries = []

    for eid,name,alias,smtp in rows:
        for entry in names:
            if(name[1] == entry):
                entries.append(str(smtp[1]))

    mapi.MAPIUninitialize()
    return entries


# search through my outlook inbox and pull out names in a list from
# emails with the subject 'toolkit signup' or any similar variation of that
def mailChecker():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)

    emails = []
    messages = inbox.Items
    while messages.GetNext():

        print 'wait time...'
        time.sleep(3) # wait 5 minutes before checking
        message = messages.GetFirst()
        for mail in inbox.Items:
            subject = message.Subject

            if subject == 'toolkit signup' or subject == 'Toolkit signup' or subject == 'Toolkit Signup' \
             or subject =='toolkit sign-up' or subject == 'Toolkit sign-up' or subject == 'Toolkit Sign-up' \
             or subject == 'toolkit sign up' or subject == 'Toolkit Sign up' or subject == 'Toolkit Sign Up':
                userEmail = message.Sender.Name
                #message.Delete()
                emails.append(userEmail)

            message = messages.GetNext()

    return emails


### choose the file and return a list of all emails
### function no longer used. emails are received from the address book
##def chooseFile():
##
##    # make the TK filechooser window and get the file name
##    Tk().withdraw()
##    filename = askopenfile()
##
##    # read emails into a list
##    text_file = open(filename.name, "r")
##    emails = text_file.readlines()
##    text_file.close()
##
##    # to stop empty lines being marked as email addresses.
##    # has to have a length of at least 5 to be safe
##    results = filter(lambda x: len(x)>4, emails)
##
##    test = []
##
##    # remove any newline characters
##    for x in results:
##        test.append(x.strip('\n'))
##
##    # remove duplicates
##    test2 = list(OrderedDict.fromkeys(test))
##
##    return test2


# generates a 6 letter random string to be used
# as the initial password
def generatePassword():

    pw = ""

    for x in range(0,6):
        pw += random.choice(string.letters)

    return pw


def sendEmailConfirmation(email, username, password):

    sharedDrive = '"\\\\Huswlf0m\\enterprise\\SupplyChainManagement\\PCSS NPI Transformation\\NPI Database\\Current Version"'
    emailForQuestions = 'alex.davenport@utas.utc.com'

    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = str.lower(email)
    mail.Subject = 'NPI Toolkit Access'
    mail.Body = 'Your request to access the NPI Toolkit has been approved.\nYour username and password are as follows:\n\n'\
    'Username: ' + username + '\nPassword: ' + password + '\n\nThe Toolkit is hosted at the following location on the shared drive:\n'\
    '' + sharedDrive + '\n\n'\
    'Open the Access Database, and when attempting to access the information, you will be prompted to login.\n\n'\
    'email any questions to ' + emailForQuestions

    # if the computer is locked, automatically press the enter key to confirm the titus warning
    if not computerIsLocked():
        thread.start_new_thread(threadedFunction, ())
    mail.Send()

    print '\n\nmail sent to ' + str.lower(email)

def main():

    while 1:

        names = mailChecker()

        # get smtp addresses
        emails = getAddressBook(names)
        print 'email list is as follows: '
        print emails

        # real information to be filled in upon getting the server

        #conn = adodbapi.connect("Provider=SQLOLEDB; SERVER=xx.x.x.x; Initial Catalog=master_db;User Id=User; Password=Pass; ")
        #curs = conn.cursor()
        #adodbapi.verbose = False

        # for each email, make the username by getting the first part of the email,
        # make the pw with the random generator, and execute the create user SQL

        for user in emails:
            username = user.split('@')[0]
            password = generatePassword()
            # may need to add additional sql to grant permissions. test when server is available
            # first check if user already exists for the NPI_DB, and create them if they do not.
            sql ="USE NPI_DB\n"\
            "IF NOT EXISTS (select name FROM master.sys.server_principals WHERE name = '" + username + "')\nBEGIN\n"\
            "CREATE LOGIN """ + username + " WITH PASSWORD = '" + password + "'; \nGO\n"\
            "CREATE USER " + username + " FOR LOGIN " + username + "; \nGO"
            #curs.execute(sql)

            sendEmailConfirmation(user, username, password)

if __name__ == '__main__':
    main()
