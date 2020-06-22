"""
    Used this module to automate email sending for the 2020 Coding For Good App Dev Virtual Event
    6/1/2020 - changed to jinja2 template engine from superformatter.  gmail requires style classes
    which fooled up the superformatter.
"""

import os, sys, time, csv, smtplib, string, jinja2
from configparser import ConfigParser
# import datasheets  # if you're using Google Sheets ...
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

def TestOnly(d):
    bool = False

    if d['TAG'] == 'TEST':
        bool = True

    return bool

def HasNotApplied(d):
    """
        Return TRUE if the organization doesn't have an application date
    """
    bool = False

    if d['APPLIED'] == False:
        bool = True

    return bool


def IgnoreRejectors(d):
    """
        Return TRUE if the organization DIDN'T have a bad email or opt-out
    """
    bool = False

    if not (d['BADEMAIL'] == True or d['NOTAVAILABLE'] == True):
        bool = True

    return bool


def ReturningNotApplied(d):
    """
        Return TRUE if the Application Date is GREATER than strDateTime
    """
    bool = False

    if d['TYPE'] == 'RETURNING' and d['CONFIRMEDINTEREST'] == False and d['APPLIED'] == False:
        bool = True

    return bool


def AppliedAfter(d, date):
    """
        Return TRUE if the Application Date is GREATER than strDateTime
    """
    bool = False

    if d['APPLIED'] != False and d['APPLIED'] > date:
        bool = True

    return bool


def TypeFilter(d, strFilter):
    """
        Return TRUE if the TYPE field matches with strFilter
    """
    bool = False

    if (d['TYPE'] == strFilter):
        bool = True

    return bool


def Company(d, lstCompany):
    """
        Return TRUE if you have a match within the organizations list
    """
    bool = False

    if d['COMPANY'] in lstCompany:
        bool = True

    return bool


def NotThisCompany(d, lstCompany):
    """
        Return TRUE if you have a match within the organizations list
    """
    bool = True

    if d['COMPANY'] in lstCompany:
        bool = False

    return bool


def HasRegistered(d):
    """
        Return TRUE if the organization has registered
    """
    bool = False

    if d['APPLICATION_DATE'] != None and d['REGISTRATION_DATE'] != None and d['REMOVE_THIS_YEAR'] == False:
        bool = True

    return bool


def HasNOTRegistered(d):
    """
        Return TRUE if the organization has applied and has not registered
    """
    bool = False

    if d['APPLICATION_DATE'] != None and d['REGISTRATION_DATE'] == None and d['NTABLES'] < 5 and d[
        'REMOVE_THIS_YEAR'] == False:
        bool = True

    return bool


def StudentLedFormError(d):
    """
        Return TRUE if the PAID fee is 0, which signifies Student Led or waiver
    """
    bool = False

    if d['PAID'] == 0:
        bool = True

    return bool


def AcceptedApplication(d):
    """
        Return TRUE if they applied, we did not DROP them (REMOVE_THIS_YEAR), and they have 4 or fewer NTABLES
    """
    bool = False

    if d['APPLICATION_DATE'] != None and d['NTABLES'] < 5 and d['REMOVE_THIS_YEAR'] == False:
        bool = True

    return bool


def HaveNotApplied(d):
    """
        Return TRUE if YOU HAVE NOT applied and IF WE ARE CONSIDERING YOU this year
    """
    bool = False

    if d['APPLICATION_DATE'] == None and d['REMOVE_THIS_YEAR'] == False and d['REMOVE_PERMANENTLY'] == False and d[
        'NOTES'] == None:
        bool = True

    return bool


def LoadGoogleSheets(strSheetName, strTabName):
    """
        Note that the folder ~/.datasheets contains the Google Oath files required
        Returns a list of dictionary items representing the rows of the
        specified google sheet (strSheetName)'s tab (strTabName)
    """
    lstRows = []

    client = datasheets.Client(service=True)
    workbook = client.fetch_workbook(strSheetName)
    tab = workbook.fetch_tab(strTabName)
    df = tab.fetch_data()

    del tab
    del workbook
    del client

    lstRows = df.to_dict(orient='record')
    del df

    return lstRows


def LoadEmailList(strPathFilename):
    """
        Returns a list of dictionaries formed by parsing a CSV file (strPathFilename) with a row header.
    """
    lstReturn = []

    try:
        fp = open(strPathFilename, 'r')
    except IOError as detail:
        print("can't open input file %s. Details:%s\n" % (strPathFilename, detail))
    else:
        reader = csv.reader(fp)
        lstHeader = next(reader)  # python 3, python 2 was reader.next()
        rows = list(reader)
        del reader
        fp.close()

        for r in rows:
            d = {}
            for l in lstHeader:
                if l in ['NTABLES', 'NSTAFF', 'PAID']:  # simulate float values as per the Google Sheets
                    d[l] = float(str(r[lstHeader.index(l)]).lstrip().rstrip())
                else:
                    d[l] = str(r[lstHeader.index(l)]).lstrip().rstrip()
            lstReturn.append(d)

    return lstReturn


def SimpleEmailMessage(strToPerson, subject, html, text, lstAttachments):
    """
        Here we're using some of python's standard methods to send a simple email.
        I didn't finish this as you need to add your gmail account pw to make this work.
        Left it in here for you to discover and fix up.
    """
    cfg = ConfigParser()
    cfg.read('emailcredentials.ini')  # NEEDS TO BE IN THE SAME FOLDER AS THIS MODULE

    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = 'gsoc.codingforgood.appdev@gmail.com'
    msg['To'] = strToPerson

    msg.attach(MIMEText(text, 'plain'))
    msg.attach(MIMEText(html, 'html'))

    # provides a mean for attachment of files
    if lstAttachments != []:
        for f in lstAttachments:
            with open(f, "rb") as fil:
                part = MIMEApplication(fil.read(), Name=os.path.basename(f))
                # After the file is closed  NOT SURE IF I NEED TO EXPLICITY CLOSE THE FP IF THERE ARE MULTIPLE FILES (need to test) ...
                part['Content-Disposition'] = 'attachment; filename="%s"' % os.path.basename(f)
                msg.attach(part)
                fil.close()

    try:
        smtpObj = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        smtpObj.ehlo()
        smtpObj.login(cfg.get('email','user'), cfg.get('email','pwd'))   # Credentials found in local file emailcredentials.ini
        smtpObj.sendmail('gsoc.codingforgood.appdev@gmail.com', strToPerson,
                         msg.as_string())  # this could be a list of addresses to loop over with a pause in between

        smtpObj.close()
        del smtpObj
        print('Email Sent to: %s' % strToPerson)
        time.sleep(5)  # let the server rest a bit ...
    except smtplib.SMTPAuthenticationError as ex:
        print("SMTP Authentication Error: ", ex)

    return


if __name__ == '__main__':
    print("hello from module %s. Python version: %s" % (sys.argv[0], sys.version))
    sys.stdout.write("--------------------------------------------------------------\n")
    sys.stdout.write("Start of %s Process: %s\n\n" % (sys.argv[0], time.strftime("%H:%M:%S", time.localtime())))

    # lstContacts = LoadGoogleSheets('2019STEMExpoExternalOrganizations', 'SaveTheDate')
    # lstContacts = filter(IgnoreRejectors, lstContacts) # ALWAYS DO THIS FIRST TO FILTER OUT PURPOSEFUL REJECTORS
    # lstContacts = filter(lambda d: Company(d, ['Nonscriptum LLC','Robotics Society of Southern California','Code Ninjas','The Energy Coalition']), lstContacts)
    # # print ("len lstContacts: " + len(lstContacts))
    # for l in lstContacts:
    #     print(l)
    #
    # sys.exit(0)

#    strSubject = "GSOC Coding for Good App Dev - KICK OFF MEETING REMINDER"
#    strTemplateName = "KickoffReminder"
#    lstContacts = LoadEmailList('EventMailingList.csv')
#    lstContacts = list(filter(TestOnly, lstContacts))
#    lstAttachments = []

#    strSubject = "GSOC Coding for Good App Dev - PART 1 VIDEO"
#    strTemplateName = "Part1Video"
#    lstContacts = LoadEmailList('EventMailingList.csv')
#    lstContacts = list(filter(TestOnly, lstContacts))
#    lstAttachments = []

#    strSubject = "GSOC Coding for Good App Dev - PART 1 MEETING REMINDER"
#    strTemplateName = "Part1Reminder"
#    lstContacts = LoadEmailList('EventMailingList.csv')
#    lstContacts = list(filter(TestOnly, lstContacts))
#    lstAttachments = []

#    strSubject = "GSOC Coding for Good App Dev - PART 2 VIDEO"
#    strTemplateName = "Part2Video"
#    lstContacts = LoadEmailList('EventMailingList.csv')
#    lstContacts = list(filter(TestOnly, lstContacts))
#    lstAttachments = [] # ['/users/brad/my projects/girlscout-stem/CodingForGoodAppEvent/assets/GirlScoutsMakeTheWorldABetterPlace.png']

#    strSubject = "GSOC Coding for Good App Dev - LINK CORRECTION, PLEASE REVIEW"
#    strTemplateName = "Part2LinkCorrection"
#    lstContacts = LoadEmailList('EventMailingList.csv')
#    lstContacts = list(filter(TestOnly, lstContacts))
#    lstAttachments = [] # ['/users/brad/my projects/girlscout-stem/CodingForGoodAppEvent/assets/GirlScoutsMakeTheWorldABetterPlace.png']

#    strSubject = "GSOC Coding for Good App Dev - PART 2 MEETING REMINDER"
#    strTemplateName = "Part2Reminder"
#    lstContacts = LoadEmailList('EventMailingList.csv')
#    lstContacts = list(filter(TestOnly, lstContacts))
#    lstAttachments = []

#    strSubject = "GSOC Coding for Good App Dev - PART 3 VIDEO"
#    strTemplateName = "Part3Video"
#    lstContacts = LoadEmailList('EventMailingList.csv')
#    lstContacts = list(filter(TestOnly, lstContacts))
#    lstAttachments = []

#    strSubject = "GSOC Coding for Good App Dev - PART 3 MEETING REMINDER"
#    strTemplateName = "Part3Reminder"
#    lstContacts = LoadEmailList('EventMailingList.csv')
#    lstContacts = list(filter(TestOnly, lstContacts))
#    lstAttachments = []

#    strSubject = "GSOC Coding for Good App Dev - PART 4 VIDEO"
#    strTemplateName = "Part4Video"
#    lstContacts = LoadEmailList('EventMailingList.csv')
#    lstContacts = list(filter(TestOnly, lstContacts))
#    lstAttachments = []

#    strSubject = "GSOC Coding for Good App Dev - PART 4 MEETING REMINDER"
#    strTemplateName = "Part4Reminder"
#    lstContacts = LoadEmailList('EventMailingList.csv')
#    lstContacts = list(filter(TestOnly, lstContacts))
#    lstAttachments = []

#    strSubject = "GSOC Coding for Good App Dev - FOLLOW UP SURVEY"
#    strTemplateName = "FollowUpSurvey"
#    lstContacts = LoadEmailList('EventMailingList.csv')
#    lstContacts = list(filter(TestOnly, lstContacts))
#    lstAttachments = []

#    strSubject = "GSOC Coding for Good App Dev - YOUR FEEDBACK IS IMPORTANT!"
#    strTemplateName = "DidNotEngage"
#    lstContacts = LoadEmailList('EventMailingList_DidNotSubmitProject.csv')
#    lstContacts = list(filter(TestOnly, lstContacts))
#    lstAttachments = []

    strSubject = "GSOC Coding for Good App Dev - SUMMER APP CHALLENGE INFO"
    strTemplateName = "ThunkableSummerChallenge"
    lstContacts = LoadEmailList('EventMailingList_DidNotSubmitProject.csv')
    lstContacts = list(filter(TestOnly, lstContacts))
    lstAttachments = []

    nEmailsSent = 0
    for r in lstContacts:
        try:
            fp = open('templates/%s.html' % strTemplateName, 'rt')
        except IOError as detail:
            s = "BAD ERROR"
        else:
            s = fp.read()
            fp.close()

        template = jinja2.Template(s)  # this is associated with the file reading test
        html = template.render(r)

        try:
            fp = open('templates/%s.txt' % strTemplateName, 'rt')
        except IOError as detail:
            s = "BAD ERROR"
        else:
            s = fp.read()
            fp.close()

        template = jinja2.Template(s)  # this is associated with the file reading test
        text = template.render(r)

        SimpleEmailMessage(r['EMAIL'], strSubject, html, text, lstAttachments)
        nEmailsSent += 1

    sys.stdout.write("\n\nEnd of %s Process: %s. %5d emails sent\n" % (
    sys.argv[0], time.strftime("%H:%M:%S", time.localtime()), nEmailsSent))
    sys.stdout.write("-------------------------------------------------------------\n")
