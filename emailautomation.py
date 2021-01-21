import smtplib
import random

# constants
# both were dictated by the exercise (exam.csv & column descriptions ) 1.1 Understanding the Sample Exam Data & 2.1 Saving Data from the File
FILENAME= r'.\liveProjects\emailAutomation\exam.csv'    # path included
HEADERS = ('Email','Last_Name','First_Name','Problem_1_score','Problem_1_comments','Problem_2_score','Problem_2_comments','Problem_3_score','Problem_3_comments')
LOOKUPKEY = 'Email' # requested lookup key
LOOKUPVAL = ('Problem_1_score','Problem_1_comments','Problem_2_score','Problem_2_comments','Problem_3_score','Problem_3_comments')  # prescribed data excl Name
NAMEPARTS = ('Last_Name','First_Name')  # data needed to compose value for 'Name' key [values provided in parts]

# functions
def read_in(data):     # handling 1.1 & 2.1
    """
    opens a .csv file with data only (no headers!) and 
    merges it with a given set of headers into a dictionary.
    input : a filename of the file containing the data (here each row a student)
    AVAILABLE: a collection of appropriate headers. (Here a tuple with constants) should be available!
    output: a list of students with a dictionary (merging header as key and data as value) for each student
    """
    students = []
    for information in open(data,'rt'):
        student = {}
        ch = 0
        for column in information.strip().split(','):
            student[HEADERS[ch]] = column
            ch += 1
        students.append(student)
    # all data retrieved and made accessable via list of students organized as a dictionary per student  
    # now to create the desired lookup dictionary from it with just the data requested
    lud = {}
    for student in students:
        value_dict = {}
        # compose 'Name' out of 'First_Name' and 'Last_Name' In hindsight it is questionable if this was ment
        value_dict['Name'] = f'{student[NAMEPARTS[1]]} {student[NAMEPARTS[0]]}'
        for k,v in student.items():
            if k in LOOKUPVAL:
                value_dict[k] = v
        lud[student[LOOKUPKEY]] = value_dict
    return lud  # a lookup dictionary met key Email (given) and values requested

def create_emails(lookup_dict):
    """
    provides a list of dict with data to compose emails {sender: ,receiver: ,msg: }one for each student.
    it does compose the message in full
    input : lookup_dictionary as defined
    output: list of emails as a dict
    """
    emails=[]
    if lookup_dict:
        selected = random.choice(list(lookup_dict.keys()))
        sender_address = 'teacher@home.edu'
    for emailaddress,student in lookup_dict.items():
        email_data = {}
        email_data['sender'] = sender_address
        email_data['receiver'] = emailaddress
        message = ''
        fn =student['Name']
        sp =fn.find(' ')
        fn = fn[:sp]    # first name is used only
        message += f'From: {sender_address}\n'    # to get them visible in the actual message, 
        message += f'To: {emailaddress}\n'        # metadata is not part of the message 
        message += f'Subject: Results for the test on the book assignment\n\n'  # nor is this part of the metadata
        message += f'Dear {fn},\n\nYour score for the book assignment is broken down below by question number.\n\n'
        message += f"  1. {student['Problem_1_score']}%: {student['Problem_1_comments']}\n\n"
        message += f"  2. {student['Problem_2_score']}%: {student['Problem_2_comments']}\n\n"
        message += f"  3. {student['Problem_3_score']}%: {student['Problem_3_comments']}"
        if emailaddress == selected:
            message += f"\n\nYou’ve been randomly chosen to present a summary of the book in the next class. Looking forward to it!"
        email_data['msg'] = message     # it (message) is unicode by default
        emails.append(email_data)
    return emails

def send_emails(list_of_emails):
    """
    Deliberately separated from creating the emails for reusability purpose
    establishes a connection to a SMTP server to send e-mails
    input : list of e-mails, each e-mail as a dict with elements [sender, receiver, msg]
    # this function did not work, presumably due to company policies
    # according to some examples, i expected it to work but could not get it to work
    # I wrote a print_emails (to_terminal) instead, to check the emails created.
    
    >>>>>    found [https://mailtrap.io/]: It works fine now <<<<<
    
    open issues: (not part of this exercise)
    -   MISSING_MID         Missing Message-Id: header
    -   MISSING_DATE        Missing Date: header
    -   DKIM_ADSP_NXDOMAIN  No valid author signature and domain not in DNS
    """
    username = 'a6c4238cc6baed'   # "your smtp username here "
    password = '92f3441e250ed4'   # "your smtp password here"
    smtp_server = 'smtp.mailtrap.io:25'    # address and port for mailtrap.io SMTP-service
    with smtplib.SMTP(smtp_server) as server:
        server.starttls()
        server.login(username, password)
        for email in list_of_emails:    
            server.sendmail(email['sender'].encode('utf-8'),email['receiver'].encode('utf-8'),email['msg'].encode('utf-8')) 
            # (email_from, email_to, email_body)
            # The sendmail() expects bytes. since python 3 all strings are unicode by default. 
            # You have to tell him to encode your unicode-strings with .encode('utf-8')
    
    ###### Journey
    # on localhost no luck
    # ConnectionRefusedError: [WinError 10061] No connection could be made because the target machine actively refused it
    # I suppose it is not allowed and the firewall is actively dropping the packages. 
    # Policies.... No chance fixing that! but: Did not work #home either!
    # actually, this is not true. there is nothing listening on localhost!
    # using 'smtp.gmail.com:587' does establish a connection to the smtp server hosted by gmail that listens at port 587.
    # so actually I miss the availability of an smtp server and port to send my mail without logging in
    # found [https://mailtrap.io/]: It works fine now, except for the '\u2019' you put in the text. (the quote in 'You’ve')
    # That results in: UnicodeEncodeError: 'ascii' codec can't encode character '\u2019' in position 241: ordinal not in range(128)
    # Likewise the usage of all kind of normal characters would result in similar failure f.e. € > '\u20ac'

def print_emails(list_of_emails):   # alternative to sending
    """
    Alternative for sending since that does not seem to work
    """
    for i,email in enumerate(list_of_emails,1):
        print(f'\n\temail {i}:\n')
        print(f'+++ METADATA, e-mail header data +++')
        print(f'van: {email["sender"]}')
        print(f'aan: {email["receiver"]}')
        print(f'+++ METADATA, e-mail header data +++')
        print(f'--- start e-mail body data '+'--'*20)
        print(email["msg"])
        print(f'=== end e-mail body data =='+'=='*20)

#main
send_emails(create_emails(read_in(FILENAME)))  
# print_emails(create_emails(read_in(FILENAME))) >> no longer needed