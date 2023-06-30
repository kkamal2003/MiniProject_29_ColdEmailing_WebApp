from django.shortcuts import render

#import required libs
import pandas as pd
import smtplib,time
import os
#import required model
from mailapp.models import DATA
# Create your views here.
def index(request):
    if (request.method=='POST'):
        senders_mail=request.POST['senders_mail']
        app_pass=request.POST['app_pass']
        senders_name=request.POST['senders_name']
        message=request.POST['message']
        subject=request.POST['subject']
        file=request.FILES['file']
        ins=DATA(location=file)
        ins.save()
        file_path = 'media/'+ ins.location.name

        logs=send_email_static(senders_mail,app_pass,senders_name,message,subject,file_path)
        #delete the file from local storage 
        os.remove(file_path)

        context={
            'log':logs
        }

        return render(request,'success.html',context)

        


    return render(request,'home.html')


def send_email_static(senders_mail, app_pass, senders_name, message, subject, file_path):
    logs = []
    
    # Establish connection
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(senders_mail, app_pass)

    # Read the file entered by the user
    email_list = pd.read_excel(file_path, engine='openpyxl')

    # Print the loaded data for debugging
    print(email_list.columns)
    email_list = email_list.rename(columns={'NAME ': 'NAME'})
    # Get all the Names, Email Addresses, Subjects, and Messages
    all_names = email_list['NAME']
    all_emails = email_list['EMAIL']
    
    # Loop through the emails
    for idx in range(len(all_emails)):
        # Get each record's name, email, subject, and message
        name = all_names[idx]
        email = all_emails[idx]

        msg = f"Dear {name},\n"

        # Create the email to send
        full_email = (
            "From: {0} <{1}>\n"
            "To: {2} <{3}>\n"
            "Subject: {4}\n\n"
            "{5}{6}"
            .format(senders_name, senders_mail, name, email, subject, msg, message)
        )

        try:
            server.sendmail(senders_mail, [email], full_email)
            logs.append(f'Email to {name} ({email}) successfully sent!')
            time.sleep(1)
        except Exception as e:
            logs.append(f'Email to {name} ({email}) could not be sent due to error: {str(e)}')

    server.close()
    return logs


