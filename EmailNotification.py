import os
import win32com.client as win32

def send_email(subject, body, to_email):
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)  # 0 represents a Mail item

    mail.Subject = subject
    mail.Body = body
    mail.To = to_email

    try:
        mail.Send()
        print("Email sent successfully!")
    except Exception as e:
        print("An error occurred while sending the email:", str(e))

# Set up the email parameters
subject = "AVNET WCO list of decrypted files"
# body = "This is a notification email sent using Python and Outlook."
to_email = "baijkuma@in.ibm.com"
# to_email_list = ["recipient1@example.com", "recipient2@example.com", "recipient3@example.com"]



box_path_AP_Decr = 'C:/Users/01934L744/Box/AVNET MVP Project/WinSCP SFTP files/AP Data/AP_Decrypted/'

bodyText = f"BOX folder Path :- {box_path_AP_Decr} \n \n List of Uploaded files in AP BOX folder:-\n"
for file in os.listdir(box_path_AP_Decr):
    filePath = box_path_AP_Decr+file
    fileSize = str(round(os.path.getsize(filePath)/(1024*1024),2))+" MB"
    bodyText = bodyText+ '\n' + f'File Name :- {file}, File Size:-{fileSize} ' + '\n'
    print(bodyText)

body = bodyText
# Send the email
send_email(subject, body, to_email)