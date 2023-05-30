from flask import Flask
from flask import Blueprint,render_template,request
from views import views,request
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import openpyxl
import os
import pandas as pd




app=Flask(__name__)
app.register_blueprint(views,url_prefix="/")

@app.route("/done",methods=['get','post'])
def done():

    userName=str(request.form['userName'])
    email=str(request.form['userEmail'])
    userSub=str(request.form['userSubject'])
    userMess=str(request.form['userMessage'])
    userAppPass=str(request.form['appPassword'])
    file = request.files['uploadFile']
    
    
    df = pd.read_excel(file)
    print(df)
    emailsList = df.copy()
    new_file_path = './emailsList.xlsx'
    emailsList.to_excel(new_file_path, index=False)
    

    # Define email credentials and SMTP server
    SMTP_SERVER = 'smtp.gmail.com'
    SMTP_PORT = 587
    SMTP_USERNAME = email
    SMTP_PASSWORD = userAppPass
    # Define message parameters
    FROM_EMAIL = email
    SUBJECT = userSub
    MESSAGE = 'Hello {name},\n\n'+userMess+'\n\nBest regards,\n'+userName

    # Load email addresses from spreadsheet
    wb = openpyxl.load_workbook('emailsList.xlsx')
    ws = wb.active
    RECIPIENTS = [{'name': row[0].value, 'email': row[1].value} for row in ws.iter_rows(min_row=2)]

    # Loop through recipients and send customised email to each
    for recipient in RECIPIENTS:
        # Set up email message
        message = MIMEMultipart()
        message['From'] = FROM_EMAIL
        message['To'] = recipient['email']
        message['Subject'] = SUBJECT
        message.attach(MIMEText(MESSAGE.format(name=recipient['name']), 'plain'))

        # Connect to SMTP server and send email
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.starttls()
            smtp.login(SMTP_USERNAME, SMTP_PASSWORD)
            smtp.sendmail(FROM_EMAIL, recipient['email'], message.as_string())
        print(f"Email sent to {recipient['email']}")
    # Close the workbook when finished
    wb.close()
    
    
    file_path = "./emailsList.xlsx"
    os.remove(file_path)
    return render_template("done.html")
    
    


if __name__ == "__main__":
    app.run(debug=True, port=8000)