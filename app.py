import requests
from flask import Flask, render_template, request, redirect
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
import smtplib as sm 
import time
import threading
from datetime import datetime
import os
import random
random_int = str(random.randint(0, 10000))
app=Flask(__name__)
@app.route('/')
def home():
    return render_template("index.html")

@app.route('/achyut',methods=['GET','POST'])
def func():
    if (request.method=='POST'):
        Email_column_name=request.form['Email_column_name']
        Sender_Id=request.form['sender_id']
        Sender_Pass=request.form['sender_password']
        Subject=request.form['subject']
        Body=request.form['body']
        Time=request.form['time']
        Column_name=request.form['Email_column_name']
        if Time!='x' and Time!='X':
            l=[]
            l=Time.split("-")
            # get the desired time and date from the user
            year = int(l[4])
            month = int(l[3])
            day =int(l[2])
            hour = int(l[0])
            minute = int(l[1])
        else:
            current_time = datetime.now()
            year=current_time.year
            month=current_time.month
            day=current_time.day
            hour=current_time.hour
            minute=current_time.minute
    
    print(datetime.now())
    def run_code():
        data=pd.read_excel(f'{random_int}.xlsx')
        email_col=data.get(Column_name)
        list_of_emails=list(email_col) #email list extracted from xlsx file                                             
        os.remove(f'{random_int}.xlsx')
        #creating object of smtp 
        try:
            server=sm.SMTP("smtp.gmail.com", 587)# gmail server, port no
            server.starttls()#secure connection stablished
            server.login(Sender_Id,Sender_Pass)#sender's credentials
            from_=Sender_Id
            to_=list_of_emails
            SUBJECT=Subject
            TEXT=Body 
            message = 'Subject: {}\n\n{}'.format(SUBJECT, TEXT)
            server.sendmail(from_,to_,message)
            print("Message has been sent to emails")



        except Exception as e:
            print(e) #writes if exception found


    def schedule_code(year, month, day, hour, minute):
        # create a datetime object for the desired time
        run_time = datetime(year, month, day, hour, minute)# required time

        # calculate the delay until the desired time
        delay = (run_time - datetime.now()).total_seconds()

        # if the desired time is in the past, add 1 day to the delay
        if (-70<delay<0):
            delay=0
        if (delay <=-70):
            delay += 86400

        # pause the current thread for the calculated delay
        time.sleep(delay)

        # run the code in a new thread
        threading.Thread(target=run_code).start()
    schedule_code(year, month, day, hour, minute)
       
    return render_template('result.html')
    




@app.route('/help')
def help():
    return render_template('help.html')

@app.route('/upload', methods=['GET','POST'])
def upload():
    # Get the file from the request object
    file = request.files['file']
    
        # Save the file to the server
    file.save(f'C:/Users/91911/Desktop/Email_Project_final/{random_int}.xlsx')

    return render_template("uploaded.html")
if __name__=="__main__":
    app.run(debug=False)