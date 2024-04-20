from flask import Flask, render_template, url_for, redirect, request, flash
from flask_mail import Mail, Message
import csv
#for mail sending
import os
import pandas as pd

app = Flask(__name__)

#for mail
app.config['SECRET_KEY']='asdfghjkl123456'
app.config['MAIL_SERVER']='smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USERNAME'] = 'zinthandarhtun2018@gmail.com'
app.config['MAIL_PASSWORD'] = 'xyea vvad rubi gjzh'
app.config['MAIL_USE_TLS'] = True
mail= Mail(app)
print(__name__)

#data read from excel for skills menu 
df_skills = pd.read_excel('data.xlsx',sheet_name='Skills')
#data read from excel for experience menu
df_experience = pd.read_excel('data.xlsx',sheet_name='Experience')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/<string:page_name>')
def page(page_name='/'):
    try:
        if page_name == 'skills.html':
            return skills()  
        elif  page_name == 'experience.html':
            return experience() 
        else:
            return render_template(page_name)
    except:
        return redirect('/')

def skills():
    df_skills.fillna('',inplace=True)
    d = {}
    for i in range(len(df_skills)):
        skill = df_skills.iloc[i]['Skill']
        image = df_skills.iloc[i]['Image']
        experience = df_skills.iloc[i]['Experience']
        d[skill]={}
        d[skill]['image']=image
        d[skill]['experience']=experience
    return render_template('skills.html',skillinfo=d)

def experience():
    df_experience.fillna('',inplace=True)
    d = {}
    for i in range(len(df_experience)):
        designation = df_experience.iloc[i]['Designation']
        company = df_experience.iloc[i]['Company']
        image = df_experience.iloc[i]['Image']
        date = df_experience.iloc[i]['Date']
        info = df_experience.iloc[i]['Info']
        d[company]={}
        d[company]['designation']=designation
        d[company]['image']=image
        d[company]['date']=date
        d[company]['info']=info
    return render_template('experience.html',experienceinfo=d)

@app.route('/submit_form', methods = ['GET','POST'])
def submit():
    if request.method == "POST":
        try:
            data = request.form.to_dict()
            write_data_csv(data)
            send_email(data)
            message = "Form submitted. We will get in touch with you shortly!"
            flash("Successfully sent", "success")
            return render_template('contact.html')
        except Exception as e:
            error_message = "Failed to save data or send email: {}".format(str(e))
            return render_template('thankyou.html', message=error_message)
    else:
        message = "Form not submitted"
        return render_template('thankyou.html', message=message)

def write_data_csv(data):
    email = data.get('email')
    subject = data.get('subject')
    message = data.get('message')
    if email and subject and message:
        folder_path = 'C://PortfolioContact'
        file_path = os.path.join(folder_path, 'db.csv')
        with open(file_path, 'a', newline='') as csvfile:
            db_writer = csv.writer(csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)
            db_writer.writerow([email, subject, message])
    else:
        raise ValueError("Incomplete form data")

def send_email(data):
    email = data.get('email')
    subject = data.get('subject')
    message = data.get('message')
    if email and subject and message:
        mail_message = Message(
            subject,
            sender=email,
            recipients=['zinthandarhtun2018@gmail.com']
        )
        mail_message.body = message
        
        mail.send(mail_message)
        #return mail to sender
        confirmation_message = Message(
            "Thank you for your interest",
            sender='zinthandarhtun2018l@gmail.com',  
            recipients=[email]
        )
        confirmation_message.body = "Thank you for your interest. We'll send you the requested data later."
        
        mail.send(confirmation_message)
    else:
        raise ValueError("Incomplete email data")

if __name__ == '__main__':
   app.run(debug = True)



