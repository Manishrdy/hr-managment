import mysql.connector

from flask import Flask,render_template,request
import xlrd
import smtplib
import os
import sys
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw

mydb = mysql.connector.connect(
  host="127.0.0.1",
  user="root",
  passwd="",
  database="manish"
)
mycursor = mydb.cursor()

# mycursor.execute("CREATE TABLE register (id INT AUTO_INCREMENT PRIMARY KEY, name VARCHAR(255), email VARCHAR(255), password VARCHAR(255))")



app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def register():
    return render_template('register.html')
    

@app.route('/login',methods=['GET', 'POST'])
def login():
    if request.method == "POST":
        username = request.form['uname']
        email = request.form['em']
        password = request.form['psd']
        sql = "INSERT INTO register (username, email, password) VALUES (%s, %s, %s)"
        val = (username, email, password)
        mycursor.execute(sql, val)
        mydb.commit()
    return render_template('login.html')

@app.route('/login_admin', methods = ['POST'])
def login_admin():

    usr = request.form['name']
    password = request.form['password']
    mycursor.execute("SELECT email, password FROM register")
    myresult = mycursor.fetchall()
    for x,y in myresult:
        if(x==usr and y==password):
            return render_template('welcome.html',name=usr)
        else:
            return ('Pls check your credientials')


@app.route('/welcome')
def welcome():
    if request.method == "POST":
        name = request.form['ename']
        dob = request.form['date']
        email = request.form['em']
        mobileno = request.form['mn']
        salary = request.form['sp']
        address = request.form['adr']
        sql = "INSERT INTO add_emp (name, dob, email, mobileno, salary, address) VALUES (%s, %s, %s, %s, %s, %s)"
        val = (name, dob, email, mobileno, salary, address)
        mycursor.execute(sql, val)
        mydb.commit()
    return render_template('welcome.html')

@app.route('/employee')
def employee():
    return render_template('employee.html')

@app.route('/welcome1_store',methods = ['GET','POST'])
def welcome1_store():
    if request.method == "POST":
        name = request.form['ename']
        dob = request.form['date']
        email = request.form['em']
        mobileno = request.form['mn']
        salary = request.form['sp']
        address = request.form['adr']
        sql = "INSERT INTO add_emp (name, dob, email, mobileno, salary, address) VALUES (%s, %s, %s, %s, %s, %s)"
        val = (name, dob, email, mobileno, salary, address)
        mycursor.execute(sql, val)
        mydb.commit()
    return render_template('welcome.html')

@app.route('/checkattandance')
def attandance():
    return render_template('attandance.html')

@app.route('/check_attendance',methods = ['GET','POST'])
def attandance_check():
    if(request.method == 'POST'):
        check_user_name = request.form['username']
        sql_2 = """SELECT * from add_emp where name='%s'""" % (check_user_name)
        mycursor.execute(sql_2)
        myresult = mycursor.fetchone()    
        user_name=myresult[0]
        dateob=myresult[1]
        emailid=myresult[2]
        mobileno=myresult[3]
        salary=myresult[4]
        address=myresult[5]
        sql_3 = """SELECT * from attandance where name='%s'""" % (check_user_name)
        mycursor.execute(sql_3)
        myresult1 = mycursor.fetchone()  
        user_name = myresult1[0]
        days_p = myresult1[1]
        days_t = myresult1[2]
        total_salary = myresult1[4]
        return render_template('print.html',name=user_name,dob=dateob,email=emailid,mob=mobileno,sal=salary,address=address,name1=user_name,days_p=days_p,days_t=days_t,total_salary=total_salary)


    return render_template('attandance.html')

@app.route('/sendsalary')
def sendsalary():
    return render_template('sendsalary.html')

@app.route('/sendsalary_mail')
def sendsalary_mail():
    def shorten( text, _max ):
        t = text.split(" ")
        text = ''
        if len(t)>1:
            for i in t[:-1]:
                text += i[0] + '.'
        text += ' ' + t[-1]
        if len(text) < _max :
            return text
        else :
            return -1

    def make_certi( ID, name, basesalary, present, total, salary):
        img = Image.open("C:\\Users\\d6mr0\\Desktop\\Payroll Managment\\Pay_Bill.jpg")
        draw = ImageDraw.Draw(img)
        # Load font
        font = ImageFont.truetype("C:\\Users\\d6mr0\\Desktop\\Payroll Managment\\Pristina.TTF", 25)

        # Check sizes and if it is possible to abbreviate
        # if not the IDs are added to an error list
        if (len( name ) > 20):
            name = shorten(name, 20)
    #    if (len( department ) > 30):
    #        department = shorten(department, 30)

        if name == -1 :
            return -1
        else:
            # Insert text into image template
            draw.text((300, 175), name, (255,255,255), font=font )
            draw.text((300, 215), str(basesalary), (255,255,255), font=font )
            draw.text((300, 255), str(present), (255,255,255), font=font )
            draw.text((300, 295), str(total), (255,255,255), font=font )
            draw.text((300, 335), str(salary), (255,255,255), font=font )

            if not os.path.exists('Pay Bills') :
                os.makedirs( 'Pay Bills' )

            # Save as a JPG
            img.save('Pay Bills\\'+str(ID)+'.jpg', resolution=300.0)
            return 'Pay Bills\\'+str(ID)+'.jpg'

    def email_certi( filename, receiver ):
        username = "d6mr07"
        password = "*********"
        sender = username + '@gmail.com'

        msg = MIMEMultipart()
        msg['Subject'] = 'Pay Bill'
        msg['From'] = username+'@gmail.com'
        msg['Reply-to'] = username + '@gmail.com'
        msg['To'] = receiver

        # That is what u see if dont have an email reader:
        msg.preamble = 'Multipart massage.\n'
        
        # Body
        part = MIMEText( "Dear Employee,\nPlease find attached Pay Bill.\n\nRegards\nHR." )
        msg.attach( part )

        # Attachment
        part = MIMEApplication(open(filename,"rb").read())
        part.add_header('Content-Disposition', 'attachment', filename = os.path.basename(filename))
        msg.attach( part )

        # Login
        server = smtplib.SMTP( 'smtp.gmail.com:587' )
        server.starttls()
        server.login( username, password )

        # Send the email
        server.sendmail( msg['From'], msg['To'], msg.as_string() )

    if __name__ == "__main__":
        error_list = []
        error_count = 0

        os.chdir(os.path.dirname(os.path.abspath((sys.argv[0]))))

        # Read data from an excel sheet from row 2
        Book = xlrd.open_workbook('C:\\Users\\d6mr0\\Desktop\\Payroll Managment\\Send_Salary.xlsx')
        WorkSheet = Book.sheet_by_name('Send_Salary')
        
        num_row = WorkSheet.nrows - 1
        row = 0
        while row < num_row:
            row += 1
            
            ID = WorkSheet.cell_value( row, 0 )
            name = WorkSheet.cell_value( row, 1 )
            receiver = WorkSheet.cell_value( row, 2 )
            basesalary = WorkSheet.cell_value( row, 3 )
            present = WorkSheet.cell_value( row, 4 )
            total = WorkSheet.cell_value( row, 5 )
            salary = WorkSheet.cell_value( row, 6 )

            filename = make_certi( ID, name, basesalary, present, total, salary)

            if filename != -1:
                email_certi( filename, receiver )
                print("Sent to " + ID )
        # Add to error list
            else:
                error_list.append( ID )
                error_count += 1

        print(str(error_count) + " Errors- List:" + ','.join(error_list))
        
        # Make certificate and check if it was successful
    return render_template('sendsalary.html')

@app.route('/manage',methods = ['GET','POST'])
def manage():
    return render_template('choose.html')

@app.route('/manage10')
def manage10():
    return render_template('update_pick.html')

@app.route('/delete')
def delete():
    return render_template('delete.html')

@app.route('/manage_details',methods = ['GET','POST'])
def manage_details():
    if(request.method == 'POST'):
        check_user_name = request.form['username']
        sql_2 = """SELECT * from add_emp where name='%s'""" % (check_user_name)
        mycursor.execute(sql_2)
        myresult = mycursor.fetchone()    
        user_name=myresult[0]
        dateob=myresult[1]
        emailid=myresult[2]
        mobileno=myresult[3]
        salary=myresult[4]
        address=myresult[5]
        return render_template('manage1.html',name=user_name,dob=dateob,email=emailid,mob=mobileno,sal=salary,address=address)

@app.route('/delete_employee',methods = ['GET','POST'])
def delete_employee():
    if(request.method == 'POST'):
        check_user_name = request.form['username']
        sql_2 = """DELETE from add_emp where name='%s'""" % (check_user_name)
        mycursor.execute(sql_2)
        mydb.commit()
        return render_template('manage.html')

@app.route('/update_mobileno',methods = ['GET','POST'])
def update_mobileno():
    return render_template('update_mobileno.html')

@app.route('/update_mobileno1',methods = ['GET','POST'])
def update_mobileno1():
        if(request.method == 'POST'):
            check_user_name = request.form['username']
            check_user_mobileno = request.form['mobileno']
            sql_2 = "UPDATE add_emp set mobileno= %s WHERE name= %s"
            val = (check_user_mobileno, check_user_name)
            mycursor.execute(sql_2,val)
            mydb.commit()
            return render_template('welcome.html')

@app.route('/update_salary',methods = ['GET','POST'])
def update_salary():
    return render_template('update_salary.html')

@app.route('/update_salary1',methods = ['GET','POST'])
def update_salary1():
        if(request.method == 'POST'):
            check_user_name = request.form['username']
            check_user_salary = request.form['salary']
            sql_2 = "UPDATE add_emp set salary= %s WHERE name= %s"
            val = (check_user_salary, check_user_name)
            mycursor.execute(sql_2,val)
            mydb.commit()
            return render_template('welcome.html')

@app.route('/update_email',methods = ['GET','POST'])
def update_email():
    return render_template('update_email.html')

@app.route('/update_email1',methods = ['GET','POST'])
def update_email1():
        if(request.method == 'POST'):
            check_user_name = request.form['username']
            check_user_email = request.form['email']
            sql_2 = "UPDATE add_emp set email= %s WHERE name= %s"
            val = (check_user_email, check_user_name)
            mycursor.execute(sql_2,val)
            mydb.commit()
            return render_template('welcome.html')

@app.route('/update_address',methods = ['GET','POST'])
def update_address():
    return render_template('update_address.html')

@app.route('/update_address1',methods = ['GET','POST'])
def update_address1():
        if(request.method == 'POST'):
            check_user_name = request.form['username']
            check_user_address = request.form['address']
            sql_2 = "UPDATE add_emp set address= %s WHERE name= %s"
            val = (check_user_address, check_user_name)
            mycursor.execute(sql_2,val)
            mydb.commit()
            return render_template('welcome.html')

@app.route('/crud',methods = ['GET','POST'])
def crud():
    return render_template('crud.html')

if(__name__ == '__main__'):
    app.run(debug = True, port = 8000)