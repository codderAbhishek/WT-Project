from logging import NullHandler
from multiprocessing import connection
import os
from subprocess import call
from threading import current_thread
from turtle import onclick, onscreenclick
from flask import *
import sqlite3
from datetime import MAXYEAR, date
import xlwt
from xlwt import Workbook
import pandas as pd
import stat
app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/add")
def add_student():
    return render_template("add_student.html")

@app.route("/saverecord",methods = ["POST","GET"])
def saveRecord():
    
    msg = "msg"
    if request.method == "POST":
        try:
            name = request.form["name"]
            quantity = request.form["email"]
            payment = request.form["gender"]
            contact = request.form["contact"]
            dob =date.today()
            address = request.form["address"]
            

            
            with sqlite3.connect("student_detials.db") as connection:

                cursor = connection.cursor()
                #cursor.execute("DBCC CHECKIDENT ('Student_Info', RESEED, 1)")
                cursor.execute("INSERT into Student_Info (name, quantity, payment, contact, dob, address) values (?,?,?,?,?,?)",(name, quantity, payment, contact, dob, address))
                connection.commit()
                msg = "Student detials successfully Added"
        except:
            connection = sqlite3.connect("/student_detials.db")
            connection.rollback()
            msg = "We can not add Student detials to the database"
        finally:
            return render_template("success_record.html",msg = msg)
            


@app.route("/delete")
def delete_student():
    return render_template("delete_student.html")


@app.route("/view")
def student_info():
    connection = sqlite3.connect("student_detials.db")
    connection.row_factory = sqlite3.Row
    cursor = connection.cursor()
    cursor.execute("select * from Student_Info")
    rows = cursor.fetchall()
    df = pd.read_sql_query("SELECT * from Student_Info", connection)
    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
    df.to_excel(os.path.join(ROOT_DIR,'customer.xls'))
    






    return render_template("student_info.html",rows = rows)
    


@app.route("/deleterecord",methods = ["POST"])
def deleterecord():
    id = request.form["id"]
    with sqlite3.connect("student_detials.db") as connection:

        cursor = connection.cursor()
        cursor.execute("select * from Student_Info where id=?", (id,))
        rows = cursor.fetchall()
        if not rows == []:

            cursor.execute("delete from Student_Info where id = ?",(id,))
            cursor.execute("UPDATE SQLITE_SEQUENCE SET SEQ=0 WHERE NAME='Student_Info'")
            msg = "Student detial successfully deleted"
            return render_template("delete_record.html", msg=msg)

        else:
            msg = "can't be deleted"
            return render_template("delete_record.html", msg=msg)

@app.route("/edit")
def edit():
    
   
    return render_template('edit.html')


@app.route("/edit",methods = ["POST"])
def editecord():
    x='test'
    try:
        sqliteConnection = sqlite3.connect('student_detials.db')
        quantity=request.form['id2']
        a=request.form['id']
        li=[(quantity,a)]
        
        cursor = sqliteConnection.cursor()
        print("Connected to SQLite")

        sqlite_update_query = """Update Student_Info set quantity = ? where id = ?"""
        cursor.executemany(sqlite_update_query, li)
        sqliteConnection.commit()
        x="Records updated successfully"
        sqliteConnection.commit()
        cursor.close()


    except sqlite3.Error as error:
        X="Failed to update multiple records of sqlite table"
    finally:
        if sqliteConnection:
            sqliteConnection.close()
            print("The SQLite connection is closed")

    return """<script type="text/javascript">
alert("{{ x }}");
</script><a href="/">Go back to home page</a>"""

if __name__ == "__main__":
    app.run(debug = True)  
