import sqlite3
import tkinter as tk
from tkinter import ttk
import smtplib
from email.mime.text import MIMEText
import ctypes

root = tk.Tk()
root.title("STUDENTS")


def send_emails():
    try:

        conn = sqlite3.connect('studentsdatabase.db')
        cursor = conn.cursor()

        cursor.execute("SELECT email from students")
        res = cursor.fetchall()

        for email in res:
            recipient = email[0]
            subject = "Grade"

            cursor.execute("SELECT student_id FROM students WHERE email = (?)", email)
            id = cursor.fetchone()

            cursor.execute("SELECT "
                       "l_1, l_2, l_3, h_1, h_2, h_3, h_4, h_5, "
                       "h_6, h_7, h_8, h_9, h_10, project, grade "
                       "FROM grades WHERE student_id = (?)",id)

            g = cursor.fetchall()
            g = g[0]

            l1 = g[0]
            l2 = g[1]
            l3 = g[2]
            h1 = g[3]
            h2 = g[4]
            h3 = g[5]
            h4 = g[6]
            h5 = g[7]
            h6 = g[8]
            h7 = g[9]
            h8 = g[10]
            h9 = g[11]
            h10 = g[12]
            proj = g[13]
            gr = g[14]

            body = f"Your grades\n list 1: {l1}\n list 2: {l2}\n list 3: {l3}\n" \
                   f"homework 1: {h1}\n" \
                   f"homework 2: {h2}\n" \
                   f"homework 3: {h3}\n" \
                   f"homework 4: {h4}\n" \
                   f"homework 5: {h5}\n" \
                   f"homework 6: {h6}\n" \
                   f"homework 7: {h7}\n" \
                   f"homework 8: {h8}\n" \
                   f"homework 9: {h9}\n" \
                   f"homework 10: {h10}\n" \
                   f"project: {proj}\n" \
                   f"full grade: {gr}\n"
            sender=""
            password=""

            msg = MIMEText(body)
            msg['Subject'] = subject
            msg['From'] = sender
            msg['To'] = recipient
            smtp_server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            smtp_server.login(sender, password)
            smtp_server.sendmail(sender, recipient, msg.as_string())
            smtp_server.quit()

            cursor.execute("UPDATE students SET status='mailed' WHERE student_id = (?)", id)
            conn.commit()
    except sqlite3.Error as e:
        print(f"Error: {e}")
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()
        ctypes.windll.user32.MessageBoxW(0, "All emails Sent", "EMAILS", 0)


def open_details_window(event):
    selected_item = treeview.focus()

    if selected_item:
        item_data = treeview.item(selected_item)
        item_values = item_data["values"]

        details_window = tk.Toplevel(root)
        details_window.title("DETAILS")

        id_label = ttk.Label(details_window, text="ID:")
        id_label.pack()
        id_entry = ttk.Entry(details_window)
        id_entry.insert(0, item_values[0])
        id_entry.config(state="disabled")
        id_entry.pack()

        l1_label = ttk.Label(details_window, text="List 1:")
        l1_label.pack()
        l1_entry = ttk.Entry(details_window)
        l1_entry.insert(0, item_values[5])
        l1_entry.pack()

        l2_label = ttk.Label(details_window, text="List 2:")
        l2_label.pack()
        l2_entry = ttk.Entry(details_window)
        l2_entry.insert(0, item_values[6])
        l2_entry.pack()

        l3_label = ttk.Label(details_window, text="List 3:")
        l3_label.pack()
        l3_entry = ttk.Entry(details_window)
        l3_entry.insert(0, item_values[7])
        l3_entry.pack()

        h1_label = ttk.Label(details_window, text="Homework 1:")
        h1_label.pack()
        h1_entry = ttk.Entry(details_window)
        h1_entry.insert(0, item_values[8])
        h1_entry.pack()

        h2_label = ttk.Label(details_window, text="Homework 2:")
        h2_label.pack()
        h2_entry = ttk.Entry(details_window)
        h2_entry.insert(0, item_values[9])
        h2_entry.pack()

        h3_label = ttk.Label(details_window, text="Homework 3:")
        h3_label.pack()
        h3_entry = ttk.Entry(details_window)
        h3_entry.insert(0, item_values[10])
        h3_entry.pack()

        h4_label = ttk.Label(details_window, text="Homework 4:")
        h4_label.pack()
        h4_entry = ttk.Entry(details_window)
        h4_entry.insert(0, item_values[11])
        h4_entry.pack()

        h5_label = ttk.Label(details_window, text="Homework 5:")
        h5_label.pack()
        h5_entry = ttk.Entry(details_window)
        h5_entry.insert(0, item_values[12])
        h5_entry.pack()

        h6_label = ttk.Label(details_window, text="Homework 6:")
        h6_label.pack()
        h6_entry = ttk.Entry(details_window)
        h6_entry.insert(0, item_values[13])
        h6_entry.pack()

        h7_label = ttk.Label(details_window, text="Homework 7:")
        h7_label.pack()
        h7_entry = ttk.Entry(details_window)
        h7_entry.insert(0, item_values[14])
        h7_entry.pack()

        h8_label = ttk.Label(details_window, text="Homework 8:")
        h8_label.pack()
        h8_entry = ttk.Entry(details_window)
        h8_entry.insert(0, item_values[15])
        h8_entry.pack()

        h9_label = ttk.Label(details_window, text="Homework 9:")
        h9_label.pack()
        h9_entry = ttk.Entry(details_window)
        h9_entry.insert(0, item_values[16])
        h9_entry.pack()

        h10_label = ttk.Label(details_window, text="Homework 10:")
        h10_label.pack()
        h10_entry = ttk.Entry(details_window)
        h10_entry.insert(0, item_values[17])
        h10_entry.pack()

        p_label = ttk.Label(details_window, text="Project:")
        p_label.pack()
        p_entry = ttk.Entry(details_window)
        p_entry.insert(0, item_values[18])
        p_entry.pack()

        def update():
            l1 = l1_entry.get()
            l2 = l2_entry.get()
            l3 = l3_entry.get()
            h1 = h1_entry.get()
            h2 = h2_entry.get()
            h3 = h3_entry.get()
            h4 = h4_entry.get()
            h5 = h5_entry.get()
            h6 = h6_entry.get()
            h7 = h7_entry.get()
            h8 = h8_entry.get()
            h9 = h9_entry.get()
            h10 = h10_entry.get()
            p = p_entry.get()
            id = id_entry.get()
            conn = sqlite3.connect('studentsdatabase.db')
            cursor = conn.cursor()

            cursor.execute("UPDATE grades SET l_1=?,l_2=?,l_3=?,"
                           "h_1=?,h_2=?,h_3=?,h_4=?,h_5=?,"
                           "h_6=?,h_7=?,h_8=?,h_9=?,h_10=?,"
                           "project=? WHERE student_id=?",(l1,l2,l3,h1,h2,h3,h4,h5,h6,h7,h8,h9,h10,p,id))
            conn.commit()
            try:
                if h1 == "None":
                    h1 = 0
                if h2 == "None":
                    h2 = 0
                if h3 == "None":
                    h3 = 0
                if h4 == "None":
                    h4 = 0
                if h5 == "None":
                    h5 = 0
                if h6 == "None":
                    h6 = 0
                if h7 == "None":
                    h7 = 0
                if h8 == "None":
                    h8 = 0
                if h9 == "None":
                    h9 = 0
                if h10 == "None":
                    h10 = 0

                var = (int(h1)+int(h2)+int(h3)+int(h4)+int(h5)+int(h6)+int(h7)+int(h8)+int(h9)+int(h10))/10

                res = 0
                numoflo = 0
                if 70 > var >= 60.0:
                    if int(l1) < 20 and numoflo < 1:
                        l1 = 20
                        numoflo += 1
                    elif int(l2) < 20 and numoflo < 1:
                        l2 = 20
                        numoflo += 1
                    elif int(l3) < 20 and numoflo < 1:
                        l2 = 20
                        numoflo += 1
                elif 80 > var >= 70.0:
                    if int(l1) < 20 and numoflo < 2:
                        l1 = 20
                        numoflo += 1
                    elif int(l2) < 20 and numoflo < 2:
                        l2 = 20
                        numoflo += 1
                    elif int(l3) < 20 and numoflo < 2:
                        l2 = 20
                        numoflo += 1
                elif var >= 80.0:
                    l1 = 20
                    l2 = 20
                    l3 = 20

                res = int(l1)+int(l2)+int(l3)+int(p)

                grade = 0.0
                if res <= 50:
                    grade = 2
                elif 50 < res <= 60:
                    grade = 3
                elif 60 < res <= 70:
                    grade = 3.5
                elif 70 < res <= 80:
                    grade = 4
                elif 80 < res <= 90:
                    grade = 4.5
                elif 90 < res <= 100:
                    grade = 5

                cursor.execute("UPDATE grades SET grade = (?) WHERE student_id = (?)", (grade,id))
                conn.commit()
            except ValueError as e:
                cursor.execute("UPDATE grades SET grade = (?) WHERE student_id = (?)", ("None", id))
                conn.commit()

            cursor.close()
            conn.close()

            load_data()
            details_window.destroy()

        def delete():
            conn = sqlite3.connect('studentsdatabase.db')
            cursor = conn.cursor()

            id = id_entry.get()
            cursor.execute("DELETE FROM students WHERE student_id=?",(id))
            cursor.execute("DELETE FROM grades WHERE student_id=?",(id))
            conn.commit()
            cursor.close()
            conn.close()

            load_data()
            details_window.destroy()

        up_button = ttk.Button(details_window, text="UPDATE", command=update)
        up_button.pack()

        del_button = ttk.Button(details_window, text="DELETE", command=delete)
        del_button.pack()


def fetch_data():
    conn = sqlite3.connect('studentsdatabase.db')
    cursor = conn.cursor()
    cursor.execute("SELECT students.student_id, firstname, surname, email, status, "
                   "l_1, l_2, l_3, h_1, h_2, h_3, h_4, h_5, "
                   "h_6, h_7, h_8, h_9, h_10, project, grade "
                   "FROM students "
                   "INNER JOIN grades ON students.student_id = grades.student_id")
    result = cursor.fetchall()
    cursor.close()
    conn.close()
    return result


def load_data():
    data = fetch_data()
    treeview.delete(*treeview.get_children())
    for row in data:
        treeview.insert("","end",values=(row[0],row[1],row[2],row[3],row[4],row[5],
                                         row[6],row[7],row[8],row[9],row[10],row[11],
                                         row[12],row[13],row[14],row[15],row[16],
                                         row[17],row[18],row[19]))


def open_new_student_window():
    new_window = tk.Toplevel(root)
    new_window.title("ADD NEW STUDENT")
    name_label = ttk.Label(new_window, text="Name:")
    name_label.pack()
    name_entry = ttk.Entry(new_window)
    name_entry.pack()
    surname_label = ttk.Label(new_window, text="Surname:")
    surname_label.pack()
    surname_entry = ttk.Entry(new_window)
    surname_entry.pack()
    email_label = ttk.Label(new_window, text="Email:")
    email_label.pack()
    email_entry = ttk.Entry(new_window)
    email_entry.pack()
    def add_new():
        name = name_entry.get()
        surname = surname_entry.get()
        email = email_entry.get()
        conn = sqlite3.connect('studentsdatabase.db')
        cursor = conn.cursor()
        cursor.execute("INSERT INTO students (firstname, surname, email) VALUES (?,?,?)",(name,surname,email))
        conn.commit()

        cursor.execute("SELECT student_id FROM students WHERE email = (?)", (email,))
        studid = cursor.fetchone()
        cursor.execute("INSERT INTO grades(student_id) VALUES (?)", (studid))

        conn.commit()
        cursor.close()
        conn.close()

        load_data()

        new_window.destroy()

    add_button = ttk.Button(new_window, text="CREATE", command=add_new)
    add_button.pack()


button = ttk.Button(root, text="ADD STUDENT", command=open_new_student_window)

mail_but = ttk.Button(root, text="SEND EMAILS", command=send_emails)

treeview = ttk.Treeview(root)
treeview["columns"] = ("id", "name", "surname", "email", "status",
                       "list 1", "list 2", "list 3",
                       "homework 1", "homework 2", "homework 3", "homework 4", "homework 5",
                       "homework 6", "homework 7", "homework 8", "homework 9", "homework 10",
                       "project", "grade")
treeview.column("#0", width=0)
treeview.heading("id", text="ID")
treeview.column("id", width=10)
treeview.heading("name", text="NAME")
treeview.column("name", width=80)
treeview.heading("surname", text="SURNAME")
treeview.column("surname", width=80)
treeview.heading("email", text="EMAIL")
treeview.column("email", width=80)
treeview.heading("status", text="STATUS")
treeview.column("status", width=80)
treeview.heading("list 1", text="LIST 1")
treeview.column("list 1", width=80)
treeview.heading("list 2", text="LIST 2")
treeview.column("list 2", width=80)
treeview.heading("list 3", text="LIST 3")
treeview.column("list 3", width=80)
treeview.heading("homework 1", text="HOMEWORK 1")
treeview.column("homework 1", width=80)
treeview.heading("homework 2", text="HOMEWORK 2")
treeview.column("homework 2", width=80)
treeview.heading("homework 3", text="HOMEWORK 3")
treeview.column("homework 3", width=80)
treeview.heading("homework 4", text="HOMEWORK 4")
treeview.column("homework 4", width=80)
treeview.heading("homework 5", text="HOMEWORK 5")
treeview.column("homework 5", width=80)
treeview.heading("homework 6", text="HOMEWORK 6")
treeview.column("homework 6", width=80)
treeview.heading("homework 7", text="HOMEWORK 7")
treeview.column("homework 7", width=80)
treeview.heading("homework 8", text="HOMEWORK 8")
treeview.column("homework 8", width=80)
treeview.heading("homework 9", text="HOMEWORK 9")
treeview.column("homework 9", width=80)
treeview.heading("homework 10", text="HOMEWORK 10")
treeview.column("homework 10", width=80)
treeview.heading("project", text="PROJECT")
treeview.column("project", width=80)
treeview.heading("grade", text="GRADE")
treeview.column("grade", width=80)
treeview.bind("<Double-1>",open_details_window)

treeview.pack()
button.pack()
mail_but.pack()
load_data()
root.mainloop()