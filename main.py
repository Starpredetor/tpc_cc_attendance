from flask import Flask, request, render_template, redirect, url_for, send_file, flash
import re
from excel_handler import add_attendance, generate_new_attendance_workbook


app = Flask(__name__)
app.secret_key = 'supersecretkey'


#main home route
@app.route('/', methods=['GET', 'POST'])
def home():
    return render_template('home.html')

        
#handler for marking attendance
@app.route('/mark_attendance', methods=['GET', 'POST'])
def mark_attendance():
    error_message = None
    success_message = None
    
    if request.method == 'POST':
        submitted_text = request.form['text']
        session = request.form.getlist('session')
        print(session)
        #session check
        if not session:
            error_message = "Please select at least one session."
            return render_template('mark_attendance.html', error_message=error_message, success_message=success_message)
        ses = 3
        
        if len(session) == 1:
            if session[0] == 'Morning':
                ses = 1
            elif session[0] == 'Afternoon':
                ses = 2
        else:
            ses = 3
        # Roll number validation
        if not submitted_text.strip():
            error_message = "Input cannot be empty."
        elif not re.match(r'^(22|23)[A-Z]{2}\d{4}$', submitted_text):
            error_message = "Invalid roll number format."
        else:
            roll_number = submitted_text
            if add_attendance(roll_number, session=ses) == 0:
                error_message = f"Roll number {roll_number} not found in the attendance sheet."
            else:
                success_message = f"Attendance marked for roll number {roll_number}"
    
    return render_template('mark_attendance.html', error_message=error_message, success_message=success_message)



#Handler for Creation of new workbook
@app.route('/handler', methods=['POST'])
def handler():
    action = request.form.get('action')
    if action == 'create_attendance_sheet':
        message = None
        if request.method == 'POST':
           res = generate_new_attendance_workbook()
        if res == 0:
            return render_template('home.html', error_message="Sheet already created for today.") 
        else:
            return render_template('home.html', success_message=f"Attendance sheet created successfully with the name {res}")
            


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
