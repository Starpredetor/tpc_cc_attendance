from flask import Flask, request, render_template, redirect, url_for, send_file, flash
import openpyxl
import os
from werkzeug.utils import secure_filename
import re
from excel_handler import add_attendance


app = Flask(__name__)
app.secret_key = 'supersecretkey'



@app.route('/')
def home():
    return render_template('home.html')


@app.route('/mark_attendance', methods=['GET', 'POST'])
def mark_attendance():
    error_message = None
    success_message = None

    if request.method == 'POST':
        submitted_text = request.form['text']

        if not submitted_text.strip():
            error_message = "Input cannot be empty."
        elif not re.match(r'^(22|23)[A-Z]{2}\d{4}$', submitted_text):
            error_message = "Invalid roll number format."
        else:
            roll_number = submitted_text
            success_message = add_attendance(roll_number)

    return render_template('mark_attendance.html', error_message=error_message, success_message=success_message)



if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
