from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
from email_validator import validate_email, EmailNotValidError

app = Flask(__name__)

def validate_emails(df):
    valid_emails = []
    invalid_emails = []
    for index, row in df.iterrows():
        email = row['Email']
        try:
            # Validate email
            validate_email(email)
            valid_emails.append(row)
        except EmailNotValidError:
            invalid_emails.append(row)
    return valid_emails, invalid_emails

def save_to_excel(df, filename):
    df.to_excel(filename, index=False)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return redirect(request.url)
    
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    
    if file:
        filename = file.filename
        if filename.endswith('.xlsx'):
            try:
                # Read Excel file
                df = pd.read_excel(file)
                # Validate emails and split into valid and invalid
                valid_emails, invalid_emails = validate_emails(df)
                # Save valid and invalid emails to Excel sheets
                save_to_excel(pd.DataFrame(valid_emails), 'valid_emails.xlsx')
                save_to_excel(pd.DataFrame(invalid_emails), 'invalid_emails.xlsx')
                return "Files processed successfully! Check your directory for the results."
            except Exception as e:
                return "Error processing file: " + str(e)
        else:
            return "Only Excel files (.xlsx) are allowed."
    return redirect(request.url)

if __name__ == '__main__':
      app.run(host='0.0.0.0', port=10000)
