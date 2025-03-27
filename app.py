from flask import Flask, render_template, request, redirect
import pandas as pd
from datetime import datetime

app = Flask(__name__)

# Create a DataFrame to store data
file_name = "entries.xlsx"
columns = ["Type", "Date", "Amount", "Comment"]

# Check if the Excel file already exists
try:
    df = pd.read_excel(file_name)
except FileNotFoundError:
    df = pd.DataFrame(columns=columns)

# Route to display the landing page and handle the form
@app.route('/', methods=['GET', 'POST'])
def index():
    global df
    
    if request.method == 'POST':
        # Get form data
        entry_type = request.form['type']
        date = request.form['date']
        amount = request.form['amount']
        comment = request.form['comment']
        
        # Convert date to datetime object
        date_obj = datetime.strptime(date, '%Y-%m-%d')
        
        # Append new entry to the DataFrame
        new_entry = {"Type": entry_type, "Date": date_obj, "Amount": amount, "Comment": comment}
        df = df.append(new_entry, ignore_index=True)
        
        # Save to Excel
        df.to_excel(file_name, index=False)
    
    # Render the page with current entries
    entries = df.to_dict(orient='records')
    return render_template('index.html', entries=entries)

if __name__ == '__main__':
    app.run(debug=True)
