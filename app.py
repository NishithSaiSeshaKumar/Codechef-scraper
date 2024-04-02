"""from flask import Flask, render_template, request, send_file
import pandas as pd
from attempting import attempt

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if a file was uploaded
        if 'file' not in request.files:
            return render_template('index.html', error='No file uploaded')
        
        file = request.files['file']
        if file.filename == '':
            return render_template('index.html', error='No file selected')

        # Read the Excel file and extract data
        try:
            df = pd.read_excel(file)
            df = df.head(500)
            df.rename(columns={df.columns[-1]: 'CCID'}, inplace=True)
            df['CCID'] = df['CCID'].astype(str)
            users = df["CCID"].tolist()  # Assuming usernames are in the "CCID" column
            contest_no = request.form.get('contest_no')
        except Exception as e:
            return render_template('index.html', error=f'Error reading file: {str(e)}')

        # Call the attempt function to process the data
        results, errors = attempt(users, contest_no)

        # Update the DataFrame with the attempt results
        df[f"CONTEST {contest_no}"] = results

        # Save the DataFrame to an Excel file
        output_file = 'results.xlsx'
        df.to_excel(output_file, index=False)

        # Return the Excel file for download
        return send_file(output_file, as_attachment=True)

    return render_template('index.html')

if __name__ == '__main__':
    app.run()
"""
from flask import Flask, render_template, request, send_file
import pandas as pd
from attempting import attempt,analyze_attendance_and_apply_colors

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if a file was uploaded
        if 'file' not in request.files:
            return render_template('index.html', error='No file uploaded')
        
        file = request.files['file']
        if file.filename == '':
            return render_template('index.html', error='No file selected')

        # Read the Excel file and extract data
        try:
            df = pd.read_excel(file)
            nos = int(request.form.get('no_of_students'))  # Get the number of rows to process
            df = df.head(nos)  # Limit rows to first 500
            df.rename(columns={df.columns[-1]: 'CCID'}, inplace=True)
            df['CCID'] = df['CCID'].astype(str)
            users = df["CCID"].tolist()  # Assuming usernames are in the "CCID" column
            contest_no = request.form.get('contest_no')
        except Exception as e:
            return render_template('index.html', error=f'Error reading file: {str(e)}')

        # Call the attempt function to process the data
        results, errors = attempt(users, contest_no)

        # Update the DataFrame with the attempt results
        df[f"CONTEST {contest_no}"] = results

        # Save the DataFrame to an Excel file
        output_file = 'results.xlsx'
        df.to_excel(output_file, index=False)
        analyze_attendance_and_apply_colors(output_file, output_file,contest_no)
        # Return the Excel file for download
        return send_file(output_file, as_attachment=True, download_name='results.xlsx')  # Delete file after download

    return render_template('index.html')

if __name__ == '__main__':
    app.run()
