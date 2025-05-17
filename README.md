# Codechef Scraper

## Overview
The **Codechef Scraper** is a web-based application designed to analyze and process data related to CodeChef contests. It allows users to upload an Excel file containing user data, fetch contest participation details for each user, and generate a detailed report with attendance analysis and visual enhancements. The application is built using Python and Flask, with support for Excel file processing and web scraping.

## Features
1. **Excel File Upload**: Users can upload an Excel file containing user details.
2. **Contest Participation Analysis**: The application fetches contest participation details for each user from CodeChef.
3. **Attendance Analysis**: Generates a detailed attendance report, including attempted, not attempted, and attendance percentage for each section.
4. **Customizable Input**: Users can specify the contest number and the number of students to process.
5. **Styled Excel Output**: The output Excel file includes color-coded rows and styled headers for better readability.
6. **Downloadable Results**: The processed Excel file is available for download.

## How It Works
1. **Upload Excel File**: The user uploads an Excel file containing user details.
2. **Input Contest Details**: The user specifies the contest number and the number of students to process.
3. **Data Processing**:
   - The application reads the Excel file and extracts user data.
   - It fetches contest participation details for each user using web scraping.
   - Attendance analysis is performed, and the results are added to the Excel file.
4. **Generate Output**:
   - The processed data is saved to a new Excel file.
   - The file is styled with colors and formatting for better presentation.
5. **Download Results**: The user can download the processed Excel file.

## Project Structure
- **`app.py`**: The main Flask application that handles file uploads, processes data, and serves the output file.
- **`attempting.py`**: Contains the core logic for fetching contest participation details and performing attendance analysis.
- **`templates/index.html`**: The HTML template for the web interface, allowing users to upload files and input contest details.
- **`vercel.json`**: Configuration file for deploying the application on Vercel.
- **`requirements.txt`**: Lists the Python dependencies required for the project.

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/codechef-scraper.git
   cd codechef-scraper
2. Install dependencies:
    Run the application:
    Access the application in your browser at [http://127.0.0.1:5000](http://127.0.0.1:5000).
## Deployment
The project is configured for deployment on Vercel. The vercel.json file specifies the build and routing configuration.
### Dependencies
- Flask
- pandas
- openpyxl
- requests
- beautifulsoup4
### Usage
1. Open the application in your browser.
2. Upload an Excel file containing user details.
3. Enter the contest number and the number of students to process.
4. Submit the form to process the data.
5. Download the generated Excel file with the results.
