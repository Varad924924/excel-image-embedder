Streamlit Excel Image Embedder 
A simple yet powerful web application built with Streamlit to automate the tedious task of embedding images into Excel spreadsheets. This tool allows users to upload an Excel file and a ZIP file of images, and it will automatically insert each image into the correct row based on a matching ID.

Demo
 Live App: https://virtual-excel-image-embedder.streamlit.app/


Screenshot
Here's a look at the application's user-friendly interface.


Features
Sidebar Layout: A clean and intuitive interface with all controls on the sidebar.

File Upload: Supports .xlsx for spreadsheets and .zip for image archives.

Data Preview: Instantly displays the first few rows of the uploaded Excel file for verification.

Automated Embedding: Matches images (e.g., part_12345.jpg) to the corresponding 'Property ID' in the Excel file.

Custom Formatting: Automatically adjusts column widths and row heights for a clean, professional-looking report.

One-Click Processing: Embed all images with a single button click.

Easy Reset & Download: Reset the app with one click and download the final, modified Excel file.

How to Run Locally
To run this application on your local machine, please follow these steps.

Prerequisites
Python 3.8+

pip package manager

Installation & Setup
Clone the repository:

Bash

git clone https://github.com/Varad924924/excel-image-embedder.git
cd YOUR_REPOSITORY_NAME
Install the required libraries:

Bash

pip install -r requirements.txt
Run the Streamlit app:

Bash

streamlit run app.py
The application will open in your default web browser.

Technologies Used
Python: The core programming language.

Streamlit: For building and deploying the interactive web application.

Pandas: For reading and previewing the Excel data.

Openpyxl: For the core functionality of reading, writing, and manipulating .xlsx files.

Pillow: For handling image processing and resizing.
