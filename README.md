# Ultrasonic-Testing-Excel-Automatic-Form-Filling
A Python tool for automating construction data reporting. It extracts inspection data from Word templates, categorizes structural components (columns, beams, supports), and generates formatted Excel reports. Designed to boost efficiency, reduce errors, and streamline construction documentation workflows.
Automated Construction Data Reporting

A Python-based automation tool designed to streamline inspection data processing in construction projects.
It extracts measurement data from Word templates, categorizes structural components (columns, beams, supports), and generates formatted Excel reports with preserved styles.
This project was originally developed during an internship to improve real-world engineering workflows, demonstrating skills in Python automation, data processing, and construction documentation.

✨ Key Features
📄 Word Data Extraction – Reads inspection data directly from Word templates
🏗️ Structural Categorization – Handles steel columns, beams, supports with automatic recognition
📊 Excel Report Generation – Produces pre-formatted Excel reports ready for printing
🔍 Instrument Detection – Identifies probe models and measurement details automatically
🗓️ Date Bucketing & Floor Segmentation – Organizes records by floors and measurement dates
⚡ Workflow Efficiency – Reduces manual input and ensures accuracy in construction documentation

📂 Repository Structure
├── 探伤excel填表.py       # Main Python script  
├── eg.docx               # Example Word template  
├── eg excel.xlsx           # Example Excel template  
├── README.md             # Project documentation  

🚀 How to Use
Prepare a Word template with inspection data (see eg.docx as reference).
Run the Python script 探伤excel填表.py.
The program extracts data, categorizes by component, and generates a formatted Excel report.
Review and print the final report directly from Excel.

🏗️ Background & Purpose
This project was created during an engineering internship to handle repetitive inspection reporting tasks.
By automating the workflow, the tool significantly reduces manual workload, improves data accuracy, and speeds up reporting — making it a valuable asset in real-world construction environments.

📜 License
This project is intended for educational and professional portfolio purposes.
You are free to reference and adapt it for personal use.
