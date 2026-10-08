# Excel Data Cleaner

A Python desktop application that automatically reads customer data from SQL Server, cleans and validates the data, generates a professional Excel report, and provides a simple graphical user interface.

---

## 📌 Project Overview

Excel Data Cleaner is a data-cleaning automation tool built with Python.

The application connects to Microsoft SQL Server, reads customer data, performs several data-quality checks, removes invalid records, and exports the cleaned dataset to an Excel file.

The project also provides a graphical user interface (GUI) with:

- Real-time cleaning progress
- Cleaning statistics
- Success and error messages
- Excel file opening
- Logging
- Error handling

---

## 🚀 Features

### Database

- Connects to Microsoft SQL Server using `pyodbc`
- Reads customer data using SQL queries
- Uses Windows Trusted Connection

### Data Cleaning

The application automatically:

- Removes duplicate rows
- Removes rows with missing required values
- Normalizes email addresses
- Validates email addresses
- Removes invalid email records

### Excel Export

The application generates:

`customers_filtered.xlsx`

The Excel workbook contains:

- `Cleaned Data`
- `Cleaning Report`

The workbook is automatically formatted with:

- Professional headers
- Auto-adjusted column widths
- Filters
- Frozen header row

### Reporting

The application generates a cleaning report containing:

- Rows before cleaning
- Duplicate rows removed
- Rows with missing values removed
- Invalid email rows removed
- Rows after cleaning

### Logging

Application activity is saved in:

`cleaner.log`

Log files use rotation to prevent the log file from growing indefinitely.

### Graphical User Interface

The application includes a Tkinter GUI with:

- Clean professional interface
- Start cleaning button
- Real-time progress percentage
- Cleaning status
- Cleaning report
- Open Excel button
- Error handling

---

## 🛠️ Technologies Used

- Python
- Pandas
- PyODBC
- OpenPyXL
- Tkinter
- Microsoft SQL Server
- Excel

---

## 📂 Project Structure

```text
excel-data-cleaner/
├── app.py
├── cleaner.py
├── data.csv
├── requirements.txt
├── README.md
└── .gitignore
```
