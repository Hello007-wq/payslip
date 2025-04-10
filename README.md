# payslip
# Payslip Generator

A Python script to generate and send payslips via email.

## Table of Contents

- [Description](#description)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Description

This script reads employee data from an Excel file, calculates the net salary, generates a PDF payslip, and sends it via email.

## Features

- Reads employee data from an Excel file.
- Calculates net salary based on basic salary, allowances, and deductions.
- Generates PDF payslips with colors for better readability.
- Sends payslips via email using SMTP.

## Prerequisites

- Python 3.x
- Required Python libraries: `pandas`, `fpdf`, `smtplib`, `python-dotenv`
- A Gmail account with "Less secure app access" enabled or an app-specific password.
- An Excel file containing employee data.

## Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/Hello007-wq/payslip.git
   cd payslip

2. Install required libraries
   pip install pandas fpdf python-dotenv

3. Set up environment variables:
   .Create a .env file in the root directory with the following content:
   EMAIL=your_email@example.com
   EMAIL_PASSWORD=your_email_password_or_app_specific_password

4. Prepare the excel file:
   .Ensure your Excel file is named employees.xlsx and is located in the root directory.
   .The Excel file should have the following columns: Name, Email, Basic Salary, Allowances, Deductions.

## Usage

1. Run the Script:
   python pay_slip.py

2. Check the output:
   .PDF payslips will be generated in the payslips directory.
   .Emails will be sent to the employees with their respective payslips attached.

  Contributing
  Contributions are welcome! Please open an issue or submit a pull request with your changes.
