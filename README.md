# Simple Attendance Tracker

This is a Python script for tracking attendance, sending email notifications to absentees, and checking if the attendance is above or below a specified threshold. It uses the `openpyxl` library for Excel file handling and the `smtplib` library for sending emails.

## Features

- Load attendance data from an Excel file.
- Save attendance data to the same Excel file.
- Send email notifications to absentees.
- Check if the attendance is above or below a specified threshold.

## Prerequisites

- Python 3.x installed.
- Required Python libraries can be installed using `pip`:
```
pip install openpyxl
```


## Configuration

Before using the script, make sure to configure the following variables in the code:

- `attendance_file`: The path to the Excel file containing attendance data.
- `threshold`: The attendance percentage threshold for sending notifications.
- `email_config`: Email configuration, including the SMTP server, sender's email address, and email message.

## Usage

1. Run the `track_attendance` function in the script to track attendance and send email notifications to absentees:
   ```
   python attendance_tracker.py
   ```


2. The script will load attendance data, calculate attendance percentage, send email notifications to absentees, and print whether the attendance is above or below the threshold.


## Acknowledgments

- [openpyxl](https://openpyxl.readthedocs.io/) for Excel file handling.
- [smtplib](https://docs.python.org/3/library/smtplib.html) for sending email notifications.
