import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Global variables
excel_file_path = "attendance.xlsx"
email_sender = "your_email@gmail.com"
email_password = "your_password"
email_subject = "Attendance Report"
threshold = 80  # Adjust this threshold as needed
absentees = []

# Initialize the Excel sheet and load data
def initialize_excel():
    global wb, sheet
    try:
        wb = openpyxl.load_workbook(excel_file_path)
        sheet = wb.active
    except Exception as e:
        print(f"Error initializing Excel: {str(e)}")

# Save the Excel sheet after updating
def save_excel():
    try:
        wb.save(excel_file_path)
        wb.close()
    except Exception as e:
        print(f"Error saving Excel: {str(e)}")

# Function to track attendance and update the absentees list
def track_attendance(student_name, attendance_percentage):
    global absentees
    if attendance_percentage < threshold:
        absentees.append(student_name)
        print(f"{student_name} is marked absent. Absentees: {absentees}")

# Function to send an email with attendance report
def send_email():
    global absentees
    try:
        message = MIMEMultipart()
        message["From"] = email_sender
        message["Subject"] = email_subject

        body = f"Attendance report for today:\n\n"
        if not absentees:
            body += "All students are present."
        else:
            body += "Absent students:\n"
            for student in absentees:
                body += f"- {student}\n"

        message.attach(MIMEText(body, "plain"))

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(email_sender, email_password)

        server.sendmail(email_sender, email_sender, message.as_string())
        server.quit()
        print("Email sent successfully.")
    except Exception as e:
        print(f"Error sending email: {str(e)}")

if __name__ == "__main__":
    initialize_excel()
    # Simulate attendance updates - you can replace this with your data
    track_attendance("Student1", 75)
    track_attendance("Student2", 90)
    track_attendance("Student3", 70)
    save_excel()
    send_email()
