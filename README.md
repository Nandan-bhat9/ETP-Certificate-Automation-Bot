# ETP-Certificate-Automation-Bot
This bot takes in the Excel list and auto-generates the certificate and sends it via Email using  UIPath
This automation project generates personalised certificates for participants based on data from an Excel file and emails them automatically.
It is built using UiPath Studio, and it demonstrates real-world automation covering:
•	Excel Automation
•	Word Document Automation
•	PDF creation
•	Email Dispatch with Attachments
•	Dynamic Text Replacement
Features
Read participant data from Excel- Reads Name, USN, Email, Event, Date.
Generate certificates from a PPT template- Replaces placeholders such as:
•	Title. FirstName LastName → student name
•	NAME OF EVENT → event name
•	DD/MM/YYYY → event date
Export each certificate as a PDF
Send certificates via email
Error-handling and logging
Displays progress and success messages.
Setup Instructions:
Install Required UiPath Packages
From Manage Packages → Official:
•	UiPath.Excel.Activities
•	UiPath.Word.Activities
•	UiPath.Mail.Activities
•	UiPath.System.Activities
Place Required Files
Ensure the following are in the project root:
•	Template Certificate
•	Excel file with participant names
•	cert_temp folder
•	certificates folder
Configure Email
Inside Uc1_certificateMail.xaml:
•	Replace your_email with your Gmail ID
•	Use Gmail App Password
Output
•	Automatically generated PDF certificates in /certificates/ and sent via email.
