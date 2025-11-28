# ETP Certificate Automation Bot (UiPath)

This UiPath automation bot reads participant information from an Excel file, generates personalized certificates using a PPT/Word template, converts them into PDF, and sends them directly to their email inbox.

This project is developed as part of the **ETP (Employability Training Program)**.

---

##  Features

###  Excel Automation
- Reads participant details such as **Name, USN, Email, Event, Date**
- Validates data rows before processing

###  Certificate Generation
- Loads a certificate **PPT/Word template**
- Dynamically replaces placeholders:
  - `Title. Firstname Lastname` → Participant Name  
  - `NAME OF EVENT` → Event Name  
  - `DD/MM/YYYY` → Event Date  

###  PDF Export
- Automatically converts the generated certificate to **PDF format**
- Saves all PDFs in the `certificates/` folder

###  Email Automation
- Sends each participant their certificate via SMTP
- Supports Gmail App Password authentication
- Customizable email body using a `.txt` template

###  Logging & Status Messages
- Shows processing status for each participant
- Displays final success message when completed

---

##  Project Structure

---

## 🔧 Tech Stack

| Component | Technology Used |
|----------|-----------------|
| Automation Tool | UiPath Studio (Windows) |
| Document Editing | PowerPoint / Word Automation |
| PDF Export | Office + UiPath |
| Email Automation | SMTP (Gmail) |
| Language | VB Expressions inside UiPath |

---

## 🚀 How It Works (Workflow Overview)

1. Load participant data from Excel  
2. Loop through each row  
3. For each participant:
   - Load template
   - Replace text fields
   - Export certificate as PDF
   - Generate email message
   - Attach PDF and send via email  
4. Process repeats until all rows are completed  
5. Display success notification  

---

##  Sample Input (Excel)

| Name       | USN         | Email              | Event            | Date       |
|------------|-------------|--------------------|------------------|------------|
| John Doe   | 4SO21CS001  | john@gmail.com     | Python Workshop  | 21/11/2025 |
| Jane Smith | 4SO21CS002  | jane@gmail.com     | AI Bootcamp      | 21/11/2025 |

---

##  Sample Email Output

---

## 🛠️ Setup Instructions

### 1️⃣ Install Required UiPath Packages
From **Manage Packages → Official**:
- UiPath.Excel.Activities  
- UiPath.Word/PowerPoint.Activities  
- UiPath.Mail.Activities  
- UiPath.System.Activities  

### 2️⃣ Configure Email
Inside `Send SMTP Mail Message`:
- Server: `smtp.gmail.com`
- Port: `587`
- Enable SSL: `True`
- Use **Gmail App Password** (not normal password)

### 3️⃣ Place Required Files
Ensure these exist:
- `Certificate Sample.pptx`
- `Certificate Text.txt`
- `List.xlsx`

### 4️⃣ Run the Bot
Click **Run** inside UiPath Studio.  
All certificates will be generated and emailed automatically.


---

---





