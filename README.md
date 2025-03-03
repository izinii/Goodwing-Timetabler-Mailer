# Goodwing-Timetabler-Mailer 📅✉️

## 📌 Overview
Goodwing-Timetabler-Mailer is an **automated email system** that extracts university timetables from an Excel file and sends personalized emails to professors with their schedules attached. The system **preserves the original Excel formatting**, ensuring a clean and professional timetable delivery.
This functionality can be seen as a **plugin** for the **Goodwing-Timetabler** project, which generates university timetables *(see the Goodwing-Timetabler repository)*.

## 🚀 Features
- **Automated professor detection** → Extracts professor names and generates their email.
- **Email generation** → Converts names to email format (`firstname.lastname@devinci.fr`).
- **Accent removal** → Ensures correct email formatting by removing special characters.
- **Timetable extraction** → Filters the Excel file to send only the relevant weeks.
- **Excel formatting preservation** → Keeps styles, colors, and structure.
- **Email automation** → Sends personalized emails with the correct attachments.

---

## To run the code

### **1- Set up your email credentials**
```bash
git clone https://github.com/your-username/Goodwing-Timetabler-Mailer.git
cd Goodwing-Timetabler-Mailer
