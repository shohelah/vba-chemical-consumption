# vba-chemical-consumption
In this project, I developed a VBA script to collect consumption data for 'Chemicals/Dyes' and 'Dyes' categories from multiple Excel sheets and calculate the total (sum) for each product code. The script automates header identification and presents results in a Summary sheet.
# 🧪 Excel VBA Project: Chemicals/Dyes Consumption Analyzer

This project includes an Excel VBA script that automates the process of summarizing chemical and dye consumption from multiple worksheets in a workbook. It identifies headers dynamically and creates a summary sheet with total consumption by product code.

---

## 🔍 Project Overview

🔸 **Project Name:** Chemicals/Dyes Consumption Analyzer  
🔸 **Language:** VBA (Excel Macro)  
🔸 **Platform:** Microsoft Excel  
🔸 **Purpose:** Automate data collection and summarization from various sheets containing chemical usage data.

---

## 📂 Features

- ✅ Automatically scans all sheets (excluding "Summary")
- ✅ Detects "Name of Chemical/Dyes" and "Dyes Items" headers
- ✅ Aggregates consumption values by product code
- ✅ Creates or refreshes a summary sheet with total values
- ✅ Handles missing headers and invalid data with error management

---

## 🛠️ Technologies Used

- VBA (Visual Basic for Applications)  
- Microsoft Excel  
- Dictionary Object for data aggregation

---

## 📋 Sample Output (in Summary Sheet)

| Product Code | Category              | Total Consumption | Sheets Counted |
|--------------|-----------------------|-------------------|----------------|
| 615390       | Sum of Chemicals/Dyes | 8.080             | All Sheets     |
| 615390       | Sum of Dyes           | 17.502            | All Sheets     |
| 413504       | Sum of Chemicals/Dyes | [value]           | All Sheets     |
| 413504       | Sum of Dyes           | [value]           | All Sheets     |

---

## 📎 How to Use

1. Open the Excel workbook containing multiple data sheets
2. Open VBA Editor (`Alt + F11`)
3. Paste the code in a module
4. Run `Summarize_Chemicals_Dyes_Totals_Fixed` macro
5. Check the newly created "Summary" sheet

---

## ⚠️ Challenges Faced & Solved

| Challenge                             | Solution                                                |
|---------------------------------------|----------------------------------------------------------|
| Dynamic header position               | Used smart loop to detect keyword position               |
| Handling missing/empty values         | Used `On Error Resume Next` and validation               |
| Overwriting existing "Summary" sheet | Automatic delete/clear with `DisplayAlerts = False`     |

---

## 📷 <img width="579" height="375" alt="image" src="https://github.com/user-attachments/assets/7a494260-3d5e-4032-aa49-1b2226f6c8f4" />
<img width="760" height="462" alt="image" src="https://github.com/user-attachments/assets/a0fd7937-f23c-4ba1-8135-f2d0566f9d8b" />
<img width="492" height="403" alt="image" src="https://github.com/user-attachments/assets/da8d8fe5-79f6-40a7-9727-b429eb3ff970" />


*(Add screenshots of your Excel sheets and summary tab here)*

---

## 🔗 Live Demo

👉 https://docs.google.com/spreadsheets/d/13q66oxDDD_DM5aUyJwx9tHXkn4Pj8lgG/edit?usp=drive_link&ouid=116403674883636880344&rtpof=true&sd=true

---

## 👨‍💻 Author

**Md. Shohel Ahmod**  
Assistant Manager, Data Analyst @ Paramount Textile PLC  
📧 shohelahmod@gmail.com | 🌐 [LinkedIn](https://www.linkedin.com/in/md-shohel-ahmod-b7650357/)

---

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
