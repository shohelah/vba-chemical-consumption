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

## 📷 Screenshots

*(Add screenshots of your Excel sheets and summary tab here)*

---

## 🔗 Live Demo

👉 [Download Sample File (.xlsm)](your_google_drive_or_github_link)

---

## 👨‍💻 Author

**Md. Shohel Ahmod**  
Assistant Manager, Supply Chain @ Paramount Textile PLC  
📧 shohelahmod@gmail.com | 🌐 [LinkedIn](https://www.linkedin.com/in/md-shohel-ahmod-b7650357/)

---

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
