# Kennedy Center Refund Form Automation Project

This project is a macro-enabled Excel tool designed to automate the process of generating refund or charge forms for student transactions at the Kennedy Center. It simplifies manual data entry and ensures accuracy, consistency, and formatting compliance for uploading to financial systems.

---

## 📌 Project Overview

- **Purpose:** Automate the student charge/refund process using form controls and VBA logic.
- **File Type:** `.xlsm` (macro-enabled Excel workbook)
- **User Input:** Student ID, Item Type, Charge/Refund Type, and Amount
- **Output:** Populated rows in the “SFS Upload” sheet with properly formatted financial data

---

## ⚙️ Features

- ✅ Dynamic row insertion based on checkbox selections (e.g., program charges or application fees)
- ✅ Validates BYU Student ID (must be 9 digits, allows leading zero)
- ✅ Ensures Amount and Item Type fields are completed before insertion
- ✅ Automatically writes data to the correct range with consistent formatting
- ✅ Prevents duplicate or empty rows by validating entries
- ✅ Compatible with the Kennedy Center’s financial upload template

---

## 🧠 Technologies Used

- **Excel VBA (Visual Basic for Applications)**
- **Excel Form Controls (TextBoxes, CheckBoxes, Buttons)**
- **Macro-enabled workbook (`.xlsm`)**

---

## 🚀 How to Use

1. Open the Excel file and enable macros.
2. Navigate to the **Add Information** form.
3. Enter:
   - 9-digit BYU Student ID
   - Item Type (12-digit code)
   - Amounts for selected charges
4. Check the appropriate boxes (Program Charges, Application Fee, etc.)
5. Click the **Submit** or **Add Charges** button to populate the “SFS Upload” sheet.
6. Use the **Clear** button to reset the form if needed.

---

## 📎 Notes

- Make sure macros are enabled when opening the file.
- The form validates input and will show messages if fields are missing.
- Row insertion aligns with Jeff’s formatting logic to ensure compatibility.

---

## 👨‍💻 Author

**Honghan (Lucas) Wang**  **Jeff Thomas** 
Data Science Student, Brigham Young University  
📫 Lucaswang718@gmail.com

---

## 📂 File

- `Kennedy Center Refund Form Automation Project.xlsm`  
  Contains all code, forms, and data formatting logic.

