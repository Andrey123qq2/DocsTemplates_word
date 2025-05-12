# 📝 Word Template Filler PowerShell Script

This PowerShell script automates filling Word document templates with user data from a CSV file. It supports multiple users (identified by surname) and prompts for any missing values required by the template.

## 📂 Features

- Load user data from a UTF-8, semicolon-delimited CSV file.
- Fill a selected Word `.docx` template with corresponding user data.
- Prompt for any template variables not found in the CSV.
- Save generated files into a specified destination folder.
- Allows filling data for multiple users in one run.

---

## 🧰 Prerequisites

- **PowerShell 5.1+**
- **Microsoft Word** installed (COM automation is used)
- CSV file must be **UTF-8 encoded** and use `;` as a delimiter

---

## 🚀 Usage

1. **Edit your config** file (e.g., `config.json`) to include:
    ```json
    {
      "CSVFile_users": "users.csv",
      "TemplatesFolder": "Templates",
      "DstFolder": "Output",
      "FileNameReplaceVar": "Surname",
      "Prompt_csv_keyfield": "Enter Surnames (comma-separated):",
      "vars_description": {
        "IssueDate": "Date of Issue",
        "SignerName": "Signed by"
      }
    }
    ```

2. **Run the script**:
    ```powershell
    .\CreateDocs_from_templates.ps1
    ```

3. **Follow prompts** to enter surnames and any missing template variables.

---

## 📄 Example: users.csv

```csv
Surname;FirstName;LastName;Position;Department
Smith;John;Smith;Accountant;Finance
Johnson;Emily;Johnson;HR Manager;Human Resources
Brown;Michael;Brown;IT Specialist;IT Department
```

## 📄 Template Example: ${Surname}_certificate.docx
```docx
EMPLOYMENT CERTIFICATE

This is to certify that ${FirstName} ${LastName} (${Surname}) is employed as a ${Position}
in the ${Department} Department.

Date of Issue: ${IssueDate}
Signed by: ${SignerName}

```