# NIS2 Audit Report Generator

This script automates the generation of a structured NIS2 compliance report from an HTML or DOCX source file. It converts raw audit exports into professional reports using pre-defined Word templates with multilingual support.

---

## 🚀 Features

* ✅ **HTML to DOCX Conversion** using LibreOffice (HTML → ODT → DOCX)
* 📑 **Template-based Word report generation** (basic / important / essential)
* 🌍 **Multilingual support**: French (`fr`), English (`en`), Dutch (`nl`)
* 🎯 **Extraction of control IDs, statuses, and observations**
* 🎨 **Color-coded compliance status** in the report tables
* 📄 **Export to PDF** using LibreOffice
* 🖥️ **Interactive CLI interface** for easy file selection

---

## 📁 Folder Structure

```bash
project_folder/
├── audits/                # Place your source .html or .docx files here
├── output/                # Generated reports will be saved here
├── templates/             # Contains the DOCX templates per language and type
│   ├── BASIC_NIS2_TEMPLATE_FR.docx
│   ├── IMPORTANT_NIS2_TEMPLATE_EN.docx
│   └── ...
├── nis2_audit_report_generator.py
└── README.md
```

---

## ⚙️ Requirements

* Python 3.7+
* [LibreOffice](https://www.libreoffice.org/) installed and available in the system PATH
```

Ensure LibreOffice is installed:

```bash
sudo apt install libreoffice

```

---

## 🧪 Usage

1. Place your `.html` or `.docx` audit file inside the `./audits` directory.

2. Run the script:

```bash
python3 nis2_audit_report_generator.py
```

3. Follow the CLI prompts:

   * Select the file to process
   * Choose the language of the template
   * Optionally define the output report title
   * Define the client name

4. The final `.docx` and `.pdf` will be generated in the `./output/<lang>/` directory.

---

## 📌 Notes

* HTML files are automatically converted to DOCX before processing.
* Template selection is based on keywords found in the source file content.
* Status tags (e.g., `{{STATUT_AC_01}}`) and observations (e.g., `{{OBSERVATION_AC_01}}`) are dynamically replaced.

---

## 🧑‍💻 Author

DeDseC-Default — Pentester & Cybersecurity Student

---
