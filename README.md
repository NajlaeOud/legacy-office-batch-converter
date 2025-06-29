#  Legacy Office Batch Converter

**Convert `.doc` and `.xls` legacy Microsoft Office files to modern `.docx` and `.xlsx` formats — fully readable by Python.**  
Built for data automation, ERP migration, and Power BI integration.



##  Why This Tool Exists

Legacy formats like `.doc` and `.xls` are **not compatible** with modern Python libraries like `python-docx`, `openpyxl`, or `pandas`.

This tool batch-processes outdated Office files, transforming them into clean, structured formats that can be:

- Parsed by Python
- Imported into Odoo, Power BI, or ERP systems
- Used in automated financial workflows



##  How It Works

Two scripts use Microsoft Office COM automation to convert files in-place.

###  Scripts included

| Script | Purpose |
|--------|---------|
| `convert_doc_to_docx.py` | Converts all `.doc` files in a folder to `.docx` |
| `convert_xls_to_xlsx.py` | Converts all `.xls` files in a folder to `.xlsx` |

Both scripts prompt the user for a folder path at runtime.

---

##  Example Usage

```bash
 Enter the full path to your folder containing .doc files:
> C:/Users/YourName/Documents/legacy_docs

 Enter the full path to your folder containing .xls files:
> C:/Users/YourName/Documents/legacy_excels
