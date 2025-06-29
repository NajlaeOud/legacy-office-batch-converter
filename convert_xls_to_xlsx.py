import os
import win32com.client  # For controlling Microsoft Excel on Windows

def convert_xls_to_xlsx(folder_path):
    try:
        excel = win32com.client.Dispatch("Excel.Application")
    except Exception as e:
        print("❌ Microsoft Excel must be installed to run this script.")
        print("Details:", e)
        return

    excel.Visible = False

    if not os.path.exists(folder_path):
        print(f"❌ Folder does not exist: {folder_path}")
        return

    for filename in os.listdir(folder_path):
        if filename.endswith(".xls") and not filename.startswith("~$"):
            xls_path = os.path.join(folder_path, filename)
            xlsx_path = os.path.splitext(xls_path)[0] + ".xlsx"

            print(f"🔄 Converting: {filename}")
            try:
                wb = excel.Workbooks.Open(xls_path)
                wb.SaveAs(xlsx_path, FileFormat=51)  # 51 = .xlsx
                wb.Close()
            except Exception as e:
                print(f"⚠️ Failed to convert {filename}: {e}")

    excel.Quit()
    print("✅ Conversion complete.")

if __name__ == "__main__":
    folder_path = input("📁 Enter the full path to your folder containing .xls files:\n> ")
    convert_xls_to_xlsx(folder_path)
