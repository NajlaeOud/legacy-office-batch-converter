import os
import comtypes.client  # For controlling Microsoft Word on Windows

def convert_doc_to_docx(folder_path):
    # Start Word in the background (invisible to the user)
    try:
        word = comtypes.client.CreateObject('Word.Application')
    except Exception as e:
        print("❌ Microsoft Word must be installed to run this script.")
        print("Details:", e)
        return

    word.Visible = False  # Keeps Word from opening in front of the user

    if not os.path.exists(folder_path):
        print(f"❌ Folder does not exist: {folder_path}")
        return

    # Loop through all files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".doc") and not filename.startswith("~$"):
            doc_path = os.path.join(folder_path, filename)
            docx_path = os.path.splitext(doc_path)[0] + ".docx"

            print(f"🔄 Converting: {filename}")
            try:
                doc = word.Documents.Open(doc_path)
                doc.SaveAs(docx_path, FileFormat=16)  # 16 = .docx format
                doc.Close()
            except Exception as e:
                print(f"⚠️ Failed to convert {filename}: {e}")

    # Close Word when done
    word.Quit()
    print("✅ Conversion complete.")

# ------------------------------
# 👇 Run this if you're using the script directly
if __name__ == "__main__":
    folder_path = input("📁 Enter the full path to your folder containing .doc files:\n> ")
    convert_doc_to_docx(folder_path)
