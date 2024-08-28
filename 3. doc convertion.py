import os
import comtypes.client


# ---- Conversion part----
def init_word():
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    return word

def doc_to_docx(directory, word):
    for root, dirs, files in os.walk(directory):
        for filename in files:
            if filename.endswith('.doc'):
                doc_path = os.path.join(root, filename)
                docx_path = doc_path + 'x'  # Change file extension to .docx
                try:
                    doc = word.Documents.Open(doc_path)
                    doc.SaveAs2(docx_path, FileFormat=16)  # FileFormat=16 for docx
                    doc.Close()
                    print(f"Converted {doc_path} to {docx_path}")
                except Exception as e:
                    print(f"Error converting {doc_path}. Reason: {e}")

def close_word(word):
    word.Quit()

base_directory = input("Enter base directory: ")
word = init_word()
doc_to_docx(base_directory, word)
close_word(word)


# ---- Optional deletion part----
# def delete_doc_files(directory):
#     for root, dirs, files in os.walk(directory):
#         for filename in files:
#             if filename.endswith('.doc'):
#                 file_path = os.path.join(root, filename)
#                 try:
#                     os.remove(file_path)
#                     print(f"Deleted {file_path}")
#                 except Exception as e:
#                     print(f"Error deleting {file_path}. Reason: {e}")
# 
# base_directory = input("Enter base directory: ")
# delete_doc_files(base_directory)

