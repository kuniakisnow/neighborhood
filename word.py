#from docx import Document
import os
import sys

for entry in os.scandir():
    if entry.is_file():
        print(entry.name)
    else:
        print(f'dir:{entry.name}')

sys.exit()

# ファイル名
word_read_file_name = "word.docx"

# ドキュメントオブジェクト作成
# <class 'docx.document.Document'>
document = Document(word_read_file_name)

# データの表示
for i in document.paragraphs:
  print(i.text)