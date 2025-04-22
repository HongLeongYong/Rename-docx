import os
import re
from docx import Document

# 修改成你要處理的資料夾路徑
folder_path = "/Users/hongleongyong/Desktop/word_rename"
target_date = "20250422"
pattern = r"2025\d{4}"

for filename in os.listdir(folder_path):
    if filename.endswith(".docx"):
        old_path = os.path.join(folder_path, filename)
        doc = Document(old_path)

        # 修改內容
        for para in doc.paragraphs:
            if re.search(pattern, para.text):
                para.text = re.sub(pattern, target_date, para.text)

        # 修改檔名
        new_filename = re.sub(pattern, target_date, filename)
        new_path = os.path.join(folder_path, new_filename)
        doc.save(new_path)

        # 如果檔名有改變就刪除舊檔
        if new_path != old_path:
            os.remove(old_path)

print("✅ 所有文件已處理完成！")
