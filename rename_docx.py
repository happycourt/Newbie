import os
from docx import Document


def extract_title_from_docx(file_path):
    doc = Document(file_path)
    # 默认将第一个段落作为标题
    title = doc.paragraphs[0].text
    return title.strip()  # 去除前后的空白字符


def rename_file_with_title(file_path):
    # 提取标题
    title = extract_title_from_docx(file_path)

    # 获取文件所在目录和扩展名
    directory, _ = os.path.split(file_path)
    _, extension = os.path.splitext(file_path)

    # 创建新的文件名
    new_file_name = f"{title}{extension}"
    new_file_path = os.path.join(directory, new_file_name)

    # 重命名文件
    os.rename(file_path, new_file_path)
    print(f"文件已重命名为: {new_file_path}")


# 示例用法
file_path = r"D:\All夹\专四专八\专四\专四词汇辨析.docx"  # 替换为你的文件路径
rename_file_with_title(file_path)
