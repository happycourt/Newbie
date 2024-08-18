from docx import Document

def test_docx(file_path):
    try:
        doc = Document(file_path)
        print("文档读取成功！")
        print(f"第一个段落: {doc.paragraphs[0].text}")
    except Exception as e:
        print(f"发生错误: {e}")

file_path = r"D:\All夹\专四专八\专四\专四词汇辨析.docx"  # 替换为你的文件路径
test_docx(file_path)
