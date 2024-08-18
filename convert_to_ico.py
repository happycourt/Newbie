from PIL import Image

# 打开PNG文件
img = Image.open("word-extractor-icon.png")

# 保存为ICO文件，包含多个常用尺寸
img.save("word-extractor-icon.ico", format='ICO', sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
