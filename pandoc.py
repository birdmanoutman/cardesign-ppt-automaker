import pypandoc

def convert_markdown_to_word(markdown_file, output_file):
    pypandoc.convert_file(markdown_file, 'docx', outputfile=output_file)

# 调用函数进行Markdown转换

path =r'/Users/birdmanoutman/Desktop/荣威D7上车体感知纪要整理.md'
convert_markdown_to_word(path, 'output.docx')