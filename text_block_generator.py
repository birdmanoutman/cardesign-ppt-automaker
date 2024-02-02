import os


def text_blocks_generator(folderPath='./inputTextFolder', blockLength=7000):
    fileNameList = os.listdir(folderPath)
    text_string = ''
    for filename in fileNameList:
        if filename.endswith('txt'):
            filePath = os.path.join(folderPath, filename)
            print(filePath)
            with open(filePath, 'r', encoding='utf-8') as f:
                text = f.read().strip()
            text_string = text_string + text
    print(len(text_string))

    block_size = blockLength # 按照长度7000进行分割。
    text_blocks = [] # 创建一个空的列表来存储分割后的文本块。

    for i in range(0, len(text_string), block_size):
        text_block = text_string[i:i+block_size] # 截取7000个字符。
        text_blocks.append(text_block) # 将文本块添加到列表中。

    return text_blocks
