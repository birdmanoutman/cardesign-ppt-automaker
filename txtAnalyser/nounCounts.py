# 输入文本链接，返回这个文本的名次词频
import jieba.posseg as pseg
from collections import Counter
from docx import Document
import re
import jieba

filePath = r'/Users/birdmanoutman/Library/CloudStorage/GoogleDrive-birdmanoutman@gmail.com/其他计算机/904Dell/sprintProjects/上海汽车论文投稿/设计战略在全球化汽车行业中的应用研究：以上汽集团MG品牌为案例 v2 简化.docx'


def read_text(filePath):
    if filePath.endswith('txt'):
        # 将您的文本替换到这里
        with open(filePath, 'r', encoding='utf-8') as file:
            text = file.read()
        return text
    elif filePath.endswith('docx'):
        doc = Document(filePath)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return '\n'.join(full_text)

text = read_text(filePath)

# 向jieba词典中添加新词MG，并指定为名词
jieba.add_word('MG', tag='n')
jieba.add_word('设计战略', tag='n')
jieba.add_word('荣威', tag='n')
jieba.add_word('飞凡', tag='n')
jieba.add_word('智己', tag='n')

# 使用jieba进行中文分词和词性标注
words = pseg.cut(text)
# 提取中文名词
cn_nouns = [word for word, flag in words if flag.startswith('n')]

# 使用正则表达式匹配英文单词
en_words = re.findall(r'\b[A-Za-z]+\b', text)
# print(en_words)

# 假设所有英文单词都算作名词（这可能不完全准确）
en_nouns = en_words

# 合并中文名词和英文名词列表
all_nouns = cn_nouns + en_nouns

# 统计名词出现频率
nouns_freq = Counter(all_nouns)

# 获取出现频率最高的20个名词
top_20_nouns = nouns_freq.most_common(30)

# 打印结果
for noun, freq in top_20_nouns:
    print(f"{noun}: {freq}")