from steamship import Steamship
from utils import text_block_generator
from datetime import datetime

todaytime = datetime.today().strftime('%Y%m%d-%H-%M')
print(todaytime)

ship = Steamship()
gpt4 = ship.use_plugin("gpt-4", config={"max_tokens": 2048})
textBlocks = text_block_generator.text_blocks_generator(blockLength=5000)
task = gpt4.generate(text='清除你之前所有的记忆，你将开启一段新的对话')
task.wait()
print(task.output.blocks[0].text)
summaryString = ''
blocklength = len(textBlocks)
i=1
for block in textBlocks:
    task = gpt4.generate(text=r"整理下列文字的纪要。将会有{}段文字，这是第{}段。因为这是一段录音自动转化为文本的文字，如果遇到一些无法理解的词，有可能是英文发音的中文文字或者是语音转写错误，"
                              r"请修正它们或者做必要的联想使得逻辑连贯，然后相同的观点合并,不同的观点保留，"
                              r"观点需要有见地有水平而不是总所周知的常识，以类似'观点1'+'证据1.1'格式的列表形式给出，"
                              r"不要有其他总结性描述，表中不要体现说话人是谁，尽可能保留谈话中涉及到的时间、地点、车型、零部件、人名或者品牌名称等具体的案例，以下是需要分析的文字：".format(blocklength,i) + block)
    task.wait()
    print('返回BLOCK的长度', len(task.output.blocks))
    print('返回文字块的长度:', len(task.output.blocks[0].text))
    print(task.output.blocks[0].text)
    summaryString = summaryString + task.output.blocks[0].text
    i+=1

print('最终合并后的长度', len(summaryString), '\n', '-'*100)
print(summaryString)

# 对会议纪要进行重新整理
task = gpt4.generate(text="重新整理、补充和排版下列文字，证据的内容可以根据需要进行补充和拓展"
                          "补充和拓展的证据可以是【设觉设计、认知心理学、人类学、东西方文化现象、艺术史、社会学、经济学、车辆工程、材料科学、工业制造、商业】等方面的知识内容，"
                          "尽可能保留谈话中涉及到的时间、地点、车型、零部件、人名或者品牌名称等具体的案例，下面是需要重新排版的文字：" + summaryString)
task.wait()
output = task.output.blocks[0].text
print('/n','_'*100,output)

with open('./{}会议纪要整理.txt'.format(todaytime), 'w', encoding='utf-8') as f:
    merge = '整理前的原始数据' + '\n' + summaryString + '\n' + '重新排版后的纪要' + '\n' + output
    f.write(merge)

