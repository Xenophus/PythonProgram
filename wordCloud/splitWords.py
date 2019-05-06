# coding = utf-8

import jieba
import codecs
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
from collections import Counter
from wordcloud import WordCloud, ImageColorGenerator

#打开评论文件，获得评论内容并对评论进行分词
file = codecs.open('luhan.txt', 'r', 'utf-8')
describe = file.read()
describeGenerator = jieba.cut(describe, cut_all=False)

newDescribe = []
file1 = codecs.open('stopWords.txt', 'r', 'utf-8')
stopWords = file1.read()

#去掉评论内容中的停用词，停用词表使用的是哈工大的
for word in describeGenerator:
                                 #手工去掉部分词语
    if word not in stopWords and word not in ['M', '想', '都', '\xa0', '做', '完']:
        newDescribe.append(word)

#统计词频
counts = Counter(newDescribe)
top20 = counts.most_common(20)
# print(top20)  #输出检查

词频展示
mask = np.array(Image.open('Xenophus.png'))   #用于展示的png文件，可替换
wc = WordCloud(background_color='white', max_words=50,font_path="C:/Windows/Fonts/simfang.ttf",min_font_size=15,max_font_size=100, mask=mask)
wc.generate_from_frequencies(counts)
image_colors = ImageColorGenerator(mask)
wc.recolor(color_func=image_colors)
plt.imshow(wc)
plt.axis('off')
plt.show()
wc.to_file("pic.png")
