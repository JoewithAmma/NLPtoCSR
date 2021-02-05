
import matplotlib.pyplot as plt
from wordcloud import WordCloud
import jieba

f=open('台電.txt','r',encoding = 'UTF-8')

text = f.read()
#结巴分词
wordlist = jieba.cut(text,cut_all=True)
wl = " ".join(wordlist)
#print(wl)#输出分词之后的txt
 
 
#把分词后的txt写入文本文件
#fenciTxt  = open("fenciHou.txt","w+")
#fenciTxt.writelines(wl)
#fenciTxt.close()
 
 
#设置词云
wc = WordCloud(font_path = "simsun.ttf")
myword = wc.generate(wl)#生成词云
 
#展示词云图
plt.imshow(myword)
plt.axis("off")
plt.show()






"""
filename = "華碩.txt"

f=open('台電.txt','r',encoding = 'UTF-8')

r=f.read()

import jieba
r = " ".join(jieba.cut(r))
print(r)
from wordcloud import WordCloud
wordcloud = WordCloud(font_path="simsun.ttf").generate(r)

import matplotlib.pyplot as plt
plt.imshow(wordcloud, interpolation='bilinear')
plt.axis("off")
"""