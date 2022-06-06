import bs4
import re

path = "John Law/"  # 文件夹名称末尾得有 /
source_filename = "index.html"
target_filename = "index2.html"

html = open(path + source_filename)
htmltext = html.read()

soup = bs4.BeautifulSoup(htmltext)

# 将所有的 \n 去掉……
htmltext = str(bs4.BeautifulSoup(htmltext)).replace("\n", "")

# <h... 之前添加空行
pttn = r'<h'
rpl = r'\n\n<h'
re.findall(pttn, htmltext)
htmltext = re.sub(pttn, rpl, htmltext)

# <div... 之前添加空行
pttn = r'<div'
rpl = r'\n\n<div'
re.findall(pttn, htmltext)
htmltext = re.sub(pttn, rpl, htmltext)

# </div> 之前添加空行
pttn = r'</div>'
rpl = r'\n\n</div>'
re.findall(pttn, htmltext)
htmltext = re.sub(pttn, rpl, htmltext)

# <p... 之前添加空行
pttn = r'<p'
rpl = r'\n\n<p'
re.findall(pttn, htmltext)
htmltext = re.sub(pttn, rpl, htmltext)

fileSave = open(path + target_filename, "w")
fileSave.write(htmltext)
print(htmltext)
