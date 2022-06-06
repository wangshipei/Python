import re


def zh_format(html):
    # 去掉半角方括号
    pttn = r'\[(.*?)\]'
    rpl = r'\1'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 直双引号转换成弯双引号
    pttn = r'\s*"(.*?)\s*"'
    rpl = r'“\1”'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 直单引号转换成弯单引号
    pttn = r"\s*'(.*?)\s*'"
    rpl = r'‘\1’'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # html tag 中被误伤的直引号
    pttn = r'=[“”"](.*?)[“”"]'
    rpl = r'="\1"'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # html 弯引号之前的空格
    pttn = r'([\u4e00-\u9fa5])([“‘])'
    rpl = r'\1 \2'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # html 弯引号之后的空格
    pttn = r'([’”])([\u4e00-\u9fa5])'
    rpl = r'\1 \2'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # html tag: <i>, <em> 转换成 <strong>
    pttn = r'(<i|<em)'
    rpl = r'<strong'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # html tag: <i>, <em> 转换成 <strong>
    pttn = r'(i>|em>)'
    rpl = r'strong>'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # html tag: strong 内部的 “”、‘’、《》、（）
    pttn = r'([《（“‘]+)<strong (.*?)>'
    rpl = r'<strong \2>\1'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    pttn = r'</strong>([》）”’。，]+)'
    rpl = r'\1</strong>'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 省略号
    pttn = r'\.{2,}\s*。*\s*'
    rpl = r'…… '
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 破折号
    pttn = r'&mdash；|&mdash;|--'
    rpl = r' —— '
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 姓名之间的 ·（重复三次）
    pttn = r'([\u4e00-\u9fa5])-([\u4e00-\u9fa5])'
    rpl = r'\1·\2'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    pttn = r'([\u4e00-\u9fa5])-([\u4e00-\u9fa5])'
    rpl = r'\1·\2'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    pttn = r'([\u4e00-\u9fa5])-([\u4e00-\u9fa5])'
    rpl = r'\1·\2'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    pttn = r'([A-Z]{1})\s*\.\s*([A-Z]{1})'
    rpl = r'\1·\2'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    pttn = r'([A-Z]{1})\s*\.\s*([\u4e00-\u9fa5])'
    rpl = r'\1·\2'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 全角百分号
    pttn = r'％'
    rpl = r'%'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 数字前的空格
    pttn = r'([\u4e00-\u9fa5])(\d)'
    rpl = r'\1 \2'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 数字后的空格，百分比 % 后的空格
    pttn = r'([\d%])([\u4e00-\u9fa5])'
    rpl = r'\1 \2'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 英文字母前的空格
    pttn = r'([\u4e00-\u9fa5])([a-zA-Z])'
    rpl = r'\1 \2'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 英文字母后的空格，百分比 % 后的空格
    pttn = r'([a-zA-Z])([\u4e00-\u9fa5])'
    rpl = r'\1 \2'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # tag 内的英文字母前的空格
    pttn = r'([\u4e00-\u9fa5])<(strong|i|em|span)>(.[a-zA-Z\d ]*?)<\/(strong|i|em|span)>'
    rpl = r'\1 <\2>\3</\4>'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # tag 内的英文字母后的空格，百分比 % 后的空格
    pttn = r'<(strong|i|em|span)>(.[a-zA-Z\d ]*?)<\/(strong|i|em|span)>([\u4e00-\u9fa5])'
    rpl = r'<\1>\2</\3> \4'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 弯引号前的逗号
    pttn = r'，([”’])'
    rpl = r'\1，'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 中文标点符号之前多余的空格
    pttn = r'([，。！？》〉】]) '
    rpl = r'\1'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 英文句号 . 与汉字之间的空格
    pttn = r'\.([\u4e00-\u9fa5])'
    rpl = r'. \1'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 左半角括号
    pttn = r'\s*\('
    rpl = r'（'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 右半角括号
    pttn = r'\)\s*'
    rpl = r'）'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    # 多余的括号（DeepL 返回文本经常出现的情况）
    pttn = r'）。）'
    rpl = r'。）'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)

    return html


path = "John Law/"  # 文件夹名称末尾得有 /
source_filename = "index3.html"  # 上一步生成的文件，成为这一步的 “源文件”
target_filename = "index4.html"

lines = open(path + source_filename, "r").readlines()

new_lines = []
for line in lines:
    if 'cn"><img ' in line:
        # 这个 if 不是通用的……
        # 是因为示例文件不知道为什么，有 img 的行未翻译但重复存在
        # 这个 if block 可注释掉……
        continue
    if ' cn"' in line:
        new_lines.append(zh_format(line))
    else:
        new_lines.append(line)

final_text = "".join(new_lines)

# 去掉多余的空行
pttn = r'\n\n\n'
rpl = r'\n\n'
re.findall(pttn, final_text)
final_text = re.sub(pttn, rpl, final_text)

# 写入文件
with open(path + target_filename, 'w') as f:
    f.write("".join(new_lines))