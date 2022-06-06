import re
import requests

auth_key = "<your DeepL API Pro authentication key>"  # 注意，要订阅的是 DeepL API Pro
target_language = "ZH"  # 当然，你可以将目标语言设置成任何 DeepL 支持的语言

path = "John Law/"  # 文件夹名称末尾得有 /
source_filename = "index2.html"  # 上一步生成的文件，成为这一步的 “源文件”
target_filename = "index3.html"


def translate(text):
    result = requests.get(
        "https://api.deepl.com/v2/translate",
        params={
            "auth_key": auth_key,
            "target_lang": target_language,
            "text": text,
            "tag_handling": "xml",  # 这个参数确保 DeepL 正确处理 html tags
        },
    )
    return result.json()["translations"][0]["text"]


def add_language_tag_en(html):
    pttn = re.compile(r'^<(.*?) class="(.*?)">', re.M)
    rpl = r'<\1 class="\2 en">'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)
    return html


def add_language_tag_cn(html):
    pttn = re.compile(r'^<(.*?) class="(.*?)">', re.M)
    rpl = r'<\1 class="\2 cn">'
    re.findall(pttn, html)
    html = re.sub(pttn, rpl, html)
    return html


lines = open(path + source_filename, "r").readlines()

new_lines = []
line_count = 0
startline = 16
endline = 4032

for line in lines:
    line_count += 1
    if line_count < startline or line_count > endline or line.strip() == '':
        new_lines.append(line)
        print(line)
        continue

    succeeded = False
    while not succeeded:
        # 以下比较粗暴的 try... except，用来防止执行过程中出现 DeepL 连接错误而导致翻译任务中断……
        try:
            line_translated = translate(line)
            # 以下一行确保将返回的字符串转换成一整行，而非含有 \n 的多行文本
            line_translated = line_translated.replace("\n", "")

            succeeded = True
        except:
            succeeded = False

    if line.strip() == line_translated.strip():
        # 返回的字符串与原字符串相同，说明 html tag 之间的内容无需翻译
        new_lines.append(line)
        print(line)
    else:
        line = add_language_tag_en(line)
        line_translated = add_language_tag_cn(line_translated)
        new_lines.append(line)
        print(line)
        new_lines.append(line_translated)
        print(line_translated)

with open(path + target_filename, 'w') as f:
    f.write("\n".join(new_lines))
