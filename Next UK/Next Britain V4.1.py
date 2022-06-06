import time  # 引入time，这个库（有的叫‘包’，有的叫‘模块’）主要用于等待浏览器响应时间或者等待网页内容加载
import re  # 这个库主要用于正则表达式
from selenium import webdriver  # 这个库主要用于控制浏览器（建议使用谷歌浏览器Chrome，所有浏览器中Chrome是最稳定，速度最快的）
from selenium.webdriver.common.by import By  # 这个库主要用于查找网页内的元素
from selenium.common.exceptions import NoSuchElementException  # 这个库主要用于报错处理
import pandas as pd  # 这个库主要用于处理excel表格
from tqdm import tqdm  # 这个库主要用于显示进度条


def next_britain(name, *urlplus):
    a1pttn = re.compile(r'-(\w*)-0')
    a1 = [re.findall(a1pttn, x1) for x1 in urlplus]
    # a1 = [re.sub(r'-(.*)=', '', x1) for x1 in urlplus]
    print(f'要爬取的所有系列为：Next Britain-{name} ', '\n', a1[0:5], '\n', a1[5:10], '\n', a1[10:15], '\n', a1[15:20])

    for a, a2 in zip(urlplus, a1):
        baseurl = 'https://www3.next.co.uk/shop/'
        url = baseurl + str(a)
        driver = webdriver.Chrome()
        driver.get(url)
        driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
        time.sleep(2)
        driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
        time.sleep(2)
        driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
        time.sleep(2)
        driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
        time.sleep(2)
        driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
        time.sleep(2)
        driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
        time.sleep(2)
        driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
        a3 = ''.join(a2)
        pros = driver.find_elements(By.XPATH, '//a[@class="Image"]')

        pic1s = []
        titles = []
        prices = []
        colors = []
        sizes = []
        compositions = []
        descriptions = []
        pic2s = []
        pic3s = []
        pic4s = []
        pic5s = []

        links = []
        pic1links = []
        pic2links = []
        pic3links = []
        pic4links = []
        pic5links = []

        for i in tqdm(pros, position=0, leave=True,
                      desc=f"正在爬取Next Britain-{name}-" + str(a3)):
            slink = i.get_attribute('href')
            try:
                driver.execute_script("window,open('');")
                driver.switch_to.window(driver.window_handles[1])
                driver.get(slink)

                try:
                    link = driver.find_element(By.XPATH, '//link[@rel="canonical"]').get_attribute('href')

                except NoSuchElementException:
                    link = ''

                try:
                    title = driver.find_element(By.XPATH, '//div[@class="Title"]').text
                except NoSuchElementException:
                    title = ''

                try:
                    price = driver.find_element(By.XPATH, '//div[@class="nowPrice branded-markdown"]').text
                except NoSuchElementException:
                    price = ''

                try:
                    color = driver.find_element(By.XPATH, '//span[@class="colourChipNameLabel"]').text
                except NoSuchElementException:
                    color = ''

                try:
                    description = driver.find_element(By.XPATH, '//div[@id="ToneOfVoice"]').text
                except NoSuchElementException:
                    description = ''

                try:
                    size = ''.join(re.findall(r'\(.*\)', title))
                except NoSuchElementException:
                    size = ''

                try:
                    composition = driver.find_element(By.ID, 'Composition').text

                except NoSuchElementException:
                    composition = ''

                piclink1 = driver.find_element(By.XPATH, '//div[@class="ThumbNailNavClip"]').get_attribute('innerHTML')
                picpttn = re.compile(r'rel="https:.*.jpg"')
                piclink2 = re.findall(picpttn, piclink1)
                piclink3 = [re.sub(r'rel="', '', x1) for x1 in piclink2]
                piclink = [re.sub(r'"', '', x2) for x2 in piclink3]

                try:
                    pic1link = ''.join(piclink[0])
                except NoSuchElementException:
                    pic1link = ''
                except IndexError:
                    pic1link = ''

                try:
                    pic2link = ''.join(piclink[1])
                except NoSuchElementException:
                    pic2link = ''
                except IndexError:
                    pic2link = ''

                try:
                    pic3link = ''.join(piclink[1])
                except NoSuchElementException:
                    pic3link = ''
                except IndexError:
                    pic3link = ''

                try:
                    pic4link = ''.join(piclink[3])
                except NoSuchElementException:
                    pic4link = ''
                except IndexError:
                    pic4link = ''

                try:
                    pic5link = ''.join(piclink[4])
                except NoSuchElementException:
                    pic5link = ''
                except IndexError:
                    pic5link = ''

                pic1s.append('')
                titles.append(title)
                prices.append(price)
                colors.append(color)
                sizes.append(size)
                compositions.append(composition)
                descriptions.append(description)
                pic2s.append('')
                pic3s.append('')
                pic4s.append('')
                pic5s.append('')
                links.append(link)
                pic1links.append(pic1link)
                pic2links.append(pic2link)
                pic3links.append(pic3link)
                pic4links.append(pic4link)
                pic5links.append(pic5link)

                driver.close()
                driver.switch_to.window(driver.window_handles[0])

                data = pd.DataFrame()

                data['Pic1'] = pic1s
                data['Title'] = titles
                data['Price'] = prices
                data['Color'] = colors
                data['Size'] = sizes
                data['Composition'] = compositions
                data['Description'] = descriptions
                data['Pic-2'] = pic2s
                data['Pic-3'] = pic3s
                data['Pic-4'] = pic4s
                data['Pic-5'] = pic5s
                data['Link'] = links
                data['Pic-1Link'] = pic1links
                data['Pic-2Link'] = pic2links
                data['Pic-3Link'] = pic3links
                data['Pic-4Link'] = pic4links
                data['Pic-5Link'] = pic5links

                data.to_excel('Next/Next Britain-' + str(name) + '-' + str(a3) + '.xlsx')

            except NoSuchElementException:
                print('此网页爬取失败：' + '\n' + slink + '\n' + '后续网页需手动处理')
                link = ''
                title = ''
                price = ''
                color = ''
                description = ''
                size = ''
                composition = ''
                pic1link = ''
                pic2link = ''
                pic3link = ''
                pic4link = ''
                pic5link = ''
                pic1s.append('')
                titles.append(title)
                prices.append(price)
                colors.append(color)
                sizes.append(size)
                compositions.append(composition)
                descriptions.append(description)
                pic2s.append('')
                pic3s.append('')
                pic4s.append('')
                pic5s.append('')
                links.append(link)
                pic1links.append(pic1link)
                pic2links.append(pic2link)
                pic3links.append(pic3link)
                pic4links.append(pic4link)
                pic5links.append(pic5link)

                driver.close()
                driver.switch_to.window(driver.window_handles[0])

                data = pd.DataFrame()

                data['Pic1'] = pic1s
                data['Title'] = titles
                data['Price'] = prices
                data['Color'] = colors
                data['Size'] = sizes
                data['Composition'] = compositions
                data['Description'] = descriptions
                data['Pic-2'] = pic2s
                data['Pic-3'] = pic3s
                data['Pic-4'] = pic4s
                data['Pic-5'] = pic5s
                data['Link'] = links
                data['Pic-1Link'] = pic1links
                data['Pic-2Link'] = pic2links
                data['Pic-3Link'] = pic3links
                data['Pic-4Link'] = pic4links
                data['Pic-5Link'] = pic5links

                data.to_excel('Next/Next Britain-' + str(name) + '-' + str(a3) + '.xlsx')

        driver.close()
        print('Next Britain' + '-' + str(name) + str(a3) + '爬取完毕！')

    print('Next Britain - ' + str(name) + '所有系列爬取完毕！')


girlname = 'Girl'
girlurlplus = [
    # "gender-newborngirls-gender-newbornunisex-gender-youngergirls-category-rompersuits-category-sleepsuits-0",
    # "gender-newborngirls-gender-newbornunisex-category-bodysuits-0",
    # "gender-newborngirls-gender-newbornunisex-gender-oldergirls-gender-youngergirls-productaffiliation-coatsandjackets-0",
    # "gender-newborngirls-gender-oldergirls-gender-youngergirls-category-dresses-0",
    # "gender-newborngirls-gender-oldergirls-gender-youngergirls-category-jeans-0",
    # "gender-newborngirls-gender-oldergirls-gender-youngergirls-productaffiliation-jumpsuitsandplaysuits-0",
    # "gender-newborngirls-gender-newbornunisex-gender-oldergirls-gender-youngergirls-productaffiliation-knitwear-0",
    # "gender-oldergirls-gender-youngergirls-productaffiliation-nightwear-0",
    # "gender-newborngirls-gender-newbornunisex-gender-oldergirls-gender-youngergirls/use-bridesmaid-use-flowergirl-use-occasionwear-use-partywear-0",
    # "gender-newborngirls-gender-newbornunisex-gender-oldergirls-gender-youngergirls-productaffiliation-outfits-0",
    # "gender-newborngirls-gender-newbornunisex-gender-oldergirls-gender-youngergirls-productaffiliation-tops/category-blouses-category-shirts-0",
    # "gender-newborngirls-gender-oldergirls-gender-youngergirls-productaffiliation-shortsandskirts-0",
    # "department-homeware-category-sleepbag-0",
    # "gender-newborngirls-gender-newbornunisex-gender-oldergirls-gender-youngergirls-productaffiliation-hosieryandsocks-0",
    # "gender-oldergirls-gender-youngergirls-productaffiliation-sportswear-0",
    # "gender-newborngirls-gender-newbornunisex-gender-oldergirls-gender-youngergirls-productaffiliation-sweatshirtsandhoodies-0",
    "gender-newborngirls-gender-oldergirls-gender-youngergirls-productaffiliation-swimwear-0",
    # "gender-newborngirls-gender-newbornunisex-gender-oldergirls-gender-youngergirls-productaffiliation-tops/category-tshirts-0",
    # "gender-newborngirls-gender-newbornunisex-gender-oldergirls-gender-youngergirls-productaffiliation-tops-0",
    # "gender-newborngirls-gender-newbornunisex-gender-oldergirls-gender-youngergirls-productaffiliation-trousersleggings-0",
    # "gender-oldergirls-gender-youngergirls-productaffiliation-underwear-0",
    # "gender-newborngirls-gender-newbornunisex-gender-oldergirls-gender-youngergirls-productaffiliation-accessories/category-bags-0",
    # "gender-newborngirls-gender-newbornunisex-gender-oldergirls-gender-youngergirls-productaffiliation-accessories/category-hairaccessories-0",
    "gender-newborngirls-gender-newbornunisex-gender-oldergirls-gender-youngergirls-productaffiliation-hatsglovesscarves-0"]
next_britain(girlname, *girlurlplus)

boyname = 'Boy'
boyurlplus = ["gender-newbornboys-gender-newbornunisex-gender-youngerboys-category-rompersuits-category-sleepsuits-0",
              "gender-newbornboys-gender-newbornunisex-category-bodysuits-0",
              "gender-newbornboys-gender-newbornunisex-gender-olderboys-gender-youngerboys-productaffiliation-coatsandjackets-0",
              "gender-newbornboys-gender-olderboys-gender-youngerboys-category-jeans-0",
              "gender-newbornboys-gender-newbornunisex-gender-olderboys-gender-youngerboys-productaffiliation-joggers-0",
              "gender-newbornboys-gender-newbornunisex-gender-olderboys-gender-youngerboys-productaffiliation-knitwear-0",
              "gender-newbornboys-gender-newbornunisex-gender-olderboys-gender-youngerboys/use-occasionwear-use-pageboy-0",
              "gender-newbornboys-gender-newbornunisex-gender-olderboys-gender-youngerboys-productaffiliation-outfits-0",
              "gender-newbornboys-gender-newbornunisex-gender-olderboys-gender-youngerboys-productaffiliation-shirts-0",
              "gender-newbornboys-gender-newbornunisex-gender-olderboys-gender-youngerboys-productaffiliation-shorts-0",
              "department-homeware-category-sleepbag-0",
              "gender-newbornboys-gender-newbornunisex-gender-olderboys-gender-youngerboys-category-socks-0",
              "gender-olderboys-gender-youngerboys-productaffiliation-sportswear-0",
              "gender-olderboys-gender-youngerboys-productaffiliation-boyssuits-0",
              "gender-newbornboys-gender-newbornunisex-gender-olderboys-gender-youngerboys-productaffiliation-sweatshirtsandhoodies-0",
              "gender-newbornboys-gender-olderboys-gender-youngerboys-productaffiliation-swimwear-0",
              "gender-newbornboys-gender-newbornunisex-gender-olderboys-gender-youngerboys-productaffiliation-tops/category-tshirts-0",
              "gender-newbornboys-gender-newbornunisex-gender-olderboys-gender-youngerboys-productaffiliation-tops-0",
              "gender-olderboys-gender-youngerboys-productaffiliation-outfits/style-coord-0",
              "gender-newbornboys-gender-newbornunisex-gender-olderboys-gender-youngerboys-productaffiliation-trousers-0",
              "gender-olderboys-gender-youngerboys-productaffiliation-underwear-0",
              "gender-newbornboys-gender-newbornunisex-gender-olderboys-gender-youngerboys-productaffiliation-accessories/category-bags-0",
              "gender-newbornboys-gender-newbornunisex-gender-olderboys-gender-youngerboys-productaffiliation-hatsglovesscarves-0",
              "gender-newbornboys-gender-newbornunisex-gender-olderboys-gender-youngerboys-productaffiliation-accessories-productaffiliation-hatsglovesscarves/category-ties-0"]

next_britain(boyname, *boyurlplus)
# 男童

babyname = 'Baby'
babyurlplus = [
    "gender-newbornboys-gender-newborngirls-gender-newbornunisex-gender-youngerboys-gender-youngergirls/category-jeans-0",
    "gender-newbornboys-gender-newborngirls-gender-newbornunisex-gender-youngerboys-gender-youngergirls/category-joggers-0",
    "gender-newbornboys-gender-newborngirls-gender-newbornunisex-gender-youngerboys-gender-youngergirls-productaffiliation-knitwear-productaffiliation-sweatshirtsandhoodies-0",
    "gender-newbornboys-gender-newborngirls-gender-newbornunisex-gender-youngerboys-gender-youngergirls/use-bridesmaid-use-flowergirl-use-occasionwear-use-pageboy-0",
    "gender-newbornboys-gender-newborngirls-gender-newbornunisex-gender-youngerboys-gender-youngergirls/category-dungarees-category-rompersuits-0",
    "gender-newbornboys-gender-newborngirls-gender-newbornunisex-gender-youngerboys-gender-youngergirls-productaffiliation-outfits-0",
    "gender-newbornboys-gender-newborngirls-gender-newbornunisex-gender-youngerboys-gender-youngergirls/category-shorts-0",
    "category-sleepbag-0",
    "gender-newbornboys-gender-newborngirls-gender-newbornunisex-gender-youngerboys-gender-youngergirls/category-socks-category-tights-0",
    "gender-newbornboys-gender-newborngirls-gender-newbornunisex-gender-youngerboys-gender-youngergirls-productaffiliation-swimwear-0",
    "gender-newbornboys-gender-newborngirls-gender-newbornunisex-gender-youngerboys-gender-youngergirls-productaffiliation-tops-0",
    "gender-newbornboys-gender-newborngirls-gender-newbornunisex-gender-youngerboys-gender-youngergirls-productaffiliation-trousersleggings-0",
    "gender-newbornboys-gender-newborngirls-gender-newbornunisex-productaffiliation-accessories/category-bags-0",
    "gender-newbornboys-gender-newborngirls-gender-newbornunisex-gender-youngerboys-gender-youngergirls-productaffiliation-accessories/category-bibs-0",
    "gender-newbornboys-gender-newborngirls-gender-newbornunisex-gender-youngerboys-gender-youngergirls-productaffiliation-hatsglovesscarves-0"]

next_britain(babyname, *babyurlplus)
# 婴儿
