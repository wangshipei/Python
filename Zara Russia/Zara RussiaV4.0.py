import time
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
from tqdm import tqdm


def zara_russia(name, *urlplus):
    a1pttn = re.compile(r'kids-(.*)-.*=')
    a1 = [re.findall(a1pttn, x1) for x1 in urlplus]
    # a1 = [re.sub(r'-(.*)=', '', x1) for x1 in urlplus]
    print(f'要爬取的所有系列为：Zara Russia-', '\n', a1[0:5], '\n', a1[5:11], '\n', a1[11:20])

    for a, a2 in zip(urlplus, a1):
        baseurl = 'https://www.zara.com/ru/en/'
        url = baseurl + str(a)
        driver = webdriver.Chrome()
        driver.get(url)
        driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
        time.sleep(2)
        driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
        time.sleep(2)
        driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
        time.sleep(2)
        a3 = ''.join(a2)
        pros = driver.find_elements(By.XPATH, '//a[@class="product-link product-grid-product__link link"]')

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

        for i in tqdm(pros, position=0, leave=True, desc="正在爬取Zara Russia-" + str(a3)):
            slink = i.get_attribute('href')
            try:
                driver.execute_script("window,open('');")
                driver.switch_to.window(driver.window_handles[1])
                driver.get(slink)

                try:
                    link = driver.find_element(By.NAME, "twitter:url").get_attribute('content')

                except NoSuchElementException:
                    link = ''

                try:
                    title = driver.find_element(By.TAG_NAME, 'h1').text
                except NoSuchElementException:
                    title = ''

                try:
                    price = driver.find_element(By.XPATH, '//span[@class="price-current__amount"]').text
                except NoSuchElementException:
                    price = ''

                try:
                    color1 = driver.find_element(By.XPATH,
                                                 '//p[@class="product-detail-selected-color product-detail-info__color"]').text
                    color = ''.join(re.findall(r'(.*) \|', color1))
                except NoSuchElementException:
                    color = ''

                try:
                    description = driver.find_element(By.TAG_NAME, 'p').text
                except NoSuchElementException:
                    description = ''

                try:
                    size = driver.find_element(By.XPATH, '//span[@class="product-detail-size-info__main-label"]').text
                except NoSuchElementException:
                    size = ''

                try:
                    composition11 = driver.find_elements(By.CLASS_NAME, 'structured-component-text-block-paragraph')[2]
                    composition12 = re.findall(r'\d*% .*', composition11.get_attribute('innerText'))
                    composition1 = ''.join(composition12)

                    composition21 = driver.find_elements(By.CLASS_NAME, 'structured-component-text-block-paragraph')[3]
                    composition22 = re.findall(r'\d*% .*', composition21.get_attribute('innerText'))
                    composition2 = ''.join(composition22)

                    composition31 = driver.find_elements(By.CLASS_NAME, 'structured-component-text-block-paragraph')[4]
                    composition32 = re.findall(r'\d*% .*', composition31.get_attribute('innerText'))
                    composition3 = ''.join(composition32)

                    composition41 = driver.find_elements(By.CLASS_NAME, 'structured-component-text-block-paragraph')[4]
                    composition42 = re.findall(r'\d*% .*', composition41.get_attribute('innerText'))
                    composition4 = ''.join(composition42)

                    composition = composition1 + '\n' + composition2 + '\n' + composition3 + '\n' + composition4

                except NoSuchElementException:
                    composition = ''
                except IndexError:
                    composition = ''

                try:
                    pic1link = ''.join(re.findall(r'(https://.*) 563w',
                                                  driver.find_elements(By.TAG_NAME, 'source')[0].get_attribute(
                                                      'srcset')))
                except NoSuchElementException:
                    pic1link = ''
                except IndexError:
                    pic1link = ''

                piclinks = driver.find_elements(By.TAG_NAME, 'source')
                if len(piclinks) >= 1:
                    try:
                        pic2link = ''.join(re.findall(r'(https://.*) 563w', piclinks[2].get_attribute('srcset')))
                    except IndexError:
                        pic2link = ''
                else:
                    pic2link = ''

                if len(piclinks) >= 2:
                    try:
                        pic3link = ''.join(re.findall(r'(https://.*) 563w', piclinks[4].get_attribute('srcset')))
                    except IndexError:
                        pic3link = ''
                else:
                    pic3link = ''

                if len(piclinks) >= 4:
                    try:
                        pic4link = ''.join(re.findall(r'(https://.*) 563w', piclinks[6].get_attribute('srcset')))
                    except IndexError:
                        pic4link = ''
                else:
                    pic4link = ''

                if len(piclinks) >= 6:
                    try:
                        pic5link = ''.join(re.findall(r'(https://.*) 563w', piclinks[8].get_attribute('srcset')))
                    except IndexError:
                        pic5link = ''
                else:
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
                data.to_excel('data/ZARA Russia' + '-' + str(a3) + '.xlsx')

            except NoSuchElementException:
                print('此网页爬取失败：' + '\n' + slink + '\n' + '后续网页需手动处理')
                break

        driver.close()
        print('ZARA Russia' + '-' + str(a3) + '爬取完毕！')

    print('ZARA Russia - ' + str(name) + '所有系列爬取完毕！')


girlname = 'Girl'
girlurlplus = ['kids-girl-outerwear-l394.html?v1=2013735&page=',
               'kids-girl-sweatshirts-l430.html?v1=2019820&page=',
               'kids-girl-tshirts-l450.html?v1=2013799&page=',
               'kids-girl-dresses-l360.html?v1=2019361&page=',
               'kids-girl-trousers-l439.html?v1=2013810&page=',
               'kids-girl-jeans-l378.html?v1=2013805&page=',
               'kids-girl-knitwear-l385.html?v1=2013765&page=',
               'kids-girl-shirts-l401.html?v1=2013743&page=',
               'kids-girl-skirts-l425.html?v1=2013750&page=',
               'kids-girl-basics-l348.html?v1=2019368&page=',
               'kids-girl-license-l2953.html?v1=2019474&page=',
               'kids-girl-looks-l388.html?v1=2013768&page=',
               'kids-girl-sporty-l1588.html?v1=2019450&page=',
               'kids-girl-beachwear-l350.html?v1=2013740&page=',
               'kids-girl-bags-l346.html?v1=2019336&page=',
               'kids-girl-underwear-l469.html?v1=2019353&page=',
               'kids-girl-accessories-l326.html?v1=2019346&page=']
zara_russia(girlname, *girlurlplus)
# 女童

boyname = 'Boy'
boyurlplus = ['kids-boy-outerwear-l231.html?v1=2019972&page=',
              'kids-boy-sweatshirts-l267.html?v1=2020490&page=',
              'kids-boy-tshirts-l286.html?v1=2019941&page=',
              'kids-boy-trousers-l274.html?v1=2019948&page=',
              'kids-boy-jeans-l218.html?v1=2019944&page=',
              'kids-boy-knitwear-l223.html?v1=2019983&page=',
              'kids-boy-shirts-l239.html?v1=2019980&page=',
              'kids-boy-suits-l266.html?v1=2020081&page=',
              'kids-boy-basics-l199.html?v1=2020059&page=',
              'kids-boy-license-l2954.html?v1=2019999&page=',
              'kids-boy-trend-8-l2355.html?v1=2020146&page=',
              'kids-boy-total-look-l4106.html?v1=2020006&page=',
              'kids-boy-bermudas-denim-l4438.html?v1=2020016&page=',
              'kids-boy-edited-l3185.html?v1=2020065&page=',
              'kids-boy-backpacks-l197.html?v1=2020467&page=',
              'kids-boy-underwear-l304.html?v1=2019968&page=',
              'kids-boy-accessories-l176.html?v1=2020029&page=']

zara_russia(boyname, *boyurlplus)
# 男童

babyboyname = 'Babyboy'
babyboyurlplus = [  # 'kids-babyboy-outerwear-l47.html?v1=2021207&page=',
    'kids-babyboy-sweatshirts-l70.html?v1=2021718&page=',
    'kids-babyboy-tshirts-l78.html?v1=2021601&page=',
    'kids-babyboy-trousers-l76.html?v1=2021656&page=',
    'kids-babyboy-jeans-l33.html?v1=2021634&page=',
    'kids-babyboy-basics-l20.html?v1=2021242&page=',
    'kids-babyboy-knitwear-l38.html?v1=2021692&page=',
    'kids-babyboy-shirts-l51.html?v1=2021600&page=',
    'kids-babyboy-trousers-overalls-l1763.html?v1=2021662&page=',
    'kids-babyboy-total-look-l3918.html?v1=2021721&page=',
    'kids-babyboy-license-l2956.html?v1=2021644&page=',
    'kids-babyboy-edited-l3186.html?v1=2021622&page=',
    'kids-babyboy-bags-l2629.html?v1=2021598&page=',
    'kids-babyboy-underwear-l84.html?v1=2021734&page=',
    'kids-babyboy-accessories-l7.html?v1=2021222&page=']

zara_russia(babyboyname, *babyboyurlplus)
# 男婴

babygirlname = 'BabyGirl'
babygirlurlplus = ['kids-babygirl-outerwear-l131.html?v1=2020600&page=',
                   'kids-babygirl-sweatshirts-l153.html?v1=2021133&page=',
                   'kids-babygirl-tshirts-l162.html?v1=2021006&page=',
                   'kids-babygirl-dresses-l108.html?v1=2021154&page=',
                   'kids-babygirl-trousers-l160.html?v1=2021073&page=',
                   'kids-babygirl-knitwear-l122.html?v1=2021104&page=',
                   'kids-babygirl-basics-l101.html?v1=2020627&page=',
                   'kids-babygirl-total-look-l3916.html?v1=2021135&page=',
                   'kids-babygirl-jeans-l115.html?v1=2021047&page=',
                   'kids-babygirl-shirts-l133.html?v1=2022609&page=',
                   'kids-babygirl-skirts-l1406.html?v1=2021030&page=',
                   'kids-babygirl-license-l2955.html?v1=2021055&page=',
                   'kids-babygirl-edited-l3183.html?v1=2021025&page=',
                   'kids-babygirl-bags-l100.html?v1=2020989&page=',
                   'kids-babygirl-underwear-l167.html?v1=2021148&page=',
                   'kids-babygirl-accessories-l90.html?v1=2020612&page=']
zara_russia(babygirlname, *babygirlurlplus)
# 女婴

newbornname = 'Newborn'
newbornurlplus = ['kids-newborn-outerwear-l507.html?v1=2021768&page=',
                  'kids-newborn-knitwear-l495.html?v1=2022004&page=',
                  'kids-newborn-bodysuits-l4208.html?v1=2021981&page=',
                  'kids-newborn-trousers-l524.html?v1=2021994&page=',
                  'kids-newborn-tshirts-l526.html?v1=2021973&page=',
                  'kids-newborn-sweatshirts-l519.html?v1=2022023&page=',
                  'kids-newborn-basics-l3946.html?v1=2046904&page=',
                  'kids-newborn-dungarees-l3096.html?v1=2021996&page=',
                  'kids-newborn-editorial-1-l491.html?v1=2021970&page=',
                  'kids-newborn-accessories-l476.html?v1=2021769&page=',
                  'kids-newborn-underwear-l528.html?v1=2022032&page=']

zara_russia(newbornname, *newbornurlplus)
# 新生儿
