# coding=utf-8
import time
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
from tqdm import tqdm


def HM_Russian(baseurl, name, **russiannames):
    print(f'当前要爬取的所有系列为：HM Russian-{name}', '\n', russiannames.values())
    for collection, russianname in russiannames.items():
        url1 = baseurl + str(collection) + str('.html?sort=stock&image-size=small&image=stillLife&offset=0&page-size=2')
        driver1 = webdriver.Chrome()
        driver1.get(url1)
        y = driver1.find_element(By.XPATH, '//div[@class="filter-pagination"]').text
        pttn = re.compile(r'\d*')
        b = int(re.findall(pttn, y)[0])

        url = baseurl + str(collection) + str(
            '.html?sort=stock&image-size=small&image=stillLife&offset=0&page-size=') + str(b)
        driver = webdriver.Chrome()
        driver.get(url)

        time.sleep(1)
        pros = driver.find_elements(By.XPATH, '//div[@class="image-container"]/a')

        pic1s = []
        titles = []
        prices = []
        colors = []
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
        for i in tqdm(pros[:b], desc="正在爬取" + 'HM Russian-' + str(name) + '-' + str(russianname)):
            slink = i.get_attribute('href')
            driver.execute_script("window,open('');")
            driver.switch_to.window(driver.window_handles[1])
            driver.get(slink)

            try:
                title = driver.find_element(By.XPATH, '//h1[@class="primary product-item-headline"]').text
            except NoSuchElementException:
                title = ''

            try:
                price = driver.find_element(By.XPATH,
                                            '//div[@class="ProductPrice-module--productItemPrice__2i2Hc"]').text
            except NoSuchElementException:
                price = ''

            try:
                color = driver.find_element(By.XPATH, '//h3[@class="product-input-label"]').text
            except NoSuchElementException:
                color = ''

            try:
                description = driver.find_element(By.XPATH,
                                                  '//div[@class="ProductDescription-module--productDescription__2mqXe"]').text
            except NoSuchElementException:
                description = ''

            try:
                link = driver.find_element(By.XPATH, '//link[@rel="canonical"]').get_attribute("href")
            except NoSuchElementException:
                link = ''

            try:
                pic1link1 = driver.find_element(By.XPATH, '//div[@class="product-detail-main-image-container"]')
                pic1link = pic1link1.find_element(By.TAG_NAME, 'img').get_attribute("src")
            except NoSuchElementException:
                pic1link = ''

            try:
                pic2link1 = driver.find_elements(By.XPATH, '//figure[@class="pdp-secondary-image pdp-image"]')[1]
                pic2link = 'https:' + pic2link1.find_element(By.TAG_NAME, 'img').get_attribute("src")
            except NoSuchElementException:
                pic2link = ''
            except IndexError:
                pic2link = ''

            try:
                pic3link1 = driver.find_elements(By.XPATH, '//figure[@class="pdp-secondary-image pdp-image"]')[2]
                pic3link = 'https:' + pic3link1.find_element(By.TAG_NAME, 'img').get_attribute("src")
            except NoSuchElementException:
                pic3link = ''
            except IndexError:
                pic3link = ''

            try:
                pic4link1 = driver.find_elements(By.XPATH, '//figure[@class="pdp-secondary-image pdp-image"]')[3]
                pic4link = 'https:' + pic4link1.find_element(By.TAG_NAME, 'img').get_attribute("src")
            except NoSuchElementException:
                pic4link = ''
            except IndexError:
                pic4link = ''

            try:
                pic5link1 = driver.find_elements(By.XPATH, '//figure[@class="pdp-secondary-image pdp-image"]')[4]
                pic5link = 'https:' + pic5link1.find_element(By.TAG_NAME, 'img').get_attribute("src")
            except NoSuchElementException:
                pic5link = ''
            except IndexError:
                pic5link = ''

            pic1s.append('')
            titles.append(title)
            prices.append(price)
            colors.append(color)
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

            data.to_excel('Russian/HM Russian ' + str(name) + '-' + str(russianname) + '.xlsx')

        print('HM Russian ' + str(name) + '-' + str(russianname) + '爬取完毕！')


girlbaseurl = 'https://www2.hm.com/ru_ru/'
girlname = 'Girls'
girlnames = {'deti/devochki/odezhda/platya': 'dresses',
             'deti/devochki/odezhda/topi': 'tops', 'deti/devochki/odezhda/svitera-dzhemperi': 'sweatshirts',
             'deti/devochki/odezhda/shtani-i-dzhinsi': 'pants',
             'deti/devochki/odezhda/yubki': 'skirts', 'deti/devochki/odezhda/shorti': 'shorts',
             'deti/devochki/odezhda/letnie-kombinezoni': 'overalls', 'deti/devochki/odezhda/naryadi-komplekti': 'suits',
             'deti/devochki/odezhda/odezhda-dlya-sna': 'pajamas',
             'deti/devochki/odezhda/nizhnee-belye': 'underwear', 'deti/devochki/odezhda/kolgotki-nosochki': 'tights',
             'deti/devochki/odezhda/kupalniki': 'swimsuit',
             'deti/devochki/verhnyaya-odezhda': 'outwears', 'deti/devochki/aksessuari': 'accessories',
             'deti/devochki/karnavalnie-kostyumi': 'playsuits'}

HM_Russian(girlbaseurl, girlname, **girlnames)
# 女童

boybaseurl = 'https://www2.hm.com/ru_ru/'
boyname = 'Boys'
boynames = {'deti/malchiki/odezhda/futbolki-i-rubashki': 'T-shirts&shirts',
            'deti/malchiki/odezhda/svitera-dzhemperi': 'Sweatshirts',
            'deti/malchiki/odezhda/bryuki-dzinsi-shtani': 'Pants',
            'deti/malchiki/odezhda/shorti': 'Shorts',
            'deti/malchiki/odezhda/kostyumi-pidzhaki': 'Classicsuits',
            'deti/malchiki/odezhda/komplekti-odezhdi-naryadi': 'Suits',
            'deti/malchiki/odezhda/odezhda-dlya-sna': 'Pajamas',
            'deti/malchiki/odezhda/nizhnee-belye': 'Underwear', 'deti/malchiki/odezhda/noski': 'Socks',
            'deti/malchiki/odezhda/kupalnie-kostyumi': 'Swimsuits', 'deti/malchiki/verhnyaya-odezhda': 'Outwear',
            'deti/malchiki/aksessuari': 'Accessories', 'deti/malchiki/karnavalnie-kostyumi': 'Playsuits'}

HM_Russian(boybaseurl, boyname, **boynames)
# 男童


newbornbaseurl = 'https://www2.hm.com/ru_ru/'
newbornname = 'Newborns'
newbornnames = {'malyshi/novorozhdennie/odezhda/naryadi-komplekti': 'Suits',
                'malyshi/novorozhdennie/odezhda/bodi': 'Bodysuits',
                'malyshi/novorozhdennie/odezhda/futbolki-topiki': 'T-shirts',
                'malyshi/novorozhdennie/odezhda/kofti-svitera': 'Sweaters',
                'malyshi/novorozhdennie/odezhda/shtanishki': 'Pants',
                'malyshi/novorozhdennie/odezhda/shortiki': 'Shorts',
                'malyshi/novorozhdennie/odezhda/platya': 'Dresses',
                'malyshi/novorozhdennie/odezhda/noski-kolgotki': 'Socks',
                'malyshi/novorozhdennie/aksessuari': 'Accessories',
                'malyshi/novorozhdennie/verhnyaya-odezhda': 'Outwears'}

HM_Russian(newbornbaseurl, newbornname, **newbornnames)
# 新生儿


babygirlbaseurl = 'https://www2.hm.com/ru_ru/'
babygirlname = 'Babygirls'
babygirlnames = {'malyshi/devochki/odezhda/platya': 'Dresses',
                 'malyshi/devochki/odezhda/naryadi-komplekti': 'Sweatshirts',
                 'malyshi/devochki/odezhda/topi': 'Tops',
                 'malyshi/devochki/odezhda/koftochki-svitera': 'Sweaters',
                 'malyshi/devochki/odezhda/dzhinsi-bryuki': 'Pants',
                 'malyshi/devochki/odezhda/bodi': 'Bodysuits',
                 'malyshi/devochki/odezhda/letnie-kombinezoni': 'Overall',
                 'malyshi/devochki/odezhda/shorti': 'Shorts',
                 'malyshi/devochki/odezhda/pizhami': 'Sleepsuits',
                 'malyshi/devochki/odezhda/kolgotki-nosochki': 'Socks',
                 'malyshi/devochki/odezhda/kupalnie-kostyumi': 'Swimsuits',
                 'malyshi/devochki/verhnyaya-odezhda': 'Outwears',
                 'malyshi/devochki/aksessuari': 'Accessories'}

HM_Russian(babygirlbaseurl, babygirlname, **babygirlnames)
# 女婴


babyboybaseurl = 'https://www2.hm.com/ru_ru/'
babyboyname = 'Babyboys'
babyboynames = {'malyshi/malchiki/odezhda/kostyumchiki': 'Suits',
                'malyshi/malchiki/odezhda/futbolki-i-rubashechki': 'T-shirts',
                'malyshi/malchiki/odezhda/kofti-svitera': 'Sweatshirts',
                'malyshi/malchiki/odezhda/dzhinsi-bryuki': 'Pants',
                'malyshi/malchiki/odezhda/bodi': 'Bodysuits',
                'malyshi/malchiki/odezhda/letnie-kombinezoni': 'Overall',
                'malyshi/malchiki/odezhda/shorti': 'Shorts',
                'malyshi/malchiki/odezhda/pizhami': 'Sleepsuits',
                'malyshi/malchiki/odezhda/noski': 'Socks',
                'malyshi/malchiki/odezhda/kupalnie-kostyumi': 'Swimsuits',
                'malyshi/malchiki/verhnyaya-odezhda': 'Outwears',
                'malyshi/malchiki/aksessuari': 'Accessories'}

HM_Russian(babyboybaseurl, babyboyname, **babyboynames)
#  男婴
