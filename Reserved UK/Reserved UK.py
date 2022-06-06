import time
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
import pandas as pd
from tqdm import tqdm


def reserved_uk(name, *urlplus):
    # a1pttn = re.compile(r'\w*\/\w*\/')
    # a1 = [re.findall(a1pttn, x1) for x1 in urlplus]
    a1 = [re.sub(r'\w*/\w*/', '', x1) for x1 in urlplus]
    print(f'要爬取的所有系列为：Reserved UK-{name} ', '\n', a1[0:5], '\n', a1[5:10], '\n', a1[10:15], '\n', a1[15:20])

    for a, a2 in zip(urlplus, a1):
        baseurl = 'https://www.reserved.com/gb/en/'
        url = baseurl + str(a)
        driver = webdriver.Chrome()
        driver.get(url)
        time.sleep(2)
        driver.find_element(By.XPATH, '//button[@id="cookiebotDialogOkButton"]').click()
        driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
        time.sleep(2)

        a3 = ''.join(a2)
        pros = driver.find_elements(By.XPATH, '//a[@class="sc-jSgvzq eltUJY es-product-photo"]')

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
                      desc=f"正在爬取Reserved UK-{name}-" + str(a3)):
            slink = i.get_attribute('href')
            try:
                driver.execute_script("window,open('');")
                driver.switch_to.window(driver.window_handles[1])
                driver.get(slink)
                time.sleep(2)

                try:
                    driver.find_element(By.XPATH, '//div[@class="close"]').click()
                except NoSuchElementException:
                    pass
                except ElementNotInteractableException:
                    pass

                try:
                    link = driver.find_element(By.XPATH, '//link[@rel="canonical"]').get_attribute('href')

                except NoSuchElementException:
                    link = ''

                try:
                    title = driver.find_element(By.XPATH, '//h1[@class="product-name"]').text
                except NoSuchElementException:
                    title = ''

                try:
                    price = driver.find_element(By.XPATH, '//div[@class="regular-price"]').text
                except NoSuchElementException:
                    price = ''

                try:
                    color1 = driver.find_element(By.XPATH, '//section[@data-selen="color-picker"]').text
                    color = re.sub('COLOUR - ', '', color1)
                except NoSuchElementException:
                    color = ''

                try:
                    description = driver.find_element(By.XPATH, '//div[@class="sc-bYEvvW iTIqaq"]').text
                except NoSuchElementException:
                    description = ''

                try:
                    size1 = driver.find_element(By.XPATH, '//div[@class="size-picker-wrapper"]').get_attribute(
                        'data-sizes-list')
                    sizepttn = re.compile(r'\d?-\d? Y')
                    size = '/'.join(re.findall(sizepttn, size1))
                except NoSuchElementException:
                    size = ''

                try:
                    composition = driver.find_element(By.CLASS_NAME, 'composition-value').get_attribute('innerText')
                except NoSuchElementException:
                    composition = ''

                try:
                    pic1link = driver.find_element(By.XPATH, '//meta[@property="og:image"]').get_attribute('content')
                except NoSuchElementException:
                    pic1link = ''
                except IndexError:
                    pic1link = ''

                try:
                    driver.find_element(By.XPATH, '//section[@data-index="1"]').click()
                    time.sleep(2)
                    pic2link1 = driver.find_element(By.XPATH,
                                                    '//div[@class="item__Item-sc-181aek5-0 bWWRbz"]').get_attribute(
                        'innerHTML')
                    pic2linkpttn = re.compile(r'https.*jpg')
                    pic2link = ''.join(re.findall(pic2linkpttn, pic2link1))
                except NoSuchElementException:
                    pic2link = ''
                except IndexError:
                    pic2link = ''

                try:
                    driver.find_element(By.XPATH, '//section[@data-index="2"]').click()
                    time.sleep(2)
                    pic3link1 = driver.find_element(By.XPATH,
                                                    '//div[@class="item__Item-sc-181aek5-0 bWWRbz"]').get_attribute(
                        'innerHTML')
                    pic3linkpttn = re.compile(r'https.*jpg')
                    pic3link = ''.join(re.findall(pic3linkpttn, pic3link1))
                except NoSuchElementException:
                    pic3link = ''
                except IndexError:
                    pic3link = ''

                try:
                    driver.find_element(By.XPATH, '//section[@data-index="3"]').click()
                    time.sleep(2)
                    pic4link1 = driver.find_element(By.XPATH,
                                                    '//div[@class="item__Item-sc-181aek5-0 bWWRbz"]').get_attribute(
                        'innerHTML')
                    pic4linkpttn = re.compile(r'https.*jpg')
                    pic4link = ''.join(re.findall(pic4linkpttn, pic4link1))
                except NoSuchElementException:
                    pic4link = ''
                except IndexError:
                    pic4link = ''

                try:
                    driver.find_element(By.XPATH, '//section[@data-index="4"]').click()
                    time.sleep(2)
                    pic5link1 = driver.find_element(By.XPATH,
                                                    '//div[@class="item__Item-sc-181aek5-0 bWWRbz"]').get_attribute(
                        'innerHTML')
                    pic5linkpttn = re.compile(r'https.*jpg')
                    pic5link = ''.join(re.findall(pic5linkpttn, pic5link1))
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

                data.to_excel(f'Reserved/Reserved UK-{name}-{a3}.xlsx')

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

                data.to_excel('Reserved/Reserved UK-' + str(name) + '-' + str(a3) + '.xlsx')

        driver.close()
        print('Reserved UK' + '-' + str(name) + str(a3) + '爬取完毕！')

    print('Reserved UK - ' + str(name) + '所有系列爬取完毕！')


# girlname = 'Girl'
# girlurlplus = [
#     "girl/junior/longsleeve",
#     "girl/junior/t-shirts",
#     "girl/junior/shirts",
#     "girl/junior/sweatshirts",
#     "girl/junior/sweaters",
#     "girl/junior/trousers",
#     "girl/junior/jeans",
#     "girl/junior/skirts",
#     "girl/junior/sets",
#     "girl/junior/bags",
#     "girl/junior/garment"
# ]
# reserved_uk(girlname, *girlurlplus)

boyname = 'Boy'
boyurlplus = [
    # "boy/junior/sweatshirts",
    # "boy/junior/longsleeve",
    # "boy/junior/trousers",
    # "boy/junior/t-shirts",
    # "boy/junior/sweaters",
    # "boy/junior/shirts",
    # "boy/junior/bags",
    # "boy/junior/garment"
]
reserved_uk(boyname, *boyurlplus)
# 男童

babyname = 'Baby'
babyurlplus = [
    # "girl/baby/sets",
    "boy/baby"]
reserved_uk(babyname, *babyurlplus)
# 婴儿

newbornname = 'Newborn'
newbornurlplus = ["girl/newborn", "boy/newborn"]
reserved_uk(newbornname, *newbornurlplus)
# 新生儿
