# coding=utf-8
import time
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
from tqdm import tqdm



def HM_China(baseurl, name, *collections):
    print(f'当前要爬取的所有系列为：{name}', collections)
    for collection in collections:

        url1 = baseurl + str(collection) + str('.html?size=2&sort=stock')
        driver1 = webdriver.Chrome()
        driver1.get(url1)
        y = driver1.find_element(By.XPATH, '//div[@class="filter-pagination"]').text
        pttn = re.compile(r'\d*')
        b = int(re.findall(pttn, y)[0])

        url = baseurl + str(collection) + str('.html?size=') + str(b) + str('&sort=stock')
        driver = webdriver.Chrome()
        driver.get(url)

        time.sleep(2)
        pros = driver.find_elements(By.XPATH, '//div[@class="product-name"]/a')

        titles = []
        prices = []
        colors = []
        compositions = []
        links = []
        mainImg1s = []
        pImg11s = []
        pImg21s = []
        pImg31s = []
        pImg41s = []
        for i in tqdm(pros[:b], desc="正在爬取" + str(name) + '-' + str(collection)):
            ActionChains(driver) \
                .key_down(Keys.CONTROL) \
                .click(i) \
                .key_up(Keys.CONTROL) \
                .perform()
            driver.switch_to.window(driver.window_handles[1])

            title = driver.find_element(By.XPATH, '//h1[@class="page-title"]').text
            composition = driver.find_element(By.XPATH, '//ul[@class="product-description-list"]').text
            link = driver.find_element(By.XPATH, '//link[@rel="canonical"]').get_attribute("href")
            price = driver.find_element(By.XPATH, '//span[@data-price-type="price"]').text
            color = driver.find_element(By.XPATH, '//span[@class="current-article-title"]').text

            try:
                mainImg = driver.find_element(By.XPATH, '//div[@class="item  main-image main_img p_img_0 "]')
                mainImg1 = mainImg.find_element(By.TAG_NAME, 'img').get_attribute("src")

            except NoSuchElementException:
                mainImg1 = ''

            try:
                pImg1 = driver.find_element(By.XPATH, '//div[@class="main_img p_img_1"]')
                pImg11 = pImg1.find_element(By.TAG_NAME, 'img').get_attribute("src")

            except NoSuchElementException:
                pImg11 = ''

            try:
                pImg2 = driver.find_element(By.XPATH, '//div[@class="main_img p_img_2"]')
                pImg21 = pImg2.find_element(By.TAG_NAME, 'img').get_attribute("data-src")

            except NoSuchElementException:
                pImg21 = ''

            try:
                pImg3 = driver.find_element(By.XPATH, '//div[@class="main_img p_img_3"]')
                pImg31 = pImg3.find_element(By.TAG_NAME, 'img').get_attribute("data-src")

            except NoSuchElementException:
                pImg31 = ''

            try:
                pImg4 = driver.find_element(By.XPATH, '//div[@class="main_img p_img_4"]')
                pImg41 = pImg4.find_element(By.TAG_NAME, 'img').get_attribute("data-src")

            except NoSuchElementException:
                pImg41 = ''

            titles.append(title)
            prices.append(price)
            colors.append(color)
            compositions.append(composition)
            links.append(link)
            mainImg1s.append(mainImg1)
            pImg11s.append(pImg11)
            pImg21s.append(pImg21)
            pImg31s.append(pImg31)
            pImg41s.append(pImg41)
            driver.close()

            driver.switch_to.window(driver.window_handles[0])

            data = pd.DataFrame()
            data['名称'] = titles
            data['价格'] = prices
            data['颜色'] = colors
            data['详细信息'] = compositions
            data['商品链接'] = links
            data['主图链接'] = mainImg1s
            data['图片1链接'] = pImg11s
            data['图片2链接'] = pImg21s
            data['图片3链接'] = pImg31s
            data['图片4链接'] = pImg41s

            data.to_excel('test/HM ' + str(name) + '-' + str(collection) + '.xlsx')  # 3.此处改文件名称

        print('HM ' + str(name) + '-' + str(collection) + '爬取完毕！')


babyboybaseurl = 'https://www.hm.com.cn/zh_cn/baby-new/baby-boys/'
babyboyname = 'baby-boys'
babyboycollections = ['nightwear', 'sets-outfits', 't-shirts-shirts', 'jumpers-cardigans', 'outerwear',
                      'trousers-jeans',
                      'jumpsuits-playsuits', 'bodysuits', 'shorts', 'swimwear', 'socks', 'accessories']
HM_China(babyboybaseurl, babyboyname, *babyboycollections)
# 男婴

babygirlbaseurl = 'https://www.hm.com.cn/zh_cn/baby-new/baby-girls/'
babygirlname = 'baby-girls'
babygirlcollections = ['nightwear', 'sets-outfits', 'dresses', 'tops-t-shirts', 'jumpers-cardigans', 'outerwear',
                       'trousers-jeans', 'jumpsuits-playsuits', 'bodysuits', 'shorts', 'swimwear', 'socks-tights',
                       'accessories']
HM_China(babygirlbaseurl, babygirlname, *babygirlcollections)
# 女婴

newbornbaseurl = 'https://www.hm.com.cn/zh_cn/baby-new/newborn/'
newbornname = 'newborn'
newborncollections = ['nightwear', 'bodysuits', 'tops-t-shirts', 'jumpers-cardigans', 'dresses-skirts',
                      'trousers-leggings',
                      'outwear', 'socks-tights', 'shorts', 'accessories']
HM_China(newbornbaseurl, newbornname, *newborncollections)
# 新生儿

boybaseurl = 'https://www.hm.com.cn/zh_cn/kids-new/boys/'
boyname = 'boys'
boycollections = ['outerwear', 'trousers-jeans', 't-shirts-shirts', 'jumpers-sweatshirts', 'sets-outfits',
                  'blazers-suits',
                  'nightwear', 'underwear', 'socks', 'shorts', 'swimwear', 'accessories', 'sportswear']
HM_China(boybaseurl, boyname, *boycollections)
# 男童

boybaseurl = 'https://www.hm.com.cn/zh_cn/kids-new/girls/'
boyname = 'girls'
boycollections = ['outerwear', 'dresses', 'jumpers-sweatshirts', 'trousers-jeans', 'sets-outfits', 'tops-t-shirts',
                  'nightwear', 'underwear', 'socks-tights', 'skirts', 'shorts', 'jumpsuits-playsuits', 'swimwear',
                  'accessories', 'fancy-dress-costumes', 'sportswear']
HM_China(boybaseurl, boyname, *boycollections)
# 女童
