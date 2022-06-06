import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import ElementNotInteractableException
import pandas as pd
from tqdm import tqdm


def kiabi_russia(name, **urlplus):
    print(f'要爬取的所有系列为：Kiabi Russia-{name} ', '\n', urlplus.values())
    for a, a1 in urlplus.items():
        baseurl = 'https://www.kiabi.ru/'
        url = baseurl + str(a)
        driver = webdriver.Chrome()
        driver.get(url)
        time.sleep(2)
        try:
            driver.find_element(By.XPATH, '//div[@class="display-all"]').click()
        except NoSuchElementException:
            pass
        except ElementNotInteractableException:
            pass
        for q in range(1000, 1020):
            driver.execute_script(f'window.scrollTo(0, window.scrollY + {q});')
            time.sleep(2)
        pros = driver.find_elements(By.XPATH, '//span[@class="img productImageElement"]')

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
                      desc=f"正在爬取Kiabi Russia-{name}-{a1}"):
            slink1 = i.get_attribute('data-product-url')
            slink = 'https://www.kiabi.ru' + slink1
            try:
                driver.execute_script("window,open('');")
                driver.switch_to.window(driver.window_handles[1])
                driver.get(slink)

                try:
                    link = driver.find_element(By.XPATH, '//link[@rel="canonical"]').get_attribute('href')

                except NoSuchElementException:
                    link = ''
                except StaleElementReferenceException:
                    link = ''

                try:
                    title = driver.find_element(By.XPATH, '//h2[@id="productName"]').text
                except NoSuchElementException:
                    title = ''
                except StaleElementReferenceException:
                    title = ''

                try:
                    price = driver.find_element(By.XPATH, '//span[@class="prices"]').text
                except NoSuchElementException:
                    price = ''
                except StaleElementReferenceException:
                    price = ''
                try:
                    color = driver.find_element(By.XPATH, '//meta[@itemprop="color"]').get_attribute('content')
                except NoSuchElementException:
                    color = ''
                except StaleElementReferenceException:
                    color = ''

                try:
                    description = driver.find_element(By.XPATH, '//div[@class="product-description"]').text
                except NoSuchElementException:
                    description = ''
                except StaleElementReferenceException:
                    description = ''

                try:
                    size = driver.find_element(By.XPATH, '//div[@class="sizes_section block"]/span').get_attribute(
                        'innerText')
                except NoSuchElementException:
                    size = ''
                except StaleElementReferenceException:
                    size = ''

                try:
                    composition = driver.find_element(By.XPATH, '//div[@class="product-composition"]').text
                except NoSuchElementException:
                    composition = ''
                except StaleElementReferenceException:
                    composition = ''

                try:
                    pic1link = driver.find_element(By.XPATH, '//meta[@property="og:image"]').get_attribute('content')
                except NoSuchElementException:
                    pic1link = ''
                except IndexError:
                    pic1link = ''
                except StaleElementReferenceException:
                    pic1link = ''

                try:
                    pic2link = driver.find_elements(By.XPATH, '//li[@class="exists"]/a')[0].get_attribute('href')
                except NoSuchElementException:
                    pic2link = ''
                except IndexError:
                    pic2link = ''
                except StaleElementReferenceException:
                    pic2link = ''

                try:
                    pic3link = driver.find_elements(By.XPATH, '//li[@class="exists"]/a')[1].get_attribute('href')
                except NoSuchElementException:
                    pic3link = ''
                except IndexError:
                    pic3link = ''
                except StaleElementReferenceException:
                    pic3link = ''

                try:
                    pic4link = driver.find_elements(By.XPATH, '//li[@class="exists"]/a')[2].get_attribute('href')
                except NoSuchElementException:
                    pic4link = ''
                except IndexError:
                    pic4link = ''
                except StaleElementReferenceException:
                    pic4link = ''

                try:
                    pic5link = driver.find_elements(By.XPATH, '//li[@class="exists"]/a')[3].get_attribute('href')
                except NoSuchElementException:
                    pic5link = ''
                except IndexError:
                    pic5link = ''
                except StaleElementReferenceException:
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

                data['Pic-1'] = pic1s
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

                data.to_excel(f'Kiabi/Kiabi Russia-{name}-{a1}.xlsx')

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

                data['Pic-1'] = pic1s
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

                data.to_excel(f'Kiabi/Kiabi Russia-{name}-{a1}.xlsx')

        driver.close()
        print(f'Kiabi Russia-{name}-{a1}爬取完毕！')

    print(f'Kiabi Russia - {name}所有系列爬取完毕！')


girlname = 'Girls'
girlurlplus = {
    # "platjya-yubki-devochki_254959": "dresses",
    # "verkhnyaya-odezhda-devochki_254835": "jackets",
    # "futbolki-vodolazki-devochki_254847": "T-shirts",
    # "rubashki-bluzki-devochki_254977": "shirts",
    # "dzhinsy-devochki_254891": "jeans",
    # "bryuki-devochki_254857": "pants",
    # "kombinezony-devochki_326072": "overall",
    "tolstovki-devochki_255007": "hoodies",
    "svitery-kardigany-devochki_254827": "sweaters",
    "sport-devochki_385552": "sports",
    "leginsy-devochki_255001": "leggings",
    "ukorochennye-bryuki-shorty-devochki_255031": "shorts",
    "pizhamy-khalaty-devochki_254939": "pajams",
    "nizhnee-belje-devochki_254949": "underwears",
    "chulochno-nosochnye-izdeliya-devochki_254971": "socks",
    "aksessuary-devochki_254983": "accessories"}
kiabi_russia(girlname, **girlurlplus)
# 女童

boyname = 'Boys'
boyurlplus = {
    "futbolki-polo-malchiki_255435": "polos",
    "verkhnyaya-odezhda-malchiki_255423": "jackets",
    "dzhinsy-malchiki_255369": "jeans",
    "bryuki-malchiki_255409": "pants",
    "tolstovki-malchiki_255269": "hoodies",
    "svitery-kardigany-malchiki_255311": "sweaters",
    "rubashki-malchiki_255277": "shirts",
    "sport-malchiki_385577": "sports",
    "komplekty-malchiki_255287": "suits",
    "bermudy-shorty-malchiki_255393": "shorts",
    "ryukzaki-portfeli-malchiki_255361": "bags",
    "pizhamy-khalaty-malchiki_255337": "pajamas",
    "nizhnee-belje-malchiki_255317": "underwears",
    "noski-malchiki_255323": "socks",
    "aksessuary-malchiki_255295": "accessories"

}
kiabi_russia(boyname, **boyurlplus)
# 男童

babyname = 'Baby'
babyurlplus = {"bodi-nizhnee-belje-malyshi_255739": "overall",
               "pizhamy-kombinezony-malyshi_255657": "sleepsuits",
               "konverty-dlya-novorozhdennykh-malyshi_255725": "sleepbags",
               "verkhnyaya-odezhda-malyshi_255717": "jackets",
               "futbolki-malyshi_255679": "T-shirts",
               "svitery-zhilety-tolstovki-malyshi_255691": "sweaters",
               "platjya-yubki-malyshi_255731": "dresses",
               "komplekty-bodi-pesochniki-malyshi_255705": "suits",
               "rubashki-malyshi_255699": "shirts",
               "bryuki-dzhinsy-leginsy-malyshi_255667": "pants",
               "shorty-bermudy-malyshi_255843": "shorts",
               "chulochno-nosochnye-izdeliya-kolgotki-malyshi_255751": "socks",
               "aksessuary-malyshi_255811": "accessories",
               "sport-malyshi_385598": "sports"}

kiabi_russia(babyname, **babyurlplus)
# 婴儿
