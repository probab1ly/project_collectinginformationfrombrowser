from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from fake_useragent import UserAgent
import json
import pandas as pd
import requests
spisok = []
token = 'a981cae6a981cae6a981cae6d9aaae9f91aa981a981cae6ce7c910f1716f384606a5381'
needwords = ['МГЮА', 'Мгюа', 'Кутафина']
needwordsmb = ['Владимир Дмитриевич Никишин', 'Владимир Никишин', 'Александр Сергеевич Глазков', 'Александр Глазков']
ua = UserAgent()
user_agent = ua.random
options = webdriver.ChromeOptions()
# options.add_argument('--headless=new')
options.add_argument('--disable-popup-blocking')
options.add_argument(f'--user-agent={user_agent}')
driver = webdriver.Chrome(options=options)
# def findgoogle_vk(testword):
#     print('Запуск парсера в Google')
#     try:
#         driver.get(f'https://www.google.com/search?q=site%3Avk.com+{testword}')
#         time.sleep(10)
#         products = driver.find_elements(By.CLASS_NAME, 'yuRUbf')
#         for product in products:
#             url1 = product.find_element(By.CLASS_NAME, 'zReHs').get_attribute('href')
#             title1 = product.find_element(By.CLASS_NAME, 'LC20lb').text.split()
#             titlegood1 = ' '.join(title1)
#             if url1:
#                 driver.execute_script("window.open('');")
#                 driver.switch_to.window(driver.window_handles[-1])
#                 try:
#                     driver.get(url1)
#                     time.sleep(2)
#                     text = driver.find_element(By.XPATH, '/html/body').text
#                     if any(keyword in text for keyword in needwords) and any(keyword1 in text for keyword1 in needwordsmb):
#                         spisok.append({
#                             'url': url1,
#                             'title': titlegood1
#                         })
#                     driver.close()
#                     driver.switch_to.window(driver.window_handles[0])
#
#                 except Exception as e:
#                     print(f'Произошла ошибка 2: {e}')
#     except Exception as e:
#         print(f'Произошла ошибка 1: {e}')
def findgoogle_vk(testword):
    url = 'https://api.vk.com/method/wall.search'
    params = {
        'access_token': token,
        'v': '5.131',
        'count': 50,
        'q': testword
    }
    response = requests.get(url, params=params)
    data = response.json()
    items = data['response']['items']
def findyandex(testword):
    print('Запуск парсера в Яндекс')
    try:
        for i in range(0, 2):#количество страниц
            driver.get(f'https://ya.ru/search/?text={testword}&search_source=yaru_desktop_common&lr=11031&p={i}')
            time.sleep(12)
            products = driver.find_elements(By.CLASS_NAME, 'OrganicTitle')
            for product in products:
                title2 = product.find_element(By.CLASS_NAME, 'OrganicTitleContentSpan').text.split()
                titlegood2 = ' '.join(title2)
                url2 = product.find_element(By.CLASS_NAME, 'OrganicTitle-Link').get_attribute('href')

                if url2[:24]=='https://www.youtube.com':
                    spisok.append({
                        'url': url2,
                        'title': titlegood2
                    })

                elif url2:
                    driver.execute_script("window.open('');")
                    driver.switch_to.window(driver.window_handles[-1])
                    try:
                        driver.get(url2)
                        time.sleep(2)
                        text = driver.find_element(By.XPATH, '/html/body').text
                        if any(keyword in text for keyword in needwords):
                            spisok.append({
                                'url': url2,
                                'title': titlegood2
                            })
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                    except Exception as e:
                        print(f'Произошла ошибка 4: {e}')
                else:
                    print('Адрес не найден')

    except Exception as e:
        print(f'Произошла ошибка 3: {e}')
    finally:
        driver.quit()
def save_to_json(w, sp):
    try:
        if w == 'Владимир Дмитриевич Никишин':
            with open('vova.json', 'w', encoding='utf-8') as file1:
                json.dump(sp, file1, indent=4, ensure_ascii=False)
        elif w == 'Александр Сергеевич Глазков':
            with open('aleks.json', 'w', encoding='utf-8') as file2:
                json.dump(sp, file2, indent=4, ensure_ascii=False)
        else:
            with open('test.json', 'w', encoding='utf-8') as file:
                json.dump(sp, file, indent=4, ensure_ascii=False)
        print('Данные сохранены в Json')
    except Exception as e:
        print(f'Произошла ошибка на этапе сохранения в JSON: {e}')
def save_to_excel(sp, fil='parserexcel.xlsx'):
    try:
        df = pd.DataFrame(sp)
        writer = pd.ExcelWriter(fil, engine='openpyxl')
        df.to_excel(writer, index=False, sheet_name='Parse')
        worksheet = writer.sheets['Parse']
        worksheet.column_dimensions['A'].width = 140
        worksheet.column_dimensions['B'].width = 160
        writer.close()
        print('Данные сохранены в Excel')
    except Exception as e:
        print(f'Произошла ошибка на этапесохранения в Excel: {e}')
def main():
    print('Введите нужное словосочетание')
    word = input()
    findgoogle_vk(word)
    findyandex(word)
    finall_spisok = [dict(s) for s in {frozenset(sp.items()) for sp in spisok}]
    save_to_json(word, finall_spisok)
    save_to_excel(finall_spisok)
if __name__=='__main__':
    main()
    
# import requests
# from vk_api import vk_api
# import requests

# def search_vk_posts(query, access_token):
#     url = 'https://api.vk.com/method/wall.search'
#     params = {
#         'access_token': access_token,
#         'v': '5.131',
#         'query': query,
#         'count': 100
#     }
#     response = requests.get(url, params=params)
#     data = response.json()

#     post_links = []
#     if 'response' in data:
#         for post in data['response']['items']:
#             post_links.append(f"https://vk.com/wall{post['owner_id']}_{post['id']}")

#     return post_links

# access_token = 'ваш_токен_доступа'
# query = 'ваши_ключевые_слова'
# links = search_vk_posts(query, access_token)
# print(links)

