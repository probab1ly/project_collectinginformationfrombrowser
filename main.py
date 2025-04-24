from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from fake_useragent import UserAgent
import json
import pandas as pd
import requests
import vk_api
from vk_api.exceptions import ApiError

spisok = []
token = 'a981cae6a981cae6a981cae6d9aaae9f91aa981a981cae6ce7c910f1716f384606a5381'
needwords = ['МГЮА', 'Мгюа', 'Кутафина']
query = ['Владимир Дмитриевич Никишин', 'Никишин Владимир Дмитриевич', 'Александр Сергеевич Глазков', 'Глазков Александр Сергеевич']
ua = UserAgent()
user_agent = ua.random
options = webdriver.ChromeOptions()
# options.add_argument('--headless=new')
options.add_argument('--disable-popup-blocking')
options.add_argument(f'--user-agent={user_agent}')
driver = webdriver.Chrome(options=options)

access_token = 'vk1.a.TG0toYM7T8fb1FOxFT-Vtxu-3Db4XVtdUbzhJeRnkDRx58mIujNZtci-S1jgFawYMpnqoWd6dmRs-slHgmoYdsES-mk7-ziTe9Wy-uHOs04ksIytCPpFcn-5QYnjmDOSofpiSCRCOwMDpEnBlzJd0JCARaqXVtBH8A_A4s-D8zbJQDHOb6Eok8C99uiWtmjf2t87gbKwyMbaiiLvBRhqlg'
vk_session = vk_api.VkApi(token=access_token)
vk = vk_session.get_api()

def findvk(word):
    try:
        results = vk.newsfeed.search(q=word, count=200)
        items = results['items']
        for item in items:
            if 'owner_id' in item and 'id' in item:
                text = item.get('text').strip()
                if any(r in text for r in query) and any(r1 in text for r1 in needwords):
                    owner_id = item['owner_id']
                    id = item['id']
                    post_url = f'https://vk.com/wall{owner_id}_{id}'
                    if post_url!='None':
                        spisok.append({
                            'url': post_url,
                            'title': 'vk'
                        })
        print('Парсинг VK успешно завершён')
    except ApiError as e:
        print(f'Ошибка vkapi: {e}')

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
                if 'youtube' not in url2:
                    driver.execute_script("window.open('');")
                    driver.switch_to.window(driver.window_handles[-1])
                    driver.get(url2)
                    time.sleep(2)
                    try:
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
        print('Парсинг Yandex успешно завершён')
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
def save_to_excel(w, sp):
    try:
        if w == 'Владимир Дмитриевич Никишин':
            df = pd.DataFrame(sp)
            writer = pd.ExcelWriter('parserexcelVL.xlsx', engine='openpyxl')
            df.to_excel(writer, index=False, sheet_name='ParseVL')
            worksheet = writer.sheets['ParseVL']
            worksheet.column_dimensions['A'].width = 85
            worksheet.column_dimensions['B'].width = 160
            writer.close()
            print('Данные сохранены в Excel(VL)')
        elif w == 'Александр Сергеевич Глазков':
            df = pd.DataFrame(sp)
            writer = pd.ExcelWriter('parserexcelAL.xlsx', engine='openpyxl')
            df.to_excel(writer, index=False, sheet_name='ParseAL')
            worksheet = writer.sheets['ParseAL']
            worksheet.column_dimensions['A'].width = 85
            worksheet.column_dimensions['B'].width = 160
            writer.close()
            print('Данные сохранены в Excel(AL)')
        else:
            df = pd.DataFrame(sp)
            writer = pd.ExcelWriter('parserexcel.xlsx', engine='openpyxl')
            df.to_excel(writer, index=False, sheet_name='Parse')
            worksheet = writer.sheets['Parse']
            worksheet.column_dimensions['A'].width = 85
            worksheet.column_dimensions['B'].width = 160
            writer.close()
            print('Данные сохранены в Excel')
    except Exception as e:
        print(f'Произошла ошибка на этапе сохранения в Excel: {e}')
def main():
    print('Введите нужное словосочетание')
    word = input()
    findvk(word)
    findyandex(word)
    finall_spisok = [dict(s) for s in {frozenset(sp.items()) for sp in spisok}] #убираем повторения
    save_to_json(word, finall_spisok)
    save_to_excel(word, finall_spisok)

if __name__=='__main__':
    main()
