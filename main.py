from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from fake_useragent import UserAgent
import json
spisok = []
needwords = ['МГЮА', 'Мгюа', 'Кутафина']
def find(testword):
    try:
        ua = UserAgent()
        user_agent = ua.random
        options = webdriver.ChromeOptions()
        # options.add_argument('--headless=new')
        options.add_argument('--disable-popup-blocking')
        options.add_argument(f'--user-agent={user_agent}')
        driver = webdriver.Chrome(options=options)

        for i in range(0, 6):
            driver.get(f'https://ya.ru/search/?text={testword}&search_source=yaru_desktop_common&lr=11031&p={i}')
            print(f'Обрабатывается страница: {i+1}')
            time.sleep(12)

            products = driver.find_elements(By.CLASS_NAME, 'OrganicTitle')
            for product in products:
                title = product.find_element(By.CLASS_NAME, 'OrganicTitleContentSpan').text.split()
                titlegood = ' '.join(title)
                url = product.find_element(By.CLASS_NAME, 'OrganicTitle-Link').get_attribute('href')

                if url and url[:24]!='https://www.youtube.com':
                    driver.execute_script("window.open('');")
                    driver.switch_to.window(driver.window_handles[-1])
                    try:
                        driver.get(url)
                        time.sleep(2)
                        text = driver.find_element(By.XPATH, '/html/body').text
                        if any(keyword in text for keyword in needwords):
                            spisok.append({
                                'url': url,
                                'title': titlegood
                            })
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                    except Exception as e:
                        print(f'Произошла ошибка 2: {e}')
    except Exception as e:
        print(f'Произошла ошибка 1: {e}')
    finally:
        driver.quit()
def main():
    print('Введите нужное словосочетание')
    word = input()
    find(word)
    # spisok1 = [s for s in spisok if s['url'][:23]!='https://yabs.yandex.ru']
    finall_spisok = [dict(s) for s in {frozenset(sp.items()) for sp in spisok}]

    if word == 'Владимир Дмитриевич Никишин':
        with open('vova.json', 'w', encoding='utf-8') as file1:
            json.dump(finall_spisok, file1, indent=4, ensure_ascii=False)
    elif word == 'Александр Сергеевич Глазков':
        with open('aleks.json', 'w', encoding='utf-8') as file2:
            json.dump(finall_spisok, file2, indent=4, ensure_ascii=False)
    else:
        with open('test.json', 'w', encoding='utf-8') as file:
            json.dump(finall_spisok, file, indent=4, ensure_ascii=False)

if __name__=='__main__':
    main()