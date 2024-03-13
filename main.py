from playwright.sync_api import sync_playwright, Playwright
from fake_useragent import UserAgent as ua
import requests as r
import json
from parsel import Selector as sel
import csv
import os.path
import pandas as pd


class ParserOzon:
    def __init__(self):
        pass
        
    def get_items(self, url, longtide, latitude):
        self.url = url
        self.longtide = longtide
        self.latitude = latitude
        urls = []
        with sync_playwright() as pw:
            print("Ждите процесс парсинга...")
            browser = pw.chromium.launch(executable_path="pw-browsers\chromium-1097\chrome-win\chrome.exe")    
            iph = pw.devices['iPhone 11']
            options = {"geolocation": {"longitude": self.longtide, "latitude": self.latitude},
                       "permissions": ["geolocation"]}
            context = browser.new_context(**iph, **options)
            page = context.new_page()
            page.goto(url)
            countItem = 1
            if page.wait_for_selector("#layoutPage > div.b2.b4 > div.v8x.xv9 > div").is_visible():
                page.wait_for_selector("#layoutPage > div.b2.b4 > div.v8x.xv9 > div")
                data = sel(page.content())
                countElemsInfo = str(data.xpath('//*/div[contains(@data-widget, "fulltextResultsHeader")]/div/text()[contains(string(), "товар")]').get()) 
                countElems = ""
                if countElemsInfo == None:
                    pass
                    print("по вашему запросу ничего ")
                else:
                    for i in range(0, len(countElemsInfo)):
                        if countElemsInfo[i].isdigit() == True:
                            countElems += countElemsInfo[i]
                if int(countElems) <= 36:
                    while countItem < int(countElems):
                        while page.locator(f"#paginatorContent > div > div > div:nth-child({countItem})").is_visible() is False:
                            page.mouse.wheel(0, 300)
                            page.wait_for_selector(f"#paginatorContent > div > div > div:nth-child({countItem}) > div > div.v3i > a")
                        elem = page.query_selector(f"#paginatorContent > div > div > div:nth-child({countItem}) > div > div.v3i > a")
                        urls.append(elem.get_attribute("href"))
                        countItem += 1
                else:       
                    while page.wait_for_selector("#paginatorContent > div > div > div:nth-child(1) > div > div.v3i > a").is_visible() is False:
                        while page.locator(f"#paginatorContent > div > div > div:nth-child({countItem})").is_visible() is False:
                            page.mouse.wheel(0, 300)
                            page.wait_for_selector(f"#paginatorContent > div > div > div:nth-child({countItem}) > div > div.v3i > a")
                        elem = page.query_selector(f"#paginatorContent > div > div > div:nth-child({countItem}) > div > div.v3i > a")
                        countElems = len(page.query_selector_all("#paginatorContent > div > div > div > div > div > a"))
                        if elem is None:
                            break
                        else:
                            urls.append(elem.get_attribute("href"))
                        countItem += 1
                        if countItem > countElems:
                            break
            browser.close()
            print("Ссылки собраны.")
            return urls
            
        
        
    def get_descr(self, links):
        self.links = links
        for i in range(0 ,len(links)-1):
            iTem = "https://www.ozon.ru" + links[i].split("?")[0] + "features/"
            with sync_playwright() as pw:
                browser = pw.chromium.launch(executable_path="pw-browsers\chromium-1097\chrome-win\chrome.exe")
                dev = pw.devices['iPhone 11']
                context = browser.new_context(**dev)
                page = context.new_page()
                page.goto(iTem)
                page.wait_for_load_state("domcontentloaded")
                data = sel(page.content())
                price = str(data.xpath('//button/span/div/div/div/span/span/text()').get()).replace('\u2009', "")
                lInk = "https://www.ozon.ru" + links[i].split("?")[0]
                file_writer.writerow({"Цена": price, "Город": city, "Запрос": query, "Ссылка": lInk})
                print("Записаны данные по ссылке " + iTem)
                browser.close()
            
            
            
if __name__ == "__main__":
    print("Чтобы прекратить парсинг нажмите клавишу 1. Для просмотра полученных данных остановите работу программы...")
    pathRead = input("Введите путь к xlsx-файлу с названиями товаров и городами: ")
    maxCountPages = int(input("Введите макс. кол-во страниц: "))
    cities = []
    queris = []
    if not os.path.exists(pathRead):
        print("Файл не существует")
        exit()
    else:
        inp_data = pd.ExcelFile(pathRead)
        for sheet_name in ["товары", "Лист2"]:
            df = pd.read_excel(inp_data, sheet_name=sheet_name)
            data = pd.DataFrame(df).values.tolist()
            if sheet_name == 'товары':
                queris.extend(data[:])
            elif sheet_name == 'Лист2':
                cities.extend(data[:])
        for city in cities:
            resp = r.get(f'https://photon.komoot.io/api/?q={city[0]}&limit=1&lang=de', headers={'user-agent': f'{ua().random}'})
            longtide = resp.json()['features'][0]['geometry']['coordinates'][0]
            latitude = resp.json()['features'][0]['geometry']['coordinates'][1]
            try:
                if not os.path.exists("Парсинг Ozon.csv"):
                    with open("Парсинг Ozon.csv", "w", encoding="utf-8") as f:
                        names = ["Цена", "Город", "Запрос", "Ссылка"]
                        file_writer = csv.DictWriter(f, fieldnames=names, lineterminator="\r")
                        file_writer.writeheader()
            except:
                print("Файл уже существует")
                pass
            finally:
                for query in queris:
                    with open("Парсинг Ozon.csv", "a", encoding="utf-8") as f:
                        pageCount = 1
                        while pageCount <= maxCountPages:
                            file_writer = csv.DictWriter(f, fieldnames=["Цена", "Город", "Запрос", "Ссылка"], lineterminator="\r")
                            url = f"https://www.ozon.ru/search/?text={query[0]}&from_global=true&page={pageCount}"
                            html = ParserOzon().get_items(url, longtide, latitude)
                            if html == []:
                                pass
                            else:
                                ParserOzon().get_descr(html)
                            pageCount += 1
                


