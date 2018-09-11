import csv
import re
import io
import os
import xlsxwriter
import urllib.request
from urllib.parse import quote
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

PChome_search_URL = 'https://ecshweb.pchome.com.tw/search/v3.3/?q='


def get_web_page(url, query):
    # scroll_fuc 模擬 瀏覽器 往下滾動
    def scroll_fuc(driver):
        flag = 0
        limi = 2000
        count = 0
        while flag == 0:
            js = "var q=document.documentElement.scrollTop=%s" % limi
            driver.execute_script(js)
            if not count % 5:
                print(u'資料加載中...')
            try:
                the_element = EC.visibility_of_element_located((By.CSS_SELECTOR, 'div#cart.site_cart.fixed_position'))
                assert the_element(driver)
                flag = 0
                limi += 1500
                count += 1
            except:
                flag = 1

    query = quote(query, 'utf-8')
    options = webdriver.ChromeOptions()
    # """
    options.add_argument('headless')
    options.add_argument('disable-gpu')
    options.add_argument('window-size=1600,900')
    # """
    driver = webdriver.Chrome(chrome_options=options)
    source = None
    try:
        print(u'模擬 瀏覽器進入PCHome搜尋頁..')
        driver.implicitly_wait(10)
        driver.get(url + query + '&scope=24h')
        print(u'網頁讀取中..')
        scroll_fuc(driver)

        print(u'回傳解析')
        source = driver.page_source
        driver.quit()
    except WebDriverException:
        print('Cannot navigate to invalid URL !')
        driver.quit()
        exit()

    return source


def collect(coll_row, price):
    item = dict()
    item['name'] = coll_row.find('h5', 'prod_name').a.text
    item['name'] = re.sub("[:.]", '', item['name'])
    item['price'] = price
    item['link'] = 'https:' + coll_row.find('h5', 'prod_name').a['href']
    item['img_link'] = 'https:' + coll_row.find('dd', 'c1f').a.img['src']
    item['img_file'] = item['img_link'].split('/')[-1]
    data = urllib.request.urlopen(item['img_link']).read()
    file = open(item['img_file'], "wb")
    file.write(data)
    file.close()

    return item


# 取得收詢價格等資料
def get_items(dom, price_max, price_min):
    its = list()
    if not price_min:
        price_min = '0'
    soup = BeautifulSoup(dom, 'html5lib')
    rows = soup.find_all('dl', 'col3f')
    print(u'收集商品相關資料..')

    for row in rows:
        get_price = int(row.find('span', 'value').text)
        if not price_max or (price_max == '-1'):
            if get_price > int(price_min):
                coll_data = collect(row, get_price)
                its.append(coll_data)
        else:
            if (get_price > int(price_min)) and (get_price < int(price_max)):
                coll_data = collect(row, get_price)
                its.append(coll_data)

    return its


def write_excel(rows, name):
    workbook = xlsxwriter.Workbook('PChome_查價' + name + '.xlsx')  # 建立excel
    worksheet = workbook.add_worksheet(name)  # 建立sheet

    worksheet.set_column('A:A', 13)  # 设置A欄寬度為13
    worksheet.set_column('B:B', 50)
    worksheet.set_column('C:C', 7)
    worksheet.set_column('D:D', 60)
    worksheet.set_default_row(70)

    bold_fmt = workbook.add_format({'bold': True, 'valign': 'vcenter', 'align': 'center'})  # 設定格式為 '粗體字','重直置中','水平置中'
    c1_fmt = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bold': True})
    center_valign = workbook.add_format({'valign': 'vcenter'})
    center_align = workbook.add_format({'valign': 'vcenter', 'align': 'center'})

    worksheet.write('A1', u'商品圖片', bold_fmt)
    worksheet.write('B1', u'商品名稱', bold_fmt)
    worksheet.write('C1', U'價格', c1_fmt)
    worksheet.write('D1', U'連結', bold_fmt)

    row = 1
    col = 0
    for item in rows:
        worksheet.insert_image(row, col, item['img_file'], {'x_offset': 5, 'y_offset': 5, 'x_scale': 0.7, 'y_scale': 0.7})  # 插入圖片
        worksheet.write_string(row, col+1, item['name'], center_valign)
        worksheet.write_number(row, col+2, item['price'], center_align)
        worksheet.write_url(row, col+3, item['link'], center_valign)
        row += 1

    workbook.close()

    for del_file in rows:
        os.remove(del_file['img_file'])


if __name__ == '__main__':
    sName = input(u'輸入要搜尋PCHome的商品名稱:')
    uFunc = input(u'需要篩選價格請按 1 :')
    user_max = '-1'
    user_min = '0'
    if uFunc == '1':
        user_min = input(u'輸入最低價為(預設0):')
        user_max = input(u'輸入最高價為(可留空):')

    page = get_web_page(PChome_search_URL, sName)
    items = get_items(page, user_max, user_min)
    if not items:
        print(u'搜尋不到您要的商品！')
    else:
        new_item = sorted(items, key=lambda e: e.__getitem__('price'))
        max_price = max(item['price'] for item in items)
        min_price = min(item['price'] for item in items)
        print(u'價格:%s ~ %s' % (min_price, max_price))
        write_excel(new_item, sName)
        print(u"詳細查看:PChome_查價_%s.xlsx" % sName)

    print(u'按下Enter退出')
    input()
