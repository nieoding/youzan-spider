# encoding: utf-8
from __future__ import unicode_literals
from __future__ import print_function
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
import xlwt

FOREVER = 9999
URL_LOGIN = 'https://account.youzan.com/login'
URL_LIST = 'https://www.youzan.com/v2/shop/list'
URL_HOME = 'https://www.youzan.com/v4/dashboard'
URL_API_LIST = 'https://www.youzan.com/v4/trade/order/getList.json?p={}&page_size=50'
URL_DETAIL = 'https://www.youzan.com/v4/trade/order/detail?orderNo={}'


def init():
    driver = webdriver.Chrome()
    return driver


def login(driver):
    driver.get(URL_LOGIN)
    print('wait for loggin.')
    wait = WebDriverWait(driver, FOREVER)
    wait.until(
        EC.url_contains(URL_LIST)
    )
    print('login ok.')
    return driver.get_cookie('user_nickname')


def get_username(driver):
    return driver.get_cookie('user_nickname')['value']


def get_shop_list(driver):
    assert(driver.current_url.startswith(URL_LIST))
    rs = []
    for item in driver.find_elements_by_xpath("//ul[@class='dp-list']//li[starts-with(@class,'dp-item')]"):
        info = item.find_element_by_tag_name('p').text
        if '未认证' == info.split('：')[1]:
            continue
        rs.append(item.get_attribute('title'))
    return rs


def goto_shoplist(driver):
    driver.get(URL_LIST)
    print('wait for loading...')
    wait = WebDriverWait(driver, FOREVER)
    wait.until(
        EC.url_contains(URL_LIST)
    )
    print('open completed.')


def goto_shop(driver, name):
    assert(driver.current_url.startswith(URL_LIST))
    for item in driver.find_elements_by_xpath("//ul[@class='dp-list']//li[starts-with(@class,'dp-item')]"):
        if item.get_attribute('title') == name:
            item.click()
            wait = WebDriverWait(driver, FOREVER)
            wait.until(
                EC.url_contains(URL_HOME)
            )
            return


def get_order_list(driver):
    cookies = {}
    rs = []
    for item in driver.get_cookies():
        cookies[item['name']] = item['value']
    print('request page 1...')
    r = requests.get(URL_API_LIST.format(1),
                     cookies=cookies)
    data = r.json()['data']
    for item in data['list']:
        rs.append(item)
    total_count = data['totalItems']
    page_size = data['pageSize']
    total_page = int(total_count/page_size)
    if total_count % page_size:
        total_page += 1
    print('total_page:', total_page)
    for i in range(1, total_page):
        page_index = i + 1
        print('request page {}...'.format(page_index))
        r = requests.get(URL_API_LIST.format(page_index),
                         cookies=cookies)
        data = r.json()['data']
        for item in data['list']:
            rs.append(item)
    assert(len(rs) == total_count)
    return rs


def parse_order(source):
    item = source['items'][0]
    return (
        item['orderNo'],
        source['userName'],
        source['tel'],
        ' '.join((source['province'], source['city'], source['county'], source['addressDetail'])),
        source['realPay'],
        source['customer'],
        source['buyerMsg'],
        item['title'],
        item['num'],
        item['price'],
        source['stateStr'],
        source['createTime'],
        source.get('payTime', ''),
        source['shopName'],
        source.get('innerTransactionNumber', ''),
    )


def write_excel(orders, sheet_name, filename):
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet(sheet_name)
    sheet.write(0, 0, '订单号')
    sheet.write(0, 1, '收货人')
    sheet.write(0, 2, '联系电话')
    sheet.write(0, 3, '收货地址')
    sheet.write(0, 4, '实付金额')
    sheet.write(0, 5, '买家')
    sheet.write(0, 6, '买家留言')
    sheet.write(0, 7, '商品名称')
    sheet.write(0, 8, '商品数量')
    sheet.write(0, 9, '单价')
    sheet.write(0, 10, '订单状态')
    sheet.write(0, 11, '下单时间')
    sheet.write(0, 12, '付款时间')
    sheet.write(0, 13, '归属店铺')
    sheet.write(0, 14, '支付流水号')
    i = 0
    for order in orders:
        cell = parse_order(order)
        i += 1
        total = len(cell)
        for j in range(total):
            sheet.write(i, j, cell[j])
    book.save(filename)
