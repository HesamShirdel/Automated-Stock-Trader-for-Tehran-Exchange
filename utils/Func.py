#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import glob
import math


# ***********************************************************
# ************ Moving Average Calculator Function************
# ***********************************************************

def moving_average(path_data):
    path1 = path_data
    all_files = glob.glob(path1 + "/*.xls")
    buy1 = pd.read_excel('BuyList.xlsx')
    buy1.set_index('Name')
    buy1['MA13'] = 0

    k = 0
    for files in all_files:

        stock1 = pd.read_excel(files)
        stock1 = stock1.set_index(pd.DatetimeIndex(stock1['<DTYYYYMMDD>'].values))

        ma13 = pd.DataFrame()
        ma13['AM'] = stock1['<CLOSE>'].tail(30).rolling(window=13).mean()

        for i in range(0, len(buy1['Name']), 1):
            short = stock1['<TICKER>'][i]
            if buy1['Name'][i] == short[:len(short) - 2]:
                buy1.loc[i, 'MA13'] = ma13['AM'].iloc[-1]
                print('MA added to Buylist')
                k = k + 1
        if k >= len(buy1['Name']):
            break

    buy1.to_excel(r'BuyList.xlsx', index=False)


# ***********************************************************
# ************ Mofid Brokerage Buying Function***************
# ***********************************************************

def buy(stock, price, volume, money, user, passw):
    stock = stock.replace('ك', 'ک')
    stock = stock.replace('ي', 'ی')
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument(
        "--user-data-dir=C:\\Users\\Hesam\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 1")
    chrome_options.add_argument("--disable-extensions")
    driver = webdriver.Chrome('./chromedriver', options=chrome_options)
    driver.set_window_size(1920, 1080)
    driver.get("https://d.easytrader.emofid.com/")

    action = webdriver.ActionChains(driver)

    # login

    q = 1
    while q == 1:
        time.sleep(3)
        url = driver.current_url
        if url != 'https://d.easytrader.emofid.com/':
            if 'account.emofid.com/' in url:

                search_bar = driver.find_element_by_id("Username")
                search_bar.clear()
                search_bar.send_keys(user)

                pass_bar = driver.find_element_by_id("Password")
                pass_bar.clear()
                pass_bar.send_keys(passw)

                driver.implicitly_wait(20)
                time.sleep(7)
                driver.find_element_by_id("submit_btn").click()

            elif 'd.easytrader.emofid.com' in url:

                action.send_keys(Keys.ESCAPE).perform()

            else:
                print('something went wrong Login', stock)
                driver.close()
                driver.quit()
                return 0
        else:

            break

    driver.implicitly_wait(10)

    # check Balance
    time.sleep(3)
    balance = driver.find_element_by_xpath(
        '/html/body/app-root/main-layout/main/div[3]/div/div/div/market-data/div/div[2]/div/span[2]').text
    balance = balance.replace(",", "")
    balance = balance.replace(" ", "")
    balance = float(balance)

    if balance > (1.06 * money):

        # buy button
        action.send_keys(Keys.ESCAPE).perform()
        # time.sleep(2)
        driver.implicitly_wait(10)
        search = driver.find_element_by_xpath(
            '/html/body/app-root/main-layout/main/d-search-management/search-panel/div/div[1]/input')
        search.click()
        search.clear()
        search.send_keys(stock)
        driver.implicitly_wait(10)
        # time.sleep(2)
        driver.find_element_by_xpath(
            '/html/body/app-root/main-layout/main/d-search-management/search-panel/div/div[2]/div[2]/div/div[1]/div/div[contains(text(),stock)]').click()

        driver.implicitly_wait(10)
        # time.sleep(2)
        driver.find_element_by_xpath(
            '/html/body/app-root/main-layout/main/div[3]/div/div/as-split/as-split-area/app-layout-selector/app-layout2/as-split/as-split-area[2]/div/div[1]/div/button[1]').click()
        # time.sleep(2)

        driver.implicitly_wait(10)
        vol = driver.find_element_by_xpath(
            '/html/body/app-root/main-layout/main/div[3]/d-order-list/div/div[2]/div/order-form/div/div/form/div[1]/div[1]/div/dx-number-box/div/div[1]/input')
        vol.click()
        vol.clear()
        vol.send_keys(volume)
        # time.sleep(2)

        driver.implicitly_wait(10)
        pri = driver.find_element_by_xpath(
            '/html/body/app-root/main-layout/main/div[3]/d-order-list/div/div[2]/div/order-form/div/div/form/div[1]/div[2]/div/dx-number-box/div/div[1]/input')
        pri.click()
        pri.clear()
        pri.send_keys(price)
        # time.sleep(2)

        driver.implicitly_wait(10)
        # driver.find_element_by_xpath('/html/body/app-root/main-layout/main/div[3]/d-order-list/div/div[2]/div/order-form/div/div/form/div[3]/button[1]').click()
        pri.send_keys(Keys.ENTER)
        time.sleep(10)

        driver.close()
        driver.quit()

        return 1

    else:
        print('Balance is not Sufficient', stock)
        driver.close()
        driver.quit()
        return 0


# ************************************************************
# ************ Mofid Brokerage Selling Function***************
# ************************************************************

def sell(stock, price, volume, user, passw):
    stock = stock.replace('ك', 'ک')
    stock = stock.replace('ي', 'ی')
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument(
        "--user-data-dir=C:\\Users\\Hesam\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 1")
    chrome_options.add_argument("--disable-extensions")
    driver = webdriver.Chrome('./chromedriver', options=chrome_options)
    driver.set_window_size(1920, 1080)
    driver.get("https://d.easytrader.emofid.com/")

    action = webdriver.ActionChains(driver)

    # login

    q = 1
    while q == 1:
        time.sleep(3)
        url = driver.current_url
        if url != 'https://d.easytrader.emofid.com/':
            if 'account.emofid.com/Login' in url:

                search_bar = driver.find_element_by_id("Username")
                search_bar.clear()
                search_bar.send_keys(user)

                pass_bar = driver.find_element_by_id("Password")
                pass_bar.clear()
                pass_bar.send_keys(passw)

                driver.implicitly_wait(20)
                time.sleep(7)
                driver.find_element_by_id("submit_btn").click()

            elif 'd.easytrader.emofid.com' in url:

                action.send_keys(Keys.ESCAPE).perform()

            else:
                print('something went wrong Login', stock)
                driver.close()
                driver.quit()
                return 0
        else:

            break

    driver.implicitly_wait(10)

    # stock Search
    # time.sleep(2)
    action.send_keys(Keys.ESCAPE).perform()
    # time.sleep(2)
    driver.implicitly_wait(10)
    search = driver.find_element_by_xpath(
        '/html/body/app-root/main-layout/main/d-search-management/search-panel/div/div[1]/input')
    search.click()
    search.clear()
    search.send_keys(stock)
    driver.implicitly_wait(10)
    # time.sleep(2)
    driver.find_element_by_xpath(
        '/html/body/app-root/main-layout/main/d-search-management/search-panel/div/div[2]/div[2]/div/div[1]/div/div[contains(text(),stock)]').click()

    # Sell form
    # time.sleep(2)
    driver.find_element_by_xpath(
        '/html/body/app-root/main-layout/main/div[3]/div/div/as-split/as-split-area/app-layout-selector/app-layout2/as-split/as-split-area[2]/div/div[1]/div/button[2]').click()
    # time.sleep(2)

    # volume
    vol = driver.find_element_by_xpath(
        '/html/body/app-root/main-layout/main/div[3]/d-order-list/div/div[2]/div[1]/order-form/div/div/form/div[1]/div[1]/div/dx-number-box/div/div[1]/input')
    vol.click()
    vol.clear()
    vol.send_keys(volume)
    # time.sleep(2)

    # price
    pri = driver.find_element_by_xpath(
        '/html/body/app-root/main-layout/main/div[3]/d-order-list/div/div[2]/div[1]/order-form/div/div/form/div[1]/div[2]/div/dx-number-box/div/div[1]/input')
    pri.click()
    pri.clear()
    pri.send_keys(price)
    # time.sleep(2)

    # Sell Button
    # driver.find_element_by_xpath('/html/body/app-root/main-layout/main/div[3]/d-order-list/div/div[2]/div[1]/order-form/div/div/form/div[3]/button[1]').click()
    pri.send_keys(Keys.ENTER)
    time.sleep(10)

    driver.close()
    driver.quit()

    return


# ************************************************************
# ************ Buyers Power Calculator from TSETMC ***********
# ************************************************************

def power_tse(stock):

    c_options = webdriver.ChromeOptions()
    c_options.add_argument("headless")
    # c_options.add_argument("disable-gpu")
    c_options.add_argument(
        "--user-data-dir=C:\\Users\\Hesam\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 1")
    c_options.add_argument("--disable-extensions")
    chdriver = webdriver.Chrome(executable_path='./chromedriver.exe', options=c_options)
    chdriver.set_window_size(1920, 1080)

    retries = 1
    while retries <= 5:

        try:

            chdriver.get("http://www.tsetmc.com/Loader.aspx?ParTree=15")
            break

        except:

            time.sleep(5)
            chdriver.refresh()
            retries += 1
            print('Can not Open TSETMC. Retrying', stock)

    try:
        chdriver.implicitly_wait(5)
        chdriver.execute_script('ShowSearchWindow()')
        # chdriver.find_element_by_xpath('/html/body/div[3]/div[2]/a[5]').click()
        chdriver.implicitly_wait(5)
        chdriver.find_element_by_xpath('/html/body/div[5]/section/div/input').click()
        chdriver.implicitly_wait(5)
        chdriver.find_element_by_xpath('/html/body/div[5]/section/div/input').send_keys(stock)
        # chdriver.find_element_by_id('SearchKey').send_keys(keys.RETURN)

        chdriver.find_element_by_xpath(
            '/html/body/div[5]/section/div/div/div/div[2]/table/tbody/tr[1]/td[1]/a[contains(text(),stock)]').click()
        chdriver.implicitly_wait(30)
        time.sleep(15)

        hagigi_buy = chdriver.find_element_by_xpath(
            '/html/body/div[4]/form/div[3]/div[2]/div[1]/div[4]/div[1]/table/tbody/tr[2]/td[2]/div[1]').text
        hagigi_buy = hagigi_buy.replace(",", "")
        hagigi_buy = hagigi_buy.replace("M", "")
        hagigi_buy_flo = float(hagigi_buy)

        hagigi_sell = chdriver.find_element_by_xpath(
            '/html/body/div[4]/form/div[3]/div[2]/div[1]/div[4]/div[1]/table/tbody/tr[2]/td[3]/div[1]').text
        hagigi_sell = hagigi_sell.replace(",", "")
        hagigi_sell = hagigi_sell.replace("M", "")
        hagigi_sell_flo = float(hagigi_sell)

        td_buy = chdriver.find_element_by_id('e5').text
        td_buy = td_buy.replace(",", "")
        td_buy_flo = float(td_buy)

        td_sell = chdriver.find_element_by_id('e8').text
        td_sell = td_sell.replace(",", "")
        td_sell_flo = float(td_sell)

        # print(hagigi_buy_flo/td_buy_flo)
        # print(0.9*(hagigi_sell_flo/td_sell_flo))

        if (hagigi_buy_flo / td_buy_flo) > 0.9 * (hagigi_sell_flo / td_sell_flo):

            chdriver.close()
            chdriver.quit()
            return 1
        else:
            chdriver.close()
            chdriver.quit()
            print('Buy to Sell Ratio is not Optimal', stock)
            return 0

    except:
        chdriver.close()
        chdriver.quit()
        print('Can not Calculate Buy to Sell Ratio Watch', stock)
        return 1


# *************************************************************
# ************ Price Monitor For Buying Function***************
# *************************************************************

def watch_buy(stock, target, order_price, sell_price, yesterday_price):
    if 1.008 * target <= order_price <= 1.02 * target:
        if 10 < sell_price < math.floor(1.05 * yesterday_price):

#            if power_tse(stock) == 1:
#                return 1
            return 1
        else:
            print('Buy Queue is full ', stock)
    else:
        return 0

# ************************************************************
# ************ Price Monitor For Selling Function*************
# ************************************************************

def watch_sell(stock, target, ma, order_price, yesterday_price, last_price):
    if order_price != 0 and order_price > math.ceil(0.95 * yesterday_price):

        if order_price < 0.988 * ma:
            return 1

        elif order_price <= 0.975 * target:

            if last_price <= 2.3:
                return 1
            else:
                return 0

        elif order_price >= 1.03 * target:

            try:
                temp1 = pd.read_excel('SellList.xlsx')
                temp1.loc[temp1['Name'] == stock, 'Buy Price'] = round(order_price)
                temp1.to_excel(r'SellList.xlsx', index=False)
                print(stock, ' target Price Updated ')
                return 0

            except:
                print(stock, ' There was an Error with Updating target Price ')
                return 0

    else:
        print('safe forush Watch', stock)
        return 0
