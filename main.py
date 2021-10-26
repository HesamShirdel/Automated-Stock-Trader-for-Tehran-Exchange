from datetime import datetime
import time
import pandas as pd
from multiprocessing.dummy import Pool as ThreadPool
from urllib.request import Request, urlopen
import zlib
import random
import math
from utils import func

path_data = 'C:\\Users\Hesam\\Projects\\Test\\Adjusted'
# Edit the Money amount for each transaction 
money = 50000000
# Edit and add your Credentials 
user = ''
passw = ''

time_now = '8:59:59'
time_end = '12:23:00'

Dirooz = 'دیروز'
Naam = 'نماد'
Kharid = 'خرید - قیمت'
Forush = 'فروش - قیمت'
Payani = 'قیمت پایانی - درصد'


# **************************************************
# **************** BUY BOT *************************
# **************************************************

def base_buy(stock, target_price):
    print(' Started Monitoring ', stock, )

    list1 = []

    try:

        order_price = Price_list.loc[Price_list[Naam] == stock, Kharid].item()
        sell_price = Price_list.loc[Price_list[Naam] == stock, Forush].item()
        yesterday_price = Price_list.loc[Price_list[Naam] == stock, Dirooz].item()
    except:
        print('Can not get orderprice for Watch_Buy', stock)
        return 0

    if func.watch_buy(stock, target_price, order_price , sell_price , yesterday_price ) == 1:

        try:




            if sell_price < 1.025 * target_price:

                if sell_price < 1.05*math.floor(yesterday_price):
                    sell_price +=10

                sell_price_str = str(math.floor(sell_price))
                volume = round(money / (sell_price * 10)) * 10
                volume_str = str(volume)
                print('Buying ', stock)
                if func.buy(stock, sell_price_str, volume_str, money, user, passw) == 1:
                    list1.append(stock)
                    list1.append(sell_price)
                    list1.append(volume_str)

                    return list1
                else:

                    return

            else:
                print('Sellers Price is too high ', stock)

                return


        except:

            print('there was an error in Buying', stock)

            return
    else:

        time_now = datetime.now().strftime("%H:%M:%S")
        if time_now > time_end:
            print(' Market is Closed ', ' Closing ', stock)

            return

        print(' //////////Price is', 'not a match ', ' Closing ', stock)

        return




# **************************************************
# ****************** SELL BOT **********************
# **************************************************

def base_sell(stock, target_price, volume, MA):
    print(' Started Monitoring ', stock, )

    list1 = []
    try:
        yesterday_price = Price_list.loc[Price_list[Naam] == stock, Dirooz].item()
        last_price = Price_list.loc[Price_list[Naam] == stock, Payani].item()
        order_price = Price_list.loc[Price_list[Naam] == stock, Kharid].item()

    except:
        print('Can not get orderprice Watch', stock)
        return 0

    if func.watch_sell(stock, target_price, MA, order_price, yesterday_price, last_price) == 1:

        try:

            order_price_str = str(math.floor(order_price))

            print('Selling ', stock)
            func.sell(stock, order_price_str, volume, user, passw)
            list1.append(stock)
            list1.append(order_price)

            return list1





        except:
            print('Something Went Wrong with Selling', stock)

            return
    else:

        time_now = datetime.now().strftime("%H:%M:%S")
        if time_now > time_end:
            print(' Market is Closed ', ' Closing ', stock)

            return

        print(' Price is', 'not a match ', ' Closing ', stock)

        return


# **************************************************
# ****************** Main Code *********************
# **************************************************

# loop for checking market approximately every 2 minutes
while True:

    try:
        # getting online market data from Tehran Securities Exchange backend
        req = Request('http://members.tsetmc.com/tsev2/excel/MarketWatchPlus.aspx?d=0')
        req.add_header('User-Agent', 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:77.0) Gecko/20100101 Firefox/77.0')
        content = urlopen(req)

        decomp = zlib.decompress(content.read(), 16 + zlib.MAX_WBITS)

        Price_list = pd.read_excel(decomp)

        Price_list.columns = Price_list.iloc[1]

        Price_list = Price_list.drop([0, 1])

    except:
        print('server did not respond. could not get prices')

    data = pd.read_excel('Watchlist.xlsx')
    data.set_index('Name')
    bought = pd.read_excel('BuyList.xlsx')
    bought.set_index('Name')

    # removing Bought stock from active Watchlist
    for q in range(0, len(bought['Name']), 1):
        data.drop(data.loc[data['Name'] == bought['Name'][q]].index, inplace=True)

    # data.to_excel(r'Watchlist.xlsx', index = False)

    time_now = datetime.now().strftime("%H:%M:%S")
    if time_now > time_end:
        print('Market is closed')
        break

    # if you want faster execution you can increase threadpool number for parallel execution
    pool = ThreadPool(1)

    results = pool.starmap(base_buy, zip(data['Name'].tolist(), data['Target'].tolist()))

    pool.close()
    pool.join()
    results = [i for i in results if i is not None]
    results = [i for i in results if i != 0]
    print(results)

    # exporting Results to excel file
    if results:
        for i in range(0, len(results), 1):
            data.drop(data.loc[data['Name'] == results[i][0]].index, inplace=True)

        temp1 = pd.DataFrame(results, columns=['Name', 'Buy Price', 'Volume'])
        print(temp1)
        bought = pd.read_excel('BuyList.xlsx')
        bought = bought.append(temp1)
        bought.drop_duplicates(subset="Name", keep='first', inplace=True)
        print(bought)
        bought.to_excel(r'BuyList.xlsx', index=False)
        func.moving_average(path_data)

        ret = 1
        while ret < 5:
            try:
                temp3 = pd.read_excel('BuyList.xlsx')
                temp2 = pd.read_excel('SellList.xlsx')
                temp2 = temp2.append(temp3)
                temp2.drop_duplicates(subset="Name", keep='first', inplace=True)
                temp2.to_excel(r'SellList.xlsx', index=False)
                break

            except:
                time.sleep(8)
                ret += 1
                print('Can not Open write SellList. Retrying')

        del results
        print('New list of Monitoring:')
        print(data)

    # main code for Selling
    try:
        data = pd.read_excel('SellList.xlsx')
        data.set_index('Name')

        sold_df = pd.read_excel('Sold.xlsx')
        sold_df.set_index('Name')

        for q in range(0, len(sold_df['Name']), 1):
            data.drop(data.loc[data['Name'] == sold_df['Name'][q]].index, inplace=True)

        #data.to_excel(r'SellList.xlsx', index=False)

    except:
        pass
    # if you want faster execution you can increase threadpool number for parallel execution
    pool = ThreadPool(1)

    results1 = pool.starmap(base_sell,
                            zip(data['Name'].tolist(), data['Buy Price'].tolist(), data['Volume'].tolist(),
                                data['MA13'].tolist()))

    pool.close()
    pool.join()
    results1 = [i for i in results1 if i is not None]
    results1 = [i for i in results1 if i != 0]
    print(results1)

    # exporting Results to excel file
    if results1:
        for i in range(0, len(results1), 1):
            data.drop(data.loc[data['Name'] == results1[i][0]].index, inplace=True)
        temp = pd.DataFrame(results1, columns=['Name', 'Sell Price'])
        sold_df = pd.read_excel('Sold.xlsx')
        sold_df = sold_df.append(temp)
        sold_df.drop_duplicates(subset="Name", keep='first', inplace=True)
        sold_df.to_excel(r'Sold.xlsx', index=False)

    # waiting so IP is does not get blacklisted for sending too many calls on TSE server
    print('wait 1 - 60 seconds')
    time.sleep(60)
    print('wait 2 - 60 seconds')
    time.sleep(60)
    # print('wait 3 - 60 seconds')
    # time.sleep(60)
    Sleep_time = random.randint(0, 15)
    print('wait 3 - ', Sleep_time, ' secounds')
    time.sleep(Sleep_time)
