import pandas as pd
from datetime import datetime
import glob
import os

#Edit the Path to your Historical Market data folder
path = 'C:\\Users\Hesam\\Projects\\Test\\Adjusted'
allFiles = glob.glob(path + "/*.xls")

for files in allFiles:
    # deleting faulty data
    file = open(files)

    file.seek(0, os.SEEK_END)

    if file.tell() < 8192:
        file.close()
        os.remove(files)

allFiles = glob.glob(path + "/*.xls")

N = {}

results = pd.DataFrame()

Buy = pd.read_excel('BuyList.xlsx')
Buy.set_index('Name')
Buy['MA13'] = 0

Sell = pd.read_excel('SellList.xlsx')
Sell.set_index('Name')
Sell['MA13'] = 0

# \\\\\\\\\\\\\\\\\\ Historical Performance Generator /////////////

Performance = pd.read_excel('Performance.xlsx')

sold = pd.read_excel('sold.xlsx')
sold.set_index('Name')
Per_temp = pd.DataFrame(columns=['Date', 'Name', 'Sell Price', 'Buy Price', 'Percentage'])
Per_temp = Per_temp.append(sold)
Per_temp.drop_duplicates(subset="Name", keep='first', inplace=True)

for j in range(0, len(Buy['Name']), 1):
    Per_temp.loc[Per_temp['Name'] == Buy['Name'][j], 'Buy Price'] = Buy['Buy Price'][j]

for j in range(0, len(Per_temp['Name']), 1):
    Per_temp['Percentage'][j] = (((Per_temp['Sell Price'][j] - Per_temp['Buy Price'][j]) / Per_temp['Buy Price'][
        j]) * 100) - 1.5
    Per_temp['Date'][j] = datetime.now().strftime('%Y-%m-%d')

Performance = Performance.append(Per_temp)
Performance.to_excel(r'Performance.xlsx', index=False)

# \\\\\\\\\\ Delete sold from buylist ////////


for q in range(0, len(sold['Name']), 1):
    Buy.drop(Buy.loc[Buy['Name'] == sold['Name'][q]].index, inplace=True)
    Sell.drop(Sell.loc[Sell['Name'] == sold['Name'][q]].index, inplace=True)

Buy.to_excel(r'BuyList.xlsx', index=False)
Sell.to_excel(r'SellList.xlsx', index=False)

# \\\\\\\\\\\ Check every Stock //////////

for files in allFiles:

    temp = pd.read_excel(files)
    if temp['<DTYYYYMMDD>'].iloc[-1] > 20210531:
        stock = pd.read_excel(files)
        stock = stock.set_index(pd.DatetimeIndex(stock['<DTYYYYMMDD>'].values))

        ma8 = pd.DataFrame()
        ma8['AM'] = stock['<CLOSE>'].rolling(window=8).mean()
        ma13 = pd.DataFrame()
        ma13['AM'] = stock['<CLOSE>'].rolling(window=13).mean()
        ma21 = pd.DataFrame()
        ma21['AM'] = stock['<CLOSE>'].rolling(window=21).mean()

        # \\\\\\\\\\\\\ Moving average ////////////

        for i in range(0, len(Sell['Name']), 1):
            short = stock['<TICKER>'][i]
            if Sell['Name'][i] == short[:len(short) - 2]:
                Sell.loc[i, 'MA13'] = ma13['AM'].iloc[-1]
                print('MA added to Selllist', Sell['Name'][i])

        exp1 = stock['<CLOSE>'].ewm(span=12, adjust=False).mean()
        exp2 = stock['<CLOSE>'].ewm(span=26, adjust=False).mean()

        stock['macd'] = exp1 - exp2
        cl = stock['<CLOSE>'].iloc[-1]

        mm = stock['macd'].tail(45).max()
        if abs(stock['macd'].tail(45).min()) > mm:
            mm = abs(stock['macd'].tail(45).min())


        t = -1
        data = stock
        results = pd.DataFrame()
        for i in range(len(data) - 5, len(data) - 14, -1):
            if t != 0:

                if (data['<HIGH>'][i + 2] >= data['<HIGH>'][i]) and (
                        data['<HIGH>'][i + 2] >= data['<HIGH>'][i + 1]) and (
                        data['<HIGH>'][i + 2] >= data['<HIGH>'][i + 3]) and (
                        data['<HIGH>'][i + 2] >= data['<HIGH>'][i + 4]):  # find up fractal
                    if (data['<HIGH>'][i + 2] > ma21['AM'][i + 2]) and (data['<HIGH>'][i + 2] > ma13['AM'][i + 2]) and (
                            data['<HIGH>'][i + 2] > ma8['AM'][i + 2]):  # check if it is over the Alligators
                        t = 0
                        if abs(data['macd'].iloc[-1]) < (0.19 * mm):  # check macd
                            if (data['<HIGH>'][i + 2] > cl) and (data['<HIGH>'][i + 2] < 1.049 * cl) and (
                                    data['<HIGH>'][i + 2] < 99999):  # check if it will be in range next day
                                if (data['<HIGH>'][i + 4] - data['<LOW>'][i + 4] >= 0.01 * data['<CLOSE>'][i + 4]) or (data['<HIGH>'][i + 3] - data['<LOW>'][i + 3] >= 0.01 * data['<CLOSE>'][i + 3]) or (data['<HIGH>'][i + 2] - data['<LOW>'][i + 2] >= 0.01 * data['<CLOSE>'][i + 2]):  # check if the stock has movement//////
                                    print(data['<TICKER>'][i + 2], data['<HIGH>'][i + 2])
                                    sname = data['<TICKER>'][i + 2]

                                    if not any(chrt.isdigit() for chrt in sname):
                                        sname = sname[:len(sname) - 2]
                                        if sname[-1] != 'ح':
                                            N[sname] = data['<HIGH>'][i + 2]

    # except:
    #   pass

results = pd.DataFrame(list(N.items()), columns=['Name', 'Target'])
results.set_index('Name')

ret = 1
while ret < 8:
    try:
        for k in range(0, len(Buy['Name']), 1):
            # if Buy['Name'][l] in results :

            results.drop(results.loc[results['Name'] == Buy['Name'][k]].index, inplace=True)
        break

    except:
        ret += 1
        print('that error again')

print(results)

results.to_excel(r'Watchlist.xlsx', index=False)
Buy.to_excel(r'BuyList.xlsx', index=False)
Sell.to_excel(r'SellList.xlsx', index=False)
sold_tmp = pd.DataFrame(columns=['Name', 'Sell Price'])
sold_tmp.to_excel(r'sold.xlsx', index=False)
