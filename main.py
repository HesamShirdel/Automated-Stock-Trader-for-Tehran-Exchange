from utils.Functions import robot
from datetime import datetime
import time
import pandas as pd
from multiprocessing.dummy import Pool as ThreadPool
from urllib.request import Request, urlopen
import zlib
import random
import math



#////////////////////////////////////////////////////////
#////////////// IMPORTANT : EDIT THIS SECTION  /////////
#//////////////////////////////////////////////////////

path_data='Path to stock market Data, example C:\\Users\\Hesam\\Projects\\Test'

# how much money would you like to invest in each trade? edit money (IR Rial)
money=50000000

# your credentials for login in Mofid Trader website
user='USERNAME'
passw='PASSWORD'
#////////////////////////////////////////////////////////
#///////////////////////////////////////////////////////
#//////////////////////////////////////////////////////


# value initiation

time_now ='8:59:59'
time_end ='12:23:00'


Dirooz='دیروز'
Naam='نماد'
Kharid='خرید - قیمت'
Forush='فروش - قیمت'
Payani='قیمت پایانی - درصد'

#**************************************************
#**************** BUY BOT *************************
#**************************************************
def EntryPriceBuy(Stock,TargetPrice):
   
    print(' Started Monitoring ', Stock,)
    
       
    list1=[]
    
    
    try:
        
        order_price = Price_list.loc[Price_list[Naam]==Stock , Kharid].item()
        
    except:
        print('Can not get orderprice for Watch_Buy', Stock)
        return 0 
    
    if robot.Watch_Buy(TargetPrice,order_price)==1:

        
        try:

            sell_price = Price_list.loc[Price_list[Naam]==Stock , Forush].item()
            sell_price_str=str(math.floor(sell_price))  
            print(sell_price_str)

            if sell_price>10 and sell_price<= math.floor(1.05*(Price_list.loc[Price_list[Naam]==Stock , Dirooz].item())): 
                if sell_price<1.025*TargetPrice:

                    volume =  round(money/(sell_price*10))*10
                    volume_str=str(volume)
                    print('Buying ', Stock)
                    if robot.Buy (Stock,sell_price_str,volume_str)==1:
                        list1.append(Stock)
                        list1.append(sell_price)
                        list1.append(volume_str)

                        return list1
                    else:

                        return

                else:
                    print('Sellers Price is too high ', Stock)

                    return 
            else:
                print('Buy Queue is full ', Stock)

        except:

            print('there was an error in Buying', Stock)

            return
    else:

        
        time_now=datetime.now().strftime("%H:%M:%S")
        if time_now>time_end:
            
            print(' Market is Closed ',' Closing ', Stock)

            return



        print(' //////////Price is', 'not a match ',' Closing ', Stock)

        return
    
    
    return



#**************************************************
#****************** SELL BOT **********************
#**************************************************

def EntryPriceSell(Stock,TargetPrice,Volume,MA):
   
    print(' Started Monitoring ', Stock,)
         
    
    list1=[]
    try: 
        
        order_price = Price_list.loc[Price_list[Naam]==Stock , Kharid].item()

    except:
        print('Can not get orderprice Watch', Stock)
        return 0  
    
    
    if robot.Watch_Sell(TargetPrice,MA,order_price)==1:

        try:

            order_price = Price_list.loc[Price_list[Naam]==Stock , Kharid].item()
            order_price_str=str(math.floor(order_price))


            print('Selling ', Stock)
            robot.Sell (Stock,order_price_str,Volume)
            list1.append(Stock)
            list1.append(order_price)


            return list1


        except:
            print('Something Went Wrong with Selling', Stock)

            return
    else:

        
        time_now=datetime.now().strftime("%H:%M:%S")
        if time_now>time_end:
            
            print(' Market is Closed ',' Closing ', Stock)

            return


        print(' Price is', 'not a match ',' Closing ', Stock)

        return
    
    
    return





#/////////////////////// Main Code \\\\\\\\\\\\\\\\\\\\\\\\



#loop for checking market approximately every 2 minutes
z=0
while z!=1:   
    
    #getting online market data from Tehran Securities Exchange backend
    try:
        
        req = Request('http://members.tsetmc.com/tsev2/excel/MarketWatchPlus.aspx?d=0')
        req.add_header('User-Agent', 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:77.0) Gecko/20100101 Firefox/77.0')
        content = urlopen(req)

        decomp =zlib.decompress(content.read(), 16+zlib.MAX_WBITS)


        Price_list=pd.read_excel(decomp)

        Price_list.columns = Price_list.iloc[1]

        Price_list=Price_list.drop([0,1])
        
    except:
        print('server did not respond. could not get prices')
    
    data=pd.read_excel('Watchlist.xlsx')
    data.set_index('Name')
    bought= pd.read_excel('BuyList.xlsx')
    bought.set_index('Name')
    
    #removing Bought stock from active Watchlist
    for l in range(0,len(bought['Name']),1):

            data.drop(data.loc[data['Name']==bought['Name'][l]].index, inplace=True)


    
    
    time_now=datetime.now().strftime("%H:%M:%S")
    if time_now>time_end:
        z=1;
    
    #if you want faster execution you can increase threadpool number for parallel execution
    pool = ThreadPool(1)
    
    #running Buy function
    results = pool.starmap(EntryPriceBuy, zip(data['Name'].tolist(), data['Target'].tolist()))

    pool.close()
    pool.join()
    results = [i for i in results if i is not None]
    print(results)
   
    #exporting Results to excel file
    if results:
        for i in range(0,len(results),1):
            data.drop(data.loc[data['Name']==results[i][0]].index, inplace=True)
            
        temp1=pd.DataFrame(results , columns = ['Name','Buy Price','Volume']) 
        print(temp1)
        bought= pd.read_excel('BuyList.xlsx')
        bought=bought.append(temp1)
        #print("after append",bought)
        bought.drop_duplicates(subset ="Name",keep = 'first', inplace = True)
        print(bought)
        bought.to_excel(r'BuyList.xlsx', index = False)
        Moving_Average()
        
        ret=1
        while ret<5:
            try:
                temp3=pd.read_excel('BuyList.xlsx')
                temp2=pd.read_excel('SellList.xlsx')
                temp2=temp2.append(temp3)
                temp2.drop_duplicates(subset ="Name",keep = 'first', inplace = True)
                temp2.to_excel(r'SellList.xlsx', index = False)
                break
                
            except:
                time.sleep(8)
                ret +=1
                print('Can not Open write SellList. Retrying', Stock)
        
        del results
        print('New list of Monitoring:')
        print(data)
        
    
    
    try:
        data=pd.read_excel('SellList.xlsx')
        data.set_index('Name')

        sold_df=pd.read_excel('Sold.xlsx')
        sold_df.set_index('Name')
        
        #removing sold stock from active Selllist
        for l in range(0,len(sold_df['Name']),1):

            data.drop(data.loc[data['Name']==sold_df['Name'][l]].index, inplace=True)


        data.to_excel(r'SellList.xlsx', index = False)

        

                    
    except:
        pass
    
    #if you want faster execution you can increase threadpool number for parallel execution
    pool = ThreadPool(1)

    results1 = pool.starmap(EntryPriceSell, zip(data['Name'].tolist(), data['Buy Price'].tolist(),data['Volume'].tolist(),data['MA13'].tolist()))

    pool.close()
    pool.join()
    results1 = [i for i in results1 if i is not None]
    print(results1)
    #exporting Results to excel file
    if results1:
        for i in range(0,len(results1),1):
            data.drop(data.loc[data['Name']==results1[i][0]].index, inplace=True)
        temp = pd.DataFrame(results1 , columns = ['Name','Sell Price'])
        sold_df=pd.read_excel('Sold.xlsx')
        sold_df=sold_df.append(temp)
        sold_df.drop_duplicates(subset ="Name",keep = 'first', inplace = True)
        sold_df.to_excel(r'Sold.xlsx', index = False)
    
    
    
    
    # waiting so IP is does not get blacklisted for sending too many calls on TSE server
    print('wait 1 - 60 secounds')
    time.sleep(60)
    print('wait 2 - 60 secounds')
    time.sleep(60)
    #print('wait 3 - 60 secounds')
    #time.sleep(60)
    Sleep_time=random.randint(0,15)
    print('wait 3 - ', Sleep_time , ' secounds')
    time.sleep(Sleep_time)
   
    

    


print('Purchased Stock:')
print(bought)



