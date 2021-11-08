# Automated Stock Trader for Tehran Exchange
This robot trades stock with a swing strategy using Mofid Securities brokrage platform

it is recommended to use Python platforms like Anaconda so that libraries are included by default
https://www.anaconda.com/products/individual
# Disclaimer: 
USE AT YOUR OWN RISK. Creators do not take any responsibility for your trading results

# Library requirements:
Pandas

Selenium (using Google Chrome)

# How do i run it?
create a chrome profile for the robot so that cookies are saved and captcha is avoided

1. Download market historical data (read next section)
2. Run watchlist_generator.py (or click Watchlist.bat in windows) to generate watchlist 
3. Run main.py (or click Main.bat in windows) to start trading (stocks from watchlist)

note: if you are using .bat files, CMD does not support farsi characters. you can change the font but even then it will still be shown as left to right language 

# How do i get Historical Market data for Iran stock Market?
use TSE Client

you can get it from here:
http://cdn.tsetmc.com/utils/getFile.aspx?Idn=16617



suggestions and feature requests are most welcome
