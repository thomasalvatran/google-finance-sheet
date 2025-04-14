This is short project to create a dynamic sheet to store the stocks symbol and change according to current price, closed price and open price. Finally find the lowest price in the last 3 months with the date. It is using google financing function GOOOLEFINANCING togther with javacript for parse date from yahoo to get the lowest of the price during the day not the closed price. With the help of AI ChatGPT on the Javascript. I retire 2 years my coding is rusty I did not do a lot of Javascript but C/C++ and PHP/Python with MySQL that the laguages I worked. Only a few functions of JS is using need to check on Work Sheet go to the cell double click to see which JS has been using for each cell. e.i =getTSLAMin3MonthsYahooLow(A2) A2 is cell number passing to JS function called getTSLAMin3MonthsYahooLow the data is from Yahoo Financing since Google only find the data in the sheet only we need to parse the data from outside such as Yahoo financing so we do not need to create the data in the working sheet must easy clean and simple otherwise the sheet will be very big with all the data from 3 months time. Please note Yahoo only provide the date in 1 hr time frame it does not create 1m or 5m time frame as the technical analyses that we see in the tradingview because if yahoo does every 1m the data will be huge when parse to Google Sheet see eexaple
![image](https://github.com/user-attachments/assets/a105cfbc-260e-423b-adfd-2680060f1ffb)
![image](https://github.com/user-attachments/assets/662e3806-8e3a-48bd-b0fc-ad7c18f924f1)
Cannot change the orders of project (pin) it is incovenient, I put here so github need to fix it in the future. ChatGPT-AI agreed with me. If I have the time will give the feedback to github to fix it and it should be a simple fix.

![image](https://github.com/user-attachments/assets/0a04d40b-d5f3-4a1b-bbd3-d68e7e57b58c)

You're right — when today is Monday, the formula might not work because GOOGLEFINANCE tries to get Friday's close, but Google Finance often doesn't return data for the most recent past trading day until later in the day (especially early Monday).

Let’s fix this by modifying your formula to check for the latest available close price within the last few days, not just relying on one day.

✅ Robust Fix:
Try this updated formula — it checks up to 5 days back to find the most recent valid close price:

excel
Copy
Edit
=INDEX(
  FILTER(
    GOOGLEFINANCE(A3, "close", TODAY()-5, TODAY()),
    NOT(ISNA(INDEX(GOOGLEFINANCE(A3, "close", TODAY()-5, TODAY()), , 2)))
  ),
  ROWS(FILTER(
    GOOGLEFINANCE(A3, "close", TODAY()-5, TODAY()),
    NOT(ISNA(INDEX(GOOGLEFINANCE(A3, "close", TODAY()-5, TODAY()), , 2)))
  )),
  2
)
🔍 What it does:
Fetches close prices for the last 5 days.

Filters out any rows where the close price is N/A.

Returns the most recent available close price (last valid row).

This should work reliably, even on Mondays, holidays, or early mornings when data hasn’t updated yet.

Want a simplified or optimized version too?










