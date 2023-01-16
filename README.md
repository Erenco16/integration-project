# integration-project
I wrote a Python script which updates the prices on dionaks.com according to the data on evan.com.tr using Selenium and Pandas.

First I extract both .xlsx files from evan.com.tr's and dionaks.com's admin panel. Then, I compare these two excel files and extract a new excel file with the price and rebate data from evan.com.tr. Finally, I submit an .xls file to the website using the same admin panel again.
