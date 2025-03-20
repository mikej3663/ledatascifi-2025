# Open Asset Pricing

An unreal project that contains info on portfolio returns for hundreds of asset pricing anomalies (at the monthly or daily level) and 
stock-level signals. 

Using this dataset, you can replicate many classic asset pricing papers. 

You can also hunt for ways to use signals to create portfolios that make money. 

You should check out:
1. The website for the project: https://www.openassetpricing.com, which includes 
    - info on the Data page
    - a list of [signals](https://drive.google.com/file/d/1Sev9s6cPFUGgxp1pFiej0lGzpsMqJCI2/view) and info about them 
    - Code they used to make the dataset
    - Sample code showing some ways to use the data
    - A partial list of [academic studies  using the dataset](https://www.openassetpricing.com/featured-in/). Looking at this will give you more ideas about what's possible
1. [The python package that downloads the data in python](https://github.com/mk0417/open-asset-pricing-download). This makes it easy to use. Some examples of using this package:
    1.  START HERE, a must: [https://github.com/mk0417/open-asset-pricing-download/blob/master/examples/quick_tour.ipynb](quick tour)
    2.  [Combining it with stock price returns from CRSP](https://github.com/mk0417/open-asset-pricing-download/blob/master/examples/merge_signals_with_crsp.ipynb) (where pros get stock returns, not yfinance): 
    3.  [Using with Fama French factors](https://github.com/mk0417/open-asset-pricing-download/blob/master/examples/merge_portfolios_with_ff_factors.ipynb) and testing portfolio anamolies
    4.  [:star: Machine Learning example :star:](https://github.com/mk0417/open-asset-pricing-download/blob/master/examples/ML_portfolio_example.ipynb)
    5.  The OpenAP_anomaly_plot.ipynb file in this folder. [Link that might work](OpenAP_anomaly_plot.ipynb)   This file plots, for a given anomaly, long-short portfolio returns. 