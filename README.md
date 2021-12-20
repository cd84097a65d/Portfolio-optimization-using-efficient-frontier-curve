# Portfolio optimization using efficient frontier curve.
The algorithm takes time series from finance.yahoo.com and / or ariva.de and plots an efficient frontier curve as a function of risk tolerance. The portfolio with maximal Sharpe ratio is supposed to be the best optimized portfolio. 
The portfolio optimization calculations are based on [Matlab](web.stanford.edu/~wfsharpe/mat/gqp.txt) algorithm of William F. Sharpe. The basic article of William F. Sharpe: William F. Sharpe, "[An Algorithm for Portfolio Improvement](https://www.gsb.stanford.edu/faculty-research/working-papers/algorithm-portfolio-improvement)", in Advances in Mathematical Programming and Financial Planning, JAI Press, Inc., 1987, pp. 155-170.
The portfolio optimization is performed in two steps:
1. "1. Update inputs" – cleans sheets "Time series", "Covariation", "Returns" and takes time series for one year from finance.yahoo.com or ariva.de.
2. "2. Optimize portfolio" – takes temporary data from "Time series", "Covariation", "Returns" and optimizes asset weights to plot efficient frontier curve. The portfolio with maximal Sharpe ratio is supposed to be the best optimized portfolio.

Inputs:
- Tickers (column A): Contains tickers. This information is necessary only for finance.yahoo.com. If you want to use also the historical data from "ariva.de", then simply put here the WKN or other description to distinguish between the assets. The updating of time series is iterating through the assets until the empty cell in column A is reached.
URL (column B): In case of American shares it must contain "Yahoo". If you want to obtain time series from "ariva.de", then it must contain the URL to desired historical data from ariva.de (see example in the table).
- Lower boundary (LBD, column C): The asset portion cannot be lower than this value (in percent). It can be only positive number. The sum over all LBDs cannot be higher than 100%. The error message will be shown if the sum of LBDs is higher than 100% and the calculation will be finished. This number is very convenient if you do not want to touch some assets in your portfolio. For example, if you want to leave the portion of Apple shares to 25% in your portfolio and optimize only the rest of assets, then set LBD for Apple equal to 25%. In this case the portion of Apple after portfolio optimization will be equal or higher than 25%. The default setting of LBD is 0%. In this case no restrictions on LBD will be performed.
- Upper boundary (UBD, column D): The asset portion cannot be higher than this value (in percent). It can be only positive number. Restricts maximal portion of assets in portfolio. The default setting of UBD is 100%. In this case no UBD restrictions will be performed. The UBD values are used to restrict the portion of assets in portfolio. For example, if the optimization tries to increase the portion of Apple to 50%, but you want to increase the number of assets and restrict the portion of Apple shares to 5%, then set UBD for Apple to 5%. Using UBD you can switch off optimization of some assets by setting their UBDs to 0%. The UBD must be always higher than LBD, otherwise the script will be stopped with an error message.
- Weights (column F): Contains portions in your portfolio. If you performed portfolio optimization, then this column contains optimized weights of assets in your portfolio. You can also check the characteristics of your current portfolio if you copy the portions of shares in your portfolio to the corresponding positions in column E and press "Recalculate portfolio characteristics". Please note, that the sum of all portions must be exactly 100%. If not, the program will show you an error and will stop calculations.

Outputs:
- Weights (column F): Contains assets portions in your portfolio after portfolio optimization.
- Sharpe ratios (column E): Contains calculated Sharpe ratios for all assets. These values can be compared to the Sharpe ratios from internet. During the debugging of the code I compared the calculated Sharpe ratios with values from internet and found that they are quite close to each other (R < 0.5%). The difference can come from removed time series values (to calculate covariance we have to have all time series filled with numbers, no empty cells are allowed).
- Obtained weights are sorted in ascending order and shown in columns N (ticker) and O (sorted weights).

Other sheets:
- "Time series", "Covariation", "Returns": These sheets contain temporary arrays of time series, covariations and returns obtained during "1. Update inputs". These sheets are used at the second step "2. Optimize portfolio". "1. Update inputs" completely removes all values from all these sheets.
- "Efficient frontier": Contains calculated efficient frontier for representation on the graph. 

Used colors:
Light blue: Value can be changed by user.
Light orange: Value is changed by script.
Light violet: Value can be changed by script, but can be also changed by user (see description of weights).
Turquoise: Cell contains formula.

Known problems:
1.  Finance.yahoo.com returns not always correct time series. The time series of NVDA (NVIDIA) were found to have the split inside. Therefore, the NVIDIA time series at German market (NVD.DE) were taken instead. It is very important to check the time series manually visually on the graph by selecting desired ticker from the dropdown list (cells Q1:T1).
2. Time series for one year are taken (TimeSeriesLength_days = 365). This parameter can be increased / decreased only in case of tickers from finance.yahoo.com. If you are using the time series also from ariva.de, the time series for one year will be taken. This can be corrected, but needs more time.
3. Calculation of risk tolerance for manually changed assets weights is not performed. So, if you will press "Recalculate portfolio characteristics", the value of risk tolerance will be set to "undefined".

