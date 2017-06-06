# Risk Scoring for Financial Assets
The project aims to provide risk scoring for CDAX companies. The scoring is done by taking the performance measures of the companies in certain parameters. 
## Getting Started
These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.
### Prerequisites
The project is built on python and requires the following
>python 2.7 
>XLRD and XLWT
>Numpy
>Scipy
>Sklearn
>Random

### Installing
Clone the project
>git clone https://gitlab.lrz.de/ga47baz/FinancialModel.git

#### Using the data cleaning
The data cleaning code is in Java. It can be found in the Data cleaning folder of the project. This is used to replace NULL values and also to values which are set to 0. 
> Provide the column of your dataset that needs to be cleaned. 
> Mention the column(s) of the dataset to parse for Data cleaning. 

The project is an export of a PyCharm project. To run it successfully on your own dataset, you can modify the data access as follows. 
The project risk_scoring contains the project. 
1. Use the Data_Collection folder to populate the dataset to run the risk_scoring. We use the following parameters.
    - Beta
    - VAR
    - Volatility
    - Gearing (FY)
    - EBITDA (FY)
    - ROCE (FY)

2. Run the corresponding Python file to store the yearly data for the parameter. Edit the column numbers to match the parameter required. 
3. Once the parameters are calculated. Use the src folder in risk_scoring for performing risk_scoring.
    - Run the fy_parameters.py
    - Next, run the risk_scoring_fy_params.py be mentioning the FY params. 
    - Run the risk_scoring_others.py for the Volatility and VAR. 
    - Run the risk_scoring_beta.py to calculate risk for BETA. 
    - Finally, run total_risk.py to get the risk results. 


## Results
The project achieves the task of providing risk values for the CDAX companies. The results were crosschecked with online stock market portals and were in sync with real world data. 
The results for the calculations can be found in the risk_results folder. 

## Further developments
1. The prediction of risk was done using kmeans - unsupervised learning as there was no prior classification. 
2. The results can be used to perform supervised learning on the CDAX companies or other datasets as well. 
3. Due to certain issues with the quality of the data, the yearly calculations were hard coded with row numbers as they were not symmetrical. On cleaner datasets, this should be avoided and can be 
replaced.

## Acknowledgements
We would like to thank the TUM chair of Entrpreuneurial Finance for giving us this project and  for guiding us till the project completion. 