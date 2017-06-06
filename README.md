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