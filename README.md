# stocks-analysis
Module 2
# Analyzing Stocks with VBA
## Overview of Project
This project is a VBA analysis of stocks data across 12 different Green Companies. The analysis is conducted on spreadsheets of Green Company stock data for years 2017 and 2018, each encompassing 3012 rows excluding the header. For each company during each year, VBA code is written to report the total volume of stock trades and the return on investment, or percent change in the stock's value for that year. The VBA code was originally written with a nested for loop using an array of the 12 different Green Companies before it was refactored for improved performance with a new index variable and all values reported from arrays.
### Purpose
The purpose of this project is to generate working VBA code for stocks analysis of data in excel. Additionally, the purpose of this project is to seek advantages and disadvantages for different VBA coding methods in this data analysis. The methods compared are: 1) an array of stock company data ran through a for loop with nested for loop and 2) a for loop indexed with a variable through four different arrays (three of which were non-array variables in the first method). To compare the methods, the code execution times are recorded.
### Background
VBA, or Visual Basic for Application, is a programming language extention integrated with the microsoft office suite (https://corporatefinanceinstitute.com/resources/excel/study/vba-in-excel/). It is useful for automating key procedures and processes in order to increase efficiency (https://corporatefinanceinstitute.com/resources/excel/study/vba-in-excel/). Additionally, VBA provides access to functions beyond what is available in Microsoft Office applications (https://corporatefinanceinstitute.com/resources/excel/study/vba-in-excel/). In finance, VBA is an invaluable tool used to manipulate large amounts of data for streamlined reports and data visualizations (https://www.investopedia.com/terms/v/visual-basic-for-applications-vba.asp).
## Results
Figure 1: Original code execution runtime for 2017
![](Resources/2017%20(pre%20refactoring).png)
Figure 2: Refactored code execution runtime for 2017
![](Resources/VBA_Challenge_2017.png)
Figure 3: Original code execution runtime for 2018
![](Resources/2018%20(pre%20refactor).png)
Figure 4: Refactored code execution runtime for 2018
![](Resources/VBA_Challenge_2018.png)
Using (refactored code execution time - original code execution time) / original code execution time * 100, we find an 84% faster execution for 2017 stocks analysis and 83% faster execution for 2018 stocks analysis with the same outputs.
## Summary

### Refactoring Code

### Original and Refactored Code
