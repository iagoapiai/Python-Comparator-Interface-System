# Comparator-Interface-System
The Python script showcases an interface featuring several functions, with the primary function being the comparison of two tables to display variations. Additionally, the script includes a web scraping function designed to extract desired data from the company's website, saving it in an Excel file for future comparisons. 

This script can be improved, but I chose to keep it as one of my first works of data manipulation with interface. I want to leave it exactly as I did at the time so I can see how much I have evolved!

It is important to note that this code will only function correctly when the website elements and spreadsheet data are appropriately modified, as well as the directories adjusted.

## :warning: Prerequisites

- [Pandas](https://pandas.pydata.org/docs/)

- [OS](https://docs.python.org/3/library/os.html)

- [Boto3](https://boto3.amazonaws.com/v1/documentation/api/latest/index.html)

- [Datetime](https://docs.python.org/3/library/datetime.html)
 
- [PySimpleGUI](https://www.pysimplegui.org/en/latest/)

- [Ctypes](https://docs.python.org/3/library/ctypes.html)

- [Pyperclip](https://pyperclip.readthedocs.io/en/latest/)

- [Time](https://docs.python.org/3/library/time.html)

- [Pathlib](https://docs.python.org/3/library/pathlib.html)

- [Openpyxl](https://openpyxl.readthedocs.io/en/stable/)

- [Selenium](https://www.selenium.dev/documentation/)


## WHAT THE SCRIPT DO:
- The interface displays three possible actions: compare, import data, and map companies.
- **Compare**: Opens a new tab where you can select two worksheets to compare. It then displays the treated worksheets in an interface with various functions.
- **Import data**: Scrapes the site to collect information from 47 companies and saves the data in a spreadsheet for future comparison.
- **Map companies**: Maps the companies present on the main website.

## OBSERVATION

This code will only work if you redo the modifications of the dataframe, it also needs to contain a folder called 'statusbolt' to save the excel file! 

## SOME PICTURES FROM THE SCRIPT

<p align="center">
    <img src="https://github.com/iagoapiai/Comparator-Interface-System/assets/116030785/a6f02e09-a378-4027-8225-340e59e5a497" width="600" height="750">
    <img src="https://github.com/iagoapiai/Comparator-Interface-System/assets/116030785/d61e7eb0-07ef-465e-872b-9de971dffec4" width="600" height="750">
</p>

Personal project, made to increase efficiency in my work! ❤️



