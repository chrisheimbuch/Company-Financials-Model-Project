# Company-Financials-Model-Project

Dataset: https://www.kaggle.com/datasets/atharvaarya25/financials

## Overview
This project showcases how Python can streamline monthly reporting, budgeting, and forecasting for companies using ERP system reports. It minimizes human error and speeds up data processing compared to Excel-only workflows.

## Features
- **Data Cleaning**: Processes a Kaggle financial dataset with extensive cleaning and preprocessing using Python (pandas, etc.).
- **Reporting**: Generates meaningful reports and visualizations embedded in an Excel file.
- **Forecasting**: Includes a 2015 forecast and roll-up for Excel-dependent workflows.
- **Presentation**: Provides a deck comparing 2013, 2014, and 2015 forecast, with opportunities, threats, and recommendations.

## Usage
Explore the code to see how Python can enhance your reporting processes.

Important files:
Sales Deck CH Consulting: Powerpoint Deck to showcase and simulate a presentation to senior management / C-suite executives.
fpa_model_edits: Working file with cleaned data, a forecast I created, waterfall graph, assumptions, and summaries. This is the main working file.
fpa_model: A file that is generated based on the original excel file "Financials.csv" via Python with different tabs for insights and graphs and cleaned data.
research.ipynb: This is a research file where I conduct multiple different data cleaning techniques and visuals to understand the data better.
main.py: Executable python file to automatically generate the new file, "fpa_model", which cleans and generates a new files with important insights on the dataset.

The research file goes through extensive testing and methods to clean the data. The main.py file is an executable with the push of a button will clean the data and generate the excel reports and visuals in seconds which can be repeatable for the same type of report, save time, and reduce human error.
