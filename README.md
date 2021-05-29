# HS-Forecast-Melter

At a CPG company selling handsoap, part of the supply planning process required taking manufacturing quantity data pivoted by item and date, and converting this data into a melted format with single columns for quantity, date, and item.

MS Excel didn't have straighforward functionality to perform this operation without using the more verbose VBA language.

Although this is simple operation in python, the additional task of updating the order status by cell color introduced some exciting complexity.

This script takes the original forecast and produces the melted data with updated status in a new Excel file.

To use from the command line, cd into the directory with the original forecast .xlsx file, and run `python3 fcmelter.py filename.xlsx`
