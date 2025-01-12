# Supercharge Excel with Python

While Excel is a powerful tool for data management and analysis, it often feels limited when dealing with complex data transformations, large datasets, and repetitive tasks. Here is where Python comes into play. By integrating Python libraries into your workflow, you can unlock its true potential.

See Our Full Range of Plans - Find the Right Plan
workspace.google.com
See Our Full Range of Plans - Find the Right Plan
Ad
In this post, I will go over powerful Python libraries that can automate tedious tasks, fly through advanced analytics, build interactive dashboards, and even create stunning visualizations in Excel. These Python libraries help you tackle any challenge with efficiency and finesse.

In case you are not aware of it, Excel for Windows supports (currently rolling out) a core set of Python libraries from Anaconda. Aside from core libraries, you can import more libraries through Anaconda. You only need to use a Python import statement to complete the process.


An image showing a laptop displaying the Python Package Index website on a web browser.
An image showing a laptop displaying the Python Package Index website on a web browser.
© Provided by XDA Developers
Related
5 Python scripts anyone can use to boost their productivity
If you've heard of Python but don't know where to start with it, these five scripts can help boost your productivity.

Pyexcel
Manage different file formats in no time

Python-library-to-extend-Excel-functionality 3
Python-library-to-extend-Excel-functionality 3
If you often deal with different file formats, the Pyexcel library simplifies working with workbook data. It offers a single API for reading, writing, and tweaking data in various spreadsheet formats, including CSV, XLS, XLSX, and ODS. As always, you can easily integrate it with other Python libraries like Pandas, where you load data from a spreadsheet using Pyexcel and use Pandas to analyze and manipulate it.

Squarify
Create treemaps

Python-library-to-extend-Excel-functionality 4
Python-library-to-extend-Excel-functionality 4
Does your Excel sheet have a hierarchical structure (like categories, subcategories, and individual items)? Treemaps can be a handy way to visualize the relationship within that hierarchy. And here is where the Squarify library comes into play. You can use the code below to create a treemap where the size of each rectangle corresponds to the values in the sizes list.

import squarify

import matplotlib.pyplot as plt

# Sample data (could be derived from Excel)

sizes = [50, 25, 15, 10]

labels = ["Category A", "Category B", "Category C", "Category D"]

# Create the treemap

squarify.plot(sizes=sizes, label=labels, alpha=.8)

plt.axis('off')

plt.show()

read more
Openpyxl
Interact with your Excel files like a pro

Python-library-to-extend-Excel-functionality 5
Python-library-to-extend-Excel-functionality 5
Openpyxl Python library is designed to read and write Excel files. You can read data, edit existing content, create new sheets, and even write data back to Excel files without opening the software in the first place. Power users can even combine openpyxl with other Python libraries like pandas for a smooth data analysis workflow.

Related video: Create Google Sheets Automatically using AI (Howfinity)
second, and just like that, I got this template that
Current Time 0:18
/
Duration 4:25
Howfinity
Create Google Sheets Automatically using AI
0
View on Watch
View on Watch
For instance, you can create a simple Excel file (people) with a table containing names and ages.

from openpyxl import Workbook

# Create a new workbook

workbook = Workbook()

# Get the active worksheet

worksheet = workbook.active

# Add some data

worksheet["A1"] = "Name"

worksheet["B1"] = "Age"

worksheet["A2"] = "Alice"

worksheet["B2"] = 30

worksheet["A3"] = "Bob"

worksheet["B3"] = 25

# Save the workbook

workbook.save("people.xlsx")

read more
Matplotlib
Visualize your Excel data

Python-library-to-extend-Excel-functionality 2
Python-library-to-extend-Excel-functionality 2
If you find the default Excel charts insufficient for your workflow, use the Matplotlib library. It offers a vast array of chart types beyond the standard ones in Excel. You can include scatter plots (check the screenshot above), histograms, heat maps, 3D plots, and more. You can tweak the chart appearance with minute details, from colors and labels to axes and legends.

Flight Attendant Reveals How To Fly Business Class For The Price of Economy
Travel Savings Tools
Flight Attendant Reveals How To Fly Business Class For The Price of Economy
Ad
Matplotlib also goes a step ahead with interactive plots that allow zooming, panning, and exploring data in more detail. This is quite helpful for large datasets or when you want to present data in an engaging way. For example, you can use the code below to create a line graph showing the sales trend over five months.

import matplotlib.pyplot as plt

# Sample data months = ["Jan", "Feb", "Mar", "Apr", "May"] sales = [1500, 1800, 1600, 2100, 2300]

# Create the plot plt.plot(months, sales)

# Add labels and title plt.xlabel("Month") plt.ylabel("Sales") plt.title("Monthly Sales")

# Show the plot plt.show()

read more

A screen with python code running, with led lighting in the background
A screen with python code running, with led lighting in the background
© Provided by XDA Developers
Related
5 reasons why I use Python instead of Excel for visualizing data
A data visualization upgrade you can’t miss

Pandas
Data manipulation powerhouse

Python-library-to-extend-Excel-functionality 1
Python-library-to-extend-Excel-functionality 1
While Excel is ideal for basic spreadsheet tasks, pandas take it to the next level when you need to do more with your data. Built on top of NumPy, it handles millions of rows with speed and efficiency. It offers a range of tools for cleaning, transforming, and analyzing data. You can easily filter, sort, aggregate, pivot, perform complex calculations, and go far beyond Excel’s built-in functions.

Finally... Gorgeous Metal Roofs That Are Affordable For Everyone
Metal Roof Nation
Finally... Gorgeous Metal Roofs That Are Affordable For Everyone
Ad
You can write scripts to process data, generate reports, and automate repetitive tasks to save time and reduce errors. Pandas is also flexible enough to let you work with data from various sources, not just Excel files. For example, this Python code snippet uses the Pandas library to read data from an Excel file and display the first five rows.

import pandas as pd

# Read the Excel file into a pandas DataFrame df = pd.read_excel("sales_data.xlsx")

# Print the first 5 rows of the DataFrame print(df.head())

read more
Become an Excel wizard
Python integration in Excel is a game-changer for anyone dealing with large datasets. What are you waiting for? Whether you are a data analyst, a business professional, or simply a spreadsheet guru, integrate these Python libraries into your workflow and gain deeper insights from your data. While you are at it, check out our separate post to find some interesting Python libraries that anyone can use.
