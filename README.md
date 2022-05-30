# MPR_Reader
Python code to read Master Production records and output an Excel file containing some of the data inside

**What is this code for?**

A Master Production Record, or MPR, is a word document containing information on a chemical product; including its reference number, the materials used to make it, expected time to completion, instructions on how to make the product, and other information. This data is kept inside multiple tables inside the word document, and is formatted for human rather than machine readability. The goal of this project was to make a program that could read these documents (~1400 word files), then output an excel file that contained relevant data (expected yield, cycle time, starting material information, ect.) while avoiding any duplication of data. This data could then be compared to information already in the company database in order to correct outdated entries and determine which MPRs most needed updating, among other issues.

**What does this code do?**

doc_to_docx.py:
This program converts the MPR files from their .doc to a .docx format, so that it is easier to manipulate later, and saves the new document in the same location

MPR_scraper.py:
This program reads in the word files and extracts data from two relevant tables, the starting materials table and the yield/cycle time table. This information is then converted into a pandas dataframe and the ExcelWriter functionality is used to create and save a new Excel file.
