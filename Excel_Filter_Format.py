""" This function is designed to allow the split of data based on 2 discrete filters, it then
apply formats and adds a sum column for the number columns. """

import pandas as pd


def split_all(file, area, group="unknown", year="unknown"):     # Defined to allow re use with any file
    x = area.lower()    # Accommodate for "area" written in either case.
    allData = pd.read_excel(file)   # Read the excel file and assign it to a dataframe allData

    # Use an if statement to define whether the dataframe will be filtered for 2 areas, in this case
    # east or west data. Then rename frame data. the long number is the filter to seperate what is in each area.
    if x == "east":
        data = allData[allData["column2"] < 2000000000000]
    elif x == "west":
        data = allData[allData["column2"] >= 2000000000000]

    data = data.copy()  # Make a copy of the original dataframe to allow columns to be added
    
    data['Sum'] = data.iloc[:, 3:].sum(axis=1)  # Add a sum column for the number columns.

    new_name = str(x)+"_"+str(group)+"_"+str(year)+".xlsx"  # Define the name of the new file

    writer = pd.ExcelWriter(new_name, engine='xlsxwriter',
                            datetime_format='dd/mm/yy', date_format='dd/mm/yy')     # Create an xlsxwriter and set date format

    data.to_excel(writer, sheet_name="sheet1", index=False)     # Use the xlsxwriter to create new excel document, removes the index

    workbook = writer.book  # Define the workbook of the new excel
    worksheet = writer.sheets["sheet1"]     # Define the first sheet of the excel

    format_column2 = workbook.add_format({"num_format": "0"})  # Create format for the second column
    format_date = workbook.add_format({"align": "left"})    # Create format for date column
    format_Numbers = workbook.add_format({"num_format": "0.00"})     # Create number and sum column format

    worksheet.set_column("A:A", 40) # Apply formats to all columns in the document
    worksheet.set_column("B:B", 14, format_column2)
    worksheet.set_column("C:C", 18, format_date)
    worksheet.set_column("D:AY", 8, format_Numbers)
    worksheet.set_column("AZ:AZ", 8, format_Numbers)

    writer.save()   # Save the final excel document

    print("\n" + "FILE " + str(new_name) + " CREATED")  # Announce the document is created and ready to open


split_all("File", "east", "all sites", "2017") # Filters a file for a east area
split_all("File", "west", "all sites", "2017") # Filters a file for a west area

