# TestConsoleApp

========================================================================

Task List:

1. Take a console application
   
2. Pull the data from SQL database using ADO.NET

3. Write the data in Text file, in a drive folder
   
4. Write the data in excel file, in a drive folder
   
5. Read the files from the folder drive and show file list in console output
   

========================================================================

DataExportConsoleApp
DataExportConsoleApp is a C# console application designed to extract data from an SQL database using ADO.NET and perform data export operations. The application provides a user-friendly interface in the console to interactively collect data from the user and save it to various file formats for further analysis and reporting.

Features:

Database Interaction: The application establishes a connection to an SQL database, allowing users to input data such as name, roll, and subject. It then securely saves this data to the "Students" table in the "StudentTestDb" database.

Text File Export: Users have the option to enter data and save it to a text file. Each input is appended as a new line in the "StudentData.txt" file, located in a specified drive folder.

Excel File Export: The application enables users to enter data and save it to an Excel file. The data is written to an existing worksheet named "Sheet1" or a user-defined worksheet. Each entry is appended to the next empty row, making it easy to view and manage the exported data.

File List Display: After performing export operations, the program displays a list of files present in the specified drive folder. Users can see the names of all the files in the folder, including the text and Excel files generated during the data export.

Usage:

Launch the console application, and a user-friendly prompt will ask you to enter the Name, Roll, and Subject for a student.

The program will securely store this information in the "Students" table of the "StudentTestDb" SQL database.

Upon entering data, you have the option to save it to a text file. Each input will be appended as a new line in the "StudentData.txt" file, located in a specified drive folder.

Additionally, you can save the entered data to an Excel file. The data will be added to an existing worksheet or a new worksheet with a user-defined name.

After performing export operations, the program will display a list of files present in the specified drive folder, so you can easily verify the exported data.
