# Python - Generate Excel file from SQL Query

The following program is a python program that will generate an excel file based on a predefined query. The program works as follows:

1. Connects to the SQL database
2. Asks the user what they would like to search for (in this case it was a search for production data based on either a device serial number or batch number)
3. Once the user makes their selection it request the device number or batch number
4. The user enters the batch/serial number
5. An excel file is created with the data named with todays date and the number entered.
6. The program prompts you to where the file is saved.

![image](https://user-images.githubusercontent.com/76489588/183917556-801c9a37-de66-4221-94f1-3978aa7f86ee.png)

This program was created for a factory line which allowed the client to search for all the line data on a device after it had be created. They could simply enter the device number and see all the process data for that part as it passed through the factory line. This is very useful when you have a faulty part and need to see what area of the process went wrong.
