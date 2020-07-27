# Excel VBA Macro for Explorer Report validation


Here is the  procedure to use the macro
***************************************

1. Open the given Macro-enabled workbook 'Report_Validation'.

2. By default, it will open with the sheet 'Report validation’ If not go to that sheet.

3. Now you can see the list of available tests ( For now only one test is available to compare the columns between uploaded and exported files).

4. Click on the button against the scenario which you want to test. For example the button against the ‘Header row validation’ test has been clicked.


Execution flow of 'Header row validation' test :
****************************************************

1. Message box will be displayed with the buttons ‘Yes’ and ‘No’ to let you select the format of the uploaded file. It can be either CSV or Excel.
   
   Click on ‘YES’ if you want to select the CSV file
   
   Click on ‘No’ if you want to select the Excel file
   
2. Select the uploaded file from your computer which was used to create the project.

3. Again the ‘Open/Browse file’ window will be opened to let you select the exported excel report from the explorer. 

4. Sit back and let the macro work for you. In a few seconds you will be automatically redirected to the ‘Header_Row’ sheet where the validation results are captured.

5. In this test,  Columns available in the uploaded file will be listed in the ‘Header_Row’ sheet. Also the background color will be changed to ‘RED’ if the column is not available in the exported file. Likewise it will be changed to ‘Green’ if the column is available in the exported file.


                                                                       ***
