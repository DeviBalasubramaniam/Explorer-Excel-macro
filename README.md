# Excel VBA Macro for Explorer Report validation


## Procedure to use the macro


1. Open the given Macro-enabled workbook 'Report_Validation'.

2. By default, it will open with the sheet 'Report validation’ If not go to that sheet.

3. Now you can see the list of available tests in the Table. Brief explanation about table.

   S. No  - Test Case number

   Test Name  - Name of the Test

   Test Description - Brief summary about the test

   Test data - Required data to run that test. Tests for which the data is not required marked as ‘NA’

   Execution Flag - It can be set to ‘Yes’ or ‘No’ by using the drop down list. The test will be executed only when the flag is set to “Yes”.

   Status - Execution status of the test. If the test is not included in the run then it will be marked as ‘No Run’ else based on the execution result it will be set to ‘PASS’   or ‘FAIL’.

4. Set the execution flag 'Yes' if you want to run that test. 

5. Update the test data if it's mentioned as 'required' in the test description.

6. Click on the 'Excel report validation' button below to start the execution.

7. It’s time to select an uploaded file.  Message box will be displayed with the buttons ‘Yes’ and ‘No’ to let you select the format of the uploaded file. It can be either CSV or Excel.
   
   Click on ‘YES’ if you want to select the CSV file
   
   Click on ‘No’ if you want to select the Excel file
   
8. Select the uploaded file from your computer which was used to create the project.

9. Again the ‘Open/Browse file’ window will be opened to let you select the exported excel report from the explorer. 

10. Sit back and let the macro work for you. Execution status will be updated in a few minutes.

11. To view more information about the test cases  especially for the failed ones go to the ‘Execution_Results’ sheet where the validation results are captured.

Note : Refer 'Report_Validation' in the macro enabled workbook to know more about the tests.

***
