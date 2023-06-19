# VBA_dailyWageCalc
This is the VBA code to calculate daily employer wage based on their attendant records.
The excel workbook "example.xlsm" is provided with employer names being censored.
There are 6 sheets that need to be existed in order for the code to work: "recordList", "wageResult", "timeResult","Wage","Arrive",and "Leave"
"recordList" sheet is used for containing attendant data.
"timeResult" sheet contains each employee their attending time in each day, with the first entry rounded up if the employee checked in before they should be.
"wageResult" sheet contains each employee their wage and overtime wage in each day that they work.
"Wage" sheet is for storing wage data of each employee in each day, the same manner is applied for "Arrive" and "Leave" sheets.

The code is not entirely dynamic to data input in "recordList" work; Some data have to be in their specifics column such as their ID, Date, and Time.
Some data column are unnecessary, but exist because they came from record device.

The programmer is pure novice, any feedback is appreciate, but whether I'm understanding or not is up to how much can I improve.
