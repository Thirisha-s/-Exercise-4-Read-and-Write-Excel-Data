# -Exercise-4-Read-and-Write-Excel-Data

```
Name: THIRISHA S
Reg No: 212222230160
```

## Aim:
To automate the process of reading data from an Excel file and writing data into another Excel file or a
different sheet using UiPath.

## Equipment Required:
UiPath Studio (installed on a compatible system).
Microsoft Excel application.
Excel file (e.g., "ExcelFile.xlsx") containing data for reading and writing purposes.
Computer with:
Minimum 4 GB RAM.
Minimum 2.0 GHz CPU.
Windows operating system.
.NET Framework 4.6.1 or later.

## Procedure:
Start a New Process
Open UiPath Studio.
Create a new process by selecting Process under the New Project tab.
Name the project UiPath Read and Write Excel Data and click Create to begin.

## Install Excel Activities Package:
If the Excel activities are not already available, go to Manage Packages (click on the Project panel).
Search for UiPath.Excel.Activities in the Official packages.
Click Install and then Save to include the Excel activities in the project.

## Add Excel Application Scope:
In the Activities panel, search for Excel Application Scope. Drag and drop the Excel Application Scope
activity into the main sequence window.
In the Properties pane, specify the path to the Excel file (e.g., "C:\Path\To\Your\ExcelFile.xlsx").
The Excel Application Scope will open the Excel file and keep it open while the robot interacts with it.

## Read Data from Excel:
Inside the Excel Application Scope, search for Read Range in the Activities panel.
Drag the Read Range activity inside the Excel Application Scope. Set the SheetName property to
"Sheet1" (or the sheet from which data needs to be read).
Leave the Range field blank if you want to read all data from the sheet, or specify a range like "A1:C10"
if you want specific cells.
Create a variable named dtExcelData to store the output data, which will be of type DataTable.

## Modify the Data (Optional):
If you wish to process the data after reading it, you can use activities like For Each Row to loop
through the rows of the DataTable.
Example: Use Assign activities to modify certain cell values, calculate totals, or apply filters.

## Write Data into Excel:
After processing or modifying the data, add a Write Range activity under the Excel Application Scope.
Set the SheetName to a different sheet, such as "Sheet2" (or the same sheet, if preferred).
Set the DataTable property to dtExcelData to write the data back to the file.
Specify the Range (e.g., "A1") to indicate where the data should be written.
Ensure the Add Headers option is checked if the data contains column headers that need to be
written.

## Save and Run the Workflow:
Save the project by pressing CTRL+S.
Click Run from the toolbar to execute the workflow.
UiPath will read data from Sheet1, process it (if specified), and then write the data to Sheet2 (or the
same sheet).

## Example Scenario:
Initial Excel File:
Sheet1 contains customer data: Name, Age, City.
Sheet2 is initially empty.

## UiPath Workflow Actions:
The robot reads the data from Sheet1.
(Optional) The data is modified (e.g., age is incremented by 1).
The modified data is written to Sheet2, including the headers

## Expected Outcome:
Sheet2 will now have the same data as Sheet1 (or the modified version), demonstrating how UiPath
can read and write data efficiently.

## UiPath WorkFlow:
![Screenshot 2024-09-16 220300](https://github.com/user-attachments/assets/255db21a-b888-4100-9faf-79822161ceb5)
![Screenshot 2024-09-16 220401](https://github.com/user-attachments/assets/b0c379a1-ebf5-4a83-861f-714d96b6300e)
![Screenshot 2024-09-16 220529](https://github.com/user-attachments/assets/9abbc2cc-3d0f-4ef7-ab13-0c86590a3120)
![Screenshot 2024-09-16 220546](https://github.com/user-attachments/assets/15d37e37-93e7-4806-9f78-6455aed2f249)

## Result:
The automation workflow successfully reads data from Sheet1 of the Excel file, processes it if required, and writes the data into Sheet2 (or another location) of the same Excel file. This confirms the ability of UiPath to interact with Excel for reading and writing operations.

