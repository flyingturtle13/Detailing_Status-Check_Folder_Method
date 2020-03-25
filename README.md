# Detailing Model Files in Focus Zone Folders Status Checker
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;At its fundamental purpose, this subroutine checks a folder if a file exists and last updated date per column and row header name. It requires a user to input root folder path to where files reside. </br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;For this application, detailers are to post model files to associated level folders when ready for coordination and appending to Navisworks (NWF) federated model. The subroutine automates checking if the detailing model files exist in the level folders and reports last posting date. The Excel file is then consumed in Power BI as a visual report showing detailing status (Up to what level trade detailers are at and tooltip indicating last posting date).</br>
Note: Information regarding updating the Power BI report data visuals will not be discussed.  Refer to https://synoptic.design/ for information about updating report visuals.

## Getting Started
Environment setup required to implement subroutine

* Repository Items:
  * Subroutine .cls file
  * Associated Excel worksheet implementing subroutine
  * Power BI report referencing Excel spreadsheet

* IDE:
  * Excel Macros

* Language:
  * VBA (Microsoft Visual Basic)

* Output Type:
  * Class File (.cls)

## Subroutine Development
Subroutine features and specs

* User Interface
  * Button (Update Detailing Status) to activate auto update feature
  * Paste root folder path to where files reside when prompted
  * Refresh file in Power BI to update report data
  * Power BI report tooltip indicates date of model file posting

* Subroutine Specifications
  * By default, if file does not exist cell value of level model by trade is "0"and date posting is empty
  * When file is found, cell value of level model by trade is set to "1" and date is copied to another column cell
  * Create folders for levels/focus zones for model file posting
  * Cell values in orange highlighted cells should match folder names where model files to be checked are/to be posted.
  * Revise range if number of levels (row) and disciplines (column) change. (See worfklow structure below)
  * Revise column number for recording file posted date if range is revised.
  * Model file name should use underscore ( _ ) as delimiter.
  * Cell values in blue highlighted cells should match discipline code used in the model file name. (See worfklow structure below)
    * Ex: MD ==> ProjectName_LXX_MD_CompanyCode.nwc, will use MD to match with file name

## Workflow Structure Implementing Subroutine
See below for the flow chart and map depicting overall structure and flow of information.</br></br>
**Legend**
1) Button to activate auto update.
2) Window prompt for user to paste folder path to Level/Focus Zone folders where model files are posted.
3) Level/Focus Zone folders
   * Folder name is used to map to corresponding row
4) Model files in Level/Focus Zone folders
   * Discipline Code in file naming convention used to map to corresponding column. ( _ ) used as delimiter.
   * The date file is posted is also mapped to another column in the same row.
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/77497697-2f91c780-6e0b-11ea-85bd-e072f87dd1fe.png" width="800">
</p> 

## Installing and Running Application
<p> 1. Clone or download project. </p>
<p> 2. Open FileStatusCheck_Template.xlsm. Check if macros already loaded. If not, import included CLS file.</p>
<p> 3. Create level folders are created for model files to be posted. </p>
<p> 4. Ensure level cell values in column B match folder names</p>
<p> 5. Revise Discipline Code in row 1 to match that of the model file names based on standardized file naming convention.
<p> 6. Check subroutine code and update parameters in code as needed.</p>
    * Discipline check range if correct
         <p align="center">
         <img src="https://user-images.githubusercontent.com/44215479/77499963-da58b480-6e10-11ea-9b1b-846232a82869.png" width="600">
         </p>
    * check if column is correct where model file date posted is stored.
         <p align="center">
         <img src="https://user-images.githubusercontent.com/44215479/77500083-38859780-6e11-11ea-8b79-107f847d185a.png" width="600">
         </p>
<p> 7. When setup is complete, click "Update Detailing Status" button to activate.
<p> 8. When prompted, paste root folder path to where level folders are located.
<p> 9. Update the report with correct data visual image per https://synoptic.design/. </p>
<p> 10. Ensure Data Source path is pointing to excel file: FileStatusCheck_Template.xlsm
<p> 11. Refresh report to update data visuals

## Power BI Report

- Mapping Data to Power BI Report
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/77501101-2fe29080-6e14-11ea-9138-858a6c31d9c9.png" width="600">
</p> 

- Power BI Tooltip displays date of model file posted
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/77501363-1726aa80-6e15-11ea-92ca-5a3cf78e626b.png" width="600">
</p> 

