# Detailing Model Files in Focus Zone Folders Status Checker
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;At its fundamental purpose, this subroutine checks a folder if a file exists and last updated date per column and row header name. It requires a user to input root folder path to where files reside. </br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;For this application, detailers are to post model files to associated level folders when ready for coordination and appending to Navisworks (NWF) federated model. The subroutine automates checking if the detailing model files exist in the level folders and reports last posting date. The Excel file is then consumed in Power BI as a visual report showing detailing status (Up to what level trade detailers are at and tool tip indicating last posting date).</br>
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
  * Button (Update Detailing Status) to activate feature
  * Paste root folder path to where files reside when prompted
  * Refresh file in Power BI to update report data

* Subroutine Specifications
  * Create folders for levels/focus zones for model file posting
  * Cell values in orange highlighted cells should match folder names where model files to be checked are/to be posted
  * Revise range if number of levels (row) and disciplines (column) change
  * Revise column number for recording file posted date if range is revised
  * Model file name should use underscore ( _ ) as dilimeter.
  * Cell values in blue highlighted cells should match discipline code used in the model file name
    * Ex: MD ==> ProjectName_LXX_MD_CompanyCode.nwc, will use MD to match with file name

## Workflow Structure Implementing Subroutine
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;See below for the flow chart and map depicting overall structure and flow of information.
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/74339457-f7519080-4d58-11ea-90d3-88cd95b4ca2c.png" width="800">
</p> 

## Installing and Running Application
<p> 1. Clone or download project. </p>
<p> 2. Open FileStatusCheck_Template.xlsm. Check if macros alread loaded. If not, import included CLS file.</p>
<p> 3. Create level folders are created for model files to be posted. </p>
<p> 4. Ensure level cell values in column B match folder names</p>
<p> 5. Revise Discipline Code in row 1 to match that of the model file names based on standardized file naming convention.
<p> 6. When setup is complete, click "Update Detailing Status" button to activate.
<p> 7. When prompted, paste root folder path to where level folders are located.
<p> 8. Update the report with correct data visual image per https://synoptic.design/. </p>
<p> 9. Ensure Data Source path is pointing to excel file: FileStatusCheck_Template.xlsm
<p> 10. Refresh report to update data visuals
 <p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/74343016-3e428480-4d5f-11ea-9575-30c933ad4b0b.png" width="600">
</p>

## UI Screenshots

- Main Menu
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/75637773-3e56d700-5bdd-11ea-931f-8e0367c8a795.png" width="600">
</p> 

- Game Main Window (Upated 3/21/2020)
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/77232839-a0618700-6b60-11ea-864b-8b10734657fd.png" width="600">
</p> 

- Try Again Result Window
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/75637919-364b6700-5bde-11ea-8e98-f60326ba6707.png" width="600">
</p> 

- Fail Result Window
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/75637926-44998300-5bde-11ea-8ec1-641e292f7bb7.png" width="600">
</p> 

- Success Result Window
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/75637933-52e79f00-5bde-11ea-8b95-d967df733462.png" width="600">
</p> 

- Game Info Window
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/75637944-62ff7e80-5bde-11ea-9f95-6a99318b38b2.png" width="600">
</p> 

- Guess History Window
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/75637948-6dba1380-5bde-11ea-920c-41c6be23e6f5.png" width="600">
</p> 

- Hint Window
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/77232964-896f6480-6b61-11ea-9c79-94a504e3be47.png" width="600">
</p> 
