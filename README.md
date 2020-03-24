# Folder in File Status Checker
At its fundamental purpose, this subroutine checks a folder if a file exists and last updated date per column and row header name. It requires a user to input root folder path to where files reside. </br>
For this application,  detailers are to post model files to associated level folders when ready for coordination and appending to Navisworks (NWF) federated model. 
The subroutine automates checking if the model files exists in the level folders and reports last posting date.  This file is then consumed 
in Power BI as a visual report showing detailing status (Up to what level trade detailers are at and tool tip indicating last posting date).

## Getting Started
Environment setup required to implement subroutine

*Repository Items:
  *Subroutine .bas file
  *Associated Excel worksheet implementing subroutine
  *Power BI report referencing Excel spreadsheet

* IDE:
  * Excel Macros

* Language:
  * VBA (Microsoft Visual Basic)

* Output Type:
  * Basic Files (.bas)

## Application Development
Application features and specs

* User Interface
  * 4 number combination to be guessed
  * number of remaining attempts visible
  * view history of attempt guesses including number of matched digits and number of matched digit positions

* Application Specifications
  * numbers can be duplicated
  * number range is from 0 to 7
  * Using Random Number Generator API (https://www.random.org/integers) to provide computer number combination. 

* Extensions
  * Random quote generated using API (https://api.forismatic.com/api/1.0/?method=getQuote&lang=en&format=jsonp&jsonp=?) when player is successful at decoding the combination.
  * User Hint (updated 3/20/2020): User can choose to receive a hint in the form of receiving a number not in the combination.  User can get up to a maximum of 3 hints.  In future update, receiving a hint will affect number of attempts left.

## Application Structure
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;See below for the flow chart depicting overall structure and flow of the application.  It highlights visible and backend processes.  The application begins at the Main Menu symbol.</br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;To develop Mastermind, the WPF application in visual studio serves as the foundation for the user interface. Various controls (mostly buttons) are implemented so that the user can interact and navigate through the game.  The controls also send information to the backend to process API web calls, evaluate user input, and determine results.  
<p align="center">
  <img src="https://user-images.githubusercontent.com/44215479/74339457-f7519080-4d58-11ea-90d3-88cd95b4ca2c.png" width="800">
</p> 

## Installing and Running Application
<p> 1. Clone or download project. </p>
<p> 2. Open REACH_Mastermind_Project.sln in Visual Studio 2019. </p>
<p> 3. Ensure that the library packages stated in Getting Started are installed and referenced. </p>
<p> 4. The application can then be run in debug mode. </p>
<p> 5. To view all game outcomes (especially with viewing the randomly generated quote when successful at guessing the combination) uncomment line 31 in the 03_NumberRequest class to view the API generated combination.
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