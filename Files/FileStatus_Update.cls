VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub UpdateStatus_Btn_Click()
    
    'Get User Input of folder path to check if detailer model file exists.
    Dim userInput As Variant
    userInput = InputBox("Input folder path to model files.  Checks if model file exists." & vbNewLine & "i.e. Local Drive:...\Project_Name\00 Models\3 Navisworks", "Check Model Detailing Status")
    
    Dim folderPath As String
    folderPath = userInput
    
    '______"UPDATE SHEET # AS NECESSARY"______ Apply script to Trade Detailing Status worksheet
    Worksheets(1).Activate
    
    'Declare Variables
    Dim modelRange As Range
    Dim cell As Range
    Dim level As String
    Dim disCode As String
    Dim colNum As Long
    Dim currColNum As Long
    Dim file As Object
    Dim currFolder As String
    Dim currFolderPath As String
    Dim oFSO As Object
    Dim oCurrFolder As Object
    Dim subStr() As String
    Dim ele As Variant
    
    '______"UPDATE RANGE ("L3") AS NECESSARY"______Iterate over cells to evaluate
    Set modelRange = Range("C2:L31")
    
    For Each cell In modelRange
        
        'Set cell value default to 0 (file does not exist)
        cell.Value = 0
        
        'Get level (row) and current trade to match (column)
        level = Range("B" & cell.Row).Value
        colNum = cell.Column
        disCode = Cells(1, colNum)
        
        'Create folder path by Level (row)
        currFolderPath = folderPath & "\" & level
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        Set oCurrFolder = oFSO.GetFolder(currFolderPath)
        
        For Each file In oCurrFolder.files
        
            subStr = Split(file.Name, "_")
            
            For Each ele In subStr
                
                'Check if trade model file matches current trade column
                If ele = disCode Then
                
                    currColNum = colNum
                    
                    'Set cell = 1 to indicate file exists and date last posted
                    cell.Value = 1
                    '______"UPDATE COLUMN SPACE (10 in "currColNum + 10") AS NEEDED"______Copies file posted date to empty cell
                    Cells(cell.Row, currColNum + 10).Value = file.DateLastModified
                    
                End If
            
            Next ele
        
        Next file
        
    Next cell
    
End Sub
