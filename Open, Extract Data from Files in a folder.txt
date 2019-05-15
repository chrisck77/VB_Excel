Sub SearchForFiles()
Application.ScreenUpdating = False 'stop screen flashing

    'Declare a variable to act as a generic counter
    Dim lngCount As Long
    Dim varB As String
    Dim actwindow As String
    
    gcell = Range("J3")  'counts number of rows already inputted --> ie to determine next free row
    locale = Range("E1") 'excel cell with the location of the file/s (ie C:\data\)
    Keywords = Range("E2")
    spreadsheetname = ActiveWindow.Caption 'name of this spreadsheet where this macro is stored
	
	'Windows("audit data extractor.xls").Activate 'optional to ensure excel file with macro is activated
    
	'FileSearch object
        strPath = locale 
    strFile = Dir(strPath + "*.xls") 'looking for all xls files in specific folder
     
    Do While strFile <> "" 'loop through all files (ie until none more are found)
        Set wb = Workbooks.Open(strPath & strFile) 'each opened file gets set to variable wb

                actwindow = ActiveWindow.Caption ' get name of active window
                
                Windows(actwindow).Activate 		'get name of file for later --> although could just use wb variable
                Sheets("INPUT").Select 				'select a tab
                Range("F2").Select					'select a cell
                Selection.Copy						'copy data
                Windows(spreadsheetname).Activate	'activate excel with macro to copy data to
                Range("A7").Select
                ActiveCell.Offset(0, gcell).Range("A1").Select 'select cell A7 than offset by number of rows entered to get to first free row 
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False			'paste the data as values only
                
                              
                Windows(actwindow).Activate ' 'repeat similar to above but select range of cells 
                Sheets("INPUT").Select
                Range("B41:c41").Select
                Range(Selection, Selection.End(xlDown)).Select 'this selects from current cell to end of the list (ie end of values in a column assuming no blanks)
                Selection.Copy
                Windows(spreadsheetname).Activate
                Range("A9").Select
                ActiveCell.Offset(0, gcell).Range("A1").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                              
                
                Windows(actwindow).Activate 'as above but different offset
                Sheets("INPUT").Select
                Range("E41").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
                Windows(spreadsheetname).Activate
                Range("A9").Select
                ActiveCell.Offset(0, gcell + 2).Range("A1").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False             
                              
                Windows(actwindow).Activate
                ActiveSheet.Range("A1").Copy 'clear clipboard contents so no large amount of data in clipboard message
                'ActiveWindow.Close SaveChanges:=False         'close window, don't save to automatically (avoids pop up), not needed as closing the wb variable
                Windows(spreadsheetname).Activate 'back to template with macro
                gcell = gcell + 5 'add to count (ie 5 rows used from this file, next will have the same no. of rows)
'                End If
                
        wb.Close True 'close file ready to open next in the directory list
        strFile = Dir

            Loop

Application.ScreenUpdating = True
End Sub
