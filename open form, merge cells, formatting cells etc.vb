'=========-Workbook OBject--==========
Private Sub Workbook_Activate()
Call CreateMenubar 'create a specialised menubar in excel
Call Show 			'show that menubar
End Sub
'------------------------------------
Private Sub Workbook_Open()
Sheets("SUMMARY-PRINT").Select
If IsEmpty(Range("AH11")) Then 'a flag to show if form was ran prevoiiusly, if not then run the form
	slt.Show False  
End If
Do While Not IsEmpty(ActiveCell)
	ActiveCell.Offset(1, 0).Select
Loop
'------------------------------------
End Sub

'=========-SLT from with 3 buttons--==========

Private Sub button_1_Click()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False 'disable pop-ups    
    Sheets("MAIN REPORT - PRINT2").Select 'delete the alternative report format (used on button 2)
    ActiveWindow.SelectedSheets.Delete
    Sheets("SUMMARY-PRINT").Select    
    ActiveSheet.Shapes("Text Box 1220").Select 'delete a box
    Selection.Delete
    
    Range("AF2").Value = "type_1" 'slected type_1
    Range("AH11").Value = Date	'record date created
    Range("U2").Select
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
Unload Me
End Sub
'------------------------------------------------------------------------
Private Sub button2_Click()    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False    
    Sheets("MAIN REPORT - PRINT").Select 'delete the primary report format (used on button 1)
    ActiveWindow.SelectedSheets.Delete    
       
    Sheets("SUMMARY-PRINT").Select 'some merging routines to change report fomat into style 2
    ActiveSheet.Unprotect ("pas585")
    Range("T65:Y80").Select
    Selection.UnMerge
    Selection.ClearContents    
    Rows("426:432").Select
    Range("A432").Activate
    Selection.EntireRow.Hidden = True    
    Range("C424:W425").Select
    Range("U425").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    
    Rows("412:423").Select 'hide some rows
    Range("A423").Activate
    Selection.EntireRow.Hidden = True
	
    Rows("400:403").Select
    Selection.EntireRow.Hidden = True
    Range("AE425").Select
    
    Range("AA411").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    
    Sheets("MAIN REPORT - PRINT2").Select
    Range("HZ1:IE16").Select
    Selection.Copy
           
    Sheets("SUMMARY-PRINT").Select
    Range("T65:Y80").Select
    ActiveSheet.Paste
    Range("Z65:Z87").Value = ""
       Range("AH11").Value = Date
    Range("Q2").Value = 1
    Range("Q6").Value = "NA"    
    
    Sheets("MAIN REPORT - PRINT2").Select
    Range("B3").Select
    Sheets("SUMMARY-PRINT").Select
    Range("U2").Select   

    ActiveSheet.Protect Password:="pas585", DrawingObjects:=False, Contents:=True, Scenarios:=True
    Application.EnableEvents = True
    Range("Q4").Value = "NO"
    Range("Q5").Value = "NO"
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
Unload Me 'close the form
End Sub
'------------------------------------
Private Sub CommandButton1_Click() ' 3rd button not ready yet, so just exit for now
MsgBox ("not available yet, exiting")
End Sub
