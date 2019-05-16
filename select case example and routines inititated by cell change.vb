'=========-Excel Objects Macros--==========
Private Sub TempCombo_Change() 'simple routine to grab the tempcombo box value and put into the selected cell
gju = TempCombo.Value
ActiveCell.FormulaR1C1 = gju
End Sub
'------------------------------------_________________
Private Sub Worksheet_Change(ByVal Target As Range)

Application.ScreenUpdating = False
ActiveSheet.Unprotect ("pas5853")

Dim icolor As Integer
Dim gt As String
Dim KeyCells As String
Dim KeyCells5 As String

gt = Target.Address 'address of selected cell

ptype = Range("AF2").Value
If ptype = "type1" Then
   KeyCells6 = "Q3:Q8,AC13"    
   If Not Application.Intersect(Target, Range(KeyCells6)) _ 'first of number of checks to see if the selected cell is in the range of the 'keycells, if so go to another sub-routine
   Is Nothing Then KeyCellsChanged6 (gt)
Else
   KeyCells = "Q2:Q8,AC13"
   If Not Application.Intersect(Target, Range(KeyCells)) _
   Is Nothing Then KeyCellsChanged (gt)   
End If
KeyCells5 = "J63:J69,J73:J81"
If Not Application.Intersect(Target, Range(KeyCells5)) _
Is Nothing Then keyCellsChanged5 (gt)

Application.EnableEvents = True
Application.ScreenUpdating = True
ActiveSheet.Protect Password:="pas5853", DrawingObjects:=False, Contents:=True, Scenarios:=True
End Sub
'---------------------------------------------------------------------------------
Sub KeyCellsChanged(gt)

Select Case gt
Case "$Q$2" 
	HEADS = Range(gt).Value
	Sheets("MAIN REPORT - PRINT").Select
	ActiveSheet.Unprotect ("pas585")
	Application.EnableEvents = False	  
	If HEADS = 2 Then
		Sheets("MAIN REPORT - PRINT").Range("AJ123,,AK548,AH143:AH145,AH148,AH150,,AH284").Value = "NA"
	Elseif HEADS = 3 Then
		Sheets("MAIN REPORT - PRINT").Range("AJ123,AJ159,AJ196,AJ228,AH1284,AH286,AH320:AH322,AH324,AH355:AH356,AH358,AH360").Value = ""
	End If
	ActiveSheet.Protect Password:="pas585", DrawingObjects:=False, Contents:=True, Scenarios:=True
	Application.EnableEvents = True
	Sheets("SUMMARY-PRINT").Select
    
Case "$Q$3" 
	'similar  code here 
	decay = Range(gt).Value   
    If Left(decay, 3) = "DAI" Then
	else
	end if
	Set Rng1 = Union(Range( _
        "S7,AO7,S29,AO29,T50,AO50,S72,AO72,S99,AO99,S121,AO121,S141,AO141,S163,S332,AO332" _
        ), Range( _
        "AO163,S190,AO190,S212,AO212,S234,AO234,S258,AO258,S285,AO285,S310,AO310,S353" _
        ))
	For Each Cell In Rng1
		Select Case Cell.Value
		Case 0 To 3
			Cell.Interior.ColorIndex = 50
			Cell.Font.ColorIndex = 2
			Cell.Select
		Case 4 To 4
			Cell.Interior.ColorIndex = 45
			Cell.Font.ColorIndex = 0
			Cell.Select
		Case 5 To 5
			Cell.Interior.ColorIndex = 3
			Cell.Font.ColorIndex = 0
			Cell.Select
		Case "NA"
			Cell.Interior.ColorIndex = 37
			Cell.Font.ColorIndex = 0
			Cell.Select
		Case Else
			Cell.Interior.ColorIndex = 34
			Cell.Select
		End Select
	Next
Case "$Q$4" 
	'similar  code here              
'then other cases    
End Select

End Sub
'---------------------------------------------------------------------------------

Private Sub Worksheet_selectionChange(ByVal Target As Range)
Dim str As String
Dim cboTemp As OLEObject
Dim ws As Worksheet
Set ws = ActiveSheet
On Error GoTo errHandler

Set cboTemp = ws.OLEObjects("TempCombo")
  On Error Resume Next
If cboTemp.Visible = True Then
  With cboTemp
    .Top = 50
    .Left = 10
    .ListFillRange = ""
    .LinkedCell = ""
    .Visible = False
    .Value = ""
  End With
End If

  On Error GoTo errHandler
  If Target.Validation.Type = 3 Then  'if the cell contains a data validation list then ....
    Application.EnableEvents = False    
    str = Target.Validation.Formula1 '....get the data validation formula
    str = Right(str, Len(str) - 1)
    With cboTemp 'shoe the combobox
      .Left = Target.Left
      .Top = Target.Top
      .Width = Target.Width + 25
      .Height = Target.Height + 5
      .LinkedCell = Target.Address
    End With
    
    With cboTemp
      .ListFillRange = ws.Range(str).Address
    End With
    cboTemp.Activate
  End If

exitHandler:
  Application.ScreenUpdating = True
  Application.EnableEvents = True
    Exit Sub
errHandler:
  Resume exitHandler

End Sub
'------------------------------------
Private Sub TextBox1_Change()
    'n = Cells(10, 8).Value
    n = TextBox1.Value
    'MsgBox n
    If n = 0 Then
    n = 1
    'MsgBox n
    End If
    If n = "" Then
    n = 1
    'MsgBox n
    End If
    If n = 1 Then
    Rows("11:11").rowheight = 15.75
    Rows("27:27").rowheight = 15.75
    Else
    If n Mod 2 = 0 Then
    Rows("11:11").rowheight = 15 * n / 2
    End If
    If n Mod 2 = 1 Then
    Rows("11:11").rowheight = 15 * (n + 1) / 2
    End If
    End If
    If n > 4 And n < 9 Then
    Rows("27:27").rowheight = 30
    End If
    If n > 8 Then
    Rows("27:27").rowheight = 45
    End If
End Sub