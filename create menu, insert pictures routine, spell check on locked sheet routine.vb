'=========-Modules--==========
Sub CreateMenubar()   		'create some buttons in a personalised menubar excel

    Dim iCtr As Long
    Dim MacNames As Variant
    Dim CapNamess As Variant
    Dim TipText As Variant

    Call RemoveMenubar 		'first remove the exisitng menubar

    MacNames = Array("INSERTPIC", "SPELLCHECK", "SPELLCHECK2", "location") 'name of the macros
    CapNamess = Array("INSERT PIC", "SPELL CHECK SHEET", "SELECTION", "set location to report location - DO ONCE!") 'name the buttons
    TipText = Array("INSERT PICTURE", "SPELL CHECK SHEET", "SPELL CHECK SELECTION", " makes it easier to find location for 1st picture if pictures and reports are stored in same location") 'tips when hovering over buttons

    With Application.CommandBars.Add 'add the new menu bar with buttons defined above
        .Name = "PIC/SPELL"
        .Left = 790
        .Top = 1000
        .Protection = msoBarNoProtection
        .Visible = True
        .Position = msoBarTop
        
        For iCtr = LBound(MacNames) To UBound(MacNames)
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "'" & ThisWorkbook.Name & "'!" & MacNames(iCtr)
                .Caption = CapNamess(iCtr)
                .Style = msoButtonIconAndCaption
                .FaceId = 71 + iCtr
                .TooltipText = TipText(iCtr)
            End With
        Next iCtr
    End With
	
End Sub
'------------------------------------
Sub noshow()
CommandBars("PIC/SPELL").Visible = False 'hide the menubar
End Sub
'------------------------------------
Sub Show()
CommandBars("PIC/SPELL").Visible = True 'show the menubar defined prevously
End Sub
'------------------------------------
Sub ferme()
CommandBars("PIC/SPELL").Delete 'close/delete the menubar whe moving to other application
End Sub
'---------------------------------------------------------------
Sub SPELLCHECK() 'spell check entire sheet

If ActiveSheet.Name = "SUMMARY-PRINT" Then
            If Range("CT23") = 1 Then 'routine here to check if cells have len(cell) >909, if so warns that this will not check the cell
            ret = "TEST PROCEUDRE"
            End If
            If Range("CT24") = 1 Then
            ret2 = "CONCLUSIONS"
            End If
            If Range("CT24") = 0 And Range("CT23") = 0 Then
            Else
            MsgBox "Auto spell check will not work on " & vbCr & ret & vbCr & ret2, vbExclamation, "MANUAL SPELL CHECK"
            End If
End If
    ActiveSheet.Unprotect ("pas585") 'unlock the sheet
    ActiveSheet.CheckSpelling		'do the spelling
    ActiveSheet.Protect Password:="pas585", DrawingObjects:=False, Contents:=True, Scenarios:=True 're-lock the sheet
End Sub
'---------------------------------------------------------------
Sub SPELLCHECK2() 'spell check the selected cells only

If ActiveSheet.Name = "SUMMARY-PRINT" Then
        Dim rg1 As Range 'define a range --> again check if includes certain cells known to exceed len(cell)>909 on occasion, so check those cells
        Set rg1 = Range("C23:AN37")
        Set isect = Application.Intersect(Selection, rg1)
        If isect Is Nothing Then
        Else
            If Range("CT23") = 1 Then
            ret = "TEST PROCEUDRE"
            End If
            If Range("CT24") = 1 Then
            ret2 = "CONCLUSIONS"
            End If
            If Range("CT24") = 0 And Range("CT23") = 0 Then
            Else
            MsgBox "Auto spell check will not work on " & vbCr & ret & vbCr & ret2, vbExclamation, "MANUAL SPELL CHECK"     
            End If            
        End If
End If
    ActiveSheet.Unprotect ("pas585")
    Application.ScreenUpdating = False
    Range(Selection.Address, ActiveCell.Offset(0, 1)).Select
    Application.CommandBars.FindControl(ID:=2).Execute  'spell check command
    ActiveCell.Select
    Application.ScreenUpdating = True
    ActiveSheet.Protect Password:="pas585", DrawingObjects:=False, Contents:=True, Scenarios:=True
End Sub
'------------------------------------
Sub location() 'set location to activeworkbook to find pictures stored in same folder quicker
ChDir ActiveWorkbook.Path
End Sub
'--------------------------------------------------------------------------------
Sub INSERTPIC()
ActiveSheet.Unprotect ("pas585")
Application.ScreenUpdating = False

Dim i As Integer
Dim pict As Variant
Dim ImgFileFormat As String
h = Selection.Height    'height of cell selected for where to place picture
w = Selection.Width		'width of cell
ImgFileFormat = "Image Files (*.jpg),others, tif (*.tif),*.tif, bmp (*.bmp),*.bmp" 'define image file format to look for
pict = Application.GetOpenFilename(ImgFileFormat, MultiSelect:=True) 'select the picture or pictures

If TypeName(pict) = "Boolean" Then 'exit routine if nothing selected
	ActiveSheet.Protect Password:="pas585", DrawingObjects:=False, Contents:=True, Scenarios:=True
	Application.ScreenUpdating = True
	Exit Sub
End If

For i = LBound(pict) To UBound(pict) 'loop round for each picture selected
	ActiveSheet.Pictures.INSERT(pict(i)).Selec
	
Selection.ShapeRange.LockAspectRatio = msoTrue	 'following cells re-size pictures to fit the cell, maintaining aspect ratio
If w < 50 Or h < 50 Then
	Selection.ShapeRange.Width = 300
	ret = Selection.ShapeRange.Height
		If ret > 300 Then
			Selection.ShapeRange.Height = 302.25
		End If
Else
	Selection.ShapeRange.Width = w - 4
	If Selection.ShapeRange.Height > h - 4 Then 'ret > 50 Then
		Selection.ShapeRange.Height = h - 4
	End If
End If

Selection.Copy 'follwing code to ensure pictures are inserted as images not linked images
Selection.Delete
ActiveSheet.PasteSpecial Format:="Picture (JPEG)", Link:=False, _
DisplayAsIcon:=False

Selection.ShapeRange.IncrementLeft (w - Selection.ShapeRange.Width) / 2 'then move pictures so perfectly central in the cell
Selection.ShapeRange.IncrementTop (h - Selection.ShapeRange.Height) / 2
   
Next i

ActiveSheet.Protect Password:="pas585", DrawingObjects:=False, Contents:=True, Scenarios:=True
Application.ScreenUpdating = True
End Sub

