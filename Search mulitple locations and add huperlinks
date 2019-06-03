'module for automatically searching through defined directories for test data files and then inserting their hyperlinks and/or details
'TO HAVE A TABLE WITH HYPERLINKS TO ALL TEST DATA, REPORTS FOR EACH ROW - 1 ROW PER PRODUCT

Public a As String, n, pa, pb, pc, pd, pe, ty As Integer, p_no As String

Sub Find_Multiple() 'MULIT FIND
Application.ScreenUpdating = False
Dim inputs As Variant
Dim inputs2 As Variant
Dim Locs As Variant ' location vector listing all different locations to search
Dim Report_no As Variant
Dim strWindowsFolder As String
Dim Locations As Variant
Dim keywords As Variant
Dim M, M2 As Integer

M2 = ActiveSheet.Range("C11:C200").Cells.SpecialCells(xlCellTypeConstants).Count 'number of rows
M = InputBox("how many products?, max (all) is displayed in box", Default, M2)
ReDim inputs(8, M) ' 8 x m matrix array --> m is number of products to search
ReDim inputs2(12) 'vector with generic details of product --> linked to storage locations for the product
ReDim Locs(5)
ReDim Report_no(4, 2)
ReDim Locations(7)
base_bk = ActiveWorkbook.Name

'Following data taken from excel cells for use in defining product
Sheets("Locs").Select
inputs2(1) = Range("D1") 'Product family
inputs2(2) = Range("L6") 'additional location
inputs2(3) = Range("E1") 'Product type
inputs2(4) = inputs2(2) & "\" & inputs2(1) & "\" & inputs2(3) 'directory for test1 data"
inputs2(5) = inputs2(2) & "\" & "DFP4" & "\" & inputs2(3) 'directory for test2 data"
inputs2(6) = Range("L7") 'ADDITIONAL USER INPUT CELLS TO SEARCH IN DIFFERENT LOCATIONS AND CHANGE IN FOLDER NAMING CONVENTIONS ETC
inputs2(7) = Range("F1") 
inputs2(8) = Range("G1") 
inputs2(9) = Range("L11") & Range("L12") & "\" & inputs2(7)
inputs2(10) = Range("L11") & Range("L13") & "\" & inputs2(8)
inputs2(11) = Range("L11") & Range("L12")
inputs2(12) = Range("L11") & Range("L13")

'following code to define locations for looking for the test data
Locations(0) = inputs2(4)  
Locations(1) = inputs2(5)  
Locations(2) = Range("L7") 
Locations(3) = inputs2(9)  
Locations(4) = inputs2(10) 
Locations(5) = inputs2(11) 
Locations(6) = inputs2(12) 
Locations(7) = Range("L16")
'code for report locations
Locs(0) = Range("L1") '2017
Locs(1) = Range("L2") '2018
Locs(2) = Range("L3") '2019
Sheets("Home_page").Select

a = ActiveCell.Address
For n = 1 To M
	Range(a).Offset(n - 1, 0).Select
	inputs(0, n) = ActiveCell.Value 'product number used for searching 

	'_____________'First find the report number for the product listed in one of 3 multiple excel sheets, and then add hyperlink to the report___
	j = 0
	For i = 0 To 2
		Application.Workbooks.Open Filename:=Locs(i), UpdateLinks:=False
		temp_bk = ActiveWorkbook.Name
		Sheets("Register").Select

		With Worksheets("Register").Range("c10:c1000")
			 Set c = .Find(inputs(0, n), LookIn:=xlValues)
			 If Not c Is Nothing Then
				firstAddress = c.Address
				Do
				Report_no(j, 0) = Range(c.Address).Offset(0, -1)
					If i = 0 Then
					Report_no(j, 1) = "18_" 'some code to make sure the report number has the correct digits as per the file name
					Report_no(j, 2) = "2018"
					ElseIf i = 1 Then
					Report_no(j, 1) = "17_"
					Report_no(j, 2) = "2017"
					Else
					Report_no(j, 1) = "19_"
					Report_no(j, 2) = "2019"
					End If
				j = j + 1
				Set c = .FindNext(c)
				Loop While c.Address <> firstAddress
			  End If
		DoneFinding:
		End With

		ActiveWorkbook.Close SaveChanges:=False
		Workbooks(base_bk).Activate

	Next i
		
	For k = 0 To j - 1
		le = Len(Report_no(k, 0))
			If le = 1 Then
			Report_no(k, 0) = "00" & Report_no(k, 0)
			ElseIf le = 2 Then
			Report_no(k, 0) = "0" & Report_no(k, 0)
			End If
			
		ActiveCell.Offset(k, 16).Select 'just select the column for inserting report number hyperlink
		ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
				"H:\Reports\" & Report_no(k, 2) & "\REPORT_" & Report_no(k, 0) _
				, TextToDisplay:=Report_no(k, 1) & Report_no(k, 0)
		ActiveCell.Offset(-k, -16).Select
	Next k
	'__________________________________________________________________________________________________________________________________

	pa = 0 'some variables to be passed to the sub-routine
	pb = 0
	pc = 0
	pd = 0
	pe = 0
	p_no = Right(ActiveCell, 8)
	For ty = 0 To 7 'loop through the 8 directories added above for the files
	strWindowsFolder = Locations(ty)
	Call photo_links(strWindowsFolder) 'hperlinks for reports photos
	Call test1_links(strWindowsFolder) 'hperlinks for test 1 files
	Call test2_links(strWindowsFolder) 'hperlinks for test 2 files
	Next

	Debug.Print n; " complete " & Timer
Next n
Application.ScreenUpdating = True
End Sub

'__________________________________________________________________________________________________________________________________
'__________________________________________________________________________________________________________________________________
'__________________________________________________________________________________________________________________________________

Sub photo_links(strFolderPath As String) ', a As String) 'n As Integer,
'report photos folder
    Dim objFileSystem As Object
    Dim objFileSystem2 As Object
    Dim objFolder As Object
    Dim objFolder2 As Object    
    Dim nLastRow As Integer
    Dim prev5, oldy5 As Integer
	
    On Error GoTo Handler:
    
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFileSystem.GetFolder(strFolderPath)
 
If ty = 0 Or ty = 1 Then
    'Process all folders and subfolders recursively
    If objFolder.Subfolders.Count > 0 Then
       For Each objSubFolder In objFolder.Subfolders
           'Skip the system and hidden folders
           If InStr(1, objSubFolder, p_no, 1) > 0 Then
                If ((objSubFolder.Attributes And 2) = 0) And ((objSubFolder.Attributes And 4) = 0) Then
                Set objFileSystem2 = CreateObject("Scripting.FileSystemObject")
                Set objFolder2 = objFileSystem2.GetFolder(objSubFolder.Path)
                Exit For
                End If
           End If
       Next
    End If
Else 'ty = 2
    Set objFileSystem2 = objFileSystem2
    Set objFolder2 = objFolder
End If

If objFolder2.Subfolders.Count > 0 Then
    For Each objSubFolder In objFolder2.Subfolders
        If InStr(1, objSubFolder, p_no, 1) > 0 And InStr(1, objSubFolder, "rep_photos", 1) > 0 Then
            If ((objSubFolder.Attributes And 2) = 0) And ((objSubFolder.Attributes And 4) = 0) Then
            'add location to excel
                   With ActiveSheet
                     Do While Not IsEmpty(Range(a).Offset(n - 1, 66 + pc)) ' find and empty cell
                     pc = pc + 1
                     Loop
                        
                        prev5 = 0
                        For oldy5 = 0 To pe - 1 'check link is not same as previous
                            If objSubFolder.Name = Range(a).Offset(n - 1, 66 + oldy5).Hyperlinks(1).Name Then
                            pc = pc - 1 'reset the pc cell don't skip empty cell for inserting hyperlinks
                            prev5 = 1
                            End If
                        Next
                        
                        If prev5 = 0 Then
                        Range(a).Offset(n - 1, 66 + pc).Select
                        ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:=objSubFolder.Path _
                        , TextToDisplay:=objSubFolder.Name
                        End If
                        
                   End With
            End If
        End If
    Next
End If

Handler:   'Debug.Print "folder doesn't exist" & strFolderPath
End Sub
'__________________________________________________________________________________________________________________________________
'__________________________________________________________________________________________________________________________________

Sub test1_links(strFolderPath As String) ', a As String) 'n As Integer,
'test1 folder
    Dim objFileSystem As Object
    Dim objFileSystem2 As Object
    Dim objFolder As Object
    Dim objFolder2 As Object
    Dim nLastRow As Integer
    Dim oldy, oldy2 As Integer
    Dim prev, prev2 As Integer
    Dim slink As String
    Dim c2, c As String
    Dim objhyperlink As Hyperlink    
         
    On Error GoTo Handler:
    
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFileSystem.GetFolder(strFolderPath)
 
If ty = 0 Or ty = 1 Then
    'Process all folders and subfolders recursively
    If objFolder.Subfolders.Count > 0 Then
       For Each objSubFolder In objFolder.Subfolders
           'Skip the system and hidden folders
           If InStr(1, objSubFolder, p_no, 1) > 0 Then
                If ((objSubFolder.Attributes And 2) = 0) And ((objSubFolder.Attributes And 4) = 0) Then
                Set objFileSystem2 = CreateObject("Scripting.FileSystemObject")
                Set objFolder2 = objFileSystem2.GetFolder(objSubFolder.Path)
                Exit For
                End If
           End If
       Next
    End If
Else 'ty = 2
    Set objFileSystem2 = objFileSystem2
    Set objFolder2 = objFolder
End If

If objFolder2.Subfolders.Count > 0 Then
    For Each objSubFolder In objFolder2.Subfolders
        If InStr(1, objSubFolder, p_no, 1) > 0 And InStr(1, objSubFolder, "test1", 1) > 0 Then
            If ((objSubFolder.Attributes And 2) = 0) And ((objSubFolder.Attributes And 4) = 0) Then
                   f = 0 'flag to say file found in the subfolder
                   '______________________________________
                   'now add the file hyperlinks
                    For Each objFile In objSubFolder.Files
                    If InStr(1, objFile.Path, "test1", 1) > 0 And objFile Like "*.xlsm*" Then
                    If Not (InStr(1, objFile.Path, "etest1or", 1) > 0 Or InStr(1, objFile.Path, "autest1", 1) > 0 Or InStr(1, objFile.Path, "AvLog 10", 1) > 0 Or InStr(1, objFile.Path, "~$", 1) > 0) Then
                    f = 1 'flag to say file found
                        With ActiveSheet
                        Do While Not IsEmpty(Range(a).Offset(n - 1, 56 + pa)) ' find and empty cell
                        pa = pa + 1
                        Loop
                        Range(a).Offset(n - 1, 56 + pa).Select

                        prev2 = 0
                        For oldy2 = 0 To pa - 1 'check link is not same as previous
                            If objFile.Name = Range(a).Offset(n - 1, 56 + oldy2).Hyperlinks(1).Name And objSubFolder.Name = Range(a).Offset(n - 1, 56 + oldy2).Hyperlinks(1).SubAddress Then
                            
                            pa = pa - 1 'reset the pd cell don't skip empty cell for inserting hyperlinks
                            prev2 = 1
                            End If
                        Next
                        If prev2 = 0 Then
                        Range(a).Offset(n - 1, 56 + pa).Select
                        ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:=objFile.Path _
                        , SubAddress:=objSubFolder.Name, TextToDisplay:=objFile.Name
                        End If
                        
                        End With
                    End If
                    End If
                    Next
                    '______________________________________
                   'add location to excel
                   If f = 1 Then 'only do if folder contains a file above
                        With ActiveSheet
              '              If objsubFolder.Name <> Range(a).Offset(n - 1, 68 + pd) Then
                             Do While Not IsEmpty(Range(a).Offset(n - 1, 68 + pd)) ' find and empty cell
                             pd = pd + 1
                             Loop
                                prev = 0
                                For oldy = 0 To pd - 1 'check link is not same as previous
                                    If objSubFolder.Name = Range(a).Offset(n - 1, 68 + oldy).Hyperlinks(1).Name Then
                                    pd = pd - 1 'reset the pd cell don't skip empty cell for inserting hyperlinks
                                    prev = 1
                                    End If
                                Next
                            If prev = 0 Then
                            Range(a).Offset(n - 1, 68 + pd).Select
                            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:=objSubFolder.Path _
                            , TextToDisplay:=objSubFolder.Name
                            End If
                        End With
                    End If
                   '__________________
            End If
        End If
    Next
End If

Handler:   'Debug.Print "folder doesn't exist" & strFolderPath
'Exit Sub
End Sub
'__________________________________________________________________________________________________________________________________
'__________________________________________________________________________________________________________________________________

Sub test2_links(strFolderPath As String) ', a As String) 'n As Integer,
'test2 folder and file hyperlinks add to excel
    Dim objFileSystem As Object
    Dim objFileSystem2 As Object
    Dim objFolder As Object
    Dim objFolder2 As Object
    Dim nLastRow As Integer
    Dim oldy3, oldy4 As Integer
    Dim prev3, prev4 As Integer

    On Error GoTo Handler:
    
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFileSystem.GetFolder(strFolderPath)
 
If ty = 0 Or ty = 1 Then
    'Process all folders and subfolders recursively
    If objFolder.Subfolders.Count > 0 Then
       For Each objSubFolder In objFolder.Subfolders
           'Skip the system and hidden folders
           If InStr(1, objSubFolder, p_no, 1) > 0 Then
                If ((objSubFolder.Attributes And 2) = 0) And ((objSubFolder.Attributes And 4) = 0) Then
                Set objFileSystem2 = CreateObject("Scripting.FileSystemObject")
                Set objFolder2 = objFileSystem2.GetFolder(objSubFolder.Path)
                Exit For
                End If
           End If
       Next
    End If
Else 'ty = 2
    Set objFileSystem2 = objFileSystem2
    Set objFolder2 = objFolder
End If

If objFolder2.Subfolders.Count > 0 Then
    For Each objSubFolder In objFolder2.Subfolders
        If InStr(1, objSubFolder, p_no, 1) > 0 And InStr(1, objSubFolder, "test2", 1) > 0 Then
            If ((objSubFolder.Attributes And 2) = 0) And ((objSubFolder.Attributes And 4) = 0) Then
                   f = 0 'flag to say file found in the subfolder
                   '___________________________________
                   'now add the file hyperlinks
                    For Each objFile In objSubFolder.Files
                        If objFile.Type = "JPG File" Then
                        f = 1 'flag to say file found
                        With ActiveSheet
                        Do While Not IsEmpty(Range(a).Offset(n - 1, 61 + pb)) ' find and empty cell
                        pb = pb + 1
                        Loop
                        prev3 = 0
                        For oldy3 = 0 To pb - 1 'check link is not same as previous
                            If objFile.Name = Range(a).Offset(n - 1, 61 + oldy3).Hyperlinks(1).Name And objSubFolder.Name = Range(a).Offset(n - 1, 61 + oldy3).Hyperlinks(1).SubAddress Then
                            
                            pb = pb - 1 'reset the pd cell don't skip empty cell for inserting hyperlinks
                            prev3 = 1
                            End If
                        Next
                        If prev3 = 0 Then
                        Range(a).Offset(n - 1, 61 + pb).Select
                        ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:=objFile.Path _
                        , SubAddress:=objSubFolder.Name, TextToDisplay:=objFile.Name
                        End If                        
                        
                        ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:=objFile.Path _
                        , TextToDisplay:=objFile.Name
                        End With
                        End If
                    Next
                    '______________________________________
                   'add location to excel
                   If f = 1 Then 'only do if folder contains a file above
                    With ActiveSheet
                      Do While Not IsEmpty(Range(a).Offset(n - 1, 73 + pe)) ' find and empty cell
                      pe = pe + 1
                      Loop
                        prev4 = 0
                        For oldy4 = 0 To pe - 1 'check link is not same as previous
                            If objSubFolder.Name = Range(a).Offset(n - 1, 73 + oldy4).Hyperlinks(1).Name Then
                            pe = pe - 1 'reset the pd cell don't skip empty cell for inserting hyperlinks
                            prev4 = 1
                            End If
                        Next
                        If prev4 = 0 Then
                        Range(a).Offset(n - 1, 73 + pe).Select
                        ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:=objSubFolder.Path _
                        , TextToDisplay:=objSubFolder.Name
                        End If
                    End With
                   End If
                   '_________________________
            End If
        End If
    Next
End If

Handler:   'Debug.Print "folder doesn't exist" & strFolderPath
'Exit Sub
End Sub

