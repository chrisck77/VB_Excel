'module for automatically searching through defined directories for test data files and then inserting their hyperlinks and/or details
'TO HAVE A TABLE WITH HYPERLINKS TO ALL TEST DATA, REPORTS FOR EACH ROW - 1 ROW PER PRODUCT
Public a As String, n, n2, pa, pb, pc, pd, pe, ty As Integer, p_no As String
Sub Find_Multiple() 'MULIT FIND
Application.ScreenUpdating = False
Dim inputs As Variant
Dim inputs2 As Variant
Dim Locs As Variant ' location vector wher to look for files
Dim EOT_no As Variant
Dim strWindowsFolder As String
Dim Locations As Variant
Dim keywords As Variant
Dim M, M2, n_components As Integer

If ActiveCell.Value = "" Then Exit Sub

c_row = ActiveCell.Row
myLastRow = Cells(Rows.Count, ActiveCell.Column).End(xlUp).Row 'last cell in that column with data in it
c_col = Mid(ActiveCell.Address, 2, 1)
n_components = WorksheetFunction.CountA(Range(c_col & c_row & ":" & c_col & myLastRow)) ' no of cells with data (no. of components /numbers)

M = CInt(InputBox("how many components?, max (all) is displayed in box", Default, n_components))

If M > n_components Or M = 0 Then
MsgBox ("You entered to many components, more than max OR 0 components please try again")
Exit Sub
End If

ReDim inputs(8, M) ' 8 x m matrix array --> m is number of components to search
ReDim inputs2(12) 'vector with generic details of sheet (component type etc)
ReDim Locs(5)
ReDim EOT_no(4, 2)
ReDim Locations(7)
base_bk = ActiveWorkbook.Name
base_tab = ActiveSheet.Name

Sheets("Locs").Select
inputs2(1) = Range("D1") 'ie "DFP6" -->add the component family type
inputs2(2) = Range("L6") 'ie "\\ukgil-ap15\diadem001$\Endurance Data"
inputs2(3) = Range("E1") 'ie  "HMC"
inputs2(4) = inputs2(2) & "\" & inputs2(1) & "\" & inputs2(3) ' test1 directory endurance data"
inputs2(5) = inputs2(2) & "\" & "DFP4" & "\" & inputs2(3) ' DFP4 test1 directory endurance data"
inputs2(6) = Range("L7") 'ie "\\ukgil-ap15\diadem001$\Endurance Live\To Be Archived"
inputs2(7) = Range("F1") 'Manor pod 7 naming convention for project -- ie HMC
inputs2(8) = Range("G1") 'Manor pod 8 naming convention for project -- ie HMC U components
inputs2(9) = Range("L11") & Range("L12") & "\" & inputs2(7) ' Manor pod 7 directory + project folder
inputs2(10) = Range("L11") & Range("L13") & "\" & inputs2(8) ' Manor pod 8 directory inputs2(2)+ project folder
inputs2(11) = Range("L11") & Range("L12") 'Manor pod 7 directory
inputs2(12) = Range("L11") & Range("L13") 'Manor pod 8 directory


Locations(0) = inputs2(4)  'component type DFP3 or DFP6 folder on endurance data
Locations(1) = inputs2(5)  'component type DFP4 folder on endurance data
Locations(2) = Range("L7") '\\ukgil-ap15\diadem001$\Endurance Live\To Be Archived
Locations(3) = inputs2(9)  'manor pod 7 (i.e HMC folder in manor pod 7)
Locations(4) = inputs2(10) 'manor pod 8 (i.e HMC folder in manor pod 8)
Locations(5) = inputs2(11) 'Manor pod 7 directory
Locations(6) = inputs2(12) 'Manor pod 8 directory
Locations(7) = Range("L16")
Locs(0) = Range("L1") '2018
Locs(1) = Range("L2") '2017
Locs(2) = Range("L3") '2019

'Sheets("GWM CR").Select
Sheets(base_tab).Select


a = ActiveCell.Address
blank = 0 'no's of blank cells

For n = 1 To M
    Do While Range(a).Offset(n - 1 + blank, 0) = ""
    blank = blank + 1
    Loop

n2 = n + blank
Range(a).Offset(n2 - 1, 0).Select
inputs(0, n) = ActiveCell.Value 'component number

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
        EOT_no(j, 0) = Range(c.Address).Offset(0, -1)
            If i = 0 Then
            EOT_no(j, 1) = "18_"
            EOT_no(j, 2) = "2018"
            ElseIf i = 1 Then
            EOT_no(j, 1) = "17_"
            EOT_no(j, 2) = "2017"
            Else
            EOT_no(j, 1) = "19_"
            EOT_no(j, 2) = "2019"
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

le = Len(EOT_no(k, 0))
    If le = 1 Then
    EOT_no(k, 0) = "00" & EOT_no(k, 0)
    ElseIf le = 2 Then
    EOT_no(k, 0) = "0" & EOT_no(k, 0)
    End If
    
'ActiveCell.Offset(k, 2).Value = EOT_no(k, 1) & EOT_no(k, 0)
ActiveCell.Offset(k, -1).Select
'ActiveCell.Offset(k, 3).Value = "S:\DFP\DFP3\Reports (hyperlinks to central storage)\Drafts\EOT\End of Test Reports 2018\DETR_" & EOT_no(k)
ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
        "S:\DFP\DFP3\Reports (hyperlinks to central storage)\Drafts\EOT\End of Test Reports " & EOT_no(k, 2) & "\DETR_" & EOT_no(k, 0) _
        , TextToDisplay:=EOT_no(k, 1) & EOT_no(k, 0)
ActiveCell.Offset(-k, 1).Select

Next k

    pa = 0 'for offsetting column (so don't overwrite existing data) count of number cells filled in
    pb = 0
    pc = 0
    pd = 0
    pe = 0
    p_no = Right(ActiveCell, 8)
    For ty = 0 To 7 'find files --> add hyperlinks
    strWindowsFolder = Locations(ty)
    Call photo_links(strWindowsFolder) 'hperlinks for strip down folders
        If ty = 0 Or 1 Then
        Call test1_links(strWindowsFolder) 'hperlinks for test1 folders and test1 files
        Call test2_links(strWindowsFolder) 'hperlinks for test2 folders and jpg photo
        Call test1_links2(strWindowsFolder) 'hperlinks for test1 folders and test1 files --> extra for the new fodler structure
        Call test2_links2(strWindowsFolder) 'hperlinks for test2 folders and jpg photo --> extra for the new fodler structure
        Else
        Call test1_links(strWindowsFolder) 'hperlinks for test1 folders and test1 files
        Call test2_links(strWindowsFolder) 'hperlinks for test2 folders and jpg photo
        End If
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

Sub test1_links2(strFolderPath As String) ', a As String) 'n As Integer,
'test1 folder and map file
    Dim objFileSystem As Object
    Dim objFileSystem2 As Object
    Dim objFileSystem3 As Object
    Dim objFolder As Object
    Dim objFolder2 As Object
    Dim objFolder3 As Object
    Dim objFile As Object
    Dim nLastRow As Integer
    Dim oldy, oldy2 As Integer
    Dim prev, prev2 As Integer
    Dim slink As String
    Dim c2, c As String
    Dim objhyperlink As Hyperlink
    
    On Error GoTo Handler:
    
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFileSystem.GetFolder(strFolderPath)
 
    'Process all folders and subfolders recursively
If objFolder.Subfolders.Count > 0 Then
   For Each objSubFolder In objFolder.Subfolders
       'Skip the system and hidden folders
       If InStr(1, objSubFolder, p_no, 1) > 0 Then
            If ((objSubFolder.Attributes And 2) = 0) And ((objSubFolder.Attributes And 4) = 0) Then
            Set objFileSystem3 = CreateObject("Scripting.FileSystemObject")
            Set objFolder3 = objFileSystem3.GetFolder(objSubFolder.Path)
            Exit For
            End If
       End If
   Next
End If


If objFolder3.Subfolders.Count > 0 Then
    For Each objSubFolder In objFolder3.Subfolders
        If InStr(1, objSubFolder, p_no, 1) > 0 And InStr(1, objSubFolder, "burn", 1) > 0 Then
            If ((objSubFolder.Attributes And 2) = 0) And ((objSubFolder.Attributes And 4) = 0) Then
            Set objFileSystem2 = CreateObject("Scripting.FileSystemObject")
            Set objFolder2 = objFileSystem2.GetFolder(objSubFolder.Path)
            Exit For
            End If
        End If
    Next
End If

If objFolder2.Subfolders.Count > 0 Then
    For Each objSubFolder In objFolder2.Subfolders
        If InStr(1, objSubFolder, p_no, 1) > 0 And InStr(1, objSubFolder, "charact", 1) > 0 Then
            If ((objSubFolder.Attributes And 2) = 0) And ((objSubFolder.Attributes And 4) = 0) Then
                   f = 0 'flag to say file found in the subfolder
                   '______________________________________
                   'now add the file hyperlinks
                    For Each objFile In objSubFolder.Files
                    If InStr(1, objFile.Path, "test1", 1) > 0 And objFile Like "*.xlsm*" Then
                    If Not (InStr(1, objFile.Path, "etest1or", 1) > 0 Or InStr(1, objFile.Path, "autest1", 1) > 0 Or InStr(1, objFile.Path, "Averaged Log 10", 1) > 0 Or InStr(1, objFile.Path, "~$", 1) > 0) Then
                    f = 1 'flag to say file found
                        With ActiveSheet
                        Do While Not IsEmpty(Range(a).Offset(n - 1, 46 + pa)) ' find and empty cell
                        pa = pa + 1
                        Loop
                        Range(a).Offset(n - 1, 46 + pa).Select

                        prev2 = 0
                        For oldy2 = 0 To pa - 1 'check link is not same as previous
                            If objFile.Name = Range(a).Offset(n - 1, 46 + oldy2).Hyperlinks(1).Name And objSubFolder.Name = Range(a).Offset(n - 1, 46 + oldy2).Hyperlinks(1).SubAddress Then
                            
                            pa = pa - 1 'reset the pd cell don't skip empty cell for inserting hyperlinks
                            prev2 = 1
                            End If
                        Next
                        If prev2 = 0 Then
                        Range(a).Offset(n - 1, 46 + pa).Select
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
                             Do While Not IsEmpty(Range(a).Offset(n - 1, 58 + pd)) ' find and empty cell
                             pd = pd + 1
                             Loop
                                prev = 0
                                For oldy = 0 To pd - 1 'check link is not same as previous
                                    If objSubFolder.Name = Range(a).Offset(n - 1, 58 + oldy).Hyperlinks(1).Name Then
                                    pd = pd - 1 'reset the pd cell don't skip empty cell for inserting hyperlinks
                                    prev = 1
                                    End If
                                Next
                            If prev = 0 Then
                            Range(a).Offset(n - 1, 58 + pd).Select
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

Sub test2_links2(strFolderPath As String) ', a As String) 'n As Integer,
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
 
If objFolder.Subfolders.Count > 0 Then
   For Each objSubFolder In objFolder.Subfolders
       'Skip the system and hidden folders
       If InStr(1, objSubFolder, p_no, 1) > 0 Then
            If ((objSubFolder.Attributes And 2) = 0) And ((objSubFolder.Attributes And 4) = 0) Then
            Set objFileSystem3 = CreateObject("Scripting.FileSystemObject")
            Set objFolder3 = objFileSystem3.GetFolder(objSubFolder.Path)
            Exit For
            End If
       End If
   Next
End If


If objFolder3.Subfolders.Count > 0 Then
    For Each objSubFolder In objFolder3.Subfolders
        If InStr(1, objSubFolder, p_no, 1) > 0 And InStr(1, objSubFolder, "burn", 1) > 0 Then
            If ((objSubFolder.Attributes And 2) = 0) And ((objSubFolder.Attributes And 4) = 0) Then
            Set objFileSystem2 = CreateObject("Scripting.FileSystemObject")
            Set objFolder2 = objFileSystem2.GetFolder(objSubFolder.Path)
            Exit For
            End If
        End If
    Next
End If


If objFolder2.Subfolders.Count > 0 Then
    For Each objSubFolder In objFolder2.Subfolders
        If InStr(1, objSubFolder, p_no, 1) > 0 And InStr(1, objSubFolder, "burn", 1) > 0 Then
            If ((objSubFolder.Attributes And 2) = 0) And ((objSubFolder.Attributes And 4) = 0) Then
                   f = 0 'flag to say file found in the subfolder
                   '___________________________________
                   'now add the file hyperlinks
                    For Each objFile In objSubFolder.Files
                        If objFile.Type = "JPG File" Then
                        f = 1 'flag to say file found
                        With ActiveSheet
                        Do While Not IsEmpty(Range(a).Offset(n - 1, 51 + pb)) ' find and empty cell
                        pb = pb + 1
                        Loop
                        prev3 = 0
                        For oldy3 = 0 To pb - 1 'check link is not same as previous
                            If objFile.Name = Range(a).Offset(n - 1, 51 + oldy3).Hyperlinks(1).Name And objSubFolder.Name = Range(a).Offset(n - 1, 51 + oldy3).Hyperlinks(1).SubAddress Then
                            
                            pb = pb - 1 'reset the pd cell don't skip empty cell for inserting hyperlinks
                            prev3 = 1
                            End If
                        Next
                        If prev3 = 0 Then
                        Range(a).Offset(n - 1, 51 + pb).Select
                        ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:=objFile.Path _
                        , SubAddress:=objSubFolder.Name, TextToDisplay:=objFile.Name
                        End If
                        
                        
                        ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:=objFile.Path _
                        , TextToDisplay:=objFile.Name
                        'ActiveCell.Offset(n - 1, -51 - pb).Select
                        End With
                        End If
                    Next
                    '______________________________________
                   'add location to excel
                   If f = 1 Then 'only do if folder contains a file above
                    With ActiveSheet						
                      Do While Not IsEmpty(Range(a).Offset(n - 1, 63 + pe)) ' find and empty cell
                      pe = pe + 1
                      Loop
                        prev4 = 0
                        For oldy4 = 0 To pe - 1 'check link is not same as previous
                            If objSubFolder.Name = Range(a).Offset(n - 1, 63 + oldy4).Hyperlinks(1).Name Then
                            pe = pe - 1 'reset the pd cell don't skip empty cell for inserting hyperlinks
                            prev4 = 1
                            End If
                        Next
                        If prev4 = 0 Then
                        Range(a).Offset(n - 1, 63 + pe).Select
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
