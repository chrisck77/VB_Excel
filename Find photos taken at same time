Public com_array(28000, 5) As String, count As Integer, ct As Integer, StartTimer As Double 


Sub date_taken_grab2()
'routine to find all copies of files (based on date taken --> not file created,modified etc)
'looks in folders and drills-down into all sub-folders to find all JPG files for comparison to each other
'then list name and path of each image that has a copy and also list the name and path of the copied image file(can be differt size, cropped etc)
'put this list into excel for later automatically adding photos side by side for visual comparison (to check photos are same --> may have been cropped/adjusted but still same image)

Dim Rng As Range
Dim objExif As ExifReader

Dim txtExifInfo As String
Dim strFolderPath As String
Dim objFileSystem As Object
Dim objFolder As Object

Dim FileSystem As Object
Dim HostFolder As String
Dim Locations(1 To 15, 0) As String

Erase com_array

StartTimer = Timer
ct = 1

count = -1

Locations(1, 0) = "C:\Chris files\photos_1\"
Locations(2, 0) = "other location on local or network2
'added a further 13 locations (15 in total) for comparing files.

For Locs = 1 To 15
HostFolder = Locations(Locs, 0)

'following 3 lines find all the JPG files in folders (inc sub-folders_ of the listed locations and puts in to an array (details of the files)
'recrusive routine to drill down into all folders and sub-folders

Set FileSystem = CreateObject("Scripting.FileSystemObject") 
DoFolder2 FileSystem.GetFolder(HostFolder) 
Next

Dim com_array2(15000, 5) As String
Dim count2 As Integer
count2 = -1

'next compare each JPG file to every other JPG file into main array, if there are copies than put them into array'
'new array records file name, path of original file and copied file
For compare = 0 To count
    If com_array(compare, 3) = "copy" Then
    count2 = count2 + 1
        For i = 0 To 5
        com_array2(count2, i) = com_array(compare, i)
        Next
    End If
Next

'now copy array of copied files into excel spreadsheet
If count2 > -1 Then
    Set Rng = Sheets("Main").Range("B4:G" & 4 + count2)
    'Rng = com_array
    Rng = com_array2
End If

EndTime = Timer
Debug.Print EndTime - StartTimer

End Sub

'_____________________________________________________________________________________
'this is the sub-routine for finding all JPG files with a date taken property then puts them into an array
Sub DoFolder2(Folder)
    Dim SubFolder
    For Each SubFolder In Folder.SubFolders
        DoFolder2 SubFolder
    Next
    
    Dim objFile
    For Each objFile In Folder.Files
        If objFile Like "*.jpg*" Or objFile Like "*.JPG*" Then
        ' Operate on each file
        
        picFile = objFile.Path
        picsize = objFile.Size
                      
        On Error Resume Next
        
        Set objExif = New ExifReader 'seperate class EXIF used for grabbing date taken from an image file
        objExif.Load picFile

        txtExifInfo = objExif.Tag(DateTimeOriginal)
        
            If txtExifInfo <> "" And picsize > 5000 Then 'ignore files that don't have a date taken
            count = count + 1
            
            If count > ct * 1000 Then
            ct = ct + 1
            Debug.Print Timer - StartTimer 'timer just for checking speed of this routine VS adding data direct to excel (rather than array)
            End If
            
            com_array(count, 0) = objFile.Name 
            com_array(count, 1) = txtExifInfo
            com_array(count, 2) = objFile.Path
            
                For compare = 0 To count - 1      'code for checking if there are muliple copies of a picture        
                    If com_array(count, 1) = com_array(compare, 1) Then
                    com_array(count, 3) = "copy"
                    com_array(count, 4) = com_array(compare, 0)
                    com_array(count, 5) = com_array(compare, 2)
                    End If                
                Next
            End If
        
        End If
    Next

End Sub
