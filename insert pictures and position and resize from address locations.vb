Sub count_copies()
'just count number of copies in a list in excel

a = Range("A1")

ReDim my_arr(1 To a, 1 To 3)

For i = 1 To a
    For j = 1 To 3
    my_arr(i, j) = Range("A3").Offset(i, j)
    Next
Next

For po = 1 To a
    count3 = 0
    For ps = 1 To a
        If my_arr(po, 2) = my_arr(ps, 2) Then
        count3 = count3 + 1
        End If
    Next
    Range("A" & 3 + po) = count3
Next

End Sub
'___________________________________________________________________________
'insert picture routine
'a column list of file paths is used to insert those pictures into another column
' routine is repeated twice in order to compare photos side by side

Sub INSERTPIC()
Application.ScreenUpdating = False

Dim i As Integer
Dim fname As String

a = Range("A1")

For i = 1 To a
Range("j" & 3 + i).Select
h = Selection.Height    'height of cell selected for where to place picture
w = Selection.Width     'width of cell

fname = Range("e" & 3 + i)
Set pic = ActiveSheet.Pictures.Insert(fname)
pic.Select
    
Selection.ShapeRange.LockAspectRatio = msoTrue   'following cells re-size pictures to fit the cell, maintaining aspect ratio
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
  
Next
Application.ScreenUpdating = True
End Sub
