Private Sub Worksheet_Change(ByVal Target As Range)
'selects a range of cells, checks the value and then changes cell color according to value using indices

Dim icolor As Integer
      
        
 Dim Rng1 As Range			'multiple range selections using Union command
 Set Rng1 = Union(Range( _
        "S643:U643,AP643:AR643,S664:U664,AP664:AR664,S684:U684,AP684:AR684,S704:U704,AP704:AR704,S732:U732,AP732:AR732,S755:U755,AP755:AR755,S773:U773,AP773:AR773,S791:U791,AP791:AR791,S815:U815,AP815:AR815,S834:U834,AP834:AR834,S853:U853,AP853:AR853,S874:U874" _
        ), Range( _
        "S988:U988,AP988:AR988,S1009:U1009,AP1009:AR1009,S1030:U1030,AP1030:AR1030,S1051:U1051,AP1051:AR1051,S309:U309,AP309:AR309,S330:U330,AP330:AR330,T350:V350,AP350:AR350,S370:U370,AP370:AR370,S396:U396,AP396:AR396,S416:U416,AP416:AR416,S436:U436,AP436:AR436" _
        ), Range( _
        "S562:U562,AP562:AR562,S581:U581,AP581:AR581,S600:U600,AP600:AR600,S619:U619,AP619:AR619" _
        ))
        
       
 For Each Cell In Rng1
     Select Case Cell.Value
 
            Case 0 To 2
                Cell.Interior.ColorIndex = 50
            Case 3 To 3
                Cell.Interior.ColorIndex = 43
            Case 4 To 4
                Cell.Interior.ColorIndex = 45
            Case 5 To 5
                Cell.Interior.ColorIndex = 3
            Case "NA"
                Cell.Interior.ColorIndex = 34
            Case Else
                Cell.Interior.ColorIndex = 34
        End Select

 Next

End Sub
