Sub DefiningRanges()

    'code which works in active worksheets
    ' Wszystko to samo, tylko pozycja inna
    
        Range("A1").Value = 100
        Range("A2").Value = 200
        [a3] = 300
        Cells(4, 1) = 400
        ActiveSheet.[a5] = 500
    
    'Code which works in choosen worksheet
    'Tutaj trzeba stworzyć 5 lub 6 sheetow i nazwac odpowiednio
    
    
        Worksheets(1).Cells(1, 2) = 1000
        Worksheets("MySh").[b2] = 2000
        Sheets(1).Range("b3") = 3000
        Sheets("Arkusz2").[b4] = 4000
        Sheets(Sheets.Count).[b5] = 5000
        Worksheets(Worksheets.Count).[b6] = 6000
        
    'Code which workds in choosen workbook and worksheet
    
        ThisWorkbook.Sheets(1).[c1] = 10000
        Worksheets("VBAStudie.xlsm").Sheets(2).[c2] = 20000
    
End Sub


Sub InsertWorksheet()
    Worksheets.Add
    Worksheets.Add after:=Worksheets(2)
    Worksheets.Add(after:=Worksheets(4)).Name = "MySh"
End Sub



Sub InsertComment()
    [a1].clearComment
    [a1].AddComment = "Sth"
End Sub


Sub Task1()
    'Rozciaga obszar czyli zaznacza b2, b3, b4, c2, c3, c4, d2, d3, d4
    [b2:d4] = 1
End Sub


Sub LoopStatment()

    Dim i As Long 'Long-2^32
        
    For i = 1 To 100 Step 1
        Cells(i, 1) = i
    Next i

End Sub


Sub LoopTask2()
    'Strzalka do górnego lewego rogu
    'lewa kolumna, pierwszy wiersz i przekatna
    Dim i As Long ''Long-2^32
    
        For i = 1 To 100 Step 1
            Cells(i, 1) = i
            Cells(i, i) = i
            Cells(1, i) = i
        Next i
        
End Sub



Sub Task3()

    'Piotr option 1
    'Od 1 id co 0.5 wartosci az do 100.5
    Dim i As Long ''Long-2^32
    For i = 1 To 200 Step 1
        Cells(i, 1) = i / 2 + 1 / 2
    Next i
    
    'Option math 2 'wypisanie od 1 do 100 ale z przerwa pomiedzy liczbami
   
    For i = 1 To 100 Step 1
        Cells((i * 2) - 1, 2) = i
    Next i
        
    'Option 3 VBA Logic'wypisanie od 1 do 100 ale z przerwa pomiedzy liczbami
    Dim j As Long ''Long-2^32
    For i = 1 To 200 Step 2
        j = j + 1
        Cells(i, 4) = j
    Next i

  
End Sub


Sub Quiz()

    Dim i, j As Long ''Long-2^32
    For i = 1 To 20 Step 1
        Cells((i * 4) - 3, 1) = i
        Cells((i * 4) - 3, 2) = "Question"
        Cells((i * 4) - 2, 2) = "A"
        Cells((i * 4) - 2, 3) = "B"
        Cells((i * 4) - 2, 4) = "C"
        For j = 2 To 4 Step 1
            Cells((i * 4) - 1, j) = "Answer"
        Next j
        
    Next i
End Sub



Sub Task5()
    'Stworzenie obszaru 3x3 z wartosciami od 1 do 9
    Dim column, row, i As Long
    For row = 2 To 4 Step 1
        For column = 2 To 4 Step 1
            i = i + 1
            Cells(row, column) = i
        Next column
    Next row
End Sub


Sub Task6()
    'Stworzenie obszaru 4x3 z wartosciami od 1 do 12
    Dim column, row, index As Long
    For column = 1 To 3 Step 1
        For row = 1 To 4 Step 1
            index = index + 1
            Cells(row + 1, column + 1) = index
        Next row
    Next column
End Sub


Sub TabliczkaMnozenia()
    Dim i, j, w As Long
    For i = 1 To 10 Step 1
        For j = 1 To 10 Step 1
            
            Cells(i + (10 * (j - 1)), j) = j
            Cells(i + (10 * (j - 1)), j + 1) = i
            Cells(i + (10 * (j - 1)), j + 2) = i * j
            
        Next j
       
    Next i
End Sub
