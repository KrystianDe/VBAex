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


Sub Randomizer()
    'Random number
    Dim i As Long
    
    For i = 1 To 100 Step 1
        'Random between o and 1
        Cells(i, 2) = Rnd()
        'Random between 5 and 10
        Cells(i, 3) = 5 + Rnd() * (10 - 5)
    Next i
End Sub


Sub BEPSIM()

    Dim i As Long
        
        For i = 1 To 100 Step 1
            'Put radnom number between 5 and 10 in b2 cell
            [b2] = 5 + Rnd() * (10 - 5)
            
        Next i
        
End Sub


Sub BEPSIM2()
    

    Dim nofs As Long
        nofs = InputBox("How many simulation You would like to get?")
    [d:f].Clear
    Dim i As Long
    [d1] = "No."
    [e1] = "Price"
    [f1] = "BEP"
    'drawing border
    [d1:f1].Borders.LineStyle = xlContinuous
    
        For i = 1 To nofs Step 1
            
            [b2] = 5 + Rnd() * (10 - 5)
            Cells(i + 1, 4) = i
            Cells(i + 1, 5) = [b2]
            Cells(i + 1, 6) = [b6]
            For j = 4 To 6 Step 1
                'drawing border on created cells
                Cells(i + 1, j).Borders.LineStyle = xlContinuous
            Next j
        Next i
        
End Sub

Sub EBITTemp()

    [a1] = "Data"
    [a2] = "Price"
    [a3] = "Variable cost"
    [a4] = "Fixed cost"
    [a5] = "Quantity"
    [a7] = "EBIT"
    [b2] = 10
    [b3] = 5
    [b4] = 100
    [b5] = 1000
    [a2:b5].Borders.LineStyle = xlContinuous
    [a7:b7].Borders.LineStyle = xlContinuous
    
    [b7].FormulaR1C1 = "=(r[-5]c-r[-4]c)*r[-2]c-r[-3]c"
        'Dla ---
        'Data
        'Price 10
        'Variable cost   5
        'Fixed cost  100
        'Quantity 1000
        '
        'EBIT 4900
        '----
        'R1C1 = (10 - 5) * 1000 - 100
End Sub


Sub DEBTPMT()
    'Generet template with creating new sheet
    'Sheets.Add.Name = "PMT"
    With Sheets("PMT")
    .[a1] = "Data"
    .[a2] = "Present Value"
    .[a3] = "Number of years"
    .[a4] = "Frequency of payments"
    .[a5] = "Rate"
    .[a7] = "PMT"
    .[b2] = 100000
    .[b3] = 30
    .[b4] = 12
    .[b5] = 0.08
    .[b7].FormulaR1C1 = "=PMT(r[-2]c/r[-3]c,r[-3]c*r[-4]c,-r[-5]c)"
    .[a2:b5].Borders.LineStyle = xlContinuous
    .[a7:b7].Borders.LineStyle = xlContinuous
    End With
   
    '[b7].FormulaR1C1 = "=PMT(r[-2]c/r[-3]c,r[-3]c*r[-4]c,-r[-5]c)"
    'Dla ---
    'Data
    'Present Value   100000
    'Number of years 30
    'Frequency of payments   12
    'Rate 0.08
    '
    'PMT $733.76
    '----
    'PMT TO FORMULA WBUDOWANA W VBA
    'R1C1 = PMT( 0.08/12, 12*30, -100000)
    

        
End Sub


Sub DEBTPMTSIM()

 
    Dim nofs As Long
        nofs = InputBox("How many simulation You would like to get?")
    
    [d:f].Clear
    With Sheets("PMT")
    .[d1] = "No."
    .[e1] = "Rate"
    .[f1] = "PMT"
    .[d1:f1].Borders.LineStyle = xlContinuous
    
    End With
    
        For i = 1 To nofs Step 1
            With Sheets("PMT")
            .[b5] = 0.0372 + Rnd() * (0.1 - 0.0372)
            .Cells(i + 1, 4) = i
            .Cells(i + 1, 5) = [b5]
            .Cells(i + 1, 6) = [b7]
            For j = 4 To 6 Step 1
                .Cells(i + 1, j).Borders.LineStyle = xlContinuous
            Next j
            End With
        Next i
        
End Sub



Sub WarmUpTask1()
    Dim oW As Worksheet
    Dim i, nofs As Long
        nofs = InputBox("How many samples you would like to get?")
   
    Set oW = Sheets.Add
        oW.Name = "BEP" & Sheets.Count
    
    With oW
        .[a1] = "Date"
        .[a2] = "Price"
        .[a3] = "Variable Cost"
        .[a4] = "Fixed cost"
        .[a6] = "BEP"
        .[b2] = 100
        .[b3] = 50
        .[b4] = 1000
        .[b6].FormulaR1C1 = "=r[-2]c/(r[-4]c-r[-3]c)"
        .[a1:b1].Merge
        .[a1:b4].Borders.LineStyle = 1
        .[a6:b6].Borders.LineStyle = 1
        .[d1] = "No."
        .[e1] = "Price"
        .[f1] = "Bep"
        
         For i = 1 To nofs Step 1
            .[b2] = 50 + Rnd() * (200 - 50)
            .Cells(i + 1, 4) = i
            .Cells(i + 1, 5) = [b2]
            .Cells(i + 1, 6) = [b6]
            
            For j = 4 To 6 Step 1
                .Cells(i + 1, j).Borders.LineStyle = 1
            Next j
        Next i
    End With
    
Sub WarmUpTask2()
    Dim i, nofs As Long
    nofs = InputBox("How many samples you would like to get?")
    
    For i = 1 To nofs Step 1
        [b2] = 50 + Rnd() * (200 - 50)
        Cells(i + 1, 4) = i
        Cells(i + 1, 5) = [b2]
        Cells(i + 1, 6) = [b6]
    Next i
    

   
    
End Sub
