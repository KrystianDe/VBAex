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
