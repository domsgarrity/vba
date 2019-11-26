
'For each item and customer in the list, run queries with the database to pull required information. 
Sub DoWork()
    Dim lst As Variant
    Dim lr As ListRow
    Dim PartNumber As String
    Dim Customer As String
    Dim PartMatrix() As Variant
    Dim PartContract As Variant
    Dim PartLoc() As Variant
    Dim Price() As Variant
    Dim PartBP() As Variant
    Dim subPrice() As Variant
    
    Set lst = Worksheets("Quote").ListObjects("Quotes")
    
    'Populate empty rows with customer num
    Call ApplytoColumn
    
    For Each lr In lst.ListRows
        'On Error Resume Next
            If lst.Range.Cells(lr.Index + 1, 1) <> "" And lst.Range.Cells(lr.Index + 1, 2) <> "" Then
                PartNumber = CStr(lst.Range.Cells(lr.Index + 1, 1))
                Customer = CStr(lst.Range.Cells(lr.Index + 1, 2))
                
                PartLoc = PartLocation(PartNumber)
                
                PartMatrix = PartInfo(PartNumber, Customer)
                
                Price = PriceDetermine(PartMatrix(7), PartMatrix(2), Customer, PartNumber)
                
                lst.Range.Cells(lr.Index + 1, 3) = PartLoc(2)
                lst.Range.Cells(lr.Index + 1, 4) = CustomerName(Customer)
                lst.Range.Cells(lr.Index + 1, 5) = PartMatrix(2)
                lst.Range.Cells(lr.Index + 1, 6) = PartMatrix(3)
                lst.Range.Cells(lr.Index + 1, 7) = PartMatrix(4)
                lst.Range.Cells(lr.Index + 1, 8) = PartMatrix(5)
                lst.Range.Cells(lr.Index + 1, 9) = PartMatrix(6)
                lst.Range.Cells(lr.Index + 1, 10) = PartMatrix(7)
                lst.Range.Cells(lr.Index + 1, 11) = Price(0)
                lst.Range.Cells(lr.Index + 1, 12) = Price(1)
                lst.Range.Cells(lr.Index + 1, 13) = PartLoc(0)
                lst.Range.Cells(lr.Index + 1, 14) = PartLoc(1)
                
                'At the tail end of things, it checks the effective date of the matrix value, if it's not current to the year, it will give Item Price
                If Year(PartMatrix(5)) < Year(Date) Or (Month(Date) = February And DateDiff("m", PartMatrix(5), Date) <= 3) Then
                    subPrice = PartBase(PartNumber)
                    lst.Range.Cells(lr.Index + 1, 11) = subPrice(0)
                    lst.Range.Cells(lr.Index + 1, 5) = subPrice(0)
                    lst.Range.Cells(lr.Index + 1, 8) = subPrice(1)
                    lst.Range.Cells(lr.Index + 1, 12) = "Item Price"
                Else
                    lst.Range.Cells(lr.Index + 1, 11) = Price(0)
                    
                End If
                
                lst.Range.Cells(lr.Index + 1, 16) = lst.Range.Cells(lr.Index + 1, 11).Value + 0.01
                lst.Range.Cells(lr.Index + 1, 17) = PartLoc(3)
                lst.Range.Cells(lr.Index + 1, 18) = PartLoc(3) + 5
                lst.Range.Cells(lr.Index + 1, 19) = PartLoc(4)
                
            End If
    Next
    
End Sub


Sub CleanTable()
    Dim lst As Variant

    Set lst = Worksheets("Quote").ListObjects("Quotes")
On Error GoTo Handler
    With lst
        .Range.AutoFilter
        .DataBodyRange.Offset(1).Resize(.DataBodyRange.Rows.Count - 1, .DataBodyRange.Columns.Count).Rows.Delete
        .DataBodyRange.Rows(1).SpecialCells(xlCellTypeConstants).ClearContents
    End With
Handler:

End Sub

Sub RepairTable()
    Dim lst As Variant

    Set lst = Worksheets("Quote")
    
    Call CleanTable
    
    lst.Range("L3").FormulaR1C1 = "=IFERROR(IF(ContractPrice([@Item],[@[Customer Num]]) = ""-"",  [@[UnitPrice 6]]*(([@[Percent(First Price)]]+100)/100), ContractPrice([@Item],[@[Customer Num]])),""-"")"
    lst.Range("M3").FormulaR1C1 = "=IF([@Price]=""-"", ""-"", IF(ContractPrice([@Item],[@[Customer Num]]) = ""-"",  ""Price Matrix"", ""Contract""))"
End Sub

Function FixEntry()
    Dim lst As Variant
    Dim Entry As String
    Dim First, Second, Answer As String
    Dim i As Integer
    
    Set lst = Worksheets("Quote")
    
    Entry = lst.Range("C3").Value
    First = Mid(Entry, 1, 1)
    Second = Mid(Entry, 2)
    Second = Replace(LTrim(Replace(Second, "0", " ")), " ", "0")

    
    Answer = First
    If Len(Entry) < 7 Then
        For i = 1 To (7 - Len(Entry))
            Answer = Answer & 0
        Next
        
        Answer = Answer & Second
    Else
        Answer = Entry
    End If
    'MsgBox Second
    FixEntry = Answer
    
End Function

Sub ApplytoColumn()
    Dim lst As Variant
    Dim Fixed As String
    Dim strQuote As String
    
    Set lst = Worksheets("Quote")
    eql = Chr$(61)
    

    strQuote = Chr$(34)
    Fixed = FixEntry
    
    'MsgBox Fixed
    lst.Range("C3").Formula = "=" & strQuote & Fixed & strQuote


End Sub


Function PriceDetermine(percent As Variant, basePrice As Variant, Customer As String, partNumb As String)
    Dim MatrixPrice As Variant
    Dim ContPrice As Variant
    Dim Returns(2) As Variant
    
    MatrixPrice = ((percent + 100) / 100) * basePrice
    ContPrice = ContractPrice(partNumb, Customer)
    
    If ContPrice <> 0 Then
        Returns(0) = ContPrice
        Returns(1) = "Contract"
    Else
        Returns(0) = MatrixPrice
        Returns(1) = "Matrix"
    End If
    
    PriceDetermine = Returns
    
End Function

Function PriceMethod(percent As Variant, basePrice As Variant, Customer As String, partNumb As String)
    Dim MatrixPrice As Variant
    Dim ContPrice As Variant
    
    MatrixPrice = ((percent + 100) / 100) * basePrice
    ContPrice = ContractPrice(partNumb, Customer)
    
    If ContPrice <> 0 Then
        PriceMethod = "Contract"
    Else
        PriceMethod = "Matrix"
    End If
End Function

Sub DoWorkMinnow()
    Dim lst As Variant
    Dim lr As ListRow
    Dim PartNumber As String
    Dim Customer As String
    Dim PartMatrix() As Variant
    Dim PartContract As Variant
    Dim PartLoc() As Variant
    Dim PartPB() As Variant
    Dim subPrice() As Variant
    
    Set lst = Worksheets("Quote").ListObjects("Quotes")
    
    'Populate empty rows with customer num
    Call ApplytoColumn
    
    For Each lr In lst.ListRows
        'On Error Resume Next
            If lst.Range.Cells(lr.Index + 1, 1) <> "" And lst.Range.Cells(lr.Index + 1, 2) <> "" Then
                PartNumber = CStr(lst.Range.Cells(lr.Index + 1, 1))
                Customer = CStr(lst.Range.Cells(lr.Index + 1, 2))
                
                PartMatrix = PartInfo(PartNumber, Customer)
                Price = PriceDetermine(PartMatrix(7), PartMatrix(2), Customer, PartNumber)
                PartLoc = PartLocation(PartNumber)
                
                lst.Range.Cells(lr.Index + 1, 3) = PartLoc(2)
                lst.Range.Cells(lr.Index + 1, 4) = CustomerName(Customer)
                lst.Range.Cells(lr.Index + 1, 5) = PartMatrix(2)
                lst.Range.Cells(lr.Index + 1, 6) = PartMatrix(3)
                lst.Range.Cells(lr.Index + 1, 7) = PartMatrix(4)
                lst.Range.Cells(lr.Index + 1, 8) = PartMatrix(5)
                lst.Range.Cells(lr.Index + 1, 9) = PartMatrix(6)

                lst.Range.Cells(lr.Index + 1, 11) = Price(0)
                lst.Range.Cells(lr.Index + 1, 12) = Price(1)
           
                lst.Range.Cells(lr.Index + 1, 13) = PartLoc(0)
                lst.Range.Cells(lr.Index + 1, 14) = PartLoc(1)
                Call MinnowChange(PartNumber)
                lst.Range.Cells(lr.Index + 1, 15) = GetMinnow()
                
                
                'At the tail end of things, it checks the effective date of the matrix value, if it's not current to the year, it will give Item Price
                If Year(PartMatrix(5)) < Year(Date) Or (Month(Date) = February And DateDiff("m", PartMatrix(5), Date) <= 3) Then
                    subPrice = PartBase(PartNumber)
                    lst.Range.Cells(lr.Index + 1, 11) = subPrice(0)
                    lst.Range.Cells(lr.Index + 1, 5) = subPrice(0)
                    lst.Range.Cells(lr.Index + 1, 8) = subPrice(1)
                    lst.Range.Cells(lr.Index + 1, 12) = "Item Price"
                Else
                    lst.Range.Cells(lr.Index + 1, 11) = Price(0)
                    
                End If
                
                lst.Range.Cells(lr.Index + 1, 16) = lst.Range.Cells(lr.Index + 1, 11).Value + 0.01
                lst.Range.Cells(lr.Index + 1, 17) = PartLoc(3)
                lst.Range.Cells(lr.Index + 1, 18) = PartLoc(3) + 5
                
            End If
    Next
    
End Sub
