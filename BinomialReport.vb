Public weekCnt, dayCnt As Integer


'Daily run to update report and record current items in area
Sub BinomialReport()
    
    'Make it go faster
    LudicrousMode (True)
    
    'Refresh Tables
    Workbooks("BRPT Report.xlsm").RefreshAll
    DoEvents
    
    'Copy today lists to long list with snapshot date (Week 1, Date)
    Call appendToMaster

    'Refresh tables and allow pivot charts to update
    Workbooks("BRPT Report.xlsm").RefreshAll
    DoEvents
    
    'Set report time for the next day
    setReport
    'Reactivate Workbook
    LudicrousMode (False)
    
    
    
End Sub



Sub appendToMaster()
    Dim lstThirty, lstSixty, oLstThirty, oLstSixty As Variant
    Dim lr As ListRow
    Dim LastRow As Variant, x As Long
    
    'Set the day and week counts
    weekCnt = Application.WorksheetFunction.RoundDown((DateDiff("d", DateValue("11/24/2019"), Date) / 7), 0) + 1
    dayCnt = (DateDiff("d", DateValue("11/25/2019"), Date) Mod 7) + 1
    
    Set lstThirty = Worksheets("Current By Group").ListObjects("ThirtySixty")
    Set lstSixty = Worksheets("Current By Group").ListObjects("SixtyNinty")
    Set oLstThirty = Worksheets("Collection").ListObjects("colThirty")
    Set oLstSixty = Worksheets("Collection").ListObjects("colSixty")
    
    'Append the current list of parts in area to the master lists
    Set LastRow = oLstThirty.ListRows.Add
    lstThirty.DataBodyRange.Copy
    LastRow.Range.PasteSpecial xlPasteValues
    
    Set LastRow = oLstSixty.ListRows.Add
    lstSixty.DataBodyRange.Copy
    LastRow.Range.PasteSpecial xlPasteValues
    
    
    DoEvents
    
    'Add the week and day Identifier for 30 to 60
    For Each lr In oLstThirty.ListRows
        'On Error Resume Next
        If oLstThirty.Range.Cells(lr.Index + 1, 10) = "" And oLstThirty.Range.Cells(lr.Index + 1, 2) <> "" Then
            oLstThirty.Range.Cells(lr.Index + 1, 10) = weekCnt '"wk" & weekCnt & " day" & dayCnt
            oLstThirty.Range.Cells(lr.Index + 1, 11) = dayCnt
            oLstThirty.Range.Cells(lr.Index + 1, 11) = Date
        End If
    Next
    
    'Add the week and day Identifier for 60 to 90
    For Each lr In oLstSixty.ListRows
    'On Error Resume Next
    If oLstSixty.Range.Cells(lr.Index + 1, 10) = "" And oLstSixty.Range.Cells(lr.Index + 1, 2) <> "" Then
        oLstSixty.Range.Cells(lr.Index + 1, 10) = weekCnt '"wk" & weekCnt & " d" & dayCnt
        oLstSixty.Range.Cells(lr.Index + 1, 11) = dayCnt
        oLstSixty.Range.Cells(lr.Index + 1, 11) = Date
    End If
    Next
    
End Sub

Sub setReport()
    'This sets the schedule to the designated time
    '**Note if changing the time value, set Schedule to False, Run setReport then run again with SetReport = True
    'This will elimintate the future scheduled procedure
    Application.OnTime TimeValue("05:30:15"), Procedure:="BinomialReport", Schedule:=True
End Sub
