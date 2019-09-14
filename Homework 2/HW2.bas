Attribute VB_Name = "EN"
Sub StockTest()
 ' variables
   Dim Ticker As String
   Dim TotalVolume As Double
   Dim NextTick As String
   Dim i As Double
   Dim UniqueTick As Double
   Dim LastRow As Double
   Dim ws As Worksheets
   Dim op As Single
   Dim cl As Single
   Dim rng As Range
   Dim rng2 As Range
   Dim rng3 As Range
   Dim rng4 As Range
   Dim LastUni As Double
   
   
   LastRow = Cells(Rows.Count, 1).End(xlUp).Row
   TotalVolume = 0
   UniqueTick = 1
   ' for loops
           For i = 1 To LastRow
           Ticker = Cells(i, 1).Value
           NextTick = Cells(i + 1, 1).Value
           If Ticker = NextTick Then
               TotalVolume = TotalVolume + Cells(i + 1, 7).Value
               
               ElseIf Ticker <> NextTick Then
               Cells(UniqueTick, 12).Value = TotalVolume
               Cells(UniqueTick, 9).Value = Ticker
           UniqueTick = UniqueTick + 1
               TotalVolume = Cells(i + 1, 7).Value
               If i = 1 Then
               op = Cells(i + 1, 3).Value
               End If
               
                If i <> 1 Then
                cl = Cells(i, 6).Value
                Cells(UniqueTick - 1, 10).Value = cl - op
                    If op <> 0 Then
                    Cells(UniqueTick - 1, 11).Value = cl / op - 1
                    End If
                    op = Cells(i + 1, 3).Value
                End If
               
               End If
       Next i
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 9).Value = "Ticker Symbol"
    
    LastUni = Cells(Rows.Count, 9).End(xlUp).Row
    
    Set rng = Range(Cells(2, 10), Cells(LastUni, 10))
    rng.FormatConditions.Delete
    ' add greater than condition
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Color = rgbLimeGreen
    End With


    ' add less than condition
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = rgbRed
    End With

    rng.NumberFormat = "0.00"
    
    Set rng2 = Range(Cells(2, 11), Cells(LastUni, 11))
    rng2.NumberFormat = "0.00%"
    
    Set rng3 = Range(Cells(2, 9), Cells(LastUni, 9))
    
    Set rng4 = Range(Cells(2, 12), Cells(LastUni, 12))
    
    Dim great As Double
    Dim least As Double
    Dim gtick As String
    Dim ltick As String
    Dim highvol As Double
    Dim hvname As String
    Dim rngform As Range
    Dim rngform2 As Range
    
    
    great = Application.WorksheetFunction.Max(rng2)
    least = Application.WorksheetFunction.Min(rng2)
    gtick = Application.WorksheetFunction.Index(rng3, Application.WorksheetFunction.Match(great, rng2, 0))
    ltick = Application.WorksheetFunction.Index(rng3, Application.WorksheetFunction.Match(least, rng2, 0))
    highvol = Application.WorksheetFunction.Max(rng4)
    hvname = Application.WorksheetFunction.Index(rng3, Application.WorksheetFunction.Match(highvol, rng4, 0))
        
    Cells(2, 17).Value = great
    Cells(2, 16).Value = gtick
    Cells(3, 17).Value = least
    Cells(3, 16).Value = ltick
    Cells(4, 17).Value = highvol
    Cells(4, 16).Value = hvname
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    Set rngform = Range(Cells(2, 17), Cells(3, 17))
    rngform.NumberFormat = "0.00%"
    
    Set rngform2 = Range("Q4")
    rngform2.NumberFormat = "0,000"
    
    ActiveSheet.Columns("A:Q").AutoFit
    
End Sub
Sub AcrossSheets()
Attribute AcrossSheets.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'
Dim ws As Worksheet
   For Each ws In Worksheets
   ws.Activate
   Call StockTest
   Next ws



'
End Sub
