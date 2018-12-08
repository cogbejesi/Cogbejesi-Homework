Attribute VB_Name = "Module1"
Sub Stocks()

Dim Ticker As String

Dim Total_Volume As Double

Total_Volume = 0

Dim Summary_Table_Row As Integer

Summary_Table_Row = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    For I = 2 To lastrow

        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
    
        Ticker = Cells(I, 1).Value
    
        Total_Volume = Total_Volume + Cells(I, 7).Value
    
        Range("J" & Summary_Table_Row).Value = Ticker
    
        Range("K" & Summary_Table_Row).Value = Total_Volume
    
        Summary_Table_Row = Summary_Table_Row + 1
    
        Total_Volume = 0
    
    Else
    
        Total_Volume = Total_Volume + Cells(I, 7).Value
        
    End If
    
Next I


End Sub
