Attribute VB_Name = "Module1"
Sub Stock_market()
  Dim Ticker As String
  Dim Ticker_next As String
  Dim CurrentRow As Double
  Dim CurrentCol As Double
  Dim nextCurrentRow As Double
  Dim TotalVolume As Double
  Dim WorksheetNames As String
  Dim OpenPrice As Double
  Dim ClosePrice As Double
  Dim YearlyChange As Double
  Dim PercentChange As Double
  Dim MaxTotalVolune As Double
  Dim Max_Ticker_name As String
  Dim G_Percent_inc As Double
  Dim G_Percent_dec As Double
  Dim G_increase_Ticker As String
  Dim G_descrease_Ticker As String
  
  ' Initialize variable.
  CurrentRow = 2
  nextCurrentRow = CurrentRow + 1
  CurrentCol = 1
  TotalVolume = 0
  YearlyChange = 0
  PercentChange = 0
  Volume_Position = 7
  Open_price_pos = 3
  close_price_pos = 6
  output_row = 2
  output_col = 9
  
 

  
  'totalRows = Cells(1, 1).End(xlDown).Row
  'MsgBox (totalRows)
  
  For Each Active_WS In Worksheets
   LastRow = Active_WS.Cells(Rows.Count, 1).End(xlUp).Row
   sheetName = Active_WS.Name
   MsgBox (LastRow)
   MsgBox (sheetName)
   ' Set up the header
     'initialize variable
   CurrentRow = 2
   Volume_Position = 7
   nextCurrentRow = CurrentRow + 1
   CurrentCol = 1
   Ticker = ""
   Ticker_next = ""
   output_row = 2
   output_col = 9
   MaxTotalVolune = 0
   MaxTicker_name = ""
   G_Percent_inc = 0
   G_Percent_dec = -1
   
 
   
   Active_WS.Cells(CurrentRow - 1, Volume_Position + 2) = "Ticker"
   Active_WS.Cells(CurrentRow - 1, Volume_Position + 3) = "YearlyChange"
   Active_WS.Cells(CurrentRow - 1, Volume_Position + 4) = "PercentChange"
   Active_WS.Cells(CurrentRow - 1, Volume_Position + 5) = "Total Stock Volume"
   Active_WS.Cells(CurrentRow - 1, Volume_Position + 10) = "Ticker"
   Active_WS.Cells(CurrentRow - 1, Volume_Position + 11) = "Value"
    
   Active_WS.Cells(CurrentRow, Volume_Position + 8) = "Greatest % Increase"
   Active_WS.Cells(CurrentRow + 1, Volume_Position + 8) = "Greatest % Decrease"
   Active_WS.Cells(CurrentRow + 2, Volume_Position + 8) = "Greatest Total Volume"
  
   

 Do While CurrentRow <= LastRow
 
  'initialize variable
 
  OpenPrice = 0
  ClosePrice = 0
 
    OpenPrice = Active_WS.Cells(CurrentRow, Open_price_pos)
  'MsgBox (OpenPrice)
  'initialize variables for before going to next ticker
  
     Ticker = Active_WS.Cells(CurrentRow, CurrentCol).Value
     Ticker_next = Active_WS.Cells(nextCurrentRow, CurrentCol).Value
     TotalVolume = 0
     YearlyChange = 0
     PercentChange = 0
   
    ' Add the volume for same Ticker
    Do While (Ticker = Ticker_next)
  
        ' Calculate the Total Volume
        TotalVolume = TotalVolume + Active_WS.Cells(CurrentRow, Volume_Position).Value
        
            ' Get the next ticker values
            CurrentRow = CurrentRow + 1
            nextCurrentRow = CurrentRow + 1
            Ticker = Active_WS.Cells(CurrentRow, CurrentCol).Value
            Ticker_next = Active_WS.Cells(nextCurrentRow, CurrentCol).Value
             
         Loop ' End of while loop
        
        If Ticker <> Ticker_next Then
             ' Add the last row
              TotalVolume = TotalVolume + Active_WS.Cells(CurrentRow, Volume_Position).Value
               
           ' closing price
            ClosePrice = Active_WS.Cells(CurrentRow, 6)
            
            YearlyChange = ClosePrice - OpenPrice
           
                
            If OpenPrice <> 0 Then
                PercentChange = FormatNumber((YearlyChange / OpenPrice) * 100, 2)
            End If
            
            If MaxTotalVolune < TotalVolume Then
                MaxTotalVolune = TotalVolume
                MaxTicker_name = Ticker
             End If
             
             If G_Percent_inc < PercentChange Then
                G_Percent_inc = PercentChange
                G_increase_Ticker = Ticker
                
              ElseIf G_Percent_dec > PercentChange Then
                G_Percent_dec = PercentChange
                G_descrease_Ticker = Ticker
                
              End If
           
            ' output to excel
            Active_WS.Cells(output_row, output_col) = Ticker
            Active_WS.Cells(output_row, output_col + 1) = YearlyChange
            Active_WS.Cells(output_row, output_col + 2) = (CStr(PercentChange) & "%")
            Active_WS.Cells(output_row, output_col + 3) = TotalVolume
            
            If YearlyChange > 0 Then
              ' Fill Green color
              Active_WS.Cells(output_row, output_col + 1).Interior.ColorIndex = 4
              
             Else
                 Active_WS.Cells(output_row, output_col + 1).Interior.ColorIndex = 3
             End If
            
            output_row = output_row + 1
            
            
        End If

        
            CurrentRow = nextCurrentRow
            nextCurrentRow = CurrentRow + 1
        
          
   Loop
     ' Initialize variable.
     MsgBox (" Max Total Volume" & MaxTotalVolune)
     MsgBox (" Max Ticker Name" & MaxTicker_name)
      MsgBox (" Max G_Percent_inc " & G_Percent_inc)
      MsgBox (" Max G_Percent_inc " & G_Percent_dec)
      MsgBox (" dec Ticker Name" & G_descrease_Ticker)
      
  CurrentRow = 2
  nextCurrentRow = CurrentRow + 1
  
     Active_WS.Cells(CurrentRow, Volume_Position + 11) = (CStr(G_Percent_inc) & "%")
     Active_WS.Cells(CurrentRow, Volume_Position + 10) = G_increase_Ticker
     Active_WS.Cells(CurrentRow + 1, Volume_Position + 11) = (CStr(G_Percent_dec) & "%")
     Active_WS.Cells(CurrentRow + 1, Volume_Position + 10) = G_descrease_Ticker
     Active_WS.Cells(CurrentRow + 2, Volume_Position + 11) = MaxTotalVolune
     Active_WS.Cells(CurrentRow + 2, Volume_Position + 10) = MaxTicker_name
  
  CurrentCol = 1
  TotalVolume = 0
  YearlyChange = 0
  PercentChange = 0
  MaxTotalVolune = 0
  MaxTicker_name = ""
   G_increase_Ticker = ""
   G_Percent_inc = 0
   G_Percent_dec = 0
   G_descrease_Ticker = ""

   
   Next Active_WS
       
End Sub

