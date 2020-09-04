Sub WallStreet_All_Worksheets()
    
    'defines xSh as a worksheet
    Dim xSh As Worksheet
    
    'turns the screen updating off to hide the columns in the worksheets so the macro runs faster
    Application.ScreenUpdating = False
    
    'loops through the worksheets available in the excel file to use the following the function
    For Each xSh In Worksheets
        
        xSh.Select
        
        'starts the function WallStreet which has the bulk of the code
        Call WallStreet
    
    Next
    
    'turns the screen updating back on once the macro ends
    Application.ScreenUpdating = True

End Sub

Sub WallStreet()

    'defining the variables

    Dim ticker As String
        
    Dim yearly_change As Double
    
    Dim percent_change As Double
    
    Dim closepriceEY As Double
    
    Dim openpriceSY As Double
    
    Dim stock_volume As Double
    
    Dim max_increase As Double
    Dim min_increase As Double
    Dim total_volume As Double
    
    'sets the range for the percent change to be percentages
    Range("K:K").NumberFormat = "0.00%"
    
    'setting a variable to count the cells which will be used for determing the open value at start of year
    Dim countcell As Integer
    
    countcell = 0
    
    'defines the row length and size
    Dim LR As Long
  
        LR = Range("A:A").SpecialCells(xlCellTypeLastCell).Row
        
    'setting the close price, open price and stock volume as 0 for calculations later
    closepriceEY = 0
    openpriceSY = 0
    stock_volume = 0
    
    Dim summary_table As Integer
        
    summary_table = 2
        
        Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        

        For i = 2 To LR
        
               
                        
          'comparing the ticker value to the next
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    
            'set the ticker name
            ticker = Cells(i, 1).Value
                       
           'set the close and open price end of year (which collects the last one before the change)
            closepriceEY = Cells(i, 6).Value
            
            openpriceSY = Cells(i - countcell, 3).Value
                   
                                     
            'calculate the percent change and sets a condition so there isn't a problem when dividing by 0
            
                If openpriceSY = 0 Then
                    
                    percent_change = 0
                
                Else
                    
                yearly_change = closepriceEY - openpriceSY
                
                percent_change = (closepriceEY - openpriceSY) / openpriceSY
                
               End If
                               
            'start adding to the stock volume
            stock_volume = stock_volume + Cells(i, 7).Value
            
            countcell = countcell + 1
            
            ' Creating the summary table
            Range("I" & summary_table).Value = ticker
            
            Range("J" & summary_table).Value = yearly_change
            
            Range("K" & summary_table).Value = percent_change
                      
            Range("L" & summary_table).Value = stock_volume
        
            
            'adding a row to the summary table
            
            summary_table = summary_table + 1
                                                
            'reset the stock volume and counter
        
            stock_volume = 0
            countcell = 0
          
                         
            ' if the cell immediately following a row is the same ticker
            Else
            
            'add to countcell
            countcell = countcell + 1
            
                                    
            ' add to stock volume
             stock_volume = stock_volume + Cells(i, 7).Value
                       
             
            End If
                           
              
        Next i
        
   'calling the function that has yearly change > 0 as green and yearly change < 0 as red
        
   Call Colour
                               
        'creating the headings for the next table
        
        Range("P1:Q1").Value = Array("Ticker", "Value", "Percent Change", "Total Stock Volume")
        Range("O2").Value = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
        Range("O3").Value = Array("Greatest % Decrease")
        Range("O4").Value = Array("Greatest Total Volume")
        
        'find the max values of total stock volume and yearly percent change and the yearly minimum percent change
       
        total_volume = WorksheetFunction.Max(Range("L:L"))
        max_increase = WorksheetFunction.Max(Range("K:K"))
        min_increase = WorksheetFunction.Min(Range("K:K"))
        
        'placing these max and min values onto the worksheet
        
        Cells(4, 17).Value = total_volume
        Cells(2, 17).Value = max_increase
        Cells(3, 17).Value = min_increase
        
        'setting the max and min yearly percent change format
        
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 17).NumberFormat = "0.00%"
        
        'matching these max and min values with the correct ticker
        
        Range("P2").Value = "= Index(I:I, match(Q2, K:K,0))"
        Range("P3").Value = "= Index(I:I, match(Q3, K:K,0))"
        Range("P4").Value = "= Index(I:I, match(Q4, L:L,0))"
 

End Sub

Sub Colour()

LR3 = Range("I:I").SpecialCells(xlCellTypeLastCell).Row

For i = 2 To LR3

        'loops through the values in the yearly change column and turns the cell green if positive or red if negative
               
               If Cells(i, 10).Value < 0 Then
            
                Cells(i, 10).Interior.ColorIndex = 3
            
               ElseIf Cells(i, 10).Value > 0 Then
               Cells(i, 10).Interior.ColorIndex = 4
               
               ElseIf Cells(i, 10).Value = 0 Then
               Cells(i, 10).Interior.ColorIndex = 0
               
                                                 
          End If
   Next i

End Sub




