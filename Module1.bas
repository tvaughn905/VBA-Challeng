Attribute VB_Name = "Module1"
Sub Module2ChallengeHW()
Attribute Module2ChallengeHW.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Module2ChallengeHW Macro
'

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim TCount As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim PerChange As Double
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatVol As Double
        
   
        WorksheetName = ws.Name
        
       
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        TCount = 2
        
        j = 2
        
        
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            For i = 2 To LastRowA
            
               
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
              
                ws.Cells(TCount, 9).Value = ws.Cells(i, 1).Value
                
                ws.Cells(TCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                   
                    If ws.Cells(TCount, 10).Value < 0 Then
                
                   
                    ws.Cells(TCount, 10).Interior.ColorIndex = 3
                
                    Else
               
                    ws.Cells(TCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                   
                    If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                   
                    ws.Cells(TCount, 11).Value = Format(PerChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
               
                ws.Cells(TCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
               
                TCount = TCount + 1
                
                
                j = i + 1
                
                End If
            
            Next i
            
       
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
       
        
        GreatVol = ws.Range("L2").Value
        GreatIncr = ws.Range("K2").Value
        GreatDecr = ws.Range("K2").Value
        
            For i = 2 To LastRowI
             
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Range("P4").Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
                If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                ws.Range("P2").Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncr = GreatIncr
                
                End If
                
                If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecr = GreatDecr
                
                End If
                
            ws.Range("Q2").Value = Format(GreatIncr, "Percent")
            ws.Range("Q3").Value = Format(GreatDecr, "Percent")
            ws.Range("Q4").Value = Format(GreatVol, "Scientific")
            
            Next i
       
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub


'



