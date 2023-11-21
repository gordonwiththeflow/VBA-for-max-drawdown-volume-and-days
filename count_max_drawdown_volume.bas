Attribute VB_Name = "Module1"

Public Function MAXDD(profit)
  Dim ThisSum, MaxSum
    Dim i As Long
    
    ThisSum = 0
    MaxSum = 0
    
    For i = 1 To profit.Rows.Count
    
        ThisSum = ThisSum + profit(i)
    
        If ThisSum <= MaxSum Then
            
           MaxSum = ThisSum
        
        ElseIf ThisSum >= 0 Then
        
           ThisSum = 0
        
        End If
    
    Next i
    
    MAXDD = MaxSum
End Function

