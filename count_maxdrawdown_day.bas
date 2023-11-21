Attribute VB_Name = "Module2"

Public Function MAXDD_P(profit)
    Dim ThisSum, MaxSum
    Dim i As Long
    Dim j As Long
    
    ThisSum = 0
    MaxSum = 0
    j = 0
    MaxP = 0
    
    For i = 1 To profit.Rows.Count
    
        ThisSum = ThisSum + profit(i)
    
        If ThisSum >= MaxSum Then
            
           MaxSum = ThisSum
           j = 0
        
        ElseIf ThisSum < MaxSum Then
            j = j + 1
           If j > MaxP Then
            MaxP = j
           End If
        End If
    
    Next i
    
    MAXDD_P = MaxP
End Function

