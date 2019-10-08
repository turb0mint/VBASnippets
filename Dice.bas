Public Function Dice(NumDie As Integer) As Integer
On Error GoTo error:

    Dim result As Integer
    
    
    Randomize
    
    If (NumDie > 0) Then
        result = 0
        
        For i = 1 To NumDie
            result = result + Int((6 - 1 + 1) * Rnd + 1)
            Debug.Print result
        Next
    Else
        Debug.Print "Enter positive number"
        Err.Raise 10, , "Enter positive Number of Dice"
        
    End If
    
    
    Dice = result
    
    Exit Function
    
error:
MsgBox Err.Description
Exit Function


    
    
End Function