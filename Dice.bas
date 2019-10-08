Public Function Dice(NumDie As Integer)
	Dim result As Integer
	
	Debug.Print "Running"
	
	Randomize
	
	If (NumDie > 0) Then
		result = 0
		
		For i = 1 To NumDie
			result = result + Int((6 - 1 + 1) * Rnd + 1)
			Debug.Print result
		Next
	Else
		Debug.Print "Enter positive number"
	End If
	
	
	NumDie = result
	
	
End Function