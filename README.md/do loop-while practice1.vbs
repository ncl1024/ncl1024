For num = 100 To 999 step 1
    digit1 = Int(num / 100)
    digit2 = Int((num Mod 100) / 10)
    digit3 = num Mod 10
    
    sumOfCubes = (digit1 ^ 3) + (digit2 ^ 3) + (digit3 ^ 3)
    
    If sumOfCubes = num Then
        msgbox(num)
    End If
Next

