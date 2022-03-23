Function ConvertChar(n)
    'Input: integer 0 <= n <= 701
    'Output: string from "A" (0) to "ZZ" (701)
    maximum = 26
    number = n\maximum
    rest = n mod maximum
    If number < 1 Then
        ConvertChar = chr(n +65)
    Else
        ConvertChar = chr(number + 64) & chr(rest + 65)
    End If
End Function
