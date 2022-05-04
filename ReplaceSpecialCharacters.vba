' Extracted from https://answers.microsoft.com/en-us/msoffice/forum/all/removing-special-characters-with-regular/d62d50b7-8586-4f08-ac7d-c5212929074a

Function ReplaceSpecialCharacters(ByVal strVal As String) As String
    Dim X As Long
    For X = 1 To Len(strVal)
        If Mid(strVal, X, 1) Like "[!0-9A-Za-z]" Then Mid(strVal, X) = Chr(1)
    Next
    ReplaceSpecialCharacters = LCase(Replace(strVal, Chr(1), ""))
End Function
