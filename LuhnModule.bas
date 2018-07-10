Attribute VB_Name = "LuhnModule"
Public Function LUHN(number As String)
    number = MULTIPLY_SECONDS(number)
    Dim sum As Integer
    sum = SUM_DIGITS(number)
    LUHN = sum Mod 10
End Function

Private Function MULTIPLY_SECONDS(number As String)
    Dim result As String
    Dim i As Integer
    For i = 1 To Len(number)
        Dim digit As Integer
        digit = (Asc(Mid(number, (Len(number) - i + 1), 1)) - Asc("0"))
        If i Mod 2 = 1 Then
            digit = digit * 2
        End If
        result = CStr(digit) + result
    Next
    MULTIPLY_SECONDS = result
End Function

Private Function SUM_DIGITS(number As String)
    Dim sum As Integer
    sum = 0
    For i = 1 To Len(number)
        sum = sum + (Asc(Mid(number, i, 1)) - Asc("0"))
    Next
    SUM_DIGITS = sum
End Function

