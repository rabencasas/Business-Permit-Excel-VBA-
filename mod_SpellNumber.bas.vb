Attribute VB_Name = "mod_SpellNumber"
Option Explicit
'Main Function
Function SpellNumber(ByVal MyNumber)
    Dim PESOS, Cents, Temp
    Dim DecimalPlace, Count
    ReDim Place(9) As String
    Place(2) = "THOUSAND "
    Place(3) = "MILLION "
    Place(4) = "BILLION "
    Place(5) = "TRILLION "
    ' String representation of amount.
    MyNumber = Trim(Str(MyNumber))
    ' Position of decimal place 0 if none.
    DecimalPlace = InStr(MyNumber, ".")
    ' Convert cents and set MyNumber to dollar amount.
    If DecimalPlace > 0 Then
        Cents = Left(Mid(MyNumber, DecimalPlace + 1) & _
                  "00", 2)
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    Count = 1
    Do While MyNumber <> ""
        Temp = GetHundreds(Right(MyNumber, 3))
        If Temp <> "" Then PESOS = Temp & Place(Count) & PESOS
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop
    Select Case PESOS
        Case ""
            PESOS = ""
        Case "ONE "
            PESOS = "ONE "
         Case Else
            PESOS = PESOS
    End Select
    Select Case Cents
        Case ""
            Cents = "PESOS ONLY"
        Case "One"
            Cents = "AND 1/100 PESOS ONLY"
        Case Else
            Cents = "PESOS AND " & Cents & "/100 ONLY"
    End Select
    SpellNumber = PESOS & Cents
End Function
      
' Converts a number from 100-999 into text
Function GetHundreds(ByVal MyNumber)
    Dim result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    ' Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        result = GetDigit(Mid(MyNumber, 1, 1)) & "HUNDRED "
    End If
    ' Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        result = result & GetTens(Mid(MyNumber, 2))
    Else
        result = result & GetDigit(Mid(MyNumber, 3))
    End If
    GetHundreds = result
End Function
      
' Converts a number from 10 to 99 into text.
Function GetTens(TensText)
    Dim result As String
    result = ""           ' Null out the temporary function value.
    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...
        Select Case Val(TensText)
            Case 10: result = "TEN "
            Case 11: result = "ELEVEN "
            Case 12: result = "TWELVE "
            Case 13: result = "THIRTEEN "
            Case 14: result = "FOURTEEN "
            Case 15: result = "FIFTEEN "
            Case 16: result = "SIXTEEN "
            Case 17: result = "SEVENTEEN "
            Case 18: result = "EIGHTEEN "
            Case 19: result = "NINETEEN "
            Case Else
        End Select
    Else                                 ' If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: result = "TWENTY "
            Case 3: result = "THIRTY "
            Case 4: result = "FORTY "
            Case 5: result = "FIFTY "
            Case 6: result = "SIXTY "
            Case 7: result = "SEVENTY "
            Case 8: result = "EIGHTY "
            Case 9: result = "NINETY "
            Case Else
        End Select
        result = result & GetDigit _
            (Right(TensText, 1))  ' Retrieve ones place.
    End If
    GetTens = result
End Function
     
' Converts a number from 1 to 9 into text.
Function GetDigit(Digit)
    Select Case Val(Digit)
        Case 1: GetDigit = "ONE "
        Case 2: GetDigit = "TWO "
        Case 3: GetDigit = "THREE "
        Case 4: GetDigit = "FOUR "
        Case 5: GetDigit = "FIVE "
        Case 6: GetDigit = "SIX "
        Case 7: GetDigit = "SEVEN "
        Case 8: GetDigit = "EIGHT "
        Case 9: GetDigit = "NINE "
        Case Else: GetDigit = ""
    End Select
End Function



