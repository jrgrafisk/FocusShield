Attribute VB_Name = "modTransfers"
Option Explicit
' Recognise transactions that move money between the user's own accounts.
' Port of budget_core/transfers.py.

Public Const MIN_ACCOUNT_DIGITS As Long = 6

Public Function NormalizeAccountNumber(ByVal raw As String) As String
    Dim i As Long, ch As String, out As String
    out = ""
    For i = 1 To Len(raw)
        ch = Mid$(raw, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    NormalizeAccountNumber = out
End Function

Public Function ParseAccountNumbers(ByVal raw As String) As Collection
    Dim out As New Collection
    If Trim$(raw) = "" Then
        Set ParseAccountNumbers = out
        Exit Function
    End If
    Dim parts() As String, i As Long
    Dim cleaned As String
    cleaned = Replace(raw, ";", ",")
    cleaned = Replace(cleaned, "/", ",")
    parts = Split(cleaned, ",")
    For i = LBound(parts) To UBound(parts)
        Dim digits As String
        digits = NormalizeAccountNumber(parts(i))
        If Len(digits) >= MIN_ACCOUNT_DIGITS Then out.Add digits
    Next i
    Set ParseAccountNumbers = out
End Function

Public Function TextMentionsAccount(ByVal txText As String, ByVal accountNumbers As Collection) As Boolean
    If accountNumbers Is Nothing Then
        TextMentionsAccount = False
        Exit Function
    End If
    If accountNumbers.Count = 0 Then
        TextMentionsAccount = False
        Exit Function
    End If
    Dim tokens() As String
    tokens = Split(SqueezeText(txText), " ")
    Dim i As Long, digits As String, n As Variant
    For i = LBound(tokens) To UBound(tokens)
        digits = NormalizeAccountNumber(tokens(i))
        If Len(digits) >= MIN_ACCOUNT_DIGITS Then
            For Each n In accountNumbers
                If InStr(digits, CStr(n)) > 0 Then
                    TextMentionsAccount = True
                    Exit Function
                End If
            Next n
        End If
    Next i
    For i = LBound(tokens) To UBound(tokens) - 1
        If IsAllDigits(tokens(i)) And IsAllDigits(tokens(i + 1)) Then
            Dim combined As String
            combined = tokens(i) & tokens(i + 1)
            For Each n In accountNumbers
                If InStr(combined, CStr(n)) > 0 Then
                    TextMentionsAccount = True
                    Exit Function
                End If
            Next n
        End If
    Next i
    TextMentionsAccount = False
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    If s = "" Then
        IsAllDigits = False
        Exit Function
    End If
    Dim i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then
            IsAllDigits = False
            Exit Function
        End If
    Next i
    IsAllDigits = True
End Function
