Attribute VB_Name = "modTextUtils"
Option Explicit
' Small text helpers shared by the parsing/detection modules.
' Port of budget_core/textutils.py.
'
' All accented characters below are built from ChrW$() Unicode code points
' rather than typed as literal characters. The VBA editor's "Import File"
' is not reliably UTF-8-safe, so a literal "ae"/"o-slash"/"a-ring" in the
' source risks turning into mojibake on import; ChrW$(230) always produces
' the correct character at runtime regardless of how the .bas file itself
' was read in.

Public Function SqueezeText(ByVal s As String) As String
    ' Collapse whitespace but keep the original characters.
    Dim i As Long, ch As String, prevSpace As Boolean, out As String
    s = Trim$(s)
    out = ""
    prevSpace = False
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            If Not prevSpace Then
                out = out & " "
                prevSpace = True
            End If
        Else
            out = out & ch
            prevSpace = False
        End If
    Next i
    SqueezeText = out
End Function

' Normalise text for keyword comparisons: lower-case, strip accents and
' fold Nordic letters, so differently-accented spellings of the same word
' all compare equal.
Public Function FoldText(ByVal s As String) As String
    Dim result As String
    result = LCase$(s)
    result = ReplaceAccented(result)
    FoldText = SqueezeText(result)
End Function

Private Function ReplaceAccented(ByVal s As String) As String
    ' The Nordic letters fold to two characters, so handle those first.
    s = Replace(s, ChrW$(230), "ae")     ' ae
    s = Replace(s, ChrW$(248), "o")      ' o-slash
    s = Replace(s, ChrW$(229), "a")      ' a-ring
    s = Replace(s, ChrW$(223), "ss")     ' sharp s

    Dim plainSrc As String, plainDst As String
    plainSrc = ChrW$(225) & ChrW$(224) & ChrW$(226) & ChrW$(227) & ChrW$(228) & _
              ChrW$(232) & ChrW$(233) & ChrW$(234) & ChrW$(235) & _
              ChrW$(236) & ChrW$(237) & ChrW$(238) & ChrW$(239) & _
              ChrW$(241) & _
              ChrW$(242) & ChrW$(243) & ChrW$(244) & ChrW$(245) & ChrW$(246) & _
              ChrW$(249) & ChrW$(250) & ChrW$(251) & ChrW$(252) & _
              ChrW$(253) & ChrW$(255) & ChrW$(231)
    plainDst = "aaaaa" & "eeee" & "iiii" & "n" & "ooooo" & "uuuu" & "yy" & "c"

    Dim i As Long, c As String, p As Long, out As String
    out = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        p = InStr(plainSrc, c)
        If p > 0 Then
            out = out & Mid$(plainDst, p, 1)
        Else
            out = out & c
        End If
    Next i
    ReplaceAccented = out
End Function

Public Function IsBlankRow(ByVal arr As Variant) As Boolean
    ' True when every cell in a 1D array/row is empty or whitespace.
    Dim i As Long
    If Not IsArray(arr) Then
        IsBlankRow = (Trim$(CStr(arr)) = "")
        Exit Function
    End If
    For i = LBound(arr) To UBound(arr)
        If Trim$(CStr(arr(i))) <> "" Then
            IsBlankRow = False
            Exit Function
        End If
    Next i
    IsBlankRow = True
End Function
