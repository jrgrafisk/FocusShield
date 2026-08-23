VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColumnStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Simple per-column statistics used by the column detectors.
' Port of budget_core.columns.ColumnStats.

Public Index As Long
Public header As String
Public Values As Collection
Public Filled As Collection
Public FillRatio As Double
Public DateRatio As Double
Public NumericRatio As Double
Public DecimalRatio As Double
Public AvgLength As Double
Public DistinctRatio As Double
Public LongIntRatio As Double

Public Sub Init(ByVal idx As Long, ByVal headerName As String, ByVal values As Collection)
    Index = idx
    header = headerName
    Set Values = values
    Set Filled = New Collection

    Dim v As Variant, s As String
    For Each v In values
        s = CStr(v)
        If Trim$(s) <> "" Then Filled.Add s
    Next v

    Dim total As Double
    total = Filled.Count
    If total = 0 Then total = 1

    If Values.Count > 0 Then
        FillRatio = Filled.Count / CDbl(Values.Count)
    Else
        FillRatio = 0
    End If

    Dim dateCount As Long, numCount As Long, decCount As Long, lenSum As Double
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim longIntCount As Long

    For Each v In Filled
        s = CStr(v)
        If LooksLikeDate(s) Then dateCount = dateCount + 1
        If LooksLikeAmount(s) Then
            numCount = numCount + 1
            If InStr(s, ",") > 0 Or InStr(s, ".") > 0 Then decCount = decCount + 1
            Dim digitsOnly As String
            digitsOnly = Replace(s, " ", "")
            If IsDigitsOnly(digitsOnly) And Len(digitsOnly) >= 8 Then longIntCount = longIntCount + 1
        End If
        lenSum = lenSum + Len(s)
        Dim key As String
        key = LCase$(Trim$(s))
        If Not seen.Exists(key) Then seen(key) = True
    Next v

    DateRatio = dateCount / total
    NumericRatio = numCount / total
    Dim numTotal As Double
    numTotal = numCount
    If numTotal = 0 Then numTotal = 1
    DecimalRatio = decCount / numTotal
    AvgLength = IIf(Filled.Count > 0, lenSum / total, 0)
    DistinctRatio = IIf(Filled.Count > 0, seen.Count / total, 0)
    LongIntRatio = longIntCount / numTotal
End Sub

Private Function IsDigitsOnly(ByVal s As String) As Boolean
    Dim i As Long
    If s = "" Then
        IsDigitsOnly = False
        Exit Function
    End If
    For i = 1 To Len(s)
        If Not (Mid$(s, i, 1) >= "0" And Mid$(s, i, 1) <= "9") Then
            IsDigitsOnly = False
            Exit Function
        End If
    Next i
    IsDigitsOnly = True
End Function

Public Function IsDateish() As Boolean
    IsDateish = (DateRatio >= 0.6)
End Function

Public Function IsNumericCol() As Boolean
    IsNumericCol = (NumericRatio >= 0.6 And DateRatio < 0.6)
End Function
