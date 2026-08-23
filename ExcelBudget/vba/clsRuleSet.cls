VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRuleSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' An ordered collection of keyword rules. Port of budget_core.rules.RuleSet.
'
' A rule is a keyword matched case/accent-insensitively against the
' transaction text. A keyword starting with "re:" is a regular expression
' instead. Longer keywords are tried first, so "Circle K" wins over "K".

Private mRules As Collection        ' Array(keyword, category), insertion order
Private mCompiled As Collection     ' Array(sortKey, kind, needleOrRegex, category)
Private mDefault As String

Private Sub Class_Initialize()
    Set mRules = New Collection
    Set mCompiled = New Collection
    mDefault = DEFAULT_CATEGORY
End Sub

Public Sub InitEmpty(Optional ByVal defaultCategory As String = "")
    If defaultCategory <> "" Then mDefault = defaultCategory
    Set mRules = New Collection
    Compile
End Sub

Public Sub InitDefault()
    InitFromRows modRules.DefaultRulesRows()
End Sub

' rows: Collection of 2-element arrays [keyword, category] (e.g. from a
' sheet). Blank rows and a "Nogleord;Kategori" style header are skipped.
Public Sub InitFromRows(ByVal rows As Collection)
    Set mRules = New Collection
    Dim row As Variant, keyword As String, category As String
    For Each row In rows
        If IsArray(row) Then
            If UBound(row) >= 0 Then keyword = Trim$(CStr(row(0))) Else keyword = ""
            If UBound(row) >= 1 Then category = Trim$(CStr(row(1))) Else category = ""
        Else
            keyword = ""
            category = ""
        End If
        If keyword <> "" And category <> "" Then
            Dim folded As String
            folded = FoldText(keyword)
            If folded <> "nogleord" And folded <> "noegleord" And _
               folded <> "keyword" And folded <> "sogeord" And folded <> "soegeord" Then
                mRules.Add Array(keyword, category)
            End If
        End If
    Next row
    Compile
End Sub

Public Property Get RuleCount() As Long
    RuleCount = mRules.Count
End Property

Public Function ToRows() As Collection
    Dim out As New Collection, r As Variant
    For Each r In mRules
        out.Add Array(r(0), r(1))
    Next r
    Set ToRows = out
End Function

' -- persistence (semicolon-delimited text file) --------------------------

Public Sub SaveToFile(ByVal path As String)
    Dim dir As String
    dir = Left$(path, InStrRev(path, "\"))
    If dir <> "" Then EnsureFolder dir
    Dim fnum As Integer
    fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, "Nogleord;Kategori"
    Dim r As Variant
    For Each r In mRules
        Print #fnum, r(0) & ";" & r(1)
    Next r
    Close #fnum
End Sub

Public Sub LoadFromFile(ByVal path As String)
    Dim rows As New Collection
    If Len(Dir$(path)) = 0 Then
        InitDefault
        Exit Sub
    End If
    Dim fnum As Integer, line As String
    fnum = FreeFile
    Open path For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, line
        If Trim$(line) <> "" Then
            Dim parts() As String
            parts = Split(line, ";")
            If UBound(parts) >= 1 Then
                rows.Add Array(parts(0), parts(1))
            End If
        End If
    Loop
    Close #fnum
    InitFromRows rows
End Sub

Private Sub EnsureFolder(ByVal path As String)
    Dim parts() As String, current As String, i As Long
    If Right$(path, 1) = "\" Then path = Left$(path, Len(path) - 1)
    parts = Split(path, "\")
    current = parts(0)
    For i = 1 To UBound(parts)
        current = current & "\" & parts(i)
        If Len(Dir$(current, vbDirectory)) = 0 Then
            On Error Resume Next
            MkDir current
            On Error GoTo 0
        End If
    Next i
End Sub

' -- editing ----------------------------------------------------------

Public Sub AddRule(ByVal keyword As String, ByVal category As String)
    keyword = Trim$(keyword)
    category = Trim$(category)
    If keyword = "" Or category = "" Then Exit Sub
    Dim folded As String
    folded = FoldText(keyword)
    Dim kept As New Collection, r As Variant
    For Each r In mRules
        If FoldText(CStr(r(0))) <> folded Then kept.Add r
    Next r
    kept.Add Array(keyword, category)
    Set mRules = kept
    Compile
End Sub

' -- matching -----------------------------------------------------------

Private Sub Compile()
    Dim tmp() As Variant
    ReDim tmp(1 To mRules.Count)
    Dim i As Long, r As Variant
    i = 0
    For Each r In mRules
        Dim keyword As String, category As String
        keyword = CStr(r(0))
        category = CStr(r(1))
        If Len(keyword) >= 3 And LCase$(Left$(keyword, 3)) = "re:" Then
            Dim patternText As String
            patternText = Trim$(Mid$(keyword, 4))
            Dim re As Object, compileOk As Boolean
            compileOk = True
            On Error Resume Next
            Err.Clear
            Set re = CreateObject("VBScript.RegExp")
            re.pattern = patternText
            re.IgnoreCase = True
            re.Global = False
            re.Test ""              ' forces the pattern to actually compile
            If Err.Number <> 0 Then compileOk = False
            On Error GoTo 0
            If compileOk Then
                i = i + 1
                tmp(i) = Array(Len(keyword), "re", re, category)
            End If
        Else
            Dim needle As String
            needle = FoldText(keyword)
            If needle <> "" Then
                i = i + 1
                If Len(needle) <= SHORT_KEYWORD And InStr(needle, " ") = 0 Then
                    tmp(i) = Array(Len(needle), "word", needle, category)
                Else
                    tmp(i) = Array(Len(needle), "kw", needle, category)
                End If
            End If
        End If
    Next r
    Dim n As Long
    n = i

    ' Stable insertion sort, longest keyword first.
    Dim j As Long, key As Variant
    For i = 2 To n
        key = tmp(i)
        j = i - 1
        Do While j >= 1
            If tmp(j)(0) < key(0) Then
                tmp(j + 1) = tmp(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        tmp(j + 1) = key
    Next i

    Set mCompiled = New Collection
    For i = 1 To n
        mCompiled.Add tmp(i)
    Next i
End Sub

Private Function IsWordChar(ByVal ch As String) As Boolean
    IsWordChar = (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9")
End Function

Private Function IsLetterAZ(ByVal ch As String) As Boolean
    IsLetterAZ = (ch >= "a" And ch <= "z")
End Function

' A folded short keyword must not match inside a longer word ("Irma" may
' not turn "firma" into groceries), but digits right after it are fine
' ("REMA1000", "NETTO8021"). Classic VBScript regex has no lookbehind, so
' this is checked by hand instead of with a regex.
Private Function WholeWordContains(ByVal haystack As String, ByVal needle As String) As Boolean
    Dim pos As Long, startPos As Long
    Dim beforeOk As Boolean, afterOk As Boolean
    startPos = 1
    Do
        pos = InStr(startPos, haystack, needle, vbBinaryCompare)
        If pos = 0 Then Exit Do
        If pos = 1 Then
            beforeOk = True
        Else
            beforeOk = Not IsWordChar(Mid$(haystack, pos - 1, 1))
        End If
        Dim afterPos As Long
        afterPos = pos + Len(needle)
        If afterPos > Len(haystack) Then
            afterOk = True
        Else
            afterOk = Not IsLetterAZ(Mid$(haystack, afterPos, 1))
        End If
        If beforeOk And afterOk Then
            WholeWordContains = True
            Exit Function
        End If
        startPos = pos + 1
    Loop
    WholeWordContains = False
End Function

Public Function Categorise(ByVal text As Variant) As String
    Dim folded As String
    folded = FoldText(CStr(text))
    If folded = "" Then
        Categorise = mDefault
        Exit Function
    End If
    Dim entry As Variant
    For Each entry In mCompiled
        Select Case entry(1)
            Case "kw"
                If InStr(1, folded, entry(2), vbBinaryCompare) > 0 Then
                    Categorise = entry(3)
                    Exit Function
                End If
            Case "word"
                If WholeWordContains(folded, entry(2)) Then
                    Categorise = entry(3)
                    Exit Function
                End If
            Case "re"
                Dim re As Object
                Set re = entry(2)
                If re.Test(CStr(text)) Then
                    Categorise = entry(3)
                    Exit Function
                End If
        End Select
    Next entry
    Categorise = mDefault
End Function

Public Function Categories() As Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim out As New Collection, r As Variant, cat As String
    For Each r In mRules
        cat = CStr(r(1))
        If Not seen.Exists(cat) Then
            seen(cat) = True
            out.Add cat
        End If
    Next r

    Dim income As Variant
    income = IncomeCategories()
    Dim incomeSet As Object
    Set incomeSet = CreateObject("Scripting.Dictionary")
    Dim k As Long
    For k = LBound(income) To UBound(income)
        incomeSet(income(k)) = True
    Next k

    Dim expenses As New Collection, one As Variant
    For Each one In out
        If Not incomeSet.Exists(CStr(one)) Then expenses.Add CStr(one)
    Next one
    SortCollection expenses

    Dim result As New Collection
    For Each one In expenses
        result.Add one
    Next one
    For k = LBound(income) To UBound(income)
        If seen.Exists(income(k)) Then result.Add income(k)
    Next k

    Dim hasDefault As Boolean, hasIgnored As Boolean, item As Variant
    For Each item In result
        If item = mDefault Then hasDefault = True
        If item = IGNORED_CATEGORY Then hasIgnored = True
    Next item
    If Not hasDefault Then result.Add mDefault
    If Not hasIgnored Then result.Add IGNORED_CATEGORY

    Set Categories = result
End Function

Private Sub SortCollection(ByRef col As Collection)
    Dim n As Long, arr() As String, i As Long, j As Long, tmp As String
    n = col.Count
    If n < 2 Then Exit Sub
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = col(i)
    Next i
    For i = 2 To n
        tmp = arr(i)
        j = i - 1
        Do While j >= 1
            If arr(j) > tmp Then
                arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j + 1) = tmp
    Next i
    Set col = New Collection
    For i = 1 To n
        col.Add arr(i)
    Next i
End Sub
