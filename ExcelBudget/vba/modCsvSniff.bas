Attribute VB_Name = "modCsvSniff"
Option Explicit
' Turn "some CSV file" into a clean clsTable. Port of budget_core/csvsniff.py.
'
' Handles byte order marks, UTF-8/UTF-16, Windows-1252/Latin-1,
' semicolon/comma/tab/pipe separators, "sep=;" hint lines, bank preambles
' above the real header, missing headers and ragged rows.
' (UTF-32 is not handled - vanishingly rare in real bank exports, and not
' cleanly supported by the ADODB.Stream charset machinery this module
' relies on for encoding-aware reads.)

Public Const SAMPLE_ROWS As Long = 200
Private mHeaderWords As Object

Private Sub EnsureHeaderWords()
    If Not mHeaderWords Is Nothing Then Exit Sub
    Set mHeaderWords = CreateObject("Scripting.Dictionary")
    Dim words As Variant, i As Long
    words = Array( _
        "dato", "date", "datum", "bogfort", "bogforingsdato", "posteringsdato", _
        "transaktionsdato", "valor", "valordato", "rentedato", "tid", "time", _
        "tekst", "text", "beskrivelse", "description", "posteringstekst", _
        "narrative", "details", "detaljer", "modtager", "afsender", "payee", _
        "merchant", "reference", "memo", "note", "navn", "name", "titel", _
        "belob", "amount", "sum", "value", "betrag", "montant", "beloeb", _
        "saldo", "balance", "beholdning", "konto", "account", "kontonummer", _
        "valuta", "currency", "type", "kategori", "category", "status", _
        "debet", "kredit", "debit", "credit", "ind", "ud", "indsat", "haevet", _
        "indbetaling", "udbetaling", "withdrawal", "deposit", "posting", _
        "transaktion", "transaction", "art", "nummer", "number", "id", _
        "afstemt", "gebyr", "fee")
    For i = LBound(words) To UBound(words)
        mHeaderWords(words(i)) = True
    Next i
End Sub

' -----------------------------------------------------------------------
' Encoding
' -----------------------------------------------------------------------

Private Function ReadAllBytes(ByVal path As String) As Variant
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.LoadFromFile path
    ReadAllBytes = stream.Read
    stream.Close
End Function

Private Function BytesStartWith(ByVal bytes() As Byte, ByVal prefix() As Byte) As Boolean
    Dim i As Long
    If UBound(bytes) < UBound(prefix) Then
        BytesStartWith = False
        Exit Function
    End If
    For i = LBound(prefix) To UBound(prefix)
        If bytes(i) <> prefix(i) Then
            BytesStartWith = False
            Exit Function
        End If
    Next i
    BytesStartWith = True
End Function

Private Function DecodeWithCharset(ByVal bytesVariant As Variant, ByVal charset As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write bytesVariant
    stream.Position = 0
    stream.Type = 2 ' adTypeText
    stream.Charset = charset
    DecodeWithCharset = stream.ReadText
    stream.Close
End Function

' Returns (text, encodingName) as a 2-element array. Never raises - the
' last resort (Windows-1252) always decodes something.
Public Function DetectEncoding(ByVal bytesVariant As Variant, Optional ByVal forcedEncoding As String = "") As Variant
    If forcedEncoding <> "" Then
        DetectEncoding = Array(DecodeWithCharset(bytesVariant, forcedEncoding), forcedEncoding)
        Exit Function
    End If
    Dim b() As Byte
    b = bytesVariant
    Dim n As Long
    n = UBound(b) - LBound(b) + 1

    If n >= 3 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then
            DetectEncoding = Array(DecodeWithCharset(bytesVariant, "utf-8"), "utf-8-sig")
            Exit Function
        End If
    End If
    If n >= 2 Then
        If b(0) = &HFF And b(1) = &HFE Then
            DetectEncoding = Array(DecodeWithCharset(bytesVariant, "unicode"), "utf-16")
            Exit Function
        End If
        If b(0) = &HFE And b(1) = &HFF Then
            DetectEncoding = Array(DecodeWithCharset(bytesVariant, "unicodeFFFE"), "utf-16")
            Exit Function
        End If
    End If

    ' UTF-16 without a BOM: lots of zero bytes at even or odd positions.
    Dim sampleLen As Long, zeroCount As Long, i As Long
    sampleLen = 4096
    If n < sampleLen Then sampleLen = n
    zeroCount = 0
    For i = 0 To sampleLen - 1
        If b(i) = 0 Then zeroCount = zeroCount + 1
    Next i
    If sampleLen > 0 And zeroCount > sampleLen \ 4 Then
        Dim evenZeros As Long
        evenZeros = 0
        For i = 0 To sampleLen - 2 Step 2
            If b(i) = 0 Then evenZeros = evenZeros + 1
        Next i
        If evenZeros > sampleLen \ 5 Then
            DetectEncoding = Array(DecodeWithCharset(bytesVariant, "unicodeFFFE"), "utf-16-be")
        Else
            DetectEncoding = Array(DecodeWithCharset(bytesVariant, "unicode"), "utf-16-le")
        End If
        Exit Function
    End If

    On Error Resume Next
    Dim text As String
    text = ""
    Err.Clear
    text = DecodeWithCharset(bytesVariant, "utf-8")
    If Err.Number = 0 And IsValidUtf8(b) Then
        On Error GoTo 0
        DetectEncoding = Array(text, "utf-8")
        Exit Function
    End If
    On Error GoTo 0

    DetectEncoding = Array(DecodeWithCharset(bytesVariant, "windows-1252"), "cp1252")
End Function

Private Function IsValidUtf8(ByRef b() As Byte) As Boolean
    ' ADODB happily decodes non-UTF-8 bytes as UTF-8 with replacement
    ' characters instead of failing, so check the byte stream by hand.
    Dim i As Long, n As Long, extra As Long
    n = UBound(b) - LBound(b) + 1
    i = 0
    Do While i < n
        Dim c As Long
        c = b(i)
        If c <= &H7F Then
            extra = 0
        ElseIf c >= &HC2 And c <= &HDF Then
            extra = 1
        ElseIf c >= &HE0 And c <= &HEF Then
            extra = 2
        ElseIf c >= &HF0 And c <= &HF4 Then
            extra = 3
        Else
            IsValidUtf8 = False
            Exit Function
        End If
        Dim k As Long
        For k = 1 To extra
            If i + k >= n Then
                IsValidUtf8 = False
                Exit Function
            End If
            If (b(i + k) And &HC0) <> &H80 Then
                IsValidUtf8 = False
                Exit Function
            End If
        Next k
        i = i + extra + 1
    Loop
    IsValidUtf8 = True
End Function

' -----------------------------------------------------------------------
' Delimiter
' -----------------------------------------------------------------------

Private Function SplitAnyChars(ByVal s As String, ByVal seps As String) As String()
    Dim i As Long, ch As String
    For i = 1 To Len(seps)
        ch = Mid$(seps, i, 1)
        s = Replace(s, ch, Chr(1))
    Next i
    SplitAnyChars = Split(s, Chr(1))
End Function

Private Function ParseCsvText(ByVal text As String, ByVal delimiter As String, _
                              Optional ByVal quotechar As String = """") As Collection
    Dim rows As New Collection
    Dim fields As New Collection
    Dim field As String
    Dim inQuotes As Boolean
    Dim i As Long, n As Long, ch As String
    n = Len(text)
    i = 1
    Do While i <= n
        ch = Mid$(text, i, 1)
        If inQuotes Then
            If ch = quotechar Then
                If i < n And Mid$(text, i + 1, 1) = quotechar Then
                    field = field & quotechar
                    i = i + 2
                Else
                    inQuotes = False
                    i = i + 1
                End If
            Else
                field = field & ch
                i = i + 1
            End If
        Else
            If ch = quotechar And field = "" Then
                inQuotes = True
                i = i + 1
            ElseIf ch = delimiter Then
                fields.Add field
                field = ""
                i = i + 1
            ElseIf ch = vbLf Then
                fields.Add field
                rows.Add ColToStringArray(fields)
                Set fields = New Collection
                field = ""
                i = i + 1
            Else
                field = field & ch
                i = i + 1
            End If
        End If
    Loop
    If field <> "" Or fields.Count > 0 Then
        fields.Add field
        rows.Add ColToStringArray(fields)
    End If
    Set ParseCsvText = rows
End Function

Private Function ColToStringArray(ByVal col As Collection) As String()
    Dim out() As String, i As Long
    ReDim out(0 To col.Count - 1)
    For i = 1 To col.Count
        out(i - 1) = col(i)
    Next i
    ColToStringArray = out
End Function

' Pick the separator that yields the most consistent, widest table.
Public Function DetectDelimiter(ByVal text As String) As String
    Dim allLines() As String, sample As String, n As Long, i As Long
    allLines = Split(text, vbLf)
    n = 0
    sample = ""
    For i = LBound(allLines) To UBound(allLines)
        If Trim$(allLines(i)) <> "" Then
            If sample <> "" Then sample = sample & vbLf
            sample = sample & allLines(i)
            n = n + 1
            If n >= 60 Then Exit For
        End If
    Next i
    If sample = "" Then
        DetectDelimiter = ";"
        Exit Function
    End If

    Dim candidates As Variant
    candidates = Array(";", ",", vbTab, "|")
    Dim best As String, bestScore As Double
    best = ""
    bestScore = 0
    Dim c As Long
    For c = LBound(candidates) To UBound(candidates)
        Dim rows As Collection
        Set rows = ParseCsvText(sample, CStr(candidates(c)))
        Dim widths As Object
        Set widths = CreateObject("Scripting.Dictionary")
        Dim r As Variant, w As Long, hasContent As Boolean
        Dim totalWidths As Long
        totalWidths = 0
        For Each r In rows
            hasContent = False
            Dim cell As Variant
            For Each cell In r
                If Trim$(CStr(cell)) <> "" Then hasContent = True: Exit For
            Next cell
            If hasContent Then
                w = UBound(r) - LBound(r) + 1
                If Not widths.Exists(w) Then widths(w) = 0
                widths(w) = widths(w) + 1
                totalWidths = totalWidths + 1
            End If
        Next r
        If totalWidths = 0 Then GoTo NextCandidate
        Dim bestWidth As Long, bestFreq As Long, key As Variant
        bestFreq = -1
        For Each key In widths.Keys
            If widths(key) > bestFreq Or (widths(key) = bestFreq And key > bestWidth) Then
                bestFreq = widths(key)
                bestWidth = key
            End If
        Next key
        If bestWidth < 2 Then GoTo NextCandidate
        Dim score As Double
        score = (bestFreq / CDbl(totalWidths)) * WorksheetFunction.Min(bestWidth, 12)
        If score > bestScore + 0.000000001 Then
            bestScore = score
            best = CStr(candidates(c))
        End If
NextCandidate:
    Next c
    If best = "" Then best = ";"
    DetectDelimiter = best
End Function

Private Function StripSepHint(ByVal text As String, ByRef hinted As String) As String
    hinted = ""
    Dim nl As Long
    nl = InStr(text, vbLf)
    Dim first As String, rest As String
    If nl = 0 Then
        StripSepHint = text
        Exit Function
    End If
    first = Left$(text, nl - 1)
    rest = Mid$(text, nl + 1)
    Dim stripped As String
    stripped = Trim$(first)
    If Len(stripped) > 0 And AscW(Left$(stripped, 1)) = 65279 Then
        stripped = Mid$(stripped, 2)
    End If
    If (Left$(stripped, 4) = "sep=") And Len(stripped) = 5 Then
        hinted = Mid$(stripped, 5, 1)
        StripSepHint = rest
        Exit Function
    End If
    StripSepHint = text
End Function

' -----------------------------------------------------------------------
' Header detection
' -----------------------------------------------------------------------

Private Function LooksLikeHeader(ByVal row() As String, ByVal following As Collection) As Boolean
    EnsureHeaderWords
    Dim nonEmpty As New Collection, i As Long, cellVal As String
    For i = LBound(row) To UBound(row)
        cellVal = Trim$(row(i))
        If cellVal <> "" Then nonEmpty.Add cellVal
    Next i
    If nonEmpty.Count = 0 Then
        LooksLikeHeader = False
        Exit Function
    End If

    For Each cellVal In nonEmpty
        If LooksLikeDate(cellVal) Then
            LooksLikeHeader = False
            Exit Function
        End If
    Next cellVal

    Dim hits As Long, folded As String, tokens() As String, t As Long
    hits = 0
    For Each cellVal In nonEmpty
        folded = FoldText(cellVal)
        If mHeaderWords.Exists(folded) Then
            hits = hits + 1
        Else
            tokens = SplitAnyChars(folded, "/-")
            Dim matched As Boolean
            matched = False
            For t = LBound(tokens) To UBound(tokens)
                If tokens(t) <> "" And mHeaderWords.Exists(tokens(t)) Then matched = True: Exit For
            Next t
            If matched Then hits = hits + 1
        End If
    Next cellVal
    If hits >= 2 Or (hits = 1 And nonEmpty.Count <= 3) Then
        LooksLikeHeader = True
        Exit Function
    End If

    Dim nonBlankFollowing As New Collection, fr As Variant
    For Each fr In following
        If Not IsBlankRow(fr) Then nonBlankFollowing.Add fr
    Next fr
    If nonBlankFollowing.Count > 0 Then
        Dim withDate As Long, numericBelow As Long
        withDate = 0: numericBelow = 0
        For Each fr In nonBlankFollowing
            Dim hasDate As Boolean, hasNum As Boolean, cv As Variant
            hasDate = False: hasNum = False
            For Each cv In fr
                If LooksLikeDate(CStr(cv)) Then hasDate = True
                If LooksLikeAmount(CStr(cv)) Then hasNum = True
            Next cv
            If hasDate Then withDate = withDate + 1
            If hasNum Then numericBelow = numericBelow + 1
        Next fr
        If withDate >= WorksheetFunction.Max(1, nonBlankFollowing.Count \ 2) Then
            LooksLikeHeader = True
            Exit Function
        End If
        Dim anyAmountInRow As Boolean
        anyAmountInRow = False
        For Each cellVal In nonEmpty
            If LooksLikeAmount(cellVal) Then anyAmountInRow = True: Exit For
        Next cellVal
        If numericBelow >= WorksheetFunction.Max(1, nonBlankFollowing.Count \ 2) And Not anyAmountInRow Then
            LooksLikeHeader = True
            Exit Function
        End If
    End If
    LooksLikeHeader = (hits > 0)
End Function

Private Function FindStart(ByVal rows As Collection, ByRef widthOut As Long) As Long
    Dim widths As Object
    Set widths = CreateObject("Scripting.Dictionary")
    Dim r As Variant, w As Long
    For Each r In rows
        If Not IsBlankRow(r) Then
            w = UBound(r) - LBound(r) + 1
            If Not widths.Exists(w) Then widths(w) = 0
            widths(w) = widths(w) + 1
        End If
    Next r
    If widths.Count = 0 Then
        widthOut = 0
        FindStart = 0
        Exit Function
    End If
    Dim bestWidth As Long, bestFreq As Long, key As Variant
    bestFreq = -1
    For Each key In widths.Keys
        If widths(key) > bestFreq Or (widths(key) = bestFreq And key > bestWidth) Then
            bestFreq = widths(key)
            bestWidth = key
        End If
    Next key
    widthOut = bestWidth

    Dim n As Long, i As Long
    n = rows.Count
    For i = 1 To n
        r = rows(i)
        If IsBlankRow(r) Or (UBound(r) - LBound(r) + 1) <> bestWidth Then GoTo ContinueLoop
        Dim following As New Collection, j As Long, cnt As Long
        cnt = 0
        For j = i + 1 To WorksheetFunction.Min(i + 4, n)
            If Not IsBlankRow(rows(j)) Then
                following.Add rows(j)
                cnt = cnt + 1
            End If
        Next j
        If following.Count = 0 Then
            FindStart = i - 1
            Exit Function
        End If
        Dim matching As Long
        matching = 0
        Dim fr As Variant
        For Each fr In following
            If (UBound(fr) - LBound(fr) + 1) = bestWidth Then matching = matching + 1
        Next fr
        If matching >= WorksheetFunction.Max(1, following.Count - 1) Then
            FindStart = i - 1
            Exit Function
        End If
ContinueLoop:
    Next i
    FindStart = 0
End Function

Private Function Uniquify(ByVal names() As String) As String()
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim out() As String, i As Long, nm As String
    ReDim out(LBound(names) To UBound(names))
    For i = LBound(names) To UBound(names)
        nm = Trim$(names(i))
        If nm = "" Then nm = "Kolonne " & (i - LBound(names) + 1)
        If seen.Exists(nm) Then
            seen(nm) = seen(nm) + 1
            nm = nm & " (" & seen(nm) & ")"
        Else
            seen(nm) = 1
        End If
        out(i) = nm
    Next i
    Uniquify = out
End Function

' Build a clsTable from already-split rows (a Collection of String() arrays).
Public Function TableFromRows(ByVal rawRows As Collection, ByVal hasHeaderIn As Variant, _
                              ByVal encodingName As String, ByVal delimiter As String, _
                              ByVal source As String) As clsTable
    Dim t As New clsTable
    t.Encoding = encodingName
    t.Delimiter = delimiter
    t.Source = source

    Dim widthOut As Long
    Dim startAt As Long
    startAt = FindStart(rawRows, widthOut)

    Dim r As Variant, idx As Long
    idx = 0
    For Each r In rawRows
        idx = idx + 1
        If idx <= startAt And Not IsBlankRow(r) Then t.Preamble.Add r
    Next r
    If t.Preamble.Count > 0 Then
        t.Notes.Add "Sprang " & t.Preamble.Count & " indledende linje(r) over."
    End If

    Dim body As New Collection
    idx = 0
    For Each r In rawRows
        idx = idx + 1
        If idx > startAt And Not IsBlankRow(r) Then body.Add r
    Next r
    If body.Count = 0 Then
        t.nColumns = 0
        Set TableFromRows = t
        Exit Function
    End If

    Dim useHeader As Boolean
    Dim first As Variant
    first = body(1)
    If VarType(hasHeaderIn) = vbBoolean Then
        useHeader = hasHeaderIn
    Else
        Dim following As New Collection, k As Long
        k = 0
        For Each r In body
            k = k + 1
            If k = 1 Then GoTo SkipFirst
            following.Add r
            If following.Count >= 20 Then Exit For
SkipFirst:
        Next r
        useHeader = LooksLikeHeader(first, following)
    End If

    Dim rest As New Collection
    idx = 0
    For Each r In body
        idx = idx + 1
        If idx > 1 Then rest.Add r
    Next r

    If useHeader Then
        Dim hdr() As String, i As Long
        ReDim hdr(LBound(first) To UBound(first))
        For i = LBound(first) To UBound(first)
            hdr(i) = Trim$(first(i))
        Next i
        t.header = Uniquify(hdr)
        Set body = rest
    Else
        t.Notes.Add DK("Ingen overskriftsr~ae~kke fundet - kolonnerne navngives automatisk.")
    End If

    Dim ragged As Long
    Dim fixedRows As New Collection
    For Each r In body
        Dim width As Long
        width = UBound(r) - LBound(r) + 1
        If width < widthOut Then
            If width < 2 Then GoTo SkipRow
            ragged = ragged + 1
            Dim padded() As String, j As Long
            ReDim padded(0 To widthOut - 1)
            For j = 0 To width - 1
                padded(j) = r(LBound(r) + j)
            Next j
            For j = width To widthOut - 1
                padded(j) = ""
            Next j
            fixedRows.Add padded
        Else
            fixedRows.Add r
        End If
SkipRow:
    Next r
    If ragged > 0 Then
        t.Notes.Add ragged & DK(" r~ae~kke(r) havde f~ae~rre kolonner end resten og blev fyldt ud.")
    End If

    Set t.Rows = fixedRows
    Dim maxLen As Long
    maxLen = widthOut
    If Not IsEmpty(t.header) Then
        If UBound(t.header) - LBound(t.header) + 1 > maxLen Then
            maxLen = UBound(t.header) - LBound(t.header) + 1
        End If
    End If
    t.nColumns = maxLen
    Set TableFromRows = t
End Function

' -----------------------------------------------------------------------
' Public entry point
' -----------------------------------------------------------------------

Public Function ReadTable(ByVal path As String, Optional ByVal forcedEncoding As String = "", _
                          Optional ByVal forcedDelimiter As String = "", _
                          Optional ByVal hasHeaderIn As Variant = Empty) As clsTable
    Dim raw As Variant
    raw = ReadAllBytes(path)
    Dim decoded As Variant
    decoded = DetectEncoding(raw, forcedEncoding)
    Dim text As String, encodingName As String
    text = decoded(0)
    encodingName = decoded(1)

    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)

    Dim hinted As String
    text = StripSepHint(text, hinted)
    Dim delimiter As String
    delimiter = forcedDelimiter
    If delimiter = "" And hinted <> "" Then delimiter = hinted
    If delimiter = "" Then delimiter = DetectDelimiter(text)

    Dim rawRows As Collection
    Set rawRows = ParseCsvText(text, delimiter)

    Dim fileName As String
    fileName = Mid$(path, InStrRev(path, "\") + 1)
    Set ReadTable = TableFromRows(rawRows, hasHeaderIn, encodingName, delimiter, fileName)
End Function
