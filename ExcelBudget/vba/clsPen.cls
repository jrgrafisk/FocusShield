VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Small styling helper wrapping one worksheet - the Excel equivalent of the
' LibreOffice extension's SheetPen. Every method takes an A1-style ref
' ("B4", "B5:E20") so the sheet-writing code stays readable.

Public ws As Worksheet

Public Sub Init(ByVal sheet As Worksheet)
    Set ws = sheet
End Sub

Public Sub Text(ByVal ref As String, ByVal value As String, _
                Optional ByVal bold As Variant, Optional ByVal size As Variant, _
                Optional ByVal color As Variant, Optional ByVal align As String = "", _
                Optional ByVal italic As Variant, Optional ByVal wrap As Variant, _
                Optional ByVal valign As String = "", Optional ByVal bg As Variant, _
                Optional ByVal fontName As String = FONT_BODY)
    Dim rng As Range
    Set rng = ws.Range(ref)
    rng.Value = value
    ApplyStyle rng, bold, size, color, align, italic, wrap, valign, bg, fontName
End Sub

Public Sub Number(ByVal ref As String, ByVal value As Double, _
                  Optional ByVal fmt As String = FMT_CURRENCY, _
                  Optional ByVal bold As Variant, Optional ByVal size As Variant, _
                  Optional ByVal color As Variant, Optional ByVal align As String = "right", _
                  Optional ByVal bg As Variant, Optional ByVal fontName As String = FONT_BODY)
    Dim rng As Range
    Set rng = ws.Range(ref)
    rng.Value = value
    rng.NumberFormat = fmt
    ApplyStyle rng, bold, size, color, align, Empty, Empty, "", bg, fontName
End Sub

Public Sub DateCell(ByVal ref As String, ByVal value As Date, _
                    Optional ByVal fmt As String = FMT_DATE, _
                    Optional ByVal bold As Variant, Optional ByVal size As Variant, _
                    Optional ByVal color As Variant, Optional ByVal align As String = "left", _
                    Optional ByVal fontName As String = FONT_BODY)
    Dim rng As Range
    Set rng = ws.Range(ref)
    rng.Value = value
    rng.NumberFormat = fmt
    ApplyStyle rng, bold, size, color, align, Empty, Empty, "", Empty, fontName
End Sub

Public Sub Formula(ByVal ref As String, ByVal formulaText As String, _
                   Optional ByVal fmt As String = FMT_CURRENCY, _
                   Optional ByVal bold As Variant, Optional ByVal size As Variant, _
                   Optional ByVal color As Variant, Optional ByVal align As String = "right", _
                   Optional ByVal bg As Variant, Optional ByVal fontName As String = FONT_BODY)
    Dim rng As Range
    Set rng = ws.Range(ref)
    rng.formula = formulaText
    If fmt <> "" Then rng.NumberFormat = fmt
    ApplyStyle rng, bold, size, color, align, Empty, Empty, "", bg, fontName
End Sub

' Formatting only - never touches the cell's value/formula.
Public Sub Style(ByVal ref As String, Optional ByVal color As Variant, _
                 Optional ByVal italic As Variant, Optional ByVal bold As Variant, _
                 Optional ByVal fmt As String = "", Optional ByVal align As String = "", _
                 Optional ByVal bg As Variant)
    Dim rng As Range
    Set rng = ws.Range(ref)
    If Not IsMissing(color) Then rng.Font.color = color
    If Not IsMissing(italic) Then rng.Font.Italic = italic
    If Not IsMissing(bold) Then rng.Font.bold = bold
    If fmt <> "" Then rng.NumberFormat = fmt
    If align = "left" Then rng.HorizontalAlignment = xlLeft
    If align = "right" Then rng.HorizontalAlignment = xlRight
    If align = "center" Then rng.HorizontalAlignment = xlCenter
    If Not IsMissing(bg) Then rng.Interior.color = bg
End Sub

Private Sub ApplyStyle(ByVal rng As Range, ByVal bold As Variant, ByVal size As Variant, _
                       ByVal color As Variant, ByVal align As String, ByVal italic As Variant, _
                       ByVal wrap As Variant, ByVal valign As String, ByVal bg As Variant, _
                       ByVal fontName As String)
    rng.Font.name = fontName
    If Not IsMissing(size) Then rng.Font.size = size
    If Not IsMissing(bold) Then rng.Font.bold = bold
    If Not IsMissing(italic) Then rng.Font.Italic = italic
    If Not IsMissing(color) Then rng.Font.color = color
    If align = "left" Then rng.HorizontalAlignment = xlLeft
    If align = "right" Then rng.HorizontalAlignment = xlRight
    If align = "center" Then rng.HorizontalAlignment = xlCenter
    If valign = "center" Then rng.VerticalAlignment = xlCenter
    If Not IsMissing(wrap) Then rng.WrapText = wrap
    If Not IsMissing(bg) Then rng.Interior.color = bg
End Sub

Public Sub Merge(ByVal ref As String)
    ws.Range(ref).Merge
End Sub

Public Sub ColumnWidth(ByVal colIndex0 As Long, ByVal width As Double)
    ws.Columns(colIndex0 + 1).ColumnWidth = width
End Sub

Public Sub RowHeight(ByVal row1 As Long, ByVal height As Double)
    ws.Rows(row1).RowHeight = height
End Sub

' A hover comment, hidden until the user points at the cell (like the
' LibreOffice version's annotations).
Public Sub Note(ByVal ref As String, ByVal noteText As String)
    If noteText = "" Then Exit Sub
    Dim rng As Range
    Set rng = ws.Range(ref)
    On Error Resume Next
    rng.ClearComments
    On Error GoTo 0
    rng.AddComment noteText
    rng.Comment.Visible = False
End Sub

Public Sub Freeze(ByVal col0 As Long, ByVal row0 As Long)
    ws.Application.Goto ws.Cells(row0 + 1, col0 + 1)
    ws.Application.ActiveWindow.FreezePanes = True
End Sub

' A dropdown list restricted to the values in sourceRef (e.g. "Konti!$B$5:$B$50").
Public Sub Dropdown(ByVal ref As String, ByVal sourceRef As String)
    Dim rng As Range
    Set rng = ws.Range(ref)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="=" & sourceRef
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    On Error GoTo 0
End Sub
