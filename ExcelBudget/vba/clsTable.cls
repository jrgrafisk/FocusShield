VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' A rectangular table of strings plus what was learned while reading it.
' Port of budget_core.csvsniff.Table.

Public Rows As Collection          ' Collection of String() arrays (0-based)
Public header As Variant           ' String() array or Empty when none
Public Encoding As String
Public Delimiter As String
Public Preamble As Collection
Public Notes As Collection
Public Source As String
Public nColumns As Long

Private Sub Class_Initialize()
    Set Rows = New Collection
    header = Empty
    Encoding = "utf-8"
    Delimiter = ";"
    Set Preamble = New Collection
    Set Notes = New Collection
    Source = ""
    nColumns = 0
End Sub

Public Function ColumnValues(ByVal index As Long, Optional ByVal limit As Long = -1) As Collection
    Dim out As New Collection, r As Variant, n As Long
    n = 0
    For Each r In Rows
        If limit >= 0 And n >= limit Then Exit For
        If index <= UBound(r) Then
            out.Add r(index)
        Else
            out.Add ""
        End If
        n = n + 1
    Next r
    Set ColumnValues = out
End Function

Public Function HeaderName(ByVal index As Long) As String
    If Not IsEmpty(header) Then
        If index <= UBound(header) Then
            If Trim$(header(index)) <> "" Then
                HeaderName = Trim$(header(index))
                Exit Function
            End If
        End If
    End If
    HeaderName = "Kolonne " & (index + 1)
End Function

Public Function Describe() As String
    Dim delimName As String
    If Delimiter = vbTab Then delimName = "TAB" Else delimName = Delimiter
    Describe = Rows.Count & DK(" r~ae~kker, ") & nColumns & DK(" kolonner, tegns~ae~t ") & _
              Encoding & DK(", skilletegn ") & delimName
End Function
