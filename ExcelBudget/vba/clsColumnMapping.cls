VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColumnMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Which column plays which role, plus the parsing settings.
' Port of budget_core.columns.ColumnMapping.
' Long fields use -1 to mean "not set" (Python's None).

Public DateCol As Long
Public TextCols As Collection      ' Collection of Long
Public AmountCol As Long
Public AmountInCol As Long
Public AmountOutCol As Long
Public BalanceCol As Long
Public CurrencyCol As Long
Public DecimalSep As String
Public dayFirst As Boolean
Public Notes As Collection

Private Sub Class_Initialize()
    DateCol = -1
    Set TextCols = New Collection
    AmountCol = -1
    AmountInCol = -1
    AmountOutCol = -1
    BalanceCol = -1
    CurrencyCol = -1
    DecimalSep = ""
    dayFirst = True
    Set Notes = New Collection
End Sub

Public Function HasAmount() As Boolean
    HasAmount = (AmountCol >= 0 Or AmountInCol >= 0 Or AmountOutCol >= 0)
End Function

Public Function IsUsable() As Boolean
    IsUsable = (DateCol >= 0 And HasAmount())
End Function
