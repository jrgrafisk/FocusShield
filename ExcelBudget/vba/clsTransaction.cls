VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' One line in the Transaktioner sheet. Port of budget_core.budget.Transaction.

Public TxDate As Date
Public Text As String
Public Amount As Double
Public Category As String
Public Kind As String            ' KIND_EXPENSE or KIND_INCOME
Public MonthKeyStr As String     ' "2024-05"
Public Currency As String
Public SourceRow As Long         ' 0-based index into the original table

Public Sub Init(ByVal d As Date, ByVal txText As String, ByVal amt As Double, _
                ByVal category As String, Optional ByVal currency As String = "", _
                Optional ByVal sourceRow As Long = 0)
    TxDate = d
    Text = txText
    Amount = amt
    Me.Category = category
    Kind = IIf(amt < 0, KIND_EXPENSE, KIND_INCOME)
    MonthKeyStr = MonthKeyOf(d)
    Me.Currency = currency
    Me.SourceRow = sourceRow
End Sub

Public Function MonthName() As String
    MonthName = MonthLabelOf(MonthKeyStr)
End Function

Public Property Get IsIncome() As Boolean
    IsIncome = (Kind = KIND_INCOME)
End Property
