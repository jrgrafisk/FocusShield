Attribute VB_Name = "modMenu"
Option Explicit
' Builds the "Budget" menu on Excel's classic menu bar (CommandBars) - this
' still works in every current Excel version even though the UI is now the
' Ribbon; it shows up as a menu item under the Add-ins tab. A modern Ribbon
' tab would need a customUI14.xml editor this environment does not have
' access to, so CommandBars is the reliable, self-contained choice: it can
' be built entirely from VBA with no extra tooling.

Private Const MENU_CAPTION As String = "Budget"

Public Sub BuildMenu()
    RemoveMenu
    Dim bar As CommandBar
    Set bar = Application.CommandBars.ActiveMenuBar
    Dim menu As CommandBarPopup
    Set menu = bar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    menu.Caption = MENU_CAPTION

    AddItem menu, DK("Import~ee~r CSV-fil til budget..."), "ImportCsv"
    AddItem menu, DK("Tilf~oe~j flere posteringer (CSV)..."), "ImportCsvAppend"
    AddItem menu, "Lav budget ud fra det aktive ark...", "ImportActiveSheet", beginGroup:=True
    AddItem menu, DK("Opdat~ee~r kategorier og budget"), "RefreshBudget", beginGroup:=True
    AddItem menu, "Konti - ret saldi...", "ShowAccounts"
    AddItem menu, DK("Opret bankbudget (~oe~jebliksbillede)"), "CreateBankBudget"
    AddItem menu, "Vis kategoriregler", "ShowRules", beginGroup:=True
    AddItem menu, "Nulstil kategoriregler", "ResetRules"
    AddItem menu, "Om Budget fra CSV", "ShowAbout", beginGroup:=True
End Sub

Public Sub RemoveMenu()
    On Error Resume Next
    Application.CommandBars.ActiveMenuBar.Controls(MENU_CAPTION).Delete
    On Error GoTo 0
End Sub

Private Sub AddItem(ByVal menu As CommandBarPopup, ByVal caption As String, _
                    ByVal macroName As String, Optional ByVal beginGroup As Boolean = False)
    Dim item As CommandBarButton
    Set item = menu.Controls.Add(Type:=msoControlButton, Temporary:=True)
    item.Caption = caption
    item.OnAction = "'" & ThisWorkbook.Name & "'!modMain." & macroName
    item.BeginGroup = beginGroup
    item.FaceId = 0
End Sub
