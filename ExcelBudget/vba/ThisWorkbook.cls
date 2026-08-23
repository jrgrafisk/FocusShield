Option Explicit
' NOTE ON SETUP: this file is NOT imported like the others. ThisWorkbook
' already exists in every workbook - double-click "ThisWorkbook" in the
' Project Explorer (under "Microsoft Excel Objects") and paste this code
' into its (empty) code window instead. See the setup guide for details.

Private Sub Workbook_Open()
    modMenu.BuildMenu
End Sub

Private Sub Workbook_AddinInstall()
    modMenu.BuildMenu
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    modMenu.RemoveMenu
End Sub

Private Sub Workbook_AddinUninstall()
    modMenu.RemoveMenu
End Sub
