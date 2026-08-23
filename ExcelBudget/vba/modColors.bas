Attribute VB_Name = "modColors"
Option Explicit
' Same palette as the LibreOffice version, converted from 0xRRGGBB to the
' 0xBBGGRR ("OLE color") Long format Excel's Font.Color/Interior.Color use.

Public Const NAVY As Long = &H604933
Public Const ORANGE As Long = &H2465F4
Public Const TEXT_GREY As Long = &H756457
Public Const MUTED As Long = &H877868
Public Const DARK As Long = &H434343
Public Const PANEL As Long = &HEFEDEB
Public Const PEACH As Long = &HEDF2FF
Public Const WHITE As Long = &HFFFFFF
Public Const BAR_GREY As Long = &HC0B7AE
Public Const LIGHT_TEXT As Long = &HCCCCCC

Public Const FONT_BODY As String = "Calibri"
Public Const FONT_TITLE As String = "Calibri Light"

Public Const FMT_CURRENCY As String = "#,##0 ""kr."""
Public Const FMT_CURRENCY2 As String = "#,##0.00 ""kr."""
Public Const FMT_SIGNED As String = "+#,##0 ""kr."";-#,##0 ""kr."""
Public Const FMT_DATE As String = "dd-mm-yyyy"
Public Const FMT_MONTH As String = "mmmm yyyy"
