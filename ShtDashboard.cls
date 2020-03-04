VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtDashboard
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 04 Feb 20
'===============================================================
Option Explicit

Private Sub UpdateTrend()
    Dim CurrMonth As Integer
    Dim SheetMonth As Integer
    Dim MonthChange As Boolean
    Dim RngTopRow As Range
    Dim RngQIP As Range
    Dim RngNotQIP As Range
    Dim RngBottomRow As Range
    Dim ChartTrend As Chart
    Dim RngTotal As Range
    Dim RngPerc As Range
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set RngQIP = Range("T65")
    Set RngNotQIP = Range("U65")
    Set RngTotal = Range("V65")
    Set RngPerc = Range("W65")
    Set RngTopRow = Range("T65:W65")
    Set ChartTrend = ShtDashboard.ChartObjects("ChartTrend").Chart
    SheetMonth = [Sheet_month]
    CurrMonth = [Curr_Month]
    
    ShtDashboard.Unprotect "2683174"
    
    If SheetMonth <> CurrMonth Then MonthChange = True
    
    If MonthChange Then
        RngTopRow.Insert shift:=xlShiftDown
        ChartTrend.SetSourceData Range("S34:W39")
    Else
        RngQIP = [QIP_Count]
        RngTotal = [CREW_COUNT]
        RngNotQIP = [CREW_COUNT] - [QIP_Count]
        RngPerc = [Overall_QIP]
    End If
    
    [Sheet_month] = CurrMonth
    
    If USER_LEVEL <> DevLvl Then ShtDashboard.Protect "2683174"
    
    Set RngTopRow = Nothing
    Set RngQIP = Nothing
    Set RngTotal = Nothing
    Set RngPerc = Nothing
    Set ChartTrend = Nothing
    Set RngNotQIP = Nothing
Exit Sub

ErrorHandler:
    
    Set RngTopRow = Nothing
    Set RngQIP = Nothing
    Set RngTotal = Nothing
    Set RngPerc = Nothing
    Set RngNotQIP = Nothing

    If USER_LEVEL <> DevLvl Then ShtDashboard.Protect "2683174"

End Sub


Private Sub Worksheet_Activate()
    UpdateTrend
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
