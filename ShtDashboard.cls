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
' v1.1.0 - Added procedure to send trend data
'---------------------------------------------------------------
' Date - 17 Mar 20
'===============================================================
Option Explicit

Private Const RNG_TREND As String = "S66:W109"
' ===============================================================
' UpdateTrend
' Updates the current trend data with current figures.  At the
' Start of each month, push back data and archive
' ---------------------------------------------------------------
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
    
    ShtDashboard.Unprotect SEC_KEY
    
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
    
    If USER_LEVEL <> DevLvl Then ShtDashboard.Protect SEC_KEY
    
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

    If USER_LEVEL <> DevLvl Then ShtDashboard.Protect SEC_KEY

End Sub

' ===============================================================
' Worksheet_Activate
' Event processing for worksheet activate
' ---------------------------------------------------------------
Private Sub Worksheet_Activate()
    UpdateTrend
End Sub

' ===============================================================
' GetTrendData
' Gets trend data and returns in a text array
' ---------------------------------------------------------------
Public Function GetTrendData() As Variant()
    Dim AryTrendData() As Variant
    
    AryTrendData = Range(RNG_TREND)
    GetTrendData = AryTrendData
End Function

' ===============================================================
' WriteTrendData
' writes Trend Data from a text file stored at the location
' specified.  Returns TRUE if an error occurs.
' ---------------------------------------------------------------
Public Function WriteTrendData(ByRef AryTrendData() As Variant) As Boolean
        
    On Error GoTo ErrorHandler
    
    ShtDashboard.Unprotect SEC_KEY
        
    Me.Range(RNG_TREND) = AryTrendData
    
    If USER_LEVEL <> DevLvl Then ShtDashboard.Protect SEC_KEY
    
    WriteTrendData = False
Exit Function

ErrorHandler:
    If USER_LEVEL <> DevLvl Then ShtDashboard.Protect SEC_KEY
    
    WriteTrendData = True
End Function

