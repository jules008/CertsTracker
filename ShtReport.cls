VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtReport
'===============================================================
' v1.0.0 - Initial Version
' v1.0.1 - Added sort to report
'---------------------------------------------------------------
' Date - 19 Mar 20
'===============================================================
Option Explicit

Public Sub PrintReport(ByRef AryReport As Variant, Title As String, Headings() As String)
    Dim i As Integer
    Range("A:E").ClearContents
    Range("A1") = Title
    
    For i = LBound(Headings) To UBound(Headings)
        Range("A2").Offset(0, i) = Headings(i)
    Next
    
    PerfSettingsOn
    ShtReport.Visible = xlSheetVisible
    
    Range("A3:E50") = AryReport
    
    Range("E3").Sort Range("E3"), xlDescending, Header:=xlYes
    Me.ExportAsFixedFormat xlTypePDF, Title & ".pdf", openafterpublish:=True
    
    If USER_LEVEL <> DevLvl Then Me.Visible = xlSheetVeryHidden
    PerfSettingsOff
End Sub

