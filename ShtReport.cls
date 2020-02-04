VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
    Me.ExportAsFixedFormat xlTypePDF, "CDC Report.PDF", openafterpublish:=True
    
    ShtReport.Visible = xlSheetVeryHidden
    PerfSettingsOff
End Sub

