VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fPrintOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub FastTrackECNPrintout_Click()
    'Purpose: Generates a PDF that has a list of fast track ECs

    DoCmd.OutputTo acOutputReport, "rFastTrackECN", acFormatPDF, "FastTrackECN.pdf", True
    DoCmd.Close acReport, "rFastTrackECN"
End Sub

Private Sub FloorBoardPrintout_Click()
    'Purpose: Generates a PDF for each manufacturing department that contains a list of problem reports that have been marked as relavent for that department.
    
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    
    Set rs = CurrentDb.OpenRecordset("tManufacturingGroups") 'Creates a connection to the manufacturing group table, which lists all departments
    
    'Generates a array of queries that include each progress of PR/ECR/EC. This allows the PDF list to be organized by progress.
    Dim QueryArray As Variant
    QueryArray = Array("qBoardListCompletedEC", "qBoardListInProgressEC", "qBoardListInProgressECR", "qBoardListRejectedPR")
    
    'Cycles through the manufacturing groups until it hits the end of the list
    rs.MoveFirst
    Do Until rs.EOF = True
    
        'Clears the tBoardList from the last used so it can be populated again with the manufacturing group. The warnings are disabled to
        'prevent pop-ups from appearing for each manufacturing group
        DoCmd.SetWarnings False
        DoCmd.RunSQL "DELETE * FROM tBoardList"
        DoCmd.SetWarnings True
    
        'Cycles through the queries and runs them for the current manufacturing group, this will populate the tBoardList
        For Each element In QueryArray
            Set qdf = CurrentDb.QueryDefs(element)
            qdf!ManufacturingGroupValue = rs!GroupName
            qdf.Execute
            qdf.Close
        Next element
        
        'Generates the PDF
        DoCmd.OpenReport "rBoardList", acViewPreview
        Reports!rBoardList.Controls!BoardListHeader.Caption = rs!GroupName & " Problem Reports"
        DoCmd.OutputTo acOutputReport, "rBoardList", acFormatPDF, "BoardList" & rs!GroupName & ".pdf", True
        DoCmd.Close acReport, "rBoardList"
        
        Set qdf = Nothing
        rs.MoveNext
    Loop
    
    DoCmd.Close acReport, "rBoardList"
End Sub

Private Sub MetricPrintout_Click()
    'Purpose: Generates a PDF that has a list metrics

    DoCmd.OutputTo acOutputReport, "rMetrics", acFormatPDF, "Metric.pdf", True
    DoCmd.Close acReport, "rMetrics"
End Sub
