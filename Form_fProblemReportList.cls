VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fProblemReportList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub OBDCID_DblClick(Cancel As Integer)
    DoCmd.OpenForm "fProblemReport", , , "OBDCID=" & OBDCID
End Sub

Private Sub AssociatedECR_DblClick(Cancel As Integer)
    DoCmd.OpenForm "fECR", , , "ID=" & AssociatedECR
End Sub

Private Sub Reviewer_AfterUpdate()
    If Not IsNull(Reviewer) Then
        If Not Form_fProblemReport.Update_Reviewer(Me.OBDCID, Me.Reviewer) Then
            Me.Reviewer = ""
        End If
    End If
End Sub
