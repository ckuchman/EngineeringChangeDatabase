VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fECRsubProblemReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub OBDCID_DblClick(Cancel As Integer)
    DoCmd.OpenForm "fProblemReport", , , "OBDCID=" & OBDCID
End Sub

