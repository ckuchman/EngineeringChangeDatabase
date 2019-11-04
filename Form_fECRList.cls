VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fECRList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ID_DblClick(Cancel As Integer)
    DoCmd.OpenForm "fECR", , , "ID=" & ID
End Sub

Private Sub AssociatedEC_DblClick(Cancel As Integer)
    DoCmd.OpenForm "fEC", , , "ID=" & AssociatedEC
End Sub
