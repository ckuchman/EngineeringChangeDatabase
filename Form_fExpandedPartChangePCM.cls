VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fExpandedPartChangePCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
    'Purpose: makes sure that the EC form properly reflects any updates done in the expanded part change PCM list

    Forms!fEC.Refresh
End Sub
