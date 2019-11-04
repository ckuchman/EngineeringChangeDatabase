VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fProblemReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function Update_Reviewer(PR As Long, ReviewerInitials As String) As Boolean
    'Purpose: Changes the reviewer after determining that the website has the PR set to the correct progress. If it is changed by somone other than the new
    'reviewer it sends a email to alert the new owner.
    
    'Commented out as this seems to not be a issue anymore - 8/11/17 CAK
    'If (DLookup("[ECRStatusID]", "dbo_tblECR", "[ECID] =" & Me.ODBCID) = 1) Then
    '   MsgBox ("This PR first needs to be taken off the New list on the dashboard before assigning to a engineer (set status to Accepted).")
    '   Update_Reviewer = False
    '   Exit Function
    'End If

    'Sends a email to the new reviewer if they had not made the change
    If DLookup("RecieveEmails", "tUserData", "Initials = '" & ReviewerInitials & "'") = True And Not (ReviewerInitials = Form_fHomePage.Initials) Then
        GenFunc.CreateEmailWithOutlook DLookup("Email", "tUserData", "Initials = '" & ReviewerInitials & "'"), "PR#" & ODBCID & " Assigned", "PR#" & ODBCID & " has been assigned to you."
    End If
    
    DoCmd.RunCommand acCmdSave
    
    'Returns that the change was successful
    Update_Reviewer = True
    
End Function

Private Sub AssociatedECR_DblClick(Cancel As Integer)
    DoCmd.OpenForm "fECR", , , "ID=" & AssociatedECR
End Sub

Private Sub AttachtoECR_Click()
    'Purpose: Attaches the EC to a existing ECR after determining that all required information is present.

    Dim ECRLink As Variant   'The ECR that the PR will attach to
    
    If (AssociatedECR <> 0) Then
        If MsgBox("This PR already has a associated ECR, there can be multiple PRs attached to one ECR but a PR should only connect to one ECR, do you want to update the connected ECR?", vbYesNo) = vbNo Then
            End
        End If
    ElseIf (IsNull(Me.ManufacturingGroup)) Then
        If MsgBox("There are no manufacturing groups selected that this impacts, are you sure this is the case?", vbYesNo) = vbNo Then
            End
        End If
    End If
       
    ECRLink = InputBox("What ECR do you want to attach the PR to?")
    
    'Determines if the ECR exists
    If ECRLink = "" Then
        'Cancel was clicked
    ElseIf Not IsNumeric(ECRLink) Then
        MsgBox ("Input must be a number")
    ElseIf DLookup("ID", "tECR", "ID=" & ECRLink) Then
        'Opens the ECR form goes to the specified ECR
        DoCmd.OpenForm "fECR", , , "ID=" & ECRLink
        
        'Populates PR will information
        AssociatedECR = ECRLink
        Progress = "ECR Created"
        Reviewer = Form_fHomePage.Initials
        If IsNull(Me.DateCompleted) Then 'Date completed is only recorded once to improve metric tracking
            DateCompleted = Now
        End If
        
        DoCmd.Save acForm, "fECR"
        DoCmd.Save acForm, "fProblemReport"
        Me.Refresh
        
        'Update the dashboard
        Call UpdateDashboard(ID, "PR", Progress)
        
        DoCmd.Close acForm, "fProblemReport"
    Else
        MsgBox ("ECR does not exists")
    End If
End Sub

Private Sub CloseForm_Click()
    DoCmd.Save acForm, "fProblemReport"
    DoCmd.Close acForm, "fProblemReport"
    
    Forms!fHomePage.SetFocus
    DoCmd.RunCommand acCmdSave
End Sub

Private Sub CreateECR_Click()
    'Purpose: Creates a new ECR and attaches this PR to it after determining that all required information is present.

    Dim ECRLink As Integer
    
    If (AssociatedECR <> 0) Then
        If MsgBox("This PR already has a associated ECR, there can be multiple PRs attached to one ECR but a PR should only connect to one ECR, do you want to update the connected ECR?", vbYesNo) = vbNo Then
            End
        End If
        End
    ElseIf (IsNull(Me.ManufacturingGroup)) Then
        If MsgBox("There are no manufacturing groups selected that this impacts, are you sure this is the case?", vbYesNo) = vbNo Then
            End
        End If
    End If
    
    'Opens the ECR form and goes to a new record to populate
    DoCmd.OpenForm "fECR", acNormal, "", "", , acNormal
    DoCmd.GoToRecord , "", acNewRec
       
    'Runs a method that populates the ECR with relevent starting information
    ECRLink = Form_fECR.CreateECR
    
    'Populates PR will information
    AssociatedECR = ECRLink
    Progress = "ECR Created"
    Reviewer = Form_fHomePage.Initials
    If IsNull(Me.DateCompleted) Then 'Date completed is only recorded once to improve metric tracking
        DateCompleted = Now
    End If

    DoCmd.Save acForm, "fProblemReport"
    
    'Update the dashboard
    Call UpdateDashboard(ID, "PR", Progress)
    
    DoCmd.Close acForm, "fProblemReport"
End Sub

Private Sub CreateStdFolder_Click()
    CreateStandardFolder "PR", Me, True
End Sub

Private Sub RejectPR_Click()
    'Purpose: Rejects the PR after determining that all required information is present.

    Dim ReviewNotesCleaned As String

    Reviewer = Form_fHomePage.Initials
    
    If (IsNull(Me.ReviewNotes)) Then
        MsgBox ("There must be review notes describing why the PR was rejected, this will be displayed to the floor")
        End
    ElseIf (IsNull(Me.ManufacturingGroup)) Then
        If MsgBox("There are no manufacturing groups selected that this impacts, are you sure this is the case?", vbYesNo) = vbNo Then
            End
        End If
    End If
    
    If (AssociatedECR <> 0) Then
        If MsgBox("This PR already has a associated ECR, would you like to remove the ECR connection?", vbYesNo) = vbYes Then
            AssociatedECR = 0
        End If
    End If
        
    'Populates PR will information
    Progress = "Rejected"
    Reviewer = Form_fHomePage.Initials
    If IsNull(Me.DateCompleted) Then 'Date completed is only recorded once to improve metric tracking
        DateCompleted = Now
    End If
    
    DoCmd.RunCommand acCmdSave
    
    'Update the dashboard
    Call UpdateDashboard(ID, "PR", Progress)
                 
    MsgBox ("PR has been rejected.")
    DoCmd.Close acForm, "fProblemReport"
End Sub

Private Sub Reviewer_Change()
    If Not Form_fProblemReport.Update_Reviewer(Me.ODBCID, Me.Reviewer) Then
        Me.Reviewer = Null
    End If
End Sub
