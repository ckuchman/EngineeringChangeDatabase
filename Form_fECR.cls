VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fECR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Private ECRSource As Integer - Don't think this is used, commenting out to see what happens

Function CreateECR() As Integer
    'Purpose: Creates a new ECR and populates it with information passed from the function calling it. Should return the ECR ID number.

    If Not (IsNull(Me.Description) Or IsNull(Me.Proposal)) Then
        If MsgBox("Something is going wrong and the new ECR is trying to overwrite a other ECR, contact Chris. Press NO to stop this process (Only Yes if Chris is present and troubleshooting)", vbYesNo) = vbNo Then
            End
        End If
    End If

    'Fills out the new ECR data with information known on creation
    Me.Owner = Form_fHomePage.Initials
    
    'Generates a secondary documentation folder if one doesn't exist
    CreateStandardFolder "ECR", Me
    
    'Makes sure the function passed back the ECR ID to the function that called it
    CreateECR = Me.ID
    
    'Shifts focus to save the ECR
    Forms!fECR.SetFocus
    DoCmd.RunCommand acCmdSave
End Function

Function CompleteECR(CompleteText As String) As Boolean
    'Purpose: Checks to make sure that the ECR is complete and takes in a string to concatinate to the error messages dependent on what the status is changing
    'to. Returns a Boolean on if the ECR is really to complete.

    'Determines that all required bits of information are filled out
    If IsNull(Me.Description) Then
        MsgBox ("Description is required " & CompleteText)
        CompleteECR = False
    ElseIf IsNull(Me.History) Then
        MsgBox ("Root Cause Analysis is required " & CompleteText)
        CompleteECR = False
    ElseIf IsNull(Me.Proposal) Then
        MsgBox ("Proposal is required " & CompleteText)
        CompleteECR = False
    ElseIf IsNull(Me.Category) Then
        MsgBox ("Category is required " & CompleteText)
        CompleteECR = False
    ElseIf Me.HoursToComplete = 0 Then
        MsgBox ("Hours to complete is required " & CompleteText)
        CompleteECR = False
    ElseIf Me.HoursToComplete > 8 And ECType = "Fast Track" Then
        MsgBox ("A EC that takes longer than 8 hours to complete should be run as a Full Track")
        CompleteECR = False
    Else
        CompleteECR = True
    End If
End Function

Public Function GenerateECRPDF(OpenPDF As Boolean)
    'Purpose: Generates a PDF for the ECR. The OpenPDF variable determines if the pdf is opened on creation

    'Hides the focus change to minimize processing time
    Application.Echo False
    
    Dim PrintType As String

    'Generates a secondary documentation folder if one doesnt exist
    CreateStandardFolder "ECR", Me

    'Creates a printer-friendly PDF version of the ECR
    DoCmd.OpenForm "fECRPrint" & PrintType, , , "ID =" & Me.ID
    DoCmd.OutputTo acOutputForm, "fECRPrint" & PrintType, acFormatPDF, "\\wfs.local\Watson\Engineering\03_Engineering\ECNs\ECR_Secondary_Documents\" & Me.ID & "\ECR_Report_" & Me.ID & ".pdf", OpenPDF
    DoCmd.Close acForm, "fECRPrint" & PrintType
    
    Application.Echo True

End Function

Private Sub AssociatedEC_DblClick(Cancel As Integer)
    DoCmd.OpenForm "fEC", , , "ID=" & AssociatedEC
End Sub

Private Sub AttachtoEC_Click()
    'Purpose: Attaches the ECR to a existing EC after determining that all required information is present.

    Dim ECLink As Variant   'The EC that the ECR will attach to
        
    ECLink = InputBox("What EC do you want to attach the PR to?")
    
    If AssociatedEC <> 0 Then
        If MsgBox("This ECR already has a associated EC, there can be multiple ECRs attached to one EC but a ECR should only connect to one EC, do you want to update the connected ECR?", vbYesNo) = vbNo Then
            End
        End If
    End If
    
    'Determines if the EC exists
    If ECLink = "" Then
        'Cancel was clicked
    ElseIf Not IsNumeric(ECLink) Then
        MsgBox ("Input must be a number")
    ElseIf CompleteECR("to attach this ECR to a EC") Then
        If DLookup("ID", "tEC", "ID=" & ECLink) Then
            'Opens the EC form goes to the specified EC
            DoCmd.OpenForm "fEC", , , "ID=" & ECLink
            
            If Forms!fEC.Progress = "Complete" Then
                MsgBox ("This EC has been completed, It can still be attached to but only if no further actions are needed to be taken.")
            ElseIf ECType = "Full Track" And Forms!fEC.ECType = "Fast Track" Then
                If Forms!fEC.Progress = "In Planning" Then
                    Forms!fEC.ECType = "Full Track"
                Else
                    MsgBox ("This ECR is a full track and is being attached to a fast track that is currently in progress. Run this ECR as a seperate EC or remake the original EC as a full track.")
                    DoCmd.Close acForm, "fEC"
                    End
                End If
            End If
                    
            'Populates the ECR will information
            AssociatedEC = ECLink
            Progress = "EC Created"
            If IsNull(Me.DateCompleted) Then 'Date completed is only recorded once to improve metric tracking
                DateCompleted = Now
            End If
    
            DoCmd.Save acForm, "fECR"
            
            'Generates a PDF of the completed ECR
            GenerateECRPDF False
            
            'Update the dashboard
            Call UpdateDashboard(ID, "ECR", Progress)
        
            DoCmd.Close acForm, "fECR"
        Else
            MsgBox ("EC does not exists")
        End If
    End If
End Sub

Private Sub CloseForm_Click()
    DoCmd.Save acForm, "fECR"
    DoCmd.Close acForm, "fECR"
    
    Forms!fHomePage.SetFocus
    DoCmd.RunCommand acCmdSave
End Sub

Private Sub CreateECRPDF_Click()
    'Purpose: Generates a PDF for the ECR through the associated function.

    GenerateECRPDF True
End Sub

Private Sub CreateStdFolder_Click()
    'Purpose: Generates the secondary information folder for the ECR if it does not exist

    CreateStandardFolder "ECR", Me
End Sub

Private Sub ExpansionEC_DblClick(Cancel As Integer)
    DoCmd.OpenForm "fEC", , , "ID=" & ExpansionEC
End Sub

Private Sub Form_Resize()
    'Purpose: Makes sure that the form is properly sized for the screen. Some extreme screen resolutions are still not fully supported.
    
    Dim WorkingPageHeight As Single, WorkingPageWidth As Single, HistHeight As Single, PoAHeight As Single, RevHeight As Single
    
    'Undersizes the form elements to fit fully on the screen without needing the scrollbars. Numbers found through trial-and-error
    
    'Height values
    WorkingPageHeight = Me.WindowHeight - 5200
    If WorkingPageHeight < 3750 Then
        WorkingPageHeight = 3750
    End If
    
    Me.History.Height = WorkingPageHeight * 0.6
    Me.ReviewerNotes.Height = WorkingPageHeight * 0.25 + 50
    
    'Width values
    WorkingPageWidth = Me.WindowWidth - 4500
    If WorkingPageWidth < 12500 Then
        WorkingPageWidth = 12500
    End If
  
    Me.History.Width = (WorkingPageWidth / 2) + 900
    Me.Proposal.Width = (WorkingPageWidth / 2) + 900
    Me.ProblemReport.Left = 200
    Me.ChangeType.Left = 200
    
    'Updates the form to reflect the new element sizes
    Me.Repaint
End Sub

Private Sub ImplementOnOrder_Click()
    'Purpose: Sets the ECR to implement on order after determining that all required information is present.

    Dim ProposalCleaned As String

    If CompleteECR("to set this ECR to Implement on Order") Then
        MsgBox ("Be sure to mark all models and studios that are impacted by this ECR with issues on ELS")
        
        'Populates the EC with information
        Progress = "Implement On Order"
        If IsNull(Me.DateCompleted) Then 'Date completed is only recorded once to improve metric tracking
            DateCompleted = Now
        End If

        DoCmd.RunCommand acCmdSave
        
        'Generates a PDF of the completed ECR
        GenerateECRPDF False
        
        'Update the dashboard
        Call UpdateDashboard(ID, "ECR", Progress)
        
        MsgBox ("ECR has been set to Implement On Order")
    End If
End Sub

Private Sub NewECR_Click()

On Error GoTo NewECR_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acNewRec
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If


NewECR_Click_Exit:
    Exit Sub

NewECR_Click_Err:
    MsgBox Error$
    Resume NewECR_Click_Exit

End Sub

Private Sub CreateEC_Click()
    'Purpose: Creates a new EC and attaches this ECR to it after determining that all required information is present.
    
    Dim ECLink As Integer
    
    If AssociatedEC <> 0 Then
        If MsgBox("This ECR already has a associated EC, there can be multiple ECRs attached to one EC but a ECR should only connect to one EC, do you want to update the connected ECR?", vbYesNo) = vbNo Then
            End
        End If
    End If
    
    If CompleteECR("to create a EC from this ECR") Then
        'Opens the EC form and goes to a new record to populate
        DoCmd.OpenForm "fEC", acNormal, "", "", , acNormal
        DoCmd.GoToRecord , "", acNewRec
        
        'Runs a method that populates the ECR with relevent starting information
        ECLink = Form_fEC.CreateEC(Me.Description, Me.History, Me.Proposal, Me.ECType)
        
        'Populates the ECR with information
        AssociatedEC = ECLink
        Progress = "EC Created"
        If IsNull(Me.DateCompleted) Then 'Date completed is only recorded once to improve metric tracking
            DateCompleted = Now
        End If

        'Saves the ECR and new EC
        DoCmd.Save acForm, "fEC"
        Me.Refresh
        DoCmd.Save acForm, "fECR"
        
        'Generates a PDF of the completed ECR
        GenerateECRPDF False
        
        'Update the dashboard
        Call UpdateDashboard(ID, "ECR", Progress)
        
        DoCmd.Close acForm, "fECR"
    End If
End Sub

Private Sub Owner_AfterUpdate()
    'Sends a email to the new owner if they had not made the change
    If DLookup("RecieveEmails", "tUserData", "Initials = '" & Me.Owner & "'") = True And Not (Me.Owner = Form_fHomePage.Initials) Then
        GenFunc.CreateEmailWithOutlook DLookup("Email", "tUserData", "Initials = '" & Me.Owner & "'"), "ECR#" & Me.ID & " Assigned", "ECR#" & Me.ID & " has been assigned to you."
    End If
    
    DoCmd.RunCommand acCmdSave
End Sub

Private Sub PendingImplementation_Click()
    'Purpose: Sets the ECR to pending implemention after determining that all required information is present.

    If CompleteECR("to set this ECR to Pending Implementation") Then
        MsgBox ("If this ECR needs to be done before a impacted product is ordered, set instead to Implement on Order")
        
        'Populates the ECR with information
        Progress = "Pending Implementation"
        If IsNull(Me.DateCompleted) Then 'Date completed is only recorded once to improve metric tracking
            DateCompleted = Now
        End If
        
        DoCmd.RunCommand acCmdSave
        
        'Generates a PDF of the completed ECR
        GenerateECRPDF False
        
        'Update the dashboard
        Call UpdateDashboard(ID, "ECR", Progress)
        
        MsgBox ("ECR has been set to Pending Implementation")
    End If
End Sub

Private Sub RejectECR_Click()
    'Purpose: Rejects the ECR after determining that all required information is present.

    Dim ReviewerNotesCleaned As String

    If IsNull(Me.ReviewerNotes) Then
        MsgBox ("There must be review notes describing why the ECR was rejected")
        End
    End If
    
    If AssociatedEC <> 0 Then
        If MsgBox("This ECR already has a associated EC, would you like to remove the EC connection?", vbYesNo) = vbYes Then
            AssociatedEC = 0
        End If
    End If
    
    'Populates the ECR with information
    Progress = "Rejected"
    If IsNull(Me.DateCompleted) Then 'Date completed is only recorded once to improve metric tracking
        DateCompleted = Now
    End If
    
    DoCmd.RunCommand acCmdSave
    
    'Generates a PDF of the rejected ECR
    GenerateECRPDF False
    
    'Update the dashboard
    Call UpdateDashboard(ID, "ECR", Progress)
    
    MsgBox ("ECR has been rejected.")
End Sub

