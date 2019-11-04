VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public Function CreateEC(Description As String, History As String, PlanOfAction As String, ECType As String) As Integer
    'Purpose: Creates a new EC and populates it with information passed from the function calling it. Should return the EC ID number.
    
    If Not (IsNull(Me.Description) Or IsNull(Me.PlanOfAction)) Then
        If MsgBox("Something is going wrong and the new EC is trying to overwrite a other EC, contact Chris. Press NO to stop this process (Only Yes if Chris is present and troubleshooting)", vbYesNo) = vbNo Then
            End
        End If
    End If
    
    'Fills out the new EC data with information known on creation
    Me.Owner = Form_fHomePage.Initials
    Me.DateCreated = Now
    Me.Description = Description
    Me.History = History
    Me.PlanOfAction = PlanOfAction
    Me.ECType = ECType
    
    'Generates a secondary documentation folder if one doesnt exist
    CreateStandardFolder "EC", Me
    
    'Makes sure the function passed back the EC ID to the function that called it
    CreateEC = Me.ID
    
    'Toggles visiblity of EC timeline based on if it is fast or full track
    ECTypeUpdate
    
    'Shifts focus to save the EC
    Forms!fEC.SetFocus
    DoCmd.RunCommand acCmdSave
End Function

Public Function ECTypeUpdate()
    'Purpose: Reveals elements only relevent to the ECR type on the EC form.
    
    If ECType = "Fast Track" Then
        Timeline.Visible = False
    Else
        Timeline.Visible = True
    End If
End Function

Public Function ProgressUpdate()
    'Purpose: Reveals elements only relevent to the EC's current status on the EC form.
    
    'Sets the focus to the EC form as no form specific is called for the changes
    Forms!fEC.ID.SetFocus
    
    'Reveals elements only relevent to the progress
    If Progress = "In Planning" Then
        MakeECActive.Visible = True
        ReviewEC.Visible = False
        PassReview.Visible = False
        FailReview.Visible = False
        FailReview.Visible = False
        CompleteEC.Visible = False
        CancelEC.Visible = True
    ElseIf Progress = "Active" Or Progress = "Failed Review" Then
        MakeECActive.Visible = False
        ReviewEC.Visible = True
        PassReview.Visible = False
        FailReview.Visible = False
        CompleteEC.Visible = False
        CancelEC.Visible = True
    ElseIf Progress = "For Review" Then
        MakeECActive.Visible = False
        ReviewEC.Visible = False
        PassReview.Visible = True
        FailReview.Visible = True
        CompleteEC.Visible = True
        CancelEC.Visible = True
    ElseIf Progress = "Passed Review" Then
        MakeECActive.Visible = False
        ReviewEC.Visible = False
        PassReview.Visible = False
        FailReview.Visible = False
        CompleteEC.Visible = True
        CancelEC.Visible = True
    Else
        MakeECActive.Visible = False
        ReviewEC.Visible = False
        PassReview.Visible = False
        FailReview.Visible = False
        CompleteEC.Visible = False
        CancelEC.Visible = False
    End If
    
    If Not Progress = "In Planning" Then
        PlanOfAction.Locked = True
        PlanOfAction.BackColor = 15921906 'Background 1, Darker 5%
        ApprovalNotes.Locked = True
        ApprovalNotes.BackColor = 15921906 'Background 1, Darker 5%
    Else
        PlanOfAction.Locked = False
        PlanOfAction.BackColor = 16777215 'Background 1
        ApprovalNotes.Locked = False
        ApprovalNotes.BackColor = 16777215 'Background 1
    End If
    
    'Makes sure that hidden buttons do not create blank gaps between buttons that do exist
    If MakeECActive.Visible = True Then
        MakeECActive.Width = 1620
    Else
        MakeECActive.Width = 0
    End If
    
    If ReviewEC.Visible = True Then
        ReviewEC.Width = 2700
    Else
        ReviewEC.Width = 0
    End If
    
    If PassReview.Visible = True Then
        PassReview.Width = 1500
    Else
        PassReview.Width = 0
    End If
    
    If FailReview.Visible = True Then
        FailReview.Width = 1500
    Else
        FailReview.Width = 0
    End If
    
    If CompleteEC.Visible = True Then
        CompleteEC.Width = 1500
    Else
        CompleteEC.Width = 0
    End If
    
    If CancelEC.Visible = True Then
        CancelEC.Width = 3360
    Else
        CancelEC.Width = 0
    End If

End Function

Public Function GenerateECPDF(OpenPDF As Boolean)
    'Purpose: Generates a PDF for the EC. The style of the pdf is dependent on the type of EC (fast or full). The OpenPDF variable determines if the pdf is
    'opened on creation

    Dim PrintType As String
    
    'Hides the focus change to minimize processing time
    Application.Echo False

    'Generates a secondary documentation folder if one doesnt exist
    CreateStandardFolder "EC", Me

    'Determines the print type needed dependent on fast or full track
    If ECType = "Full Track" Then
        PrintType = "Full"
    ElseIf ECType = "Fast Track" Then
        PrintType = "Fast"
    Else
        MsgBox ("Not a valid ECType")
        End
    End If
    
    'Creates a printer-friendly PDF version of the EC
    DoCmd.OpenForm "fECPrint" & PrintType, , , "ID =" & Me.ID
    DoCmd.OutputTo acOutputForm, "fECPrint" & PrintType, acFormatPDF, "\\wfs.local\Watson\Engineering\03_Engineering\ECNs\EC_Secondary_Documents\" & Me.ID & "\EC_Report_" & Me.ID & ".pdf", OpenPDF
    DoCmd.Close acForm, "fECPrint" & PrintType
    
    Application.Echo True
End Function

Private Sub CancelEC_Click()
    'Purpose: Cancels the EC after determining that all required information is present.
    
    If Progress = "InPlanning" Then
        If IsNull(Me.ApprovalNotes) Then
            MsgBox ("There must be approval review notes describing why the EC was cancelled")
            End
        End If
    Else
        If IsNull(Me.ReviewNotes) Then
            MsgBox ("There must be review notes describing why the active EC was cancelled")
            End
        End If
    End If
    
    If Not IsNull(DLookup("AssociatedEC", "tECPartChange", "AssociatedEC=" & Me.ID)) Then
        If MsgBox("The part changed list must be cleared before rejecting the EC. A manual clearing will leave blank records " _
        & "that will still be attached. Would you like to clear all part change records attached to this EC? Only do after the changes " _
        & "have been undone", vbYesNo) = vbYes Then
            Dim qdf As DAO.QueryDef
            Set qdf = CurrentDb.QueryDefs("qRejectedECPartChangeDelete")
            qdf!RejectedEC = Me.ID
            qdf.Execute
            Set qdf = Nothing
        Else
            MsgBox ("EC must have part change list cleared before cancelling")
            End
        End If
    End If

    'Populates the EC with information
    Progress = "Cancelled"
    If IsNull(Me.DateCompleted) Then 'Date completed is only recorded once to improve metric tracking
        DateCompleted = Now
    End If
    
    'Updates the form to reflect the new status
    ProgressUpdate
    
    DoCmd.RunCommand acCmdSave
    
    'Generates a PDF of the cancelled EC
    GenerateECPDF False
    
    'Update the dashboard
    Call UpdateDashboard(ID, "EC", Progress)
    
    MsgBox ("EC has been cancelled.")
End Sub

Private Sub CloseForm_Click()
    'Purpose: Closes the form and makes sure all changes are saved.

    DoCmd.Save acForm, "fEC"
    DoCmd.Close acForm, "fEC"
    
    'Shifts focus to save the EC through the homepage
    Forms!fHomePage.SetFocus
    DoCmd.RunCommand acCmdSave
End Sub

Private Sub CompleteEC_Click()
    'Purpose: Completes EC. This will trickle down to also complete any associated ECRs and PRs to also mark them as complete. The pdf for the EC will be
    'generated and sent out to impacted parties.

    Dim EmailTo As String, Subject As String, Body As String, Attachment As String
    
    If Progress <> "Passed Review" Then
        If IsNull(Me.Reviewer) Or Not Me.Reviewer = Form_fHomePage.Initials Then
            MsgBox ("Only the Reviewer can pass this EC, please be sure to fill out the Reviewer field")
            End
        End If
    End If
    
    'Makes sure that the timeline is complete before allowing the completion of the EC
    If Not Form_fECsubTimeline.CompletedFullTrack() Then
        'Message box will be created by function
        End
    End If

    'Sends a email to the owner that the EC is complete, will only send if the completion is occuring during the review (which can't be done by owner)
    If Progress <> "Passed Review" Then
        If DLookup("RecieveEmails", "tUserData", "Initials = '" & Owner & "'") = True And Not Progress = "Complete" Then
            GenFunc.CreateEmailWithOutlook DLookup("Email", "tUserData", "Initials = '" & Owner & "'"), "EC#" & ID & " Complete", "EC#" & ID & " has been marked as complete."
        End If
    End If

    'Populates the EC with information
    Progress = "Complete"
    If IsNull(Me.DateCompleted) Then 'Date completed is only recorded once to improve metric tracking
        DateCompleted = Now
    End If
    
    'Updates the form to reflect the new status
    ProgressUpdate
    
    'Trickles down complete progress to all subordinate ECR and PRs. Warnings turned off to prevent multiple clicks required by user.
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qCompleteECtoECR"
    DoCmd.OpenQuery "qCompleteECRtoPR"
    DoCmd.SetWarnings True
    
    DoCmd.RunCommand acCmdSave
    
    'Generates a PDF of the completed EC
    GenerateECPDF False
    
    'Sends a email out to relevant parties containing the EC PDF
    EmailTo = "ProductMgmnt@WatsonFG.com; 1stShift@WatsonFG.com; 2ndShift@WatsonFG.com; productioncontrol@WatsonFG.com; qc@WatsonFG.com; jswan@WatsonFG.com; install@WatsonFG.com; mrienstra@WatsonFG.com; cbrandt@WatsonFG.com; OrderEngineeringGroup@WatsonFG.com;"
    Subject = "EC#" & ID & ": " & Description & " - Complete"
    Body = "EC#" & ID & " has been completed. See attached PDF for details and effectivity (date change will be complete for floor)."
    Attachment = "\\wfs.local\Watson\Engineering\03_Engineering\ECNs\EC_Secondary_Documents\" & ID & "\EC_Report_" & ID & ".pdf"
    GenFunc.CreateEmailWithOutlook EmailTo, Subject, Body, Attachment
    
    'Update the dashboard
    Call UpdateDashboard(ID, "EC", Progress)
    
    MsgBox ("EC has been marked as complete.")
End Sub

Private Sub CreateStdFolder_Click()
    'Purpose: Generates the secondary information folder for the EC if it does not exist

    CreateStandardFolder "EC", Me
End Sub

Private Sub Form_Open(Cancel As Integer)
    'Purpose: On form open makes sure that the most useful tab is moved to automatically

    Me.ActiveECInfo.SetFocus
End Sub

Private Sub MakeECActive_Click()
    'Purpose: Makes the EC active after determining that all required information is present

    'Determines that all required bits of information are filled out
    If IsNull(Description) Then
        MsgBox ("Description is required before making a EC Active")
        End
    ElseIf IsNull(History) Then
        MsgBox ("History is required before making a EC Active")
        End
    ElseIf IsNull(PlanOfAction) Then
        MsgBox ("Plan of Action is required before making a EC Active")
        End
    ElseIf ECType = "Full Track" Then
        'Makes sure that the timeline is fully approved
        If Not Form_fECsubTimeline.ApprovedFullTrack() Then
            'Message box will be created by function
            End
        End If
    End If
        
    'Populates the EC with information
    Progress = "Active"
    
    'Updates the form to reflect the new status
    ProgressUpdate
    
    DoCmd.RunCommand acCmdSave
    
    'Generates a PDF of the active EC
    GenerateECPDF False
    
    'Update the dashboard
    Call UpdateDashboard(ID, "EC", Progress)
    
    MsgBox ("EC has been set to Active.")
End Sub

Private Sub CreateECPDF_Click()
    'Purpose: Generates a PDF for the EC through the associated function.

    GenerateECPDF True
End Sub

Private Sub ReviewEC_Click()
    'Purpose: Puts the EC up for review after determining that all required information is present

    'Determines that all required bits of information are filled out
    If Not (Progress = "Active" Or Progress = "Failed Review") Then
        MsgBox ("The EC must be active before submitting it for review")
        End
    ElseIf IsNull(Effectivity) Then
        MsgBox ("Effectivity is required before submitting a EC for review, this can be a order or a date that the change will be made")
        End
    ElseIf IsNull(DLookup("AssociatedEC", "tECPartChange", "AssociatedEC=" & Me.ID)) Then
        If MsgBox("Are you sure there are no models that were changed for this EC?", vbYesNo) = vbNo Then
            End
        End If
    End If

    'Sends a email to the owner that the EC is back to review if the status change was made by not the owner
    If Not IsNull(Reviewer) Then
        If DLookup("RecieveEmails", "tUserData", "Initials = '" & Reviewer & "'") = True Then
            GenFunc.CreateEmailWithOutlook DLookup("Email", "tUserData", "Initials = '" & Reviewer & "'"), "EC#" & ID & " Back for Review", "EC#" & ID & " has been marked as up for review again. You are the previous reviewer."
        End If
    End If

    'Populates the EC with information
    Progress = "For Review"
    
    'Updates the form to reflect the new status
    ProgressUpdate
    
    DoCmd.RunCommand acCmdSave
    
    'Generates a PDF of the EC for review
    GenerateECPDF False
    
    'Update the dashboard
    Call UpdateDashboard(ID, "EC", Progress)
    
    MsgBox ("EC has been put up for review.")
End Sub

Private Sub PassReview_Click()
    'Purpose: Pass the EC review after determining that all required information is present

    If IsNull(Me.Reviewer) Or Not Me.Reviewer = Form_fHomePage.Initials Then
        MsgBox ("Only the Reviewer can pass this EC, please be sure to fill out the Reviewer field")
        End
    End If
    
    'Makes sure that the timeline is fully complete
    If Not Form_fECsubTimeline.CompletedFullTrack() Then
        'Message box will be created by function
        End
    End If

    'Sends a email to the owner that the EC has passed the review
    If DLookup("RecieveEmails", "tUserData", "Initials = '" & Owner & "'") = True And Not Progress = "Complete" Then
        GenFunc.CreateEmailWithOutlook DLookup("Email", "tUserData", "Initials = '" & Owner & "'"), "EC#" & ID & " Passed Review", "EC#" & ID & " has been marked as passed review."
    End If
    
    'Populates the EC with information
    Progress = "Passed Review"
    
    'Updates the form to reflect the new status
    ProgressUpdate
    
    DoCmd.RunCommand acCmdSave
    
    'Update the dashboard
    Call UpdateDashboard(ID, "EC", Progress)
    
    MsgBox ("EC has been marked as passed review.")
End Sub

Private Sub FailReview_Click()
    'Purpose: Marks the EC as having failed the review after determining that all required information is present

    'Determines that all required bits of information are filled out
    If IsNull(Reviewer) Then
        MsgBox ("Only the Reviewer can fail this EC, please be sure to fill out the Reviewer field")
        End
    ElseIf IsNull(ReviewNotes) Then
        MsgBox ("Review notes are required detailing why the review was failed")
        End
    End If

    'Sends a email to the owner that the EC has failed the review
    If DLookup("RecieveEmails", "tUserData", "Initials = '" & Owner & "'") = True Then
        GenFunc.CreateEmailWithOutlook DLookup("Email", "tUserData", "Initials = '" & Owner & "'"), "EC#" & ID & " Review Failed", "EC#" & ID & " has been marked as failed." & Chr(13) & Chr(10) & "Failure Reason: " & ReviewNotes
    End If

    'Populates the EC with information
    Progress = "Failed Review"
    
    'Updates the form to reflect the new status
    ProgressUpdate
    
    DoCmd.RunCommand acCmdSave
    
    'Update the dashboard
    Call UpdateDashboard(ID, "EC", Progress)
    
    MsgBox ("EC has been marked as failed review.")
End Sub

Private Sub ECType_AfterUpdate()
    'Purpose: Makes sure the form properly displays the information relevent for its status when the status changes
    
    ECTypeUpdate
End Sub

Private Sub Form_Current()
    'Purpose: Makes sure the form properly displays the information relevent for its status on opening
    
    ECTypeUpdate
    ProgressUpdate
End Sub

Private Sub ExpandedPartChange_Click()
    DoCmd.OpenForm "fExpandedPartChange", acNormal, "", "[ID]=" & [ID], , acNormal
End Sub

Private Sub ExpandedPartChangePCM_Click()
    DoCmd.OpenForm "fExpandedPartChangePCM", acNormal, "", "[ID]=" & [ID], , acNormal
End Sub

Private Sub Form_Resize()
    'Purpose: Makes sure that the form is properly sized for the screen. Some extreme screen resolutions are still not fully supported.

    Dim WorkingPageHeight As Single, WorkingPageWidth As Single, HistHeight As Single, PoAHeight As Single, RevHeight As Single
    
    'Undersizes the form elements to fit fully on the screen without needing the scrollbars. Numbers found through trial-and-error
    
    'Height Values
    WorkingPageHeight = Me.WindowHeight - 5200
    If WorkingPageHeight < 5000 Then
        WorkingPageHeight = 5000
    End If
    
    Me.History.Height = WorkingPageHeight * 0.73
    Me.TabList.Top = WorkingPageHeight * 0.73 + 1600
    Me.ApprovalNotesLabel.Top = WorkingPageHeight * 0.73 + 2200
    Me.PCMChanges.Top = WorkingPageHeight * 0.73 + 2200
    Me.PlanOfActionDeviationLabel.Top = WorkingPageHeight * 0.73 + 2200
    Me.PartChangeList.Top = WorkingPageHeight * 0.73 + 2200
    Me.ExpansionECRs.Top = WorkingPageHeight * 0.73 + 2200
    Me.subECRs.Top = 1
    
    'Width Values
    WorkingPageWidth = Me.WindowWidth - 4500
    If WorkingPageWidth < 16000 Then
        WorkingPageWidth = 16000
    End If
  
    Me.History.Width = (WorkingPageWidth / 2) - 1250
    Me.PlanOfAction.Width = (WorkingPageWidth / 2) - 1250
    Me.Timeline.Left = WorkingPageWidth - 2200
    Me.TabList.Width = 20000
    
    'Updates the form to reflect the new element sizings
    Me.Repaint
End Sub


