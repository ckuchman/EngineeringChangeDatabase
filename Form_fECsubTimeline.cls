VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fECsubTimeline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public Function ApprovedFullTrack() As Boolean
    'Purpose: Determines if the full track timeline has been filled out and if it has been full approved and returns a boolean detailing if the
    'timeline is fully approved.

    'Checks if any approval is active other than the supervisor
    If ProductionControlApproval.Visible = False And ProductionControlApproval = False And _
    ProductManagementApproval.Visible = False And ProductManagementApproval = False And _
    ManufacturingEngineerApproval.Visible = False And ManufacturingEngineerApproval = False And _
    PCMApproval.Visible = False And PCMApproval = False And _
    ManufacturingApproval.Visible = False And ManufacturingApproval = False And _
    MarketingApproval.Visible = False And MarketingApproval = False And _
    OtherApproval1.Visible = False And OtherApproval1 = False And _
    OtherApproval2.Visible = False And OtherApproval2 = False And _
    OtherApproval3.Visible = False And OtherApproval3 = False Then
        If MsgBox("Are you sure that this Full Track EC requires no components on the timeline or approvals other than supervisor?", vbYesNo) = vbNo Then
            ApprovedFullTrack = False
            Exit Function
        End If
    End If
    
    'Checks if the visible approvals have all been completed
    If ProductionControlApproval.Visible = ProductionControlApproval And _
    ProductManagementApproval.Visible = ProductManagementApproval And _
    ManufacturingEngineerApproval.Visible = ManufacturingEngineerApproval And _
    PCMApproval.Visible = PCMApproval And _
    ManufacturingApproval.Visible = ManufacturingApproval And _
    MarketingApproval.Visible = MarketingApproval And _
    OtherApproval1.Visible = OtherApproval1 And _
    OtherApproval2.Visible = OtherApproval2 And _
    OtherApproval3.Visible = OtherApproval3 And _
    SupervisorApproval.Value = True Then
        ApprovedFullTrack = True
    Else
        MsgBox ("A Full Track EC can't become active until all approval checkboxes have been filled")
        ApprovedFullTrack = False
    End If
End Function


Public Function CompletedFullTrack() As Boolean
    'Purpose: Determines if the full track timeline has been completed
    
    If (UpdateVendorQuotes = "N/A" Or UpdateVendorQuotes = "Complete" Or UpdateVendorQuotes = "Cancelled") And _
    (InventoryDispositionCommunication = "N/A" Or InventoryDispositionCommunication = "Complete" Or InventoryDispositionCommunication = "Cancelled") And _
    (ArchiveModels = "N/A" Or ArchiveModels = "Complete" Or ArchiveModels = "Cancelled") And _
    (UpdateModels = "N/A" Or UpdateModels = "Complete" Or UpdateModels = "Cancelled") And _
    (ModelReview = "N/A" Or ModelReview = "Complete" Or ModelReview = "Cancelled") And _
    (PublishModels = "N/A" Or PublishModels = "Complete" Or PublishModels = "Cancelled") And _
    (ManufacturingReview = "N/A" Or ManufacturingReview = "Complete" Or ManufacturingReview = "Cancelled") And _
    (ModifyFixtures = "N/A" Or ModifyFixtures = "Complete" Or ModifyFixtures = "Cancelled") And _
    (LoadNewParts = "N/A" Or LoadNewParts = "Complete" Or LoadNewParts = "Cancelled") And _
    (UpdateRouting = "N/A" Or UpdateRouting = "Complete" Or UpdateRouting = "Cancelled") And _
    (WrenchWorkInProgress = "N/A" Or WrenchWorkInProgress = "Complete" Or WrenchWorkInProgress = "Cancelled") And _
    (ObsoleteOldPartNumbers = "N/A" Or ObsoleteOldPartNumbers = "Complete" Or ObsoleteOldPartNumbers = "Cancelled") And _
    (UpdateSOPs = "N/A" Or UpdateSOPs = "Complete" Or UpdateSOPs = "Cancelled") And _
    (UpdatingAssemblyInstructions = "N/A" Or UpdatingAssemblyInstructions = "Complete" Or UpdatingAssemblyInstructions = "Cancelled") And _
    (WatsonBulletin = "N/A" Or WatsonBulletin = "Complete" Or WatsonBulletin = "Cancelled") And _
    (UpdatingVisuals = "N/A" Or UpdatingVisuals = "Complete" Or UpdatingVisuals = "Cancelled") And _
    (OtherProgress1 = "N/A" Or OtherProgress1 = "Complete" Or OtherProgress1 = "Cancelled") And _
    (OtherProgress2 = "N/A" Or OtherProgress2 = "Complete" Or OtherProgress2 = "Cancelled") And _
    (OtherProgress3 = "N/A" Or OtherProgress3 = "Complete" Or OtherProgress3 = "Cancelled") Then
        CompletedFullTrack = True
    Else
        MsgBox ("A Full Track EC can't become completed until all tasks have been completed")
        CompletedFullTrack = False
    End If
End Function

Private Sub ProductionControlApprovalFunction()
    'Purpose: Checks approval and determines if it should be visible or not dependent on the status

    'Sets visiblity on the approval checkbox only if the step is required
    If UpdateVendorQuotes <> "N/A" Then
        ActiveApproval = True
    Else
        ActiveApproval = False
    End If
    
    'Toggles Visibility
    ProductionControlApproval.Visible = ActiveApproval
    ProductionControlApprovalLabel.Visible = ActiveApproval
End Sub

Private Sub ProductMgmtApprovalFunction()
    'Purpose: Checks approval and determines if it should be visible or not dependent on the status
    
    'Sets visiblity on the approval checkbox only if the step is required
    If InventoryDispositionCommunication <> "N/A" Or ArchiveModels <> "N/A" Or _
    UpdateModels <> "N/A" Or ModelReview <> "N/A" Or PublishModels <> "N/A" Then
        ActiveApproval = True
    Else
        ActiveApproval = False
    End If
    
    'Toggles Visibility
    ProductManagementApproval.Visible = ActiveApproval
    ProductManagementApprovalLabel.Visible = ActiveApproval
End Sub

Private Sub ManufacturingEngApprovalFunction()
    'Purpose: Checks approval and determines if it should be visible or not dependent on the status
    
    'Sets visiblity on the approval checkbox only if the step is required
    If ManufacturingReview <> "N/A" Or ModifyFixtures <> "N/A" Then
        ActiveApproval = True
    Else
        ActiveApproval = False
    End If
    
    'Toggles Visibility
    ManufacturingEngineerApproval.Visible = ActiveApproval
    ManufacturingEngineerApprovalLabel.Visible = ActiveApproval
End Sub

Private Sub PCMApprovalFunction()
    'Purpose: Checks approval and determines if it should be visible or not dependent on the status
    
    'Sets visiblity on the approval checkbox only if the step is required
    If LoadNewParts <> "N/A" Or UpdateRouting <> "N/A" Or _
    WrenchWorkInProgress <> "N/A" Or ObsoleteOldPartNumbers <> "N/A" Then
        ActiveApproval = True
    Else
        ActiveApproval = False
    End If
    
    'Toggles Visibility
    PCMApproval.Visible = ActiveApproval
    PCMApprovalLabel.Visible = ActiveApproval
End Sub

Private Sub ManufacturingApprovalFunction()
    'Purpose: Checks approval and determines if it should be visible or not dependent on the status
    
    'Sets visiblity on the approval checkbox only if the step is required
    If UpdateSOPs <> "N/A" Then
        ActiveApproval = True
    Else
        ActiveApproval = False
    End If
    
    'Toggles Visibility
    ManufacturingApproval.Visible = ActiveApproval
    ManufacturingApprovalLabel.Visible = ActiveApproval
End Sub

Private Sub MarketingApprovalFunction()
    'Purpose: Checks approval and determines if it should be visible or not dependent on the status
    
    'Sets visiblity on the approval checkbox only if the step is required
    If UpdatingAssemblyInstructions <> "N/A" Or WatsonBulletin <> "N/A" Or UpdatingVisuals <> "N/A" Then
        ActiveApproval = True
    Else
        ActiveApproval = False
    End If
    
    'Toggles Visibility
    MarketingApproval.Visible = ActiveApproval
    MarketingApprovalLabel.Visible = ActiveApproval
End Sub

Private Sub OtherApprovalFunction()
    'Purpose: Checks approval and determines if it should be visible or not dependent on the status

    'Sets visiblity on the approval checkbox only if the step is required
    If OtherProgress1 <> "N/A" Then
        ActiveApproval = True
    Else
        ActiveApproval = False
    End If
    
    'Toggles Visibility
    OtherApproval1.Visible = ActiveApproval
    OtherApprovalLabel1.Visible = ActiveApproval
    
    'Sets visiblity on the approval checkbox only if the step is required
    If OtherProgress2 <> "N/A" Then
        ActiveApproval = True
    Else
        ActiveApproval = False
    End If
    
    'Toggles Visibility
    OtherApproval2.Visible = ActiveApproval
    OtherApprovalLabel2.Visible = ActiveApproval
    
    'Sets visiblity on the approval checkbox only if the step is required
    If OtherProgress3 <> "N/A" Then
        ActiveApproval = True
    Else
        ActiveApproval = False
    End If
    
    'Toggles Visibility
    OtherApproval3.Visible = ActiveApproval
    OtherApprovalLabel3.Visible = ActiveApproval
End Sub

Private Sub Form_Current()
    'Purpose: Updates the form to have the correct elements visible

    ProductionControlApprovalFunction
    ProductMgmtApprovalFunction
    ManufacturingEngApprovalFunction
    PCMApprovalFunction
    ManufacturingApprovalFunction
    MarketingApprovalFunction
    OtherApprovalFunction
End Sub

Private Sub OtherProgress1_AfterUpdate()
OtherApprovalFunction
End Sub

Private Sub OtherProgress2_AfterUpdate()
OtherApprovalFunction
End Sub

Private Sub OtherProgress3_AfterUpdate()
OtherApprovalFunction
End Sub

Private Sub UpdateVendorQuotes_AfterUpdate()
ProductionControlApprovalFunction
End Sub

Private Sub InventoryDispositionCommunication_AfterUpdate()
ProductMgmtApprovalFunction
End Sub

Private Sub ArchiveModels_AfterUpdate()
ProductMgmtApprovalFunction
End Sub

Private Sub UpdateModels_AfterUpdate()
ProductMgmtApprovalFunction
End Sub

Private Sub ModelReview_AfterUpdate()
ProductMgmtApprovalFunction
End Sub

Private Sub PublishModels_AfterUpdate()
ProductMgmtApprovalFunction
End Sub

Private Sub ManufacturingReview_AfterUpdate()
ManufacturingEngApprovalFunction
End Sub

Private Sub ModifyFixtures_AfterUpdate()
ManufacturingEngApprovalFunction
End Sub

Private Sub LoadNewParts_AfterUpdate()
PCMApprovalFunction
End Sub

Private Sub UpdateRouting_AfterUpdate()
PCMApprovalFunction
End Sub

Private Sub WrenchWorkInProgress_AfterUpdate()
PCMApprovalFunction
End Sub

Private Sub ObsoleteOldPartNumbers_AfterUpdate()
PCMApprovalFunction
End Sub

Private Sub UpdateSOPs_AfterUpdate()
ManufacturingApprovalFunction
End Sub

Private Sub UpdatingAssemblyInstructions_AfterUpdate()
MarketingApprovalFunction
End Sub

Private Sub WatsonBulletin_AfterUpdate()
MarketingApprovalFunction
End Sub

Private Sub UpdatingVisuals_AfterUpdate()
MarketingApprovalFunction
End Sub

Private Sub OtherProgress_AfterUpdate()
OtherApprovalFunction
End Sub


