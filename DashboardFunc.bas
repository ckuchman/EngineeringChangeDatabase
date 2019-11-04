Attribute VB_Name = "DashboardFunc"
Option Compare Database

Public Sub UpdateDashboard( _
    ID As Integer, _
    Stage As String, _
    Progress As String)

    'Starts with all three levels being zeroed out so the right level can be populated
    PRID = 0
    PRProgress = ""
    ECRID = 0
    ECRProgress = ""
    ECID = 0
    ECProgress = ""
    
    'Populates the level that called the function
    If Stage = "PR" Then
        PRID = ID
        PRProgress = Progress
    ElseIf Stage = "ECR" Then
        ECRID = ID
        ECRProgress = Progress
    ElseIf Stage = "EC" Then
        ECID = ID
        ECProgress = Progress
    Else
        MsgBox ("Not a valid stage")
    End If
    
    'Determines any higher level connections and populates the appropriate variables
    If PRID <> 0 Then
        If DLookup("AssociatedECR", "tProblemReport", "ID =" & PRID) <> 0 Then
            ECRID = DLookup("AssociatedECR", "tProblemReport", "ID =" & PRID)
            ECRProgress = DLookup("Progress", "tECR", "ID =" & ECRID)
        End If
    End If
    
    If ECRID <> 0 Then
        If DLookup("AssociatedEC", "tECR", "ID =" & ECRID) <> 0 Then
            ECID = DLookup("AssociatedEC", "tECR", "ID =" & ECRID)
            ECProgress = DLookup("Progress", "tEC", "ID =" & ECID)
        End If
    End If
    
    
    'Goes through the stages and determines the appropritate SQL statement to update the children PRs. The query being edited makes sure that
    'only the PR children that are not already rejected or that are attached to a rejected ECR will be editted. (This block of code controls the upstream
    'and the query controls the downstream)
    ItrStage = Stage 'Variable to allow the prgram to increment up through the stages without overwritting the information of what the initial
    DoCmd.SetWarnings False
    Do While FinalSQL = False
        If ItrStage = "PR" Then
            If PRProgress = "ForReview" Then
                FinalSQL = True
            ElseIf PRProgress = "Rejected" Then
                Reviewer = DLookup("Reviewer", "tProblemReport", "ID =" & PRID)
                ReviewNotesCleaned = Replace(DLookup("ReviewNotes", "tProblemReport", "ID =" & PRID), """", "''") 'Makes sure " does not cause a SQL injection
                DoCmd.RunSQL "UPDATE qODBCwithPR " & _
                             "SET ECRStatusID = 3, Change = """ & ReviewNotesCleaned & """, Approved = 0, ClosedBy = """ & Reviewer & """, DateClosed = Now(), InvConcern = 'See ECR' " & _
                             "WHERE ID =" & PRID
                             
                FinalSQL = True
            ElseIf PRProgress = "ECR Created" Or PRProgress = "Complete" Then
                'Updates the dashboard with base information that is built on with the ECR/EC
                Reviewer = DLookup("Reviewer", "tProblemReport", "ID =" & PRID)
                DoCmd.RunSQL "UPDATE qODBCwithPR " & _
                             "SET ECRStatusID = 3, Change = 'ECR#' & """ & ECRID & """ & ' In Progress', Approved = -1, ClosedBy = """ & Reviewer & """, DateClosed = Now(), InvConcern = 'See ECR' " & _
                             "WHERE ID =" & PRID
                
                'The PR has a higher level so it this will loop over the next levels
                ItrStage = "ECR"
            Else
                FinalSQL = True
            End If
        ElseIf ItrStage = "ECR" Then
            If ECRProgress = "In Progress" Then
                OwnerCleaned = Replace(DLookup("Owner", "tECR", "ID =" & ECRID), """", "''") 'Makes sure " does not cause a SQL injection
                DoCmd.RunSQL "UPDATE qODBCwithECR " & _
                             "SET Change = 'ECR#' & """ & ECRID & """ & ' In Progress Owner: ' & """ & OwnerCleaned & """" & _
                             "WHERE ID =" & ECRID
            
                FinalSQL = True
            ElseIf ECRProgress = "Pending Implementation" Then
                ProposalCleaned = Replace(DLookup("Proposal", "tECR", "ID =" & ECRID), """", "''") 'Makes sure " does not cause a SQL injection
                OwnerCleaned = Replace(DLookup("Owner", "tECR", "ID =" & ECRID), """", "''") 'Makes sure " does not cause a SQL injection
                DoCmd.RunSQL "UPDATE qODBCwithECR " & _
                             "SET Change = 'ECR#' & """ & ECRID & """ & ' Pending Implementation Owner: ' & """ & OwnerCleaned & """ & ' Proposal: ' & """ & ProposalCleaned & """" & _
                             "WHERE ID =" & ECRID
                             
                FinalSQL = True
            ElseIf ECRProgress = "Implement On Order" Then
                ProposalCleaned = Replace(DLookup("Proposal", "tECR", "ID =" & ECRID), """", "''") 'Makes sure " does not cause a SQL injection
                OwnerCleaned = Replace(DLookup("Owner", "tECR", "ID =" & ECRID), """", "''") 'Makes sure " does not cause a SQL injection
                DoCmd.RunSQL "UPDATE qODBCwithECR " & _
                             "SET Change = 'ECR#' & """ & ECRID & """ & ' Implementing on Order Owner: ' & """ & OwnerCleaned & """ & ' Proposal: ' & """ & ProposalCleaned & """" & _
                             "WHERE ID =" & ECRID
                
                FinalSQL = True
            ElseIf ECRProgress = "Rejected" Then
                DoCmd.RunSQL "UPDATE qODBCwithECR " & _
                             "SET Change = 'ECR#' & """ & ECRID & """ & ' Rejected, PR will either be attached to another ECR or rejected'" & _
                             "WHERE [ID] =" & ECRID
                             
                FinalSQL = True
            ElseIf ECRProgress = "EC Created" Or ECRProgress = "Complete" Then
                ProposalCleaned = Replace(DLookup("Proposal", "tECR", "ID =" & ECRID), """", "''") 'Makes sure " does not cause a SQL injection
                OwnerCleaned = Replace(DLookup("Owner", "tECR", "ID =" & ECRID), """", "''") 'Makes sure " does not cause a SQL injection
                DoCmd.RunSQL "UPDATE qODBCwithECR " & _
                             "SET Change = 'EC#' & """ & ECID & """ & ' In Planning Owner: ' & """ & OwnerCleaned & """ & ' Proposal: ' & """ & ProposalCleaned & """" & _
                             "WHERE [ID] =" & ECRID
                
                'The ECR has a higher level so it this will loop over the next levels
                ItrStage = "EC"
            Else
                FinalSQL = True
            End If
        ElseIf ItrStage = "EC" Then
            If ECProgress = "In Planning" Then
                FinalSQL = True
            ElseIf ECProgress = "Active" Then
                PlanOfActionCleaned = Replace(DLookup("PlanOfAction", "tEC", "ID =" & ECID), """", "''") 'Makes sure " does not cause a SQL injection
                OwnerCleaned = Replace(DLookup("Owner", "tEC", "ID =" & ECID), """", "''") 'Makes sure " does not cause a SQL injection
                DoCmd.RunSQL "UPDATE qODBCwithEC " & _
                             "SET Change = 'EC#' & """ & ECID & """ & ' Active Owner: ' & """ & OwnerCleaned & """ & ' Plan Of Action: ' & """ & PlanOfActionCleaned & """" & _
                             "WHERE [ID] =" & ECID
                
                FinalSQL = True
            ElseIf ECProgress = "For Review" Or ECProgress = "Failed Review" Or ECProgress = "Passed Review" Then
                PlanOfActionCleaned = Replace(DLookup("PlanOfAction", "tEC", "ID =" & ECID), """", "''") 'Makes sure " does not cause a SQL injection
                OwnerCleaned = Replace(DLookup("Owner", "tEC", "ID =" & ECID), """", "''") 'Makes sure " does not cause a SQL injection
                EffectivityCleaned = Replace(DLookup("Effectivity", "tEC", "ID =" & ECID), """", "''") 'Makes sure " does not cause a SQL injection
                DoCmd.RunSQL "UPDATE qODBCwithEC " & _
                             "SET Change = 'EC#' & """ & ECID & """ & ' Up for Review Owner: ' & """ & OwnerCleaned & """ & ' Effectivity: ' & """ & EffectivityCleaned & """ & ' Plan Of Action: ' & """ & PlanOfActionCleaned & """" & _
                             "WHERE [ID] =" & ECID
                
                FinalSQL = True
            ElseIf ECProgress = "Cancelled" Then
                DoCmd.RunSQL "UPDATE qODBCwithEC " & _
                             "SET Change = 'EC#' & """ & ECID & """ & ' Cancelled, ECR will either be attached to another EC or rejected'" & _
                             "WHERE [ID] =" & ECID
                
                FinalSQL = True
            ElseIf ECProgress = "Complete" Then
                PlanOfActionCleaned = Replace(DLookup("PlanOfAction", "tEC", "ID =" & ECID), """", "''") 'Makes sure " does not cause a SQL injection
                OwnerCleaned = Replace(DLookup("Owner", "tEC", "ID =" & ECID), """", "''") 'Makes sure " does not cause a SQL injection
                EffectivityCleaned = Replace(DLookup("Effectivity", "tEC", "ID =" & ECID), """", "''") 'Makes sure " does not cause a SQL injection
                DoCmd.RunSQL "UPDATE qODBCwithEC " & _
                             "SET Change = 'EC#' & """ & ECID & """ & ' Complete Owner: ' & """ & OwnerCleaned & """ & ' Effectivity: ' & """ & EffectivityCleaned & """ & ' Plan Of Action: ' & """ & PlanOfActionCleaned & """" & _
                             "WHERE [ID] =" & ECID
                
                FinalSQL = True
            Else
                FinalSQL = True
            End If
            
            FinalSQL = True 'It will need to terminate at the EC level
        Else
            MsgBox ("Iteration stage is not valid: " & ItrStage)
            FinalSQL = True
        End If
    Loop
    DoCmd.SetWarnings True
End Sub
