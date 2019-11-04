VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fHomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Function HomePageFilter()
    'Purpose: Filters the home page to be setup for the active user. The user is determined from the windows profile that the user is using.

    'Turns off the visible update of the filters to heavily spead up the process of the update and make it seem more seamless
    Application.Echo False
    
    'Pulls the username from the windows account
    LocalUserPath = Environ("username")

    'Determines if the user has used this database before and sets up a account with an associated initials (that need to be filled out)
    If DCount("UserPath", "tUserData", "UserPath = '" & LocalUserPath & "'") = 0 Then
        DoCmd.GoToRecord , , acNewRec
        UserPath = LocalUserPath
        Initials = "???"
    Else
        Form.FilterOn = False
        Form.Filter = "UserPath = '" & LocalUserPath & "'"
        Form.FilterOn = True
        
        'Runs through all the sublists and sets filters based on the specification of the list
        PersonalPRForm.Form.FilterOn = False
        PersonalPRForm.Form.Filter = "([Creator] = '" & Initials & "' And ([Progress] = 'In Progress' Or [Progress] = 'Failed Review')) Or " & _
                                           "([Reviewer] = '" & Initials & "' And ([Progress] = 'In Progress' Or [Progress] = 'For Review'))"
        PersonalPRForm.Form.FilterOn = True
        PersonalPRForm.Form.OrderByOn = False
        PersonalPRForm.Form.OrderBy = "ID DESC"
        PersonalPRForm.Form.OrderByOn = True

        PersonalECRForm.Form.FilterOn = False
        PersonalECRForm.Form.Filter = "[Owner] = '" & Initials & "' And ([Progress] = 'In Progress')"
        PersonalECRForm.Form.FilterOn = True
        PersonalECRForm.Form.OrderByOn = False
        PersonalECRForm.Form.OrderBy = "ID DESC"
        PersonalECRForm.Form.OrderByOn = True

        PersonalECForm.Form.FilterOn = False
        PersonalECForm.Form.Filter = "([Owner] = '" & Initials & "' And ([Progress] = 'In Planning' Or [Progress] = 'Active' Or [Progress] = 'Failed Review' Or [Progress] = 'Passed Review' Or [Progress] = 'For Review')) Or " & _
                                           "([Reviewer] = '" & Initials & "' And [Progress] = 'For Review')"
        PersonalECForm.Form.FilterOn = True
        PersonalECForm.Form.OrderByOn = False
        PersonalECForm.Form.OrderBy = "ID DESC"
        PersonalECForm.Form.OrderByOn = True
        
        UnclaimedPRForm.Form.FilterOn = False
        UnclaimedPRForm.Form.Filter = "([Reviewer] Not Like '*' And [Progress] = 'For Review')"
        UnclaimedPRForm.Form.FilterOn = True
        UnclaimedPRForm.Form.OrderByOn = False
        UnclaimedPRForm.Form.OrderBy = "ID DESC"
        UnclaimedPRForm.Form.OrderByOn = True

        ImplementableECRForm.Form.FilterOn = False
        ImplementableECRForm.Form.Filter = "([Owner] Not Like '*' And [Progress] = 'In Progress') Or [Progress] = 'Pending Implementation' Or [Progress] = 'Implement On Order'"
        ImplementableECRForm.Form.FilterOn = True
        ImplementableECRForm.Form.OrderByOn = False
        ImplementableECRForm.Form.OrderBy = "ID DESC"
        ImplementableECRForm.Form.OrderByOn = True

        ForReviewECForm.Form.FilterOn = False
        ForReviewECForm.Form.Filter = "([Reviewer] Not Like '*' And [Progress] = 'For Review')"
        ForReviewECForm.Form.FilterOn = True
        ForReviewECForm.Form.OrderByOn = False
        ForReviewECForm.Form.OrderBy = "ID DESC"
        ForReviewECForm.Form.OrderByOn = True
        
        PRForm.Form.FilterOn = False
        PRForm.Form.OrderByOn = False
        PRForm.Form.OrderBy = "ID DESC"
        PRForm.Form.OrderByOn = True
        
        ECRForm.Form.FilterOn = False
        ECRForm.Form.OrderByOn = False
        ECRForm.Form.OrderBy = "ID DESC"
        ECRForm.Form.OrderByOn = True
        
        ECForm.Form.FilterOn = False
        ECForm.Form.OrderByOn = False
        ECForm.Form.OrderBy = "ID DESC"
        ECForm.Form.OrderByOn = True
    End If
    
    'Turns the echo back on so that any filtering done by the user will be visible
    Application.Echo True
End Function

Function HomePageResize()
    'Purpose: Makes sure that the form is properly sized for the screen. Some extreme screen resolutions are still not fully supported.

    'Determines the size of the Access window
    'The panel breaks if driven too large (multi-monitors) so this keeps it from erroring out
    If Abs(Me.WindowWidth) < 32000 Then
        HomePageWidth = Abs(Me.WindowWidth)
    Else
        HomePageWidth = 32000
    End If
    HomePageHeight = Me.WindowHeight

    'Undersizes the form to fit fully on the screen without needing the scrollbars. Numbers found through trial-and-error
    FormWidth = HomePageWidth * 0.96
    FormHeight = HomePageHeight * 0.81

    PersonalPRForm.Width = FormWidth
    PersonalPRForm.Height = FormHeight
    PersonalECRForm.Width = FormWidth
    PersonalECRForm.Height = FormHeight
    PersonalECForm.Width = FormWidth
    PersonalECForm.Height = FormHeight
    
    UnclaimedPRForm.Width = FormWidth
    UnclaimedPRForm.Height = FormHeight
    ImplementableECRForm.Width = FormWidth
    ImplementableECRForm.Height = FormHeight
    ForReviewECForm.Width = FormWidth
    ForReviewECForm.Height = FormHeight
    
    PRForm.Width = FormWidth
    PRForm.Height = FormHeight
    ECRForm.Width = FormWidth
    ECRForm.Height = FormHeight
    ECForm.Width = FormWidth
    ECForm.Height = FormHeight
    
    PersonalPR.Width = HomePageWidth * 0.975
    PersonalPR.Height = HomePageHeight * 0.85

    'Regenerates the window image to reflect the new sizing
    Me.Repaint
End Function

Private Sub OpenUserSettings_Click()
    DoCmd.OpenForm "fUserSettings", , , "UserPath='" & Environ("username") & "'"
End Sub

Private Sub Form_Open(Cancel As Integer)
    'Purpose: On the opening of the form, runs the filters and resizing to reflect the user and their screen size.

    Refresh_Click
    HomePageFilter
    HomePageResize
End Sub

Private Sub Form_Resize()
    HomePageResize
End Sub

Private Sub PartChangeSearch_Click()
    DoCmd.OpenForm "fSearchPartChange"
End Sub

Private Sub PrintOptions_Click()
    DoCmd.OpenForm "fPrintOptions"
End Sub

Private Sub Refresh_Click()
    
    Dim dbs As DAO.Database
    Dim rsQuery As DAO.Recordset
    Dim WStext As String
    Dim EmailTo As String, Subject As String, textBody As String
    
    Set dbs = CurrentDb

    'Opens a record set of the new entries on the dashboard to be converted to PRs
    Set rsQuery = dbs.OpenRecordset("qNewOBDCEntries", dbOpenDynaset, _
          dbFailOnError + dbSeeChanges)
          
    counter = 0
    textBody = "Work stoppage PR(s) have been submitted:" + vbCrLf
    
    'Cycles through the new entries and adds the issues to the email
    With rsQuery
        Do While Not .EOF
            If rsQuery![ECRUrgencyID] = 1 Then
                counter = counter + 1
                textBody = textBody + "Issue: " + rsQuery![Subject] + vbCrLf + vbCrLf
            End If
            .MoveNext
        Loop
    End With
    
    'Sends a email out alerting order engineering to the new PR
    If counter > 0 Then
        EmailTo = "ckuchman@WatsonFG.com; orderengineeringgroup@WatsonFG.com"
        Subject = "Work Stoppage Submitted"
        GenFunc.CreateEmailWithOutlook EmailTo, Subject, textBody
    End If
    
    'Creates a 'PR' in the database to link with the new one on the dashboard.
    With DoCmd
         .SetWarnings False
         .OpenQuery "qAddNewOBDCEntries"
         .SetWarnings True
    End With

    Me.Requery
    
    Me.PersonalPRForm.Form.Requery
    Me.PersonalECRForm.Form.Requery
    Me.PersonalECForm.Form.Requery
    
    Me.UnclaimedPRForm.Form.Requery
    Me.ImplementableECRForm.Form.Requery
    Me.ForReviewECForm.Form.Requery
    
    Me.PRForm.Form.Requery
    Me.ECRForm.Form.Requery
    Me.ECForm.Form.Requery
    
    HomePageFilter
End Sub

Private Sub NewECR_Click()
    'Purpose: Creates a new ECR and populates it with a owner

    Dim ECRLink As Integer

    'Opens the ECR form and goes to a new record to populate
    DoCmd.OpenForm "fECR", acNormal, "", "", , acNormal
    DoCmd.GoToRecord , "", acNewRec

    'Runs a method that populates the ECR with relevent starting information
    ECRLink = Form_fECR.CreateECR
End Sub
