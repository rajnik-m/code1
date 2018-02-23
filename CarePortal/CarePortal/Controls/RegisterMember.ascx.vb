Public Class RegisterMember
  Inherits CareWebControl

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctRegisterMember, tblDataEntry)
      'Set the labels visible to false if enabled.
      SetLabelmessage("WarningMessage1", False)
      SetLabelmessage("WarningMessage2", False)
      Session("AddContactList") = Nothing
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      If IsValid() Then
        Dim vList As New ParameterList(HttpContext.Current)
        If GetTextBoxText("MemberNumber").Length > 0 Then vList("MemberNumber") = GetTextBoxText("MemberNumber")
        If GetTextBoxText("Surname").Length > 0 Then vList("Surname") = GetTextBoxText("Surname")
        If GetTextBoxText("DateofBirth").Length > 0 Then vList("DateOfBirth") = GetTextBoxText("DateofBirth")
        If GetTextBoxText("EmailAddress").Length > 0 Then vList("EmailAddress") = GetTextBoxText("EmailAddress")
        If GetTextBoxText("Postcode").Length > 0 Then vList("Postcode") = GetTextBoxText("Postcode")
        vList("Current") = "Y"
        vList("ContactType") = "C"
        Dim vTable As DataTable = Nothing
        vTable = GetDataTable(DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftMembers, vList))
        If Not vTable Is Nothing Then
          Dim vContactNumbers As New List(Of String)
          Dim vRowDel As New List(Of Integer)
          For vRow As Integer = 0 To vTable.Rows.Count - 1
            If vContactNumbers.Contains(vTable.Rows(vRow).Item("ContactNumber").ToString) Then
              vRowDel.Add(vRow)
            Else
              vContactNumbers.Add(vTable.Rows(vRow).Item("ContactNumber").ToString)
            End If
          Next
          For vItem As Integer = 0 To vRowDel.Count - 1
            vTable.Rows(vRowDel.Item(0)).Delete()
          Next
        End If
        If vTable Is Nothing Then
          SetLabelmessage("WarningMessage1")
        ElseIf vTable.Rows.Count = 1 Then
          Dim vTableSelectData As DataTable = Nothing
          vList("ContactNumber") = vTable.Rows(0).Item("ContactNumber").ToString()
          vTableSelectData = GetDataTable(DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactRegisteredUsers, vList))
          If Not vTableSelectData Is Nothing Then
            SetLabelmessage("WarningMessage2")
          Else
            vList("UserName") = vList("EmailAddress")
            vList("EmailAddress") = vList("EmailAddress")
            vList("Password") = GeneratePassword()
            DataHelper.AddRegisteredUser(vList)
            vList("Password") = PasswordEncrypted
            'Sending Mail if User is Registered
            MailingProcess(vList)
            'Redirect to Submit Page
            GoToSubmitPage()
          End If
        ElseIf vTable.Rows.Count > 1 Then
          SetLabelmessage("WarningMessage1")
        End If
      End If
    Catch vEx As CareException
      ProcessError(vEx)
    End Try
  End Sub

  Private Sub MailingProcess(ByVal pList As ParameterList)
    ' Password mailing Process
    Dim vContentParams As New ParameterList
    vContentParams("EMail") = pList("EmailAddress")
    vContentParams("Password") = pList("Password")
    'Default Parameters Set from WPD
    Dim vEmailParams As New ParameterList(HttpContext.Current)
    vEmailParams("StandardDocument") = DefaultParameters("StandardDocument")
    vEmailParams("EMailAddress") = DefaultParameters("EMailAddress")
    vEmailParams("Name") = DefaultParameters("Name")
    DataHelper.ProcessBulkEMail(vContentParams.ToCSVFile, vEmailParams, True)
    ' Password mailing Process End
  End Sub

  Private Sub SetLabelmessage(ByVal pMessageControl As String, Optional ByVal pVisible As Boolean = True)
    If FindControlByName(Me, pMessageControl) IsNot Nothing Then
      DirectCast(FindControlByName(Me, pMessageControl), Label).Visible = pVisible
    End If
  End Sub
End Class