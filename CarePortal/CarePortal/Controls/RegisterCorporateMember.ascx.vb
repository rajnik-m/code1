Public Class RegisterCorporateMember
  Inherits CareWebControl
  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctRegisterCorporateMember, tblDataEntry)
      'Set the labels visible to false if enabled.
      SetLabelmessage("WarningMessage1", False)
      SetLabelmessage("WarningMessage2", False)
      Session("AddContactList") = Nothing
      If (Request.QueryString("ON") IsNot Nothing And Request.QueryString("AN") IsNot Nothing And Request.QueryString("EA") IsNot Nothing) AndAlso (Request.QueryString("ON").Length > 0 And Request.QueryString("AN").Length > 0 And Request.QueryString("EA").Length > 0) Then
        DirectCast(FindControlByName(Me, "Organisations"), DataGrid).Visible = False
        ProcessFindContacts(Request.QueryString("EA"), Request.QueryString("ON"), Request.QueryString("AN"))
      End If
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
        If GetTextBoxText("CompanyName").Length > 0 Then vList("Name") = GetTextBoxText("CompanyName") & "*"
        If GetTextBoxText("CorporateEmail").Length > 0 Then
          Dim vSplit As String()
          Dim vEmail As String = GetTextBoxText("CorporateEmail")
          If vEmail.Contains("@") Then
            vSplit = vEmail.Split(New Char() {"@"c})
            If vSplit.Length = 2 Then
              vEmail = vSplit(1)
              vList("EmailAddress") = "*@" & vEmail
            End If
          Else
            vList("EmailAddress") = vEmail
          End If
          If GetTextBoxText("Postcode").Length > 0 Then vList("Postcode") = GetTextBoxText("Postcode")
          vList("SystemColumns") = "Y"
          vList("WebPageItemNumber") = Me.WebPageItemNumber
          'Find Organisation which is a member that has matching Details
          Dim vTableOrg As DataTable = Nothing
          Dim vResult As String = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebMemberOrganisations, vList)
          vTableOrg = GetDataTable(vResult, True)
          'If No organisation is found
          Dim vDGR As DataGrid = TryCast(Me.FindControl("Organisations"), DataGrid)
          If Not vTableOrg Is Nothing Then
            Dim vContactNumbers As New List(Of String)
            Dim vRowDel As New List(Of Integer)
            For vRow As Integer = 0 To vTableOrg.Rows.Count - 1
              If vContactNumbers.Contains(vTableOrg.Rows(vRow).Item("ContactNumber").ToString) Then
                vRowDel.Add(vRow)
              Else
                vContactNumbers.Add(vTableOrg.Rows(vRow).Item("ContactNumber").ToString)
              End If
            Next
            For vItem As Integer = 0 To vRowDel.Count - 1
              vTableOrg.Rows(vRowDel.Item(0)).Delete()
            Next
          End If
          Dim vUrlText As String = String.Empty
          Dim vOrgNo As String
          Dim vAddressNo As String
          If Request.QueryString("ReturnURL") IsNot Nothing Then
            vUrlText = Request.QueryString("ReturnURL").ToString
          ElseIf WebPageNumber > 0 Then
            vUrlText = String.Format("default.aspx?pn={0}", WebPageNumber)
          End If
          If vTableOrg Is Nothing Then
            SetLabelmessage("WarningMessage1")
            'If more than one member organisation is found then list of organisation display in grid
          ElseIf vTableOrg.Rows.Count > 1 Then
            'Find grid and Populate 
            Dim vTableOrgAll As DataTable = Nothing
            vTableOrgAll = GetDataTable(vResult, True)
            If vDGR IsNot Nothing Then
              Dim vTableSelectData As DataTable = Nothing
              DataHelper.FillGrid(vResult, vDGR)
              Dim vColumn As New BoundColumn()
              vColumn.HeaderText = ""
              vDGR.Columns.AddAt(0, vColumn)
              vDGR.DataBind()
              Dim vSelectPos As Integer
              For vRow As Integer = 0 To vDGR.Items.Count - 1
                vOrgNo = vTableOrgAll.Rows(vRow).Item("ContactNumber").ToString
                vAddressNo = vTableOrgAll.Rows(vRow).Item("AddressNumber").ToString
                vDGR.Items(vRow).Cells(vSelectPos).Text = "<a href='" & vUrlText & "&ON=" & vOrgNo & "&AN=" & vAddressNo & "&EA=" & GetTextBoxText("CorporateEmail") & "'>Select</a>"
              Next
            End If
            'Hide the Other Controls 
            HideControls()
            ElseIf vTableOrg.Rows.Count = 1 Then
              vOrgNo = vTableOrg.Rows(0).Item("ContactNumber").ToString
              vAddressNo = vTableOrg.Rows(0).Item("AddressNumber").ToString
            ProcessRedirect(vUrlText & "&ON=" & vOrgNo & "&AN=" & vAddressNo & "&EA=" & GetTextBoxText("CorporateEmail") & "")
            End If
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

  Private Sub ProcessFindContacts(ByVal pEmailAddress As String, ByVal pOrganisationNumber As String, ByVal pAddressNumber As String)
    Dim vList As New ParameterList(HttpContext.Current)
    vList("EMailAddress") = pEmailAddress
    vList("OrganisationNumber") = pOrganisationNumber
    vList("Corporate") = "Y"
    Dim vResult As String = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftEMailContacts, vList)
    Dim vTablecCon As DataTable = Nothing
    vTablecCon = GetDataTable(vResult)
    If vTablecCon Is Nothing Then
      'If No Contact is Found 
      If Not DefaultParameters("RegisterPageNumber") Is Nothing Then
        ProcessRedirect("default.aspx?pn=" & DefaultParameters("RegisterPageNumber").ToString & "&ON=" & pOrganisationNumber & "&AN=" & pAddressNumber & "&EA=" & pEmailAddress)
      End If
      'If more than one Contact is found  with the given email address 
    ElseIf vTablecCon.Rows.Count > 1 Then
      SetLabelmessage("WarningMessage1")
      HideControls()
      'If Only one contact is found 
    ElseIf vTablecCon.Rows.Count = 1 Then
      vList("ContactNumber") = vTablecCon.Rows(0).Item("ContactNumber").ToString()
      vList("EmailAddress") = pEmailAddress
      Dim vTableRegisteredUsr As DataTable = Nothing
      vTableRegisteredUsr = GetDataTable(DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactRegisteredUsers, vList))
      If vTableRegisteredUsr Is Nothing Then
        'If there is no registered users record returned 
        vList("ContactNumber") = vTablecCon.Rows(0).Item("ContactNumber").ToString()
        vList("UserName") = vList("EmailAddress")
        vList("EmailAddress") = vList("EmailAddress")
        vList("Password") = GeneratePassword()
        DataHelper.AddRegisteredUser(vList)
        vList("Password") = PasswordEncrypted
        'Sending Mail if User is Registered
        MailingProcess(vList)
        'Redirect to Submit Page
        GoToSubmitPage()
      ElseIf vTableRegisteredUsr.Rows.Count = 1 Then
        SetLabelmessage("WarningMessage2")
        HideControls()
      End If
    End If
  End Sub
  Private Sub SetLabelmessage(ByVal pMessageControl As String, Optional ByVal pVisible As Boolean = True)
    If FindControlByName(Me, pMessageControl) IsNot Nothing Then
      DirectCast(FindControlByName(Me, pMessageControl), Label).Visible = pVisible
    End If
  End Sub

  Private Sub HideControls()
    If FindControlByName(Me, "MemberNumber") IsNot Nothing Then FindControlByName(Me, "MemberNumber").Parent.Parent.Visible = False
    If FindControlByName(Me, "CompanyName") IsNot Nothing Then FindControlByName(Me, "CompanyName").Parent.Parent.Visible = False
    If FindControlByName(Me, "CorporateEmail") IsNot Nothing Then FindControlByName(Me, "CorporateEmail").Parent.Parent.Visible = False
    If FindControlByName(Me, "Postcode") IsNot Nothing Then FindControlByName(Me, "Postcode").Parent.Parent.Visible = False
    If FindControlByName(Me, "Submit") IsNot Nothing Then FindControlByName(Me, "Submit").Parent.Parent.Visible = False
  End Sub
End Class