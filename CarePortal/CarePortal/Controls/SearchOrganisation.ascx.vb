Public Class SearchOrganisation
  Inherits CareWebControl
  Implements IMultiViewWebControl

  Dim mvOrganisationData As New DataTable

  Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctSearchOrganisation, tblDataEntry)
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vEx As PortalAccessException
      Throw vEx
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      HideControls()
      If IsPostBack Then
        If SearchClicked() Then
          BindGrid()
        Else
          If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
        End If
      Else
        If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  Private Function SearchClicked() As Boolean
    Dim vControl As Control = FindControlByName(Me, "Search")
    If vControl IsNot Nothing Then
      Dim vNameID As String = FindControlByName(Me, "Search").UniqueID
      'Search through all the Keys in the Request.Form as it will contain the button id which raised the Click event
      For Each vName As String In Request.Form.Keys
        If vName = vNameID Then
          Return True
        End If
      Next
    End If
    Return False
  End Function

  Public Sub BindGrid()
    Try
      Dim vHasError As Boolean = False
      Dim vParams As New ParameterList(HttpContext.Current)
      If GetTextBoxText("OrganisationName").Trim.Length > 0 Then vParams("Name") = "*" & GetTextBoxText("OrganisationName") & "*"
      If GetTextBoxText("EMailAddress").Trim.Length > 0 Then
        Dim vSplit As String()
        Dim vSpiltChar As String = "@"
        Dim vEmail As String = GetTextBoxText("EMailAddress")
        If vEmail.Contains("@") Then
          vSplit = vEmail.Split(New Char() {"@"c})
          If vSplit.Length = 2 Then
            vEmail = vSplit(1)
            vParams("EmailAddress") = "*@" & vEmail
          End If
        End If
      End If
      If GetTextBoxText("WebAddress").Trim.Length > 0 Then vParams("WebAddress") = GetTextBoxText("WebAddress")
      If GetTextBoxText("Address").Trim.Length > 0 Then vParams("Address") = GetTextBoxText("Address")
      If GetTextBoxText("Town").Trim.Length > 0 Then vParams("Town") = GetTextBoxText("Town")
      If GetTextBoxText("Postcode").Trim.Length > 0 Then vParams("Postcode") = GetTextBoxText("Postcode")
      Dim vShowNew As Boolean = False
      If vParams.Count > 0 Then
        If InitialParameters.Contains("IncludeMemberOrganisations") Then vParams("NoMemberOrganisations") = InitialParameters("IncludeMemberOrganisations").ToString
        If InitialParameters.Contains("MaximumOrganisations") Then vParams("NumberOfRows") = InitialParameters("MaximumOrganisations").ToString
        vParams("SystemColumns") = "Y"
        vParams("PortalOrgSearch") = "Y"
        Dim vResult As String
        vResult = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftOrganisations, vParams)
        mvOrganisationData = GetDataTable(vResult, True)

        If mvOrganisationData IsNot Nothing AndAlso mvOrganisationData.Rows.Count > 0 Then
          If mvOrganisationData.Rows.Count > IntegerValue(InitialParameters("MaximumOrganisations").ToString) Then
            'Too many organisations match your criteria
            SetControlVisible("WarningMessage1", True)
            vHasError = True
          Else
            Dim vOrgDataGrid As DataGrid
            Dim vSelectPos As Integer
            Dim vColumn As New BoundColumn()
            Dim vOrgNo As String
            Dim vAddressNo As String
            Dim vUrlText As String = ""
            FindControlByName(Me, "OrganisationData").Visible = True
            vOrgDataGrid = CType(FindControlByName(Me, "OrganisationData"), DataGrid)
            DataHelper.FillGrid(vResult, vOrgDataGrid)
            If Request.QueryString("ReturnURL") IsNot Nothing Then
              vUrlText = Request.QueryString("ReturnURL").ToString
            ElseIf mvSubmitItemNumber > 0 Then
              vUrlText = String.Format("Default.aspx?pn={0}", mvSubmitItemNumber)
            End If
            vColumn.HeaderText = "Select"
            vOrgDataGrid.Columns.AddAt(0, vColumn)
            vOrgDataGrid.DataBind()
            If vUrlText.Length = 0 AndAlso vSelectPos >= 0 Then
              vOrgDataGrid.Columns(vSelectPos).Visible = False
            Else
              For vRow As Integer = 0 To vOrgDataGrid.Items.Count - 1
                vOrgNo = mvOrganisationData.Rows(vRow).Item("OrganisationNumber").ToString
                vAddressNo = mvOrganisationData.Rows(vRow).Item("AddressNumber").ToString
                vOrgDataGrid.Items(vRow).Cells(vSelectPos).Text = "<a href='" & vUrlText & "&ON=" & vOrgNo & "&AN=" & vAddressNo & "'>Select</a>"
              Next
            End If
            vShowNew = True
          End If
        Else
          'No Organisations match your criteria
          SetControlVisible("WarningMessage2", True)
          vShowNew = True
          vHasError = True
        End If
      Else
        'At least one item must be entered to search by
        SetControlVisible("WarningMessage3", True)
        vHasError = True
      End If
      If InitialParameters.Contains("NewOrganisationPageNumber") Then
        SetControlVisible("NewOrganisation", vShowNew)
      End If
      If mvSupportsMultiView Then
        If vHasError Then
          mvMultiView.SetActiveView(mvView2)
        Else
          mvMultiView.SetActiveView(mvView1)
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try

  End Sub

  Public Overrides Sub ProcessButtonClickEvent(ByVal pValues As Object)
    Try
      ' Retrieve values from object
      Dim vButtonID As String = TryCast(pValues, String)
      If vButtonID = "NewOrganisation" Then ProcessRedirect(String.Format("Default.aspx?PN={0}", InitialParameters("NewOrganisationPageNumber")))
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  Private Sub HideControls()
    SetControlVisible("NewOrganisation", False)
    SetControlVisible("OrganisationData", False)
    SetControlVisible("WarningMessage1", False)
    SetControlVisible("WarningMessage2", False)
    SetControlVisible("WarningMessage3", False)
  End Sub

  Protected Overrides Function MultiViewGridOnTop() As Boolean
    Return False
  End Function

  Public Function GridHyperLinkVisibility() As Boolean Implements IMultiViewWebControl.GridHyperLinkVisibility
    Return True
  End Function
End Class