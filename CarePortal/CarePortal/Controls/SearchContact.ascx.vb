Public Class SearchContact
  Inherits CareWebControl
  Implements IMultiViewWebControl

  Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    Try
      If Not InWebPageDesigner() AndAlso Session("SelectedOrganisationNumber") Is Nothing Then
        If Not IsBackOfficeUser Then Throw New PortalAccessException
      End If
      InitialiseControls(CareNetServices.WebControlTypes.wctSearchContact, tblDataEntry)
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
        If Request.QueryString("CN") IsNot Nothing AndAlso Request.QueryString("CN").Length > 0 Then
          Dim vContactNumber As Integer = IntegerValue(Request.QueryString("CN"))
          Session("SelectedContactNumber") = vContactNumber
          GoToSubmitPage()
        End If
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
    Dim vHasError As Boolean = False
    Dim vShowNew As Boolean = False
    Dim vList As ParameterList = GetSearchCriteria()
    If vList.Count > 0 Then
      vList("SystemColumns") = "Y"
      vList("WebPageItemNumber") = Me.WebPageItemNumber
      'vList("StartRow") = "0" 'Rows will always start from 0 as paging is not yet supported
      vList("NumberOfRows") = InitialParameters("MaximumContacts")
      vList.AddConectionData(HttpContext.Current)
      Dim vGrd As DataGrid = CType(FindControlByName(tblDataEntry, "ContactData"), DataGrid)
      Dim vResult As String = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebContacts, vList)
      Dim vContactData As DataTable = GetDataTable(vResult, True)

      If vContactData IsNot Nothing AndAlso vContactData.Rows.Count > 0 Then
        If vContactData.Rows.Count > IntegerValue(InitialParameters("MaximumContacts").ToString) Then
          'Too many contacts match your criteria
          SetControlVisible("WarningMessage1", True)
          vHasError = True
        ElseIf vContactData.Rows.Count = 1 Then
          'Set the session and redirect the user to the next page
          Session("SelectedContactNumber") = vContactData.Rows(0)("ContactNumber")
          GoToSubmitPage()
        Else
          DataHelper.FillGrid(vResult, vGrd)
          Dim vUrlText As String = ""
          Dim vColumn As New BoundColumn()
          Dim vSelectPos As Integer
          Dim vContactNo As String
          If Request.QueryString("ReturnURL") IsNot Nothing Then
            vUrlText = Request.QueryString("ReturnURL").ToString
          ElseIf mvSubmitItemNumber > 0 Then
            vUrlText = String.Format("Default.aspx?pn={0}", WebPageNumber)
          End If
          vColumn.HeaderText = "Select"
          vGrd.Columns.AddAt(0, vColumn)
          vGrd.DataBind()
          If vUrlText.Length = 0 AndAlso vSelectPos >= 0 Then
            vGrd.Columns(vSelectPos).Visible = False
          Else
            For vRow As Integer = 0 To vGrd.Items.Count - 1
              vContactNo = vContactData.Rows(vRow).Item("ContactNumber").ToString
              vGrd.Items(vRow).Cells(vSelectPos).Text = "<a href='" & vUrlText & "&CN=" & vContactNo & "'>Select</a>"
            Next
          End If
          vGrd.Visible = True
          vShowNew = True
        End If
      Else
        SetControlVisible("WarningMessage2", True) 'no contacts match the criteria
        vShowNew = True
        vHasError = True
      End If
    Else
      SetControlVisible("WarningMessage3", True) 'no search criteria entered
      vHasError = True
    End If
    If InitialParameters.Contains("NewContactPageNumber") Then
      SetControlVisible("NewContact", vShowNew)
    End If
    If mvSupportsMultiView Then
      If vHasError Then
        mvMultiView.SetActiveView(mvView2)
      Else
        mvMultiView.SetActiveView(mvView1)
      End If
    End If
  End Sub

  Public Overrides Sub ProcessButtonClickEvent(ByVal pValues As Object)
    Try
      ' Retrieve values from object
      Dim vButtonID As String = TryCast(pValues, String)
      If vButtonID = "NewContact" Then ProcessRedirect(String.Format("Default.aspx?PN={0}", InitialParameters("NewContactPageNumber")))
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  Public Overrides Sub HandleDataGridEdit(ByVal e As DataGridCommandEventArgs)
    Try
      'Set the session and redirect the user to the next page
      Session("SelectedContactNumber") = e.CommandArgument
      GoToSubmitPage()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  ''' <summary>
  ''' Returns as parameter list for the non empty fields
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetSearchCriteria() As ParameterList
    Dim vList As New ParameterList()
    Dim vValue As String = String.Empty

    'Get values from textboxes
    Dim vFields As String() = {"Forenames", "Surname", "DateOfBirth", "MemberNumber", "EmailAddress", "NiNumber", "Town", "Postcode"}
    For Each vField As String In vFields
      vValue = GetTextBoxText(vField)
      If vValue.Length > 0 Then vList(vField) = vValue
    Next

    'Get dropdown values
    vFields = {"Title", "Sex"}
    For Each vField As String In vFields
      vValue = GetDropDownValue(vField)
      If vValue.Length > 0 Then vList(vField) = vValue
    Next
    Return vList
  End Function

  Private Sub HideControls()
    SetControlVisible("NewContact", False)
    SetControlVisible("ContactData", False)
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