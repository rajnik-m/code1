Public Class SearchDirectory
  Inherits CareWebControl
  Implements IMultiViewWebControl

  Dim mvDirectoryData As New DataTable
  Dim mvActivity1Data As New DataTable
  Dim mvActivity2Data As New DataTable
  Dim mvActivity3Data As New DataTable
  Dim mvActivity4Data As New DataTable
  Dim mvActivity5Data As New DataTable
  Dim mvActivity6Data As New DataTable
  Dim mvDoSearch As Boolean = True

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctSearchDirectory, tblDataEntry, "", "")
      If FindControlByName(Me, "WarningMessage1") IsNot Nothing Then FindControlByName(Me, "WarningMessage1").Visible = False
      If FindControlByName(Me, "WarningMessage2") IsNot Nothing Then FindControlByName(Me, "WarningMessage2").Visible = False
      If FindControlByName(Me, "WarningMessage3") IsNot Nothing Then FindControlByName(Me, "WarningMessage3").Visible = False
      If FindControlByName(Me, "DirectoryData") IsNot Nothing Then FindControlByName(Me, "DirectoryData").Visible = False
      If FindControlByName(Me, "Activity1") IsNot Nothing Then FindControlByName(Me, "Activity1").Parent.Parent.Visible = False
      If FindControlByName(Me, "Activity2") IsNot Nothing Then FindControlByName(Me, "Activity2").Parent.Parent.Visible = False
      If FindControlByName(Me, "Activity3") IsNot Nothing Then FindControlByName(Me, "Activity3").Parent.Parent.Visible = False
      If FindControlByName(Me, "Activity4") IsNot Nothing Then FindControlByName(Me, "Activity4").Parent.Parent.Visible = False
      If FindControlByName(Me, "Activity5") IsNot Nothing Then FindControlByName(Me, "Activity5").Parent.Parent.Visible = False
      If FindControlByName(Me, "Activity6") IsNot Nothing Then FindControlByName(Me, "Activity6").Parent.Parent.Visible = False

      If DefaultParameters.Contains("Activity1") Then FillDataInListBox("Activity1")
      If DefaultParameters.Contains("Activity2") Then FillDataInListBox("Activity2")
      If DefaultParameters.Contains("Activity3") Then FillDataInListBox("Activity3")
      If DefaultParameters.Contains("Activity4") Then FillDataInListBox("Activity4")
      If DefaultParameters.Contains("Activity5") Then FillDataInListBox("Activity5")
      If DefaultParameters.Contains("Activity6") Then FillDataInListBox("Activity6")

      If IsPostBack OrElse Request.QueryString("PAGE") IsNot Nothing Then
        FindDirectories()
      End If
      HandleMultiViewDisplay("Search")
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If TryCast(sender, Control).ID = "Search" AndAlso IsValid() Then
      Try
        If FindControlByName(Me, "WarningMessage1") IsNot Nothing Then FindControlByName(Me, "WarningMessage1").Visible = False
        If FindControlByName(Me, "WarningMessage2") IsNot Nothing Then FindControlByName(Me, "WarningMessage2").Visible = False
        If FindControlByName(Me, "WarningMessage3") IsNot Nothing Then FindControlByName(Me, "WarningMessage3").Visible = False
        plcHolder.Visible = True
        FindDirectories()
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub

  Private Sub FindDirectories()
    Dim vParams As New ParameterList()
    Dim vIsValidSearch As Boolean = True
    mvDoSearch = True
    If Request.QueryString("PAGE") Is Nothing Then
      If Me.FindControl("MemberNumber") IsNot Nothing AndAlso GetTextBoxText("MemberNumber").Trim.Length > 0 Then vParams("MemberNumber") = GetTextBoxText("MemberNumber")
      If Me.FindControl("Forename") IsNot Nothing AndAlso GetTextBoxText("Forename").Trim.Length > 0 Then vParams("Forenames") = GetTextBoxText("Forename")
      If Me.FindControl("Surname") IsNot Nothing AndAlso GetTextBoxText("Surname").Trim.Length > 0 Then vParams("Surname") = GetTextBoxText("Surname")
      If Me.FindControl("NINumber") IsNot Nothing AndAlso GetTextBoxText("NINumber").Trim.Length > 0 Then vParams("NiNumber") = GetTextBoxText("NINumber")
      If Me.FindControl("ContactNumber") IsNot Nothing AndAlso GetTextBoxText("ContactNumber").Trim.Length > 0 Then vParams("ContactNumber") = GetTextBoxText("ContactNumber")
      If Me.FindControl("DateOfBirth") IsNot Nothing AndAlso GetTextBoxText("DateOfBirth").Trim.Length > 0 Then vParams("DateOfBirth") = GetTextBoxText("DateOfBirth")
      If Me.FindControl("Address") IsNot Nothing AndAlso GetTextBoxText("Address").Trim.Length > 0 Then vParams("Address") = Mid(GetTextBoxText("Address"), 1, 160)
      If Me.FindControl("Town") IsNot Nothing AndAlso GetTextBoxText("Town").Trim.Length > 0 Then vParams("Town") = GetTextBoxText("Town")
      If Me.FindControl("Postcode") IsNot Nothing AndAlso GetTextBoxText("Postcode").Trim.Length > 0 Then vParams("Postcode") = GetTextBoxText("Postcode")
      If Me.FindControl("Country") IsNot Nothing AndAlso DirectCast(Me.FindControl("Country"), DropDownList).SelectedValue.Trim.Length > 0 Then vParams("Country") = DirectCast(Me.FindControl("Country"), DropDownList).SelectedValue
      If DefaultParameters.Contains("Activity1") Then AddParams(vParams, "1")
      If DefaultParameters.Contains("Activity2") Then AddParams(vParams, "2")
      If DefaultParameters.Contains("Activity3") Then AddParams(vParams, "3")
      If DefaultParameters.Contains("Activity4") Then AddParams(vParams, "4")
      If DefaultParameters.Contains("Activity5") Then AddParams(vParams, "5")
      If DefaultParameters.Contains("Activity6") Then AddParams(vParams, "6")
      If vParams.Count > 0 AndAlso mvDoSearch Then
        vParams("WebPageItemNumber") = Me.WebPageItemNumber
        Session("DirectorySearch") = vParams
        vParams.AddConectionData(HttpContext.Current)
        If UserContactNumber() > 0 And UserAddressNumber() > 0 Then
          vParams("UserID") = UserContactNumber()
          vParams("AddressNumber") = UserAddressNumber()
        End If
        vParams("ContactType") = InitialParameters("ContactType").ToString
        vParams("ViewName") = InitialParameters("DirectoryName").ToString
      Else
        vIsValidSearch = False
      End If
    ElseIf Session("DirectorySearch") IsNot Nothing Then
      vParams = CType(Session("DirectorySearch"), ParameterList)
    End If
    Dim vHasError As Boolean = False
    If vIsValidSearch Then
      vParams("SystemColumns") = "Y"
      vParams("ViewDetails") = "N"
      Dim vCount As Long
      Dim vDirectoryDataList As BaseDataList
      If FindControlByName(Me, "DirectoryData") IsNot Nothing Then
        vDirectoryDataList = CType(FindControlByName(Me, "DirectoryData"), BaseDataList)
        vCount = DataHelper.GetPagedFinderData(CareNetServices.XMLDataFinderTypes.xdftWebDirectoryEntries, vDirectoryDataList, Request, plcHolder, vParams, IntegerValue(InitialParameters("ItemsPerPage").ToString), , False, )
        If (Not InitialParameters.ContainsKey("DisplayFormat")) OrElse InitialParameters("DisplayFormat").ToString = "0" Then
          Dim vDataGrid As DataGrid = CType(vDirectoryDataList, DataGrid)
          If vCount > 0 Then
            If vCount > IntegerValue(InitialParameters("MaximumRecords").ToString) Then
              If FindControlByName(Me, "WarningMessage1") IsNot Nothing Then FindControlByName(Me, "WarningMessage1").Visible = True
              plcHolder.Visible = False
              vHasError = True
            Else
              Dim vContactNumberPos As Integer
              Dim vColumn As New BoundColumn()
              Dim vContactNo As String
              Dim vUrlText As String = ""
              FindControlByName(Me, "DirectoryData").Visible = True
              If InitialParameters.Contains("DirectoryDetailsPageNumber") Then
                vUrlText = "default.aspx?pn=" & InitialParameters("DirectoryDetailsPageNumber").ToString & "&CN="
              End If
              If vDataGrid.Columns(0).HeaderText <> "" Then vDataGrid.Columns.AddAt(0, vColumn)
              vDataGrid.DataBind()
              For vColCount As Integer = 0 To vDataGrid.Columns.Count - 1
                If TypeOf vDataGrid.Columns(vColCount) Is BoundColumn Then
                  Dim vBoundColumn As BoundColumn = DirectCast(vDataGrid.Columns(vColCount), BoundColumn)
                  If vBoundColumn.DataField = "ContactNumber" Then
                    vContactNumberPos = vColCount
                  End If
                End If
              Next
              If vUrlText.Length = 0 AndAlso vContactNumberPos >= 0 Then
                vDataGrid.Columns(vContactNumberPos).Visible = False
              Else
                For vRow As Integer = 0 To vDataGrid.Items.Count - 1
                  vContactNo = vDataGrid.Items(vRow).Cells(vContactNumberPos).Text
                  If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
                    vDataGrid.Items(vRow).Cells(0).Text = String.Format("<input type=""button"" class=""Button"" value='{0}' onclick=""location.href='Default.aspx?pn={1}&CN={2}'"">", InitialParameters("HyperlinkText").ToString, InitialParameters("DirectoryDetailsPageNumber").ToString, vContactNo)
                  Else
                    vDataGrid.Items(vRow).Cells(0).Text = "<a href='" & vUrlText & vContactNo & "'>" & InitialParameters("HyperlinkText").ToString & "</a>"
                  End If
                Next
              End If
            End If
          Else
            If FindControlByName(Me, "WarningMessage2") IsNot Nothing Then FindControlByName(Me, "WarningMessage2").Visible = True
            vHasError = True
          End If
        End If
      End If
    Else
      If FindControlByName(Me, "WarningMessage3") IsNot Nothing Then FindControlByName(Me, "WarningMessage3").Visible = True
      vHasError = True
    End If
    If vHasError AndAlso mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
  End Sub
  Public Overrides Sub ClearControls(ByVal pClearLabels As Boolean)
    ClearControls(False, Nothing)
  End Sub
  Public Overrides Sub ClearControls(ByVal pClearLabels As Boolean, ByVal pErrorLabel As Label)
    ClearChildControls(tblDataEntry, True, False, pErrorLabel)
  End Sub

  Private Sub FillDataInListBox(ByVal pListBoxId As String)
    Dim vParamList As New ParameterList(HttpContext.Current)
    vParamList("Activity") = DefaultParameters(pListBoxId).ToString
    If FindControlByName(Me, pListBoxId) IsNot Nothing Then
      FindControlByName(Me, pListBoxId).Parent.Parent.Visible = True
      Dim vListBox As ListBox = DirectCast(FindControlByName(Me, pListBoxId), ListBox)
      DataHelper.FillList(CareNetServices.XMLLookupDataTypes.xldtActivityValues, vListBox, True, vParamList, True)
    Else
      Dim vData As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtActivityValues, vParamList)
      Select Case pListBoxId
        Case "Activity1"
          mvActivity1Data = vData
        Case "Activity2"
          mvActivity2Data = vData
        Case "Activity3"
          mvActivity3Data = vData
        Case "Activity4"
          mvActivity4Data = vData
        Case "Activity5"
          mvActivity5Data = vData
        Case "Activity6"
          mvActivity6Data = vData
      End Select

    End If
  End Sub

  Private Sub AddParams(ByRef pParams As ParameterList, ByVal pActivityNumber As String)
    Dim vParamStr As New StringBuilder
    If FindControlByName(Me, "Activity" & pActivityNumber) IsNot Nothing Then
      Dim vListBox As ListBox = DirectCast(FindControlByName(Me, "Activity" & pActivityNumber), ListBox)
      If vListBox.GetSelectedIndices.Length > 0 Then
        For Each index As Integer In vListBox.GetSelectedIndices
          If index <> 0 Then
            If vParamStr.Length > 0 Then
              vParamStr.Append(",")
              vParamStr.Append(vListBox.Items(index).Value)
            Else
              vParamStr.Append(vListBox.Items(index).Value)
            End If
          End If
        Next
        If vParamStr.ToString.Trim.Length > 0 Then
          pParams("Activity" & pActivityNumber) = vParamStr
          mvDoSearch = True
          If DefaultParameters.Contains("ItemSelectType" & pActivityNumber) Then
            pParams("ItemSelectType" & pActivityNumber) = DefaultParameters("ItemSelectType" & pActivityNumber)
          Else
            pParams("ItemSelectType" & pActivityNumber) = "ANY"
          End If
        Else
          If pParams.Count <= 0 Then mvDoSearch = False
        End If
      Else
        If pParams.Count <= 0 Then mvDoSearch = False
      End If
      If DefaultParameters.Contains("Activity" & pActivityNumber) Then
        pParams("Category" & pActivityNumber) = DefaultParameters("Activity" & pActivityNumber)
      End If
    Else
      If pParams.Count <= 0 Then mvDoSearch = False
      Dim vData As New DataTable
      Select Case pActivityNumber
        Case "1"
          vData = mvActivity1Data
        Case "2"
          vData = mvActivity2Data
        Case "3"
          vData = mvActivity3Data
        Case "4"
          vData = mvActivity4Data
        Case "5"
          vData = mvActivity5Data
        Case "6"
          vData = mvActivity6Data
      End Select
      If vData.Rows.Count > 0 Then
        Dim vRowCount As Integer = 0
        For Each vDr As DataRow In vData.Rows
          If vParamStr.Length > 0 Then
            vParamStr.Append(",")
            vParamStr.Append(vDr.Item("ActivityValue").ToString)
          Else
            vParamStr.Append(vDr.Item("ActivityValue").ToString)
          End If
          vRowCount = vRowCount + 1
        Next
        If vParamStr.ToString.Trim.Length > 0 Then
          pParams("Activity" & pActivityNumber) = vParamStr
          If DefaultParameters.Contains("ItemSelectType" & pActivityNumber) Then
            pParams("ItemSelectType" & pActivityNumber) = DefaultParameters("ItemSelectType" & pActivityNumber)
          Else
            pParams("ItemSelectType" & pActivityNumber) = "ANY"
          End If
          If DefaultParameters.Contains("Activity" & pActivityNumber) Then
            pParams("Category" & pActivityNumber) = DefaultParameters("Activity" & pActivityNumber)
          End If
        End If
      End If
    End If
  End Sub

  Public Overrides Sub HandleDataListItemDataBound(ByVal e As System.Web.UI.WebControls.DataListItemEventArgs)
    If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
      Dim vDirectoryDetailsPageNumber As String = InitialParameters("DirectoryDetailsPageNumber").ToString
      If vDirectoryDetailsPageNumber.Length > 0 Then
        'Add a select link at the end
        Dim vCount As Integer = e.Item.Controls.Count
        Dim vSelectLink As New Literal
        Dim vDrv As DataRowView = CType(e.Item.DataItem, DataRowView)
        If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
          vSelectLink.Text = String.Format("<input type=""button"" class=""Button"" value='{0}' onclick=""location.href='Default.aspx?pn={1}&CN={2}'"">", InitialParameters("HyperlinkText").ToString, InitialParameters("DirectoryDetailsPageNumber"), vDrv.Row("ContactNumber"))
        Else
          vSelectLink.Text = String.Format("<a href='Default.aspx?pn={0}&CN={1}'>{2}</a>", vDirectoryDetailsPageNumber, vDrv.Row("ContactNumber"), InitialParameters("HyperlinkText").ToString)
        End If
        If vCount > 0 Then e.Item.Controls(vCount - 1).Parent.Controls.Add(vSelectLink)
      End If
    End If
  End Sub

  Protected Overrides Function MultiViewGridOnTop() As Boolean
    Return False
  End Function

  Public Function GridHyperLinkVisibility() As Boolean Implements IMultiViewWebControl.GridHyperLinkVisibility
    Return True
  End Function
End Class
