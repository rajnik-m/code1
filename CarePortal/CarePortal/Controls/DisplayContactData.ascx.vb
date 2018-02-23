Partial Public Class DisplayContactData
  Inherits CareWebControl
  Implements ICareChildWebControl

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    mvNeedsParent = True
    mvUsesHiddenContactNumber = True
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctDisplayContactData, tblDataEntry)
      Dim vDGR As DataGrid = CType(Me.FindControl("ContactData"), DataGrid)
      Dim vList As New ParameterList(HttpContext.Current)
      vList("DocumentColumns") = "Y"
      vList("SystemColumns") = "Y"
      vList("CarePortal") = "Y"
      vList.Add("WebPageItemNumber", Me.WebPageItemNumber)
      If Not vList.Contains("WPD") Then vList.Add("WPD", "Y")
      Dim vResult As String = DataHelper.SelectContactData(DirectCast(IntegerValue(InitialParameters("DataSelectionType").ToString), CareNetServices.XMLContactDataSelectionTypes), vList)
      DataHelper.FillGrid(vResult, vDGR, "")
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Public Overrides Sub ProcessContactSelection(ByVal pTable As DataTable)
    Dim vRow As DataRow = DataHelper.GetRowFromDataTable(pTable)
    Dim vContactNumber As Integer = IntegerValue(vRow("ContactNumber").ToString)
    Dim vRestriction As String = ""
    Dim vDGR As DataGrid = CType(Me.FindControl("ContactData"), DataGrid)
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vEditPageNumber As Integer = 0
    If DirectCast(IntegerValue(InitialParameters("DataSelectionType").ToString), CareNetServices.XMLContactDataSelectionTypes) = CareNetServices.XMLContactDataSelectionTypes.xcdtContactActions Then
      For Each vCareWebControl As CareWebControl In PageCareControls
        If TypeOf (vCareWebControl) Is AddAction Then vEditPageNumber = Me.WebPageNumber
      Next
      vList("IgnoreStatus") = "Y"
    End If
    If InitialParameters.ContainsKey("DataFilter") AndAlso InitialParameters("DataFilter").ToString.Length > 0 Then
      vRestriction = InitialParameters("DataFilter").ToString
      If vRestriction.Contains("+") Then vRestriction = vRestriction.Replace("+", ",")
      If vRestriction.Contains("^") Then vRestriction = vRestriction.Replace("^", "=")
    End If
    vDGR.Columns.Clear()
    vList.Add("WebPageItemNumber", Me.WebPageItemNumber)
    If Not vList.Contains("WPD") Then vList.Add("WPD", "Y")
    DataHelper.GetContactData(DirectCast(IntegerValue(InitialParameters("DataSelectionType").ToString), CareNetServices.XMLContactDataSelectionTypes), vDGR, vRestriction, vContactNumber, , vList, vEditPageNumber)
  End Sub

  Public Sub SubmitChild(ByVal pList As ParameterList) Implements ICareChildWebControl.SubmitChild
    'Nothing to do as this is a display only control
  End Sub
End Class