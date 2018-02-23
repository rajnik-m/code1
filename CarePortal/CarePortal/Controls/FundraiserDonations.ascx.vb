Partial Public Class FundraiserDonations
  Inherits CareWebControl

  Private mvFundraisingNumber As Integer

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    InitialiseControls(CareNetServices.WebControlTypes.wctFundraiserDonations, tblDataEntry, "")
    Dim vList As New ParameterList(HttpContext.Current)
    vList("WebPageNumber") = WebPageNumber
    Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactFundraisingEvents, vList)
    Dim vDGR As DataGrid = CType(Me.FindControl("FundraisingDonations"), DataGrid)
    vList("SystemColumns") = "Y"
    If vTable IsNot Nothing Then
      mvFundraisingNumber = IntegerValue(vTable.Rows(0)("ContactFundraisingNumber").ToString)
      vList("ContactFundraisingNumber") = mvFundraisingNumber
      If IntegerValue(vTable.Rows(0)("ContactNumber").ToString) <> UserContactNumber() Then
        'User is not the Fundraiser so do not allow editing the page
        For Each vCareControl As CareWebControl In PageCareControls
          If vCareControl.WebPageItemName = "EditPage" Then vCareControl.Visible = False
        Next
      End If
    Else
      vList("DocumentColumns") = "Y"
    End If
    Dim vResult As String = DataHelper.SelectFundraisingEventData(CareNetServices.XMLFundraisingEventDataSelectionTypes.xfdtAnalysis, vList)
    DataHelper.FillGrid(vResult, vDGR, "")
  End Sub


End Class