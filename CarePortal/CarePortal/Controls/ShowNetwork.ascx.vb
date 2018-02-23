Partial Public Class ShowNetwork
  Inherits CareWebControl

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctShowMyNetwork, tblDataEntry)
      Dim vDGR As DataGrid = CType(Me.FindControl("MyNetwork"), DataGrid)
      AddBoundColumn(vDGR, "ContactName", "Contact Name")
      AddBoundColumn(vDGR, "Phone", "Telephone")
      AddMemoColumn(vDGR, "Notes", "Details")
      Dim vCount As Integer
      If HttpContext.Current.User.Identity.IsAuthenticated Then
        Dim vRestriction As String = String.Format("RelationshipCode = '{0}' AND Historical<>'Y'", InitialParameters("Relationship").ToString)
        vCount = DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactLinksTo, vDGR, vRestriction, UserContactNumber)
      End If
      If vCount = 0 Then
        Dim vDT As New DataTable
        Dim vDV As New DataView(vDT)
        vDGR.DataSource = vDV
        vDGR.DataBind()
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
End Class