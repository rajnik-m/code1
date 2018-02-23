Public Class ShowExamHistory
  Inherits CareWebControl

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub
  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctShowExamHistory, tblDataEntry)
      Dim vList As New ParameterList(HttpContext.Current)
      vList("SystemColumns") = "Y"
      vList("WebPageItemNumber") = Me.WebPageItemNumber
      If InWebPageDesigner() Then
        vList("DocumentColumns") = "Y"
      Else
        vList("ContactNumber") = UserContactNumber().ToString
      End If
      Dim vDataGrid As DataGrid = TryCast(Me.FindControl("ExamHistoryData"), DataGrid)
      If vDataGrid IsNot Nothing Then
        Dim vResult As String = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebExamHistory, vList)
        DataHelper.FillGrid(vResult, vDataGrid)
        If vDataGrid.Items.Count > 0 Then
          vDataGrid.DataBind()
          DirectCast(Me.FindControl("WarningMessage"), Label).Visible = False
        Else
          vDataGrid.Visible = False
          DirectCast(Me.FindControl("WarningMessage"), Label).Visible = True
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    End Try
  End Sub

End Class