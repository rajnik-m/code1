Partial Public Class FindLocalGroup
  Inherits CareWebControl

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    InitialiseControls(CareNetServices.WebControlTypes.wctFindLocalGroup, tblDataEntry, "Postcode")
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        Dim vDGR As DataGrid = CType(Me.FindControl("LocalGroups"), DataGrid)
        Dim vList As New ParameterList(HttpContext.Current)
        AddOptionalTextBoxValue(vList, "Postcode")
        vList("SystemColumns") = "Y"
        DataHelper.GetNearest(vList, vDGR)
        'GoToSubmitPage()
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub

End Class