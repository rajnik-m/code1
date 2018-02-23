Partial Public Class RecordActivity
  Inherits CareWebControl

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    InitialiseControls(CareNetServices.WebControlTypes.wctRecordActivity, tblDataEntry, "Notes")
  End Sub

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        Dim vList As New ParameterList(HttpContext.Current)
        vList("ContactNumber") = UserContactNumber()
        vList("Activity") = InitialParameters("Activity")
        vList("ActivityValue") = GetDropDownValue("ActivityValue")
        AddOptionalTextBoxValue(vList, "ValidFrom")
        AddOptionalTextBoxValue(vList, "ValidTo")
        AddOptionalTextBoxValue(vList, "Notes")
        AddDefaultParameters(vList)
        DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctActivities, vList)
        GoToSubmitPage()
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub
End Class