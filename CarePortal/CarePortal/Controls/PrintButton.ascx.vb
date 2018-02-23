Public Class PrintButton
  Inherits CareWebControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    InitialiseControls(CareNetServices.WebControlTypes.wctPrintButton, tblDataEntry)
    End Sub

End Class