Partial Public Class ShowDefaultAddress
  Inherits CareWebControl

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctDisplayAddress, tblDataEntry)
      Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetAddressDataTable(CareNetServices.XMLAddressDataSelectionTypes.xadtAddressInformation, UserAddressNumber))
      If vRow IsNot Nothing Then
        SetTextBoxText("Address", vRow("Address").ToString)
        SetTextBoxText("Town", vRow("Town").ToString)
        SetTextBoxText("County", vRow("County").ToString)
        SetTextBoxText("Postcode", vRow("Postcode").ToString)
        SetTextBoxText("Country", vRow("CountryDesc").ToString)
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

End Class