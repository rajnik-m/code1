Partial Public Class FundraiserPageInformation
  Inherits CareWebControl

  Private mvFundraisingNumber As Integer

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    InitialiseControls(CareNetServices.WebControlTypes.wctFundraiserPageInformation, tblDataEntry, "")
    Dim vList As New ParameterList(HttpContext.Current)
    vList("WebPageNumber") = WebPageNumber
    Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactFundraisingEvents, vList)
    If vTable IsNot Nothing Then
      mvFundraisingNumber = IntegerValue(vTable.Rows(0)("ContactFundraisingNumber").ToString)
      Dim vTargetAmount As Single
      Single.TryParse(vTable.Rows(0)("TargetAmount").ToString, vTargetAmount)
      Dim vDonationTotal As Single
      Single.TryParse(vTable.Rows(0)("DonationTotal").ToString, vDonationTotal)
      Dim vGiftAidTotal As Single
      Single.TryParse(vTable.Rows(0)("GiftAidTotal").ToString, vGiftAidTotal)
      SetTextBoxText("TargetAmount", vTargetAmount.ToString("#.00"))
      SetTextBoxText("DonationTotal", vDonationTotal.ToString("#.00"))
      SetTextBoxText("GiftAidTotal", vGiftAidTotal.ToString("#.00"))
      SetTextBoxText("TotalAmount", (vDonationTotal + vGiftAidTotal).ToString("#.00"))
      Dim vControl As HtmlGenericControl = TryCast(FindControlByName(tblDataEntry, "DivProgress"), HtmlGenericControl)
      If vControl IsNot Nothing Then
        vControl.Style("width") = ((200 * vDonationTotal) / vTargetAmount).ToString & "px"
      End If
    End If
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        GoToSubmitPage(String.Format("&cfn={0}", mvFundraisingNumber))
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub

End Class