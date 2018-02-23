Public Partial Class AddMemberDD
  Inherits CareWebControl
  Implements ICareParentWebControl

  Private mvMembershipType As String = ""
  Private mvContactNumber As Integer
  Private mvStartDate As String = ""

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctAddMemberDD, tblDataEntry, "AccountName", "DirectNumber,MobileNumber")

      mvContactNumber = GetContactNumberFromParentGroup()
      If Request.QueryString("MT") IsNot Nothing AndAlso Request.QueryString("MT").Length > 0 Then mvMembershipType = Request.QueryString("MT")

      If Request.QueryString("SD") IsNot Nothing AndAlso Request.QueryString("SD").Length > 0 Then
        mvStartDate = Request.QueryString("SD").ToString
      ElseIf Session.Contents.Item("StartDate") IsNot Nothing AndAlso Session("StartDate").ToString.Length > 0 Then
        mvStartDate = Session("StartDate").ToString
      Else
        mvStartDate = ""
      End If
      SetDefaults()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Public Sub ProcessChildControls(ByVal pList As ParameterList) Implements ICareParentWebControl.ProcessChildControls
    SubmitChildControls(pList)
  End Sub

  Private Sub SetDefaults()
    If mvMembershipType.Length = 0 Then mvMembershipType = InitialParameters("MembershipType").ToString
    SetTextBoxText("MembershipType", mvMembershipType)
    SetLookupItem(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, "MembershipType", mvMembershipType)
    SetMemberBalance(DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.pm_dd), mvMembershipType, mvContactNumber, mvStartDate)
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        Dim vReturnList As ParameterList = AddNewContact()
        Dim vList As New ParameterList(HttpContext.Current)
        AddMemberParameters(vList, IntegerValue(vReturnList("ContactNumber").ToString), IntegerValue(vReturnList("AddressNumber").ToString), DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.pm_dd), mvMembershipType, mvStartDate)
        AddDDParameters(vList)
        Dim vMemberList As ParameterList = DataHelper.AddMember(vList)
        AddGiftAidDeclaration(IntegerValue(vList("PayerContactNumber").ToString))
        ProcessChildControls(vReturnList)
        If vMemberList.ContainsKey("DirectDebitNumber") AndAlso vMemberList("DirectDebitNumber").ToString.Length > 0 Then Session("DirectDebitNumber") = vMemberList("DirectDebitNumber").ToString
        If DefaultParameters.ContainsKey("PaymentFrequency") AndAlso DefaultParameters("PaymentFrequency").ToString.Length > 0 Then Session("DirectDebitPaymentFrequency") = DefaultParameters("PaymentFrequency").ToString
        If vMemberList.ContainsKey("FrequencyAmount") AndAlso vMemberList("FrequencyAmount").ToString.Length > 0 Then Session("DirectDebitFrequencyAmount") = vMemberList("FrequencyAmount").ToString
        If vMemberList.ContainsKey("DirectDebitClaimDate") AndAlso vMemberList("DirectDebitClaimDate").ToString.Length > 0 Then Session("DirectDebitClaimDate") = vMemberList("DirectDebitClaimDate").ToString
        If SubmitItemUrl.Length > 0 Then
          Dim vSubmitParams As New StringBuilder
          With vSubmitParams
            .Append("MT=")
            .Append(mvMembershipType)
            .Append("&DDN=")
            .Append(vMemberList("DirectDebitNumber").ToString)
            .Append("&MN=")
            .Append(vMemberList("MemberNumber").ToString)
            .Append("&MSN=")
            .Append(vMemberList("MembershipNumber").ToString)
            If vMemberList.ContainsKey("CardExpiryDate") AndAlso vMemberList("CardExpiryDate").ToString.Length > 0 Then
              .Append("&CED=")
              .Append(vMemberList("CardExpiryDate").ToString)
            End If
            If vMemberList.ContainsKey("MemberNumber2") Then
              .Append("&MN2=")
              .Append(vMemberList("MemberNumber2").ToString)
              .Append("&MSN2=")
              .Append(vMemberList("MembershipNumber2").ToString)
              If vMemberList("CardExpiryDate2").ToString.Length > 0 Then
                .Append("&CED2=")
                .Append(vMemberList("CardExpiryDate2").ToString)
              End If
              If mvStartDate.Length > 0 Then
                .Append("&SD=")
                .Append(mvStartDate)
              End If
            End If
          End With
          GoToSubmitPage(vSubmitParams.ToString)
        Else
          GoToSubmitPage()
        End If
      Catch vEX As ThreadAbortException
        Throw vEX
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub
End Class