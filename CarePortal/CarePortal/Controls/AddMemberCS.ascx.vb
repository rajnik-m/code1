Partial Public Class AddMemberCS
  Inherits CareWebControl
  Private mvMembershipType As String = ""
  Private mvContactNumber As Integer
  Private mvStartDate As String = ""
  Private mvSkipProcessing As Boolean

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      mvUsesHiddenContactNumber = True
      mvHiddenFields = "HiddenAddressNumber"
      InitialiseControls(CareNetServices.WebControlTypes.wctAddMemberCS, tblDataEntry)
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

  Private Sub SetDefaults()
    If mvMembershipType.Length = 0 Then mvMembershipType = InitialParameters("MembershipType").ToString
    SetTextBoxText("MembershipType", mvMembershipType)
    SetLookupItem(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, "MembershipType", mvMembershipType)
    SetMemberBalance(GetPaymentMethod, mvMembershipType, mvContactNumber, mvStartDate)
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        Dim vReturnList As ParameterList = AddNewContact(GetHiddenContactNumber() > 0)
        Dim vList As New ParameterList(HttpContext.Current)
        Dim vContactCompany As String = ""
        Dim vSalesAccount As String = ""
        Dim vPaymentPlanDetails As New ParameterList()
        AddMemberParameters(vList, IntegerValue(vReturnList("ContactNumber").ToString), IntegerValue(vReturnList("AddressNumber").ToString), GetPaymentMethod, mvMembershipType, mvStartDate)
        GetCreditCustomerDetails(vContactCompany, vSalesAccount)
        vList("Company") = vContactCompany
        vList("SalesLedgerAccount") = vSalesAccount
        AddOptionalTextBoxValue(vList, "Reference")
        AddOptionalTextBoxValue(vList, "Notes")
        Dim vMemberList As ParameterList = DataHelper.AddMember(vList)
        AddGiftAidDeclaration(IntegerValue(vReturnList("ContactNumber").ToString))

        'Generate the invoice only when the invoice flag is set else just create a membership 
        If InitialParameters.Contains("CreateInvoice") AndAlso InitialParameters("CreateInvoice").ToString = "Y" Then
          Dim vPaylist As New ParameterList(HttpContext.Current)
          AddPayerInfo(vPaylist)
          vPaylist("PaymentPlanNumber") = vMemberList("PaymentPlanNumber")
          vPaylist("Amount") = GetTextBoxText("Balance")
          vPaylist("BankAccount") = DefaultParameters("BankAccount")
          vPaylist("Source") = DefaultParameters("Source")
          AddUserParameters(vPaylist)

          Try
            vPaymentPlanDetails = DataHelper.AddPaymentPlanPayment(vPaylist)
          Catch vEx As ThreadAbortException
            Throw vEx
          Catch vEx As CareException
            SetErrorLabel(vEx.Message)
          End Try

          If vPaymentPlanDetails IsNot Nothing Then
            Dim vCreditSalesParam As New ParameterList(HttpContext.Current)
            vCreditSalesParam("BatchNumber") = vPaymentPlanDetails("BatchNumber")
            vCreditSalesParam("TransactionNumber") = vPaymentPlanDetails("TransactionNumber")
            AddPayerInfo(vCreditSalesParam)
            vCreditSalesParam("Company") = vContactCompany
            vCreditSalesParam("SalesLedgerAccount") = vSalesAccount
            AddOptionalTextBoxValue(vCreditSalesParam, "Reference")
            AddOptionalTextBoxValue(vCreditSalesParam, "Notes")
            If DefaultParameters.ContainsKey("BatchCategory") Then vCreditSalesParam("BatchCategory") = DefaultParameters("BatchCategory")
            Try
              DataHelper.ConfirmCreditSaleTransaction(vCreditSalesParam)
            Catch vEx As ThreadAbortException
              Throw vEx
            Catch vEx As CareException
              SetErrorLabel(vEx.Message)
              mvSkipProcessing = True
            End Try
          End If
        End If
        If SubmitItemUrl.Length > 0 Then
          Dim vSubmitParams As New StringBuilder
          With vSubmitParams
            .Append("MT=")
            .Append(mvMembershipType)
            .Append("&MN=")
            .Append(vMemberList("MemberNumber").ToString)
            .Append("&MSN=")
            .Append(vMemberList("MembershipNumber").ToString)
            If vMemberList.ContainsKey("MemberNumber2") Then
              .Append("&MN2=")
              .Append(vMemberList("MemberNumber2").ToString)
              .Append("&MSN2=")
              .Append(vMemberList("MembershipNumber2").ToString)
            End If
            If mvStartDate.Length > 0 Then
              .Append("&SD=")
              .Append(mvStartDate)
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

  Private Sub GetCreditCustomerDetails(ByRef pContactCompany As String, ByRef pSalesAccount As String)
    Dim vParamList As ParameterList = New ParameterList(HttpContext.Current)

    AddPayerInfo(vParamList, True)
    vParamList("Company") = DefaultParameters.OptionalValue("Company").ToString
    'Check if the contact is a credit sales customer if not then create "CreateCreditCustomer" check box is checked 
    'else display an error
    Dim vDataTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCreditCustomers, vParamList)
    If vDataTable Is Nothing Then
      Dim vCreditCustParams As New ParameterList(HttpContext.Current)
      AddPayerInfo(vCreditCustParams)
      vCreditCustParams("Company") = DefaultParameters.OptionalValue("Company").ToString
      vCreditCustParams("CreditCategory") = DefaultParameters.OptionalValue("CreditCategory")
      Try
        Dim vResult As ParameterList = DataHelper.AddCreditCustomer(vCreditCustParams)
        pSalesAccount = vResult("SalesLedgerAccount").ToString
        pContactCompany = DefaultParameters.OptionalValue("Company").ToString
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vEx As CareException
        SetErrorLabel(vEx.Message)
        mvSkipProcessing = True
      End Try
    Else
      Dim vRow As DataRow() = vDataTable.Select()
      pContactCompany = vRow(0)("Company").ToString
      pSalesAccount = vRow(0)("SalesLedgerAccount").ToString
    End If
  End Sub

  Private Sub AddPayerInfo(ByVal pList As ParameterList, Optional ByVal pAddAddressNumber As Boolean = True)
    If ParentGroup.Length > 0 Then
      pList("ContactNumber") = mvContactNumber
      If pAddAddressNumber Then pList("AddressNumber") = GetContactAddress(mvContactNumber)
    Else
      If Session("ContactNumber") IsNot Nothing AndAlso Session("AddressNumber") IsNot Nothing Then
        pList("ContactNumber") = Session("ContactNumber")
        pList("AddressNumber") = Session("AddressNumber")
      Else
        pList("ContactNumber") = UserContactNumber()
        If pAddAddressNumber Then pList("AddressNumber") = UserAddressNumber()
      End If
    End If
  End Sub

  Protected Function GetPaymentMethod() As String
    Return DataHelper.ControlValue(DataHelper.ControlTables.credit_sales_controls, DataHelper.ControlValues.payment_method)
  End Function
End Class