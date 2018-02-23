Partial Public Class MakeDonationCC
  Inherits CareWebControl
  Implements ICareParentWebControl

  Private mvFundraisingNumber As Integer
  Private mvProduct As String = ""
  Private mvRate As String = ""

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      mvUsesHiddenContactNumber = True
      mvHiddenFields = "HiddenAddressNumber"
      SupportsOnlineCCAuthorisation = True
      InitialiseControls(CareNetServices.WebControlTypes.wctMakeDonationCC, tblDataEntry, "CreditCardNumber,CardExpiryDate", "DirectNumber,MobileNumber")
      If (Request.QueryString("PR") IsNot Nothing AndAlso Request.QueryString("RA") IsNot Nothing) AndAlso (Request.QueryString("PR").Length > 0 AndAlso Request.QueryString("RA").Length > 0) Then
        'We expect both Product & Rate to be passed in, or neither
        mvProduct = Request.QueryString("PR")
        mvRate = Request.QueryString("RA")
      End If
      If Request.QueryString("cfn") IsNot Nothing AndAlso Request.QueryString("cfn").Length > 0 Then mvFundraisingNumber = IntegerValue(Request.QueryString("cfn"))
      SetDefaults()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        SetErrorLabel("")
        Dim vReturnList As ParameterList = AddNewContact(GetHiddenContactNumber() > 0)
        Dim vList As New ParameterList(HttpContext.Current)
        vList("ContactNumber") = vReturnList("ContactNumber")
        vList("AddressNumber") = vReturnList("AddressNumber")
        AddGiftAidDeclaration(IntegerValue(vList("ContactNumber").ToString))
        'Now need to take the payment
        vList("Product") = mvProduct
        vList("Rate") = mvRate
        vList("Quantity") = "1"
        vList("Amount") = GetTextBoxText("Amount")
        vList("Notes") = GetTextBoxText("Notes")
        If mvFundraisingNumber > 0 Then vList("ContactFundraisingNumber") = mvFundraisingNumber.ToString
        AddCCParameters(vList)
        AddUserParameters(vList)
        AddDefaultParameters(vList)
        Dim vSkipProcessing As Boolean
        Try
          DataHelper.AddProductSale(vList)
        Catch vEx As ThreadAbortException
          Throw vEx
        Catch vEx As CareException
          SetErrorLabel(vEx.Message)
          SetHiddenText("HiddenContactNumber", vReturnList("ContactNumber").ToString)
          SetHiddenText("HiddenAddressNumber", vReturnList("AddressNumber").ToString)
          vSkipProcessing = True
        End Try
        If vSkipProcessing = False Then
          'Need to Email
          If mvFundraisingNumber > 0 Then
            vList = New ParameterList(HttpContext.Current)
            vList("ContactFundraisingNumber") = mvFundraisingNumber.ToString
            Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactFundraisingEvents, vList))
            If vRow IsNot Nothing Then
              Dim vMessage As String = vRow("ThankYouMessage").ToString             'This is the Email message
              If vMessage.Length > 0 Then
                Dim vFromEmailAddress As String = ""
                If DefaultParameters.ContainsKey("EMailAddress") Then vFromEmailAddress = DefaultParameters("EMailAddress").ToString
                If vFromEmailAddress.Length > 0 Then
                  Dim vSubject As String = vRow("FundraisingDescription").ToString    'This is the Email subject
                  Dim vToEmailAddress As String = ToEmailAddress(vReturnList)
                  If vToEmailAddress.Length > 0 Then
                    If vMessage.Contains("%Amount") Then vMessage = vMessage.Replace("%Amount", DoubleValue(GetTextBoxText("Amount")).ToString("0.00"))
                    If vMessage.Contains("%TargetAmount") Then vMessage = vMessage.Replace("%TargetAmount", DoubleValue(vRow("TargetAmount").ToString).ToString("0.00"))
                    If vMessage.Contains("%TargetDate") Then vMessage = vMessage.Replace("%TargetDate", vRow("TargetDate").ToString)
                    If vMessage.Contains("%Description") Then vMessage = vMessage.Replace("%Description", vRow("FundraisingDescription").ToString)
                    SendEmail(vFromEmailAddress, vToEmailAddress, vSubject, vMessage)
                  End If
                Else
                  Throw New CareException(String.Format("EMail Address Undefined"))
                End If
              End If
            End If
          ElseIf DefaultParameters.Contains("StandardDocument") AndAlso DefaultParameters("StandardDocument").ToString.Length > 0 Then
            Dim vEmailParams As New ParameterList(HttpContext.Current)
            vEmailParams("StandardDocument") = DefaultParameters("StandardDocument")
            vEmailParams("EMailAddress") = DefaultParameters("EMailAddress")
            vEmailParams("Name") = DefaultParameters("Name")
            vEmailParams("Source") = DefaultParameters("Source")
            vEmailParams("CreateMailingHistory") = "Y"
            Dim vToEmailAddress As String = ToEmailAddress(vReturnList)
            If vToEmailAddress.Length > 0 Then
              Dim vContactList As New ParameterList(HttpContext.Current)
              vContactList("ContactNumber") = vReturnList("ContactNumber")
              Dim vDR As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vContactList))
              Dim vContentParams As New ParameterList
              If vDR IsNot Nothing Then
                vContentParams("EMail") = vToEmailAddress
                vContentParams("Contact Number") = vDR("ContactNumber") 'Used in MailingData-ProcessBulkEmail
                vContentParams("Contact_Number") = vDR("ContactNumber") 'Used for MergeField
                vContentParams("Address_Number") = vDR("AddressNumber")
                vContentParams("Name") = vDR("ContactName")
                vContentParams("Title") = vDR("Title")
                vContentParams("Initials") = vDR("Initials")
                vContentParams("Forenames") = vDR("Forenames")
                vContentParams("Surname") = vDR("Surname")
                vContentParams("Label_Name") = vDR("LabelName")
                vContentParams("Honorifics") = vDR("Honorifics")
                vContentParams("Salutation") = vDR("Salutation")
                vContentParams("Informal Salutation") = vDR("InformalSalutation")
                vContentParams("Position") = vDR("Position")
                vContentParams("Organisation") = vDR("OrganisationName")
                vContentParams("Address_1") = vDR("Address")
                vContentParams("Address_Line") = vDR("AddressLine")
                vContentParams("Address_Multi_Line") = vDR("AddressMultiLine")
                vContentParams("Town") = vDR("Town")
                vContentParams("County") = vDR("County")
                vContentParams("PostCode") = vDR("PostCode")
                vContentParams("Country") = vDR("CountryDesc")
                vContentParams("Country_Code") = vDR("CountryCode")
                vContentParams("Phone_Number") = vDR("PhoneNumber")
                vContentParams("Amount") = vList("Amount")

                DataHelper.ProcessBulkEMail(vContentParams.ToCSVFile, vEmailParams, True)
              End If
            End If
          End If
          ProcessChildControls(vReturnList)
          GoToSubmitPage()
        End If
      Catch vEX As ThreadAbortException
        Throw vEX
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub

  Private Function ToEmailAddress(ByVal vReturnList As ParameterList) As String
    Dim vToEmailAddress As String = ""
    If mvContactEntryHidden Then
      'Get the EmailAddress from the datatabse
      Dim vList As New ParameterList(HttpContext.Current)
      vList("ContactNumber") = vReturnList("ContactNumber")
      Dim vDT As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers, vList)
      Dim vEmailDevice As String = DataHelper.ControlValue(DataHelper.ControlValues.email_device)
      If vDT IsNot Nothing AndAlso vEmailDevice.Length > 0 Then
        vDT.DefaultView.RowFilter = String.Format("DeviceCode = '{0}' AND IsActive = 'Yes' AND Mail = 'Yes'", vEmailDevice)
        Dim vDR As DataRow = DataHelper.GetRowFromDataTable(vDT.DefaultView.ToTable)
        If vDR IsNot Nothing Then vToEmailAddress = vDR("Number").ToString
      End If
    Else
      'Get the EmailAddress from the page
      vToEmailAddress = GetTextBoxText("EMailAddress")
    End If
    Return vToEmailAddress
  End Function

  Public Sub ProcessChildControls(ByVal pList As ParameterList) Implements ICareParentWebControl.ProcessChildControls
    SubmitChildControls(pList)
  End Sub

  Private Sub SetDefaults()
    If mvProduct.Length = 0 Then
      mvProduct = InitialParameters("Product").ToString
      mvRate = InitialParameters("Rate").ToString
    End If
    SetAmountOrBalance("Amount", mvProduct, mvRate)
  End Sub
End Class