Imports System.Web.Configuration
Imports System.IO

Public Class DataHelper

  Public Enum ConfigurationOptions
    cd_joint_contact_support
    cpd_points_allow_numeric
    fp_default_blank_claim_day
    use_ajax_for_contact_names
    portal_password_case_sensitive
    portal_bookings_close_lock
    multiple_label_name_formats
  End Enum

  Public Enum ConfigurationValues
    albacs_verify
    afd_everywhere_server
    afd_software
    cd_status_mandatory
    email_smtp_server
    qas_pro_web_url
    fp_cc_authorisation_type
    fp_cc_security_code_mandatory
    pm_cc
    pm_dd
    jnr_label_name_format
    label_name_format
    default_salutation
    default_female_salutation
    default_male_salutation
    initials_format
    option_country
    portal_password_min_length
    portal_password_complexity
    portal_update_details_freq
    web_product_image_name
    protX_web_url
    protX_vendor_name
    paypal_web_url
    fp_cc_authorisation_timeout
    web_documents_directory
    fixed_renewal_M
    cd_address_change_with_branch
  End Enum

  Public Enum ControlTables
    caf_mail_order_controls
    contact_controls
    covenant_controls
    credit_sales_controls
    email_controls
    event_controls
    financial_controls
    gaye_controls
    gift_aid_controls
    legacy_controls
    marketing_controls
    membership_controls
    noncaf_mail_order_controls
    stock_movement_controls
    warehouses
    exam_controls
  End Enum

  Public Enum ControlValues
    auto_pay_claim_date_method
    direct_device
    switchboard_device
    mobile_device
    fax_device
    email_device
    web_device
    payment_method
    force_smtp_address
    End Enum

  Public Enum AccountNoVerifyResult
    avrValid = 1
    avrSortcodeValidAccountInvalid
    avrInvalid
    avrSortcodeValidAccountWarn
    avrWarning
  End Enum

  Private Shared mvControlValues As DataTable = Nothing
  Private Shared mvDevicesTable As DataTable = Nothing
  Private Shared mvNS As New CareNetServices.NDataAccess With {.Credentials = System.Net.CredentialCache.DefaultCredentials}
  Private Shared mvXS As New ExamsAccess.ExamsDataAccess With {.Credentials = System.Net.CredentialCache.DefaultCredentials}
  Private Shared mvUseWebService As Boolean = True
  Private Shared mvDatabase As String


  Public Shared ReadOnly Property Database() As String
    Get
      If mvDatabase Is Nothing Then
        If WebConfigurationManager.AppSettings("Database") IsNot Nothing Then
          mvDatabase = WebConfigurationManager.AppSettings("Database").ToString
        Else
          mvDatabase = "CDBWEBSERVER"
        End If
      End If
      Return mvDatabase
    End Get
  End Property

  Public Shared Function GetReportFile(ByVal pList As ParameterList) As String
    If mvUseWebService Then
      Dim vBuffer As Byte() = mvNS.GetReportFile(pList.XMLParameterString)
      Return GetFileFromByteArray(vBuffer)
    Else
      Return DataHelperDirect.GetReportFile(pList)
    End If

  End Function

  Public Shared Function GetFileFromByteArray(ByVal pBuffer As Byte()) As String
    Dim vFileName As String = ""
    If pBuffer IsNot Nothing Then
      If CheckBufferValid(pBuffer) Then
        vFileName = My.Computer.FileSystem.GetTempFileName
        Using vFS As New IO.FileStream(vFileName, IO.FileMode.Create)
          vFS.Write(pBuffer, 0, pBuffer.Length)
          vFS.Close()
        End Using
      End If
    Else
      'No file returned
      vFileName = My.Computer.FileSystem.GetTempFileName
    End If
    Return vFileName
  End Function

  Public Shared Function CheckBufferValid(ByVal pBuffer As Byte()) As Boolean
    If Not pBuffer Is Nothing Then
      Dim vError As String = (New UnicodeEncoding).GetString(pBuffer, 0, Math.Min(200, pBuffer.Length))
      If vError.IndexOf("<Result><ErrorMessage>") > 0 Then
        'Create a parameter list from the result - This will raise the error
        Dim vResult As New ParameterList((New UnicodeEncoding).GetString(pBuffer))
      End If
      Return True
    End If
  End Function

  Public Shared Function ControlValue(ByVal pControlTable As ControlTables, ByVal pControlValue As ControlValues) As String
    If mvControlValues Is Nothing Then mvControlValues = GetLookupData(CareNetServices.XMLLookupDataTypes.xldtControlValues)
    Dim vRows() As DataRow = mvControlValues.Select(String.Format("ControlTable = '{0}' AND ControlName = '{1}'", [Enum].GetName(GetType(ControlTables), pControlTable), [Enum].GetName(GetType(ControlValues), pControlValue)))
    If vRows.Length > 0 Then
      Return vRows(0).Item("ControlValue").ToString
    Else
      Return ""
    End If
  End Function

  Public Shared Function ControlValue(ByVal pControlValue As ControlValues) As String
    If mvControlValues Is Nothing Then mvControlValues = GetLookupData(CareNetServices.XMLLookupDataTypes.xldtControlValues)
    Dim vRows() As DataRow = mvControlValues.Select(String.Format("ControlName = '{0}'", [Enum].GetName(GetType(ControlValues), pControlValue)))
    If vRows.Length > 0 Then
      Return vRows(0).Item("ControlValue").ToString
    Else
      Return ""
    End If
  End Function

  Public Shared Function DeviceMaxLength(ByVal pDeviceCode As String) As Integer
    If mvDevicesTable Is Nothing Then mvDevicesTable = GetLookupData(CareNetServices.XMLLookupDataTypes.xldtDevices)
    Dim vMaxLength As Integer = 20
    If pDeviceCode.Length > 0 Then
      Dim vRows() As DataRow = mvDevicesTable.Select(String.Format("Device = '{0}'", pDeviceCode))
      If vRows.Length > 0 Then
        vMaxLength = IntegerValue(vRows(0).Item("MaxLength").ToString)
      End If
    End If
    Return vMaxLength
  End Function

  Public Shared Function AddItem(ByVal pType As CareNetServices.XMLMaintenanceControlTypes, ByVal pList As ParameterList) As ParameterList
    Dim vResult As String = ""
    If mvUseWebService Then
      Select Case pType
        Case CareNetServices.XMLMaintenanceControlTypes.xmctActivities
          vResult = mvNS.AddActivity(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctAddresses
          vResult = mvNS.AddAddress(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctContact
          vResult = mvNS.AddContact(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctDocument
          vResult = mvNS.AddCommunicationsLog(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctReference
          vResult = mvNS.AddExternalReference(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctFundraisingEvents
          vResult = mvNS.AddContactFundraisingEvent(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctGiftAidDeclarations
          vResult = mvNS.AddGiftAidDeclaration(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctLink
          vResult = mvNS.AddLink(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctNumber
          vResult = mvNS.AddCommunicationsNumber(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctSuppression
          vResult = mvNS.AddSuppression(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctContactAccounts
          vResult = mvNS.AddContactAccount(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctOrganisation
          vResult = mvNS.AddOrganisation(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctPosition
          vResult = mvNS.AddPosition(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctPosition
          vResult = mvNS.AddPosition(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctContactSurveys
          vResult = mvNS.AddContactSurvey(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctCPDCycles
          vResult = mvNS.AddContactCPDCycle(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctCPDPoints
          vResult = mvNS.AddContactCPDPoints(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctCPDObjectives
          vResult = mvNS.AddCpdObjective(pList.XMLParameterString)
      End Select
      Return New ParameterList(vResult)
    Else
      Return DataHelperDirect.AddItem(pType, pList)
    End If
  End Function

  Public Shared Function AddWebItem(ByVal pType As CareNetServices.XMLWebDataSelectionTypes, ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Dim vResult As String = ""
      Select Case pType
        Case CareNetServices.XMLWebDataSelectionTypes.wstPage
          vResult = mvNS.AddWebItem(pType, pList.XMLParameterString)
        Case CareNetServices.XMLWebDataSelectionTypes.wstPageItem
          vResult = mvNS.AddWebItem(pType, pList.XMLParameterString)
      End Select
      Return New ParameterList(vResult)
    Else
      Return DataHelperDirect.AddWebItem(pType, pList)
    End If
  End Function

  Public Shared Function UpdateWebItem(ByVal pType As CareNetServices.XMLWebDataSelectionTypes, ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Dim vResult As String = ""
      Select Case pType
        Case CareNetServices.XMLWebDataSelectionTypes.wstPage
          vResult = mvNS.UpdateWebItem(pType, pList.XMLParameterString)
        Case CareNetServices.XMLWebDataSelectionTypes.wstPageItem
          vResult = mvNS.UpdateWebItem(pType, pList.XMLParameterString)
      End Select
      Return New ParameterList(vResult)
    Else
      Return DataHelperDirect.UpdateWebItem(pType, pList)
    End If
  End Function

  Public Shared Function UpdateProvisionalTransaction(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Dim vResult As String = ""
      vResult = mvNS.UpdateProvisionalTransaction(pList.XMLParameterString)
      Return New ParameterList(vResult)
    Else
      Return DataHelperDirect.UpdateProvisionalTransaction(pList)
    End If
  End Function

  Public Shared Function UpdateItem(ByVal pType As CareNetServices.XMLMaintenanceControlTypes, ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Dim vResult As String = ""
      Select Case pType
        Case CareNetServices.XMLMaintenanceControlTypes.xmctAction
          vResult = mvNS.UpdateAction(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctActivities
          vResult = mvNS.UpdateActivity(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctAddresses
          pList("CarePortal") = "Y"
          vResult = mvNS.UpdateAddress(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctContact
          vResult = mvNS.UpdateContact(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctFundraisingEvents
          mvNS.UpdateContactFundraisingEvent(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctLink
          vResult = mvNS.UpdateLink(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctNumber
          vResult = mvNS.UpdateCommunicationsNumber(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctContactAccounts
          vResult = mvNS.UpdateContactAccount(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctReference
          vResult = mvNS.UpdateExternalReference(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctSuppression
          vResult = mvNS.UpdateSuppression(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctEventDelegateActivity
          vResult = mvNS.UpdateDelegateActivity(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctPosition
          vResult = mvNS.UpdatePosition(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctContactSurveys
          vResult = mvNS.UpdateContactSurvey(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctCPDCycles
          vResult = mvNS.UpdateContactCPDCycle(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctCPDPoints
          vResult = mvNS.UpdateContactCPDPoints(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctCPDObjectives
          vResult = mvNS.UpdateCpdObjective(pList.XMLParameterString)
      End Select
      Return New ParameterList(vResult)
    Else
      Return DataHelperDirect.UpdateItem(pType, pList)
    End If
  End Function

  Public Shared Function UpdateWebDocument(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Dim vResult As String = ""
      vResult = mvNS.UpdateWebDocument(pList.XMLParameterString)
      Return New ParameterList(vResult)
    Else
      Return DataHelperDirect.UpdateWebDocument(pList)
    End If
  End Function

  Public Shared Function DeleteItem(ByVal pType As CareNetServices.XMLTransactionDataSelectionTypes, ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Dim vResult As String = ""
      Select Case pType
        Case CareNetServices.XMLTransactionDataSelectionTypes.xtdtTransactionAnalysis
          vResult = mvNS.DeleteProvisionalTransaction(pList.XMLParameterString)
      End Select
      Return New ParameterList(vResult)
    Else
      Return DataHelperDirect.DeleteItem(pType, pList)
    End If
  End Function

  Public Shared Function DeleteItem(ByVal pType As CareNetServices.XMLMaintenanceControlTypes, ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Dim vResult As String = ""
      Select Case pType
        Case CareNetServices.XMLMaintenanceControlTypes.xmctActionLink
          vResult = mvNS.DeleteActionLink(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctActionTopic
          vResult = mvNS.DeleteActionSubject(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctActivities
          vResult = mvNS.DeleteActivity(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctLink
          vResult = mvNS.DeleteLink(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctNumber
          vResult = mvNS.DeleteCommunicationsNumber(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctReference
          vResult = mvNS.DeleteExternalReference(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctSuppression
          vResult = mvNS.DeleteSuppression(pList.XMLParameterString)
        Case CareNetServices.XMLMaintenanceControlTypes.xmctEventDelegateActivity
          vResult = mvNS.DeleteDelegateActivity(pList.XMLParameterString)
      End Select
      Return New ParameterList(vResult)
    Else
      Return DataHelperDirect.DeleteItem(pType, pList)
    End If
  End Function

  Public Shared Function GetAddressDataTable(ByVal pType As CareNetServices.XMLAddressDataSelectionTypes, ByVal pAddressNumber As Long) As DataTable
    If mvUseWebService Then
      Dim vList As New ParameterList(HttpContext.Current)
      vList("AddressNumber") = pAddressNumber
      Return GetDataTable(mvNS.SelectAddressData(pType, vList.XMLParameterString))
    Else
      Return DataHelperDirect.GetAddressDataTable(pType, pAddressNumber)
    End If
  End Function

  Public Shared Function GetContactDataTable(ByVal pType As CareNetServices.XMLContactDataSelectionTypes, ByVal pContactNumber As Long, Optional ByVal pList As ParameterList = Nothing) As DataTable
    If pList Is Nothing Then pList = New ParameterList(HttpContext.Current)
    pList("ContactNumber") = pContactNumber
    Return GetDataTable(SelectContactData(pType, pList))
  End Function

  Public Shared Function GetContactDataTable(ByVal pType As CareNetServices.XMLContactDataSelectionTypes, ByVal pList As ParameterList) As DataTable
    Return GetDataTable(SelectContactData(pType, pList))
  End Function

  Public Shared Function GetMembershipData(ByVal pType As CareNetServices.XMLMembershipDataSelectionTypes, ByVal pMembershipNumber As Integer, Optional ByVal pList As ParameterList = Nothing) As DataTable
    If mvUseWebService Then
      If pList Is Nothing Then pList = New ParameterList(HttpContext.Current)
      pList("MembershipNumber") = pMembershipNumber
      Dim vResult As String = mvNS.SelectMembershipData(pType, pList.XMLParameterString)
      Return GetDataTable(vResult)
    Else
      Return DataHelperDirect.GetMembershipData(pType, pMembershipNumber, pList)
    End If
  End Function

  Public Shared Function SelectContactData(ByVal pType As CareNetServices.XMLContactDataSelectionTypes, ByVal pList As ParameterList) As String
    If mvUseWebService Then
      Return mvNS.SelectContactData(pType, pList.XMLParameterString)
    Else

      Return DataHelperDirect.SelectContactData(pType, pList)
    End If
  End Function

  Public Shared Function SelectEventData(ByVal pType As CareNetServices.XMLEventDataSelectionTypes, ByVal pList As ParameterList) As String
    If mvUseWebService Then
      Return mvNS.SelectEventData(pType, pList.XMLParameterString)
    Else
      Return DataHelperDirect.SelectEventData(pType, pList)
    End If
  End Function

  Public Shared Function SelectExamData(ByVal pType As ExamsAccess.XMLExamDataSelectionTypes, ByVal pList As ParameterList) As String
    If mvUseWebService Then
      Return mvXS.SelectExamData(pType, pList.XMLParameterString)
    Else
      Return DataHelperDirect.SelectExamData(pType, pList)
    End If
  End Function

  Public Shared Sub DefaultClaimDay(ByVal pDataTable As DataTable, ByVal pBankAccount As String)
    If DataHelper.ConfigurationOption(DataHelper.ConfigurationOptions.fp_default_blank_claim_day, False) Then
      pDataTable.DefaultView.RowFilter = "(BankAccount = '" & pBankAccount & "' AND ClaimType = 'DD') Or (BankAccount = '')"
    Else
      pDataTable.DefaultView.RowFilter = "BankAccount = '" & pBankAccount & "' AND ClaimType = 'DD'"
    End If
  End Sub

  Public Shared Function SelectFundraisingEventData(ByVal pType As CareNetServices.XMLFundraisingEventDataSelectionTypes, ByVal pList As ParameterList) As String
    If mvUseWebService Then
      Return mvNS.SelectFundraisingEventData(pType, pList.XMLParameterString)
    Else
      Return DataHelperDirect.SelectFundraisingEventData(pType, pList)
    End If
  End Function

  Public Shared Function SelectTableData(ByVal pType As CareNetServices.XMLTableDataSelectionTypes, ByVal pList As ParameterList) As DataTable
    If mvUseWebService Then
      Return Utilities.GetDataTable(mvNS.SelectTableData(pType, pList.XMLParameterString), True)
    Else
      Return DataHelperDirect.SelectTableData(pType, pList)
    End If
  End Function

  Public Shared Function SelectTableDataString(ByVal pType As CareNetServices.XMLTableDataSelectionTypes, ByVal pList As ParameterList) As String
    If mvUseWebService Then
      Return mvNS.SelectTableData(pType, pList.XMLParameterString)
    Else
      Return DataHelperDirect.SelectTableDataString(pType, pList)
    End If
  End Function

  Public Shared Function GetEventDataTable(ByVal pType As CareNetServices.XMLEventDataSelectionTypes, ByVal pList As ParameterList) As DataTable
    If mvUseWebService Then
      Return GetDataTable(mvNS.SelectEventData(pType, pList.XMLParameterString))
    Else
      Return DataHelperDirect.GetEventDataTable(pType, pList)
    End If
  End Function

  Public Shared Function SelectWebDataTable(ByVal pType As CareNetServices.XMLWebDataSelectionTypes, ByVal pList As ParameterList) As DataTable
    If mvUseWebService Then
      Return GetDataTable(mvNS.SelectWebData(pType, pList.XMLParameterString))
    Else
      Return DataHelperDirect.SelectWebDataTable(pType, pList)
    End If
  End Function

  Public Shared Function AddEventBooking(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddEventBooking(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddEventBooking(pList)
    End If
  End Function

  Public Shared Function AddExamBooking(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvXS.AddExamBooking(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddExamBooking(pList)
    End If
  End Function

  Public Shared Function CalculateExamBookingPrice(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvXS.CalculateExamBookingPrice(pList.XMLParameterString))
    Else
      Return DataHelperDirect.CalculateExamBookingPrice(pList)
    End If
  End Function

  Public Shared Function AddAction(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddAction(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddAction(pList)
    End If
  End Function

  Public Shared Function AddActionLink(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddActionLink(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddActionLink(pList)
    End If
  End Function

  Public Shared Function AddActionSubject(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddActionSubject(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddActionSubject(pList)
    End If
  End Function

  Public Shared Function AddActivity(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddActivity(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddActivity(pList)
    End If
  End Function

  Public Shared Function AddEvent(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddEvent(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddEvent(pList)
    End If
  End Function

  Public Shared Function AddEventBookingOption(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddEventBookingOption(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddEventBookingOption(pList)
    End If
  End Function

  Public Shared Function UpdateEvent(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.UpdateEvent(pList.XMLParameterString))
    Else
      Return DataHelperDirect.UpdateEvent(pList)
    End If
  End Function

  Public Shared Function AddDirectDebit(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddAutoPaymentMethod(CareNetServices.ppAutoPayMethods.ppAPMDD, pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddDirectDebit(pList)
    End If
  End Function

  Public Shared Function AddLink(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddLink(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddLink(pList)
    End If
  End Function

  Public Shared Function AddSuppresion(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddSuppression(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddSuppresion(pList)
    End If
  End Function

  Public Shared Function AddMember(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddMembership(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddMember(pList)
    End If
  End Function

  Public Shared Function AddPaymentPlan(ByVal pType As CareNetServices.ppType, ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddPaymentPlan(pType, pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddPaymentPlan(pType, pList)
    End If
  End Function

  Public Shared Function AddInvoicePayment(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddInvoicePayment(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddInvoicePayment(pList)
    End If
  End Function

  Public Shared Function AddPaymentPlanPayment(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddPaymentPlanPayment(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddPaymentPlanPayment(pList)
    End If
  End Function

  Public Shared Function UpdatePaymentPlan(ByVal pType As CareNetServices.XMLPaymentPlanUpdateTypes, ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.UpdatePaymentPlan(pType, pList.XMLParameterString))
    Else
      Return DataHelperDirect.UpdatePaymentPlan(pType, pList)
    End If
  End Function

  Public Shared Function AddProductSale(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddProductSale(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddProductSale(pList)
    End If
  End Function

  Public Shared Function ConfigurationValue(ByVal pOption As ConfigurationValues) As String
    Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtConfigs)
    Dim vRows() As DataRow = vTable.Select(String.Format("ConfigName = '{0}'", [Enum].GetName(GetType(ConfigurationValues), pOption)))
    If vRows.Length > 0 Then
      Return vRows(0).Item("ConfigValue").ToString
    Else
      Return ""
    End If
  End Function

  Public Shared Function ConfigurationOption(ByVal pOption As ConfigurationOptions, Optional ByVal pDefaultValue As Boolean = False) As Boolean
    Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtConfigs)
    Dim vRows() As DataRow = vTable.Select(String.Format("ConfigName = '{0}'", [Enum].GetName(GetType(ConfigurationOptions), pOption)))
    If vRows.Length > 0 Then
      Dim vString As String = vRows(0).Item("ConfigValue").ToString
      Return vString.ToUpper.StartsWith("Y")
    Else
      Return pDefaultValue
    End If
  End Function

  Public Shared Function ConfigurationValueOption(ByVal pOption As ConfigurationValues) As Boolean
    Return BooleanValue(ConfigurationValue(pOption))
  End Function

  Public Shared Function ConfirmCardSaleTransaction(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.ConfirmCardSaleTransaction(pList.XMLParameterString))
    Else
      Return DataHelperDirect.ConfirmCardSaleTransaction(pList)
    End If
  End Function

  Public Shared Function ConfirmCreditSaleTransaction(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.ConfirmCreditSaleTransaction(pList.XMLParameterString))
    Else
      Return DataHelperDirect.ConfirmCreditSaleTransaction(pList)
    End If
  End Function

  Public Shared Function ConfirmCashSaleTransaction(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.ConfirmCashSaleTransaction(pList.XMLParameterString))
    Else
      Return DataHelperDirect.ConfirmCashSaleTransaction(pList)
    End If
  End Function

  Public Shared Function ConfirmCreditAndCardSaleTransaction(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.ConfirmCreditAndCardSaleTransaction(pList.XMLParameterString))
    Else
      Return DataHelperDirect.ConfirmCreditAndCardSaleTransaction(pList)
    End If
  End Function

  Public Shared Sub FillCombo(ByVal pType As CareNetServices.XMLLookupDataTypes, ByVal pCombo As DropDownList, Optional ByVal pAddBlankRow As Boolean = False, Optional ByVal pList As ParameterList = Nothing)
    If mvUseWebService Then
      Dim vTable As DataTable
      If pList Is Nothing Then pList = New ParameterList(HttpContext.Current)
      vTable = GetDataTable((mvNS.GetLookupData(pType, pList.XMLParameterString)))
      If vTable IsNot Nothing Then
        If pAddBlankRow Then vTable.Rows.InsertAt(vTable.NewRow(), 0)
        pCombo.DataSource = vTable
        pCombo.DataBind()
      End If
    Else
      DataHelperDirect.FillCombo(pType, pCombo, pAddBlankRow, pList)
    End If
  End Sub

  Public Shared Sub FillList(ByVal pType As CareNetServices.XMLLookupDataTypes, ByVal pListBox As ListBox, Optional ByVal pAddBlankRow As Boolean = False, Optional ByVal pList As ParameterList = Nothing, Optional ByVal pMultiSelect As Boolean = False)
    If mvUseWebService Then
      Dim vTable As DataTable
      If pList Is Nothing Then pList = New ParameterList(HttpContext.Current)
      vTable = GetDataTable((mvNS.GetLookupData(pType, pList.XMLParameterString)))
      If vTable IsNot Nothing Then
        If pAddBlankRow Then vTable.Rows.InsertAt(vTable.NewRow(), 0)
        pListBox.DataTextField = "ActivityValueDesc"
        pListBox.DataValueField = "ActivityValue"
        pListBox.DataSource = vTable
        pListBox.DataBind()
        If pMultiSelect Then pListBox.SelectionMode = ListSelectionMode.Multiple
      End If
    Else
      DataHelperDirect.FillList(pType, pListBox, pAddBlankRow, pList, pMultiSelect)
    End If
  End Sub

  Public Shared Sub FillList(ByVal pType As CareNetServices.XMLLookupDataTypes, ByVal pListBox As ListBox, ByVal pDataAndTextField As ParameterList, Optional ByVal pAddBlankRow As Boolean = False, Optional ByVal pList As ParameterList = Nothing, Optional ByVal pMultiSelect As Boolean = False)
    If mvUseWebService Then
      Dim vTable As DataTable
      If pList Is Nothing Then pList = New ParameterList(HttpContext.Current)
      vTable = GetDataTable((mvNS.GetLookupData(pType, pList.XMLParameterString)))
      If vTable IsNot Nothing Then
        If pAddBlankRow Then vTable.Rows.InsertAt(vTable.NewRow(), 0)
        pListBox.DataTextField = pDataAndTextField("TextField").ToString
        pListBox.DataValueField = pDataAndTextField("ValueField").ToString
        pListBox.DataSource = vTable
        pListBox.DataBind()
        If pMultiSelect Then pListBox.SelectionMode = ListSelectionMode.Multiple
      End If
    Else
      DataHelperDirect.FillList(pType, pListBox, pDataAndTextField, pAddBlankRow, pList, pMultiSelect)
    End If
  End Sub

  Public Shared Sub FillComboWithRestriction(ByVal pType As CareNetServices.XMLLookupDataTypes, ByVal pCombo As DropDownList, ByVal pAddBlankRow As Boolean, ByVal pList As ParameterList, ByVal pRestriction As String)
    If mvUseWebService Then
      Dim vTable As DataTable
      If pList Is Nothing Then pList = New ParameterList(HttpContext.Current)
      vTable = GetDataTable((mvNS.GetLookupData(pType, pList.XMLParameterString)))
      If vTable IsNot Nothing Then
        If pAddBlankRow Then vTable.Rows.InsertAt(vTable.NewRow(), 0)
        If pRestriction.Length > 0 Then vTable.DefaultView.RowFilter = pRestriction
        pCombo.DataSource = vTable
        pCombo.DataBind()
      End If
    Else
      DataHelperDirect.FillComboWithRestriction(pType, pCombo, pAddBlankRow, pList, pRestriction)
    End If
  End Sub

  Public Shared Function FindData(ByVal pType As CareNetServices.XMLDataFinderTypes, ByVal pList As ParameterList) As String
    If mvUseWebService Then
      Return mvNS.FindData(pType, pList.XMLParameterString)
    Else
      Return DataHelperDirect.FindData(pType, pList)
    End If
  End Function

  Public Shared Function FindDataTable(ByVal pType As CareNetServices.XMLDataFinderTypes, ByVal pList As ParameterList) As DataTable
    If mvUseWebService Then
      Return GetDataTable(mvNS.FindData(pType, pList.XMLParameterString))
    Else
      Return DataHelperDirect.FindDataTable(pType, pList)
    End If
  End Function

  Public Shared Function GetLookupData(ByVal pType As CareNetServices.XMLLookupDataTypes, Optional ByVal pList As ParameterList = Nothing, Optional ByVal pAddBlankRow As Boolean = False) As DataTable
    If pList Is Nothing Then pList = New ParameterList(HttpContext.Current)
    If pType <> CareNetServices.XMLLookupDataTypes.xldtMerchantDetails Then pList("SmartClient") = "Y"
    If mvUseWebService Then
      Return GetDataTable(mvNS.GetLookupData(pType, pList.XMLParameterString))
    Else
      Return DataHelperDirect.GetLookupData(pType, pList)
    End If
  End Function

  Public Shared Function GetMemberBalance(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.GetMemberBalance(pList.XMLParameterString))
    Else
      Return DataHelperDirect.GetMemberBalance(pList)
    End If
  End Function

  Public Shared Function GetNearest(ByVal pList As ParameterList, ByVal pGRd As DataGrid) As Integer
    Dim vResult As String = ""
    If mvUseWebService Then
      vResult = mvNS.GetNearestOrganisation(pList.XMLParameterString)
      Return FillGrid(vResult, pGRd)
    Else
      Return DataHelperDirect.GetNearest(pList, pGRd)
    End If
  End Function

  Public Shared Function GetNextPaymentData(ByVal pList As ParameterList) As DataTable
    If mvUseWebService Then
      Return GetDataTable(mvNS.GetNextPaymentData(pList.XMLParameterString))
    Else
      Return DataHelperDirect.GetNextPaymentData(pList)
    End If
  End Function

  Public Shared Function GetWebControls(ByVal pPageType As CareNetServices.WebControlTypes, ByVal pList As ParameterList) As DataTable
    If mvUseWebService Then
      Return GetDataTable(mvNS.GetWebControls(pPageType, pList.XMLParameterString))
    Else
      Return DataHelperDirect.GetWebControls(pPageType, pList)
    End If
  End Function

  Public Shared Function GetWebInfo(ByVal pList As ParameterList) As DataTable
    If mvUseWebService Then
      Return GetDataTable(mvNS.GetWebInfo(pList.XMLParameterString))
    Else
      Return DataHelperDirect.GetWebInfo(pList)
    End If
  End Function

  Public Shared Function GetWebMenus(ByVal pList As ParameterList) As DataTable
    If mvUseWebService Then
      Return GetDataTable(mvNS.GetWebMenus(pList.XMLParameterString))
    Else
      Return DataHelperDirect.GetWebMenus(pList)
    End If
  End Function

  Public Shared Function GetWebPageInfo(ByVal pList As ParameterList) As DataTable
    If mvUseWebService Then
      Return GetDataTable(mvNS.GetWebPageInfo(pList.XMLParameterString))
    Else
      Return DataHelperDirect.GetWebPageInfo(pList)
    End If
  End Function

  Public Shared Function GetWebPageItems(ByVal pList As ParameterList) As DataTable
    If mvUseWebService Then
      Return GetDataTable(mvNS.GetWebPageItems(pList.XMLParameterString))
    Else
      Return DataHelperDirect.GetWebPageItems(pList)
    End If
  End Function

  Public Shared Function Login(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.Login(pList.XMLParameterString))
    Else
      Return DataHelperDirect.Login(pList)
    End If
  End Function

  Public Shared Function LoginRegisteredUser(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.LoginRegisteredUser(pList.XMLParameterString))
    Else
      Return DataHelperDirect.LoginRegisteredUser(pList)
    End If
  End Function

  Public Shared Function MovePosition(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.MovePosition(pList.XMLParameterString))
    Else
      Return DataHelperDirect.MovePosition(pList)
    End If
  End Function

  Public Shared Function UpdateRegisteredUser(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.UpdateRegisteredUser(pList.XMLParameterString))
    Else
      Return DataHelperDirect.UpdateRegisteredUser(pList)
    End If
  End Function
  Public Shared Function AddCreditCustomer(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddCreditCustomer(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddCreditCustomer(pList)
    End If
  End Function
  Public Shared Function AddDelegateActivity(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.AddDelegateActivity(pList.XMLParameterString))
    Else
      Return DataHelperDirect.AddDelegateActivity(pList)
    End If
  End Function

  Public Shared Function GetRowFromDataTable(ByVal pTable As DataTable) As DataRow
    If pTable IsNot Nothing AndAlso pTable.Rows.Count = 1 Then
      Return pTable.Rows(0)
    Else
      Return Nothing
    End If
  End Function

  Public Shared Function UpdateDirectoryPreferences(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Return New ParameterList(mvNS.UpdateDirectoryPreferences(pList.XMLParameterString))
    Else
      Return DataHelperDirect.UpdateDirectoryPreferences(pList)
    End If
  End Function

  Public Shared Function GetContactData(ByVal pType As CareNetServices.XMLContactDataSelectionTypes, ByVal pGrd As DataGrid, Optional ByVal pDataRestriction As String = "", Optional ByVal pContactNumber As Long = 0, Optional ByVal pAddressNumber As Long = 0, Optional ByVal pList As ParameterList = Nothing, Optional ByVal pEditPageNumber As Integer = 0, Optional ByVal pDisplayEditColumn As Boolean = False, Optional ByVal pCommandNameForEditColumn As String = "Edit", Optional ByVal pEditColumnRestriction As String = "") As Integer
    If pList Is Nothing Then
      pList = New ParameterList(HttpContext.Current)
    End If
    pList("ContactNumber") = pContactNumber
    pList("SystemColumns") = "Y"
    Dim vResult As String = DataHelper.SelectContactData(pType, pList)
    Return FillGrid(vResult, pGrd, pDataRestriction, pEditPageNumber, pDisplayEditColumn, pCommandNameForEditColumn, pEditColumnRestriction)
  End Function

  Public Shared Function ProcessBulkEMail(ByVal pFileName As String, ByVal vList As ParameterList, Optional ByVal pIgnoreProcessJob As Boolean = False) As ParameterList
    If mvUseWebService Then
      Dim vResult As String = ""
      vList("IgnoreProcessJob") = BooleanString(pIgnoreProcessJob)
      Dim vBuffer As Byte()
      Using vFS As New FileStream(pFileName, FileMode.Open)
        ReDim vBuffer(CInt(vFS.Length - 1))
        vFS.Read(vBuffer, 0, CInt(vFS.Length))
        vFS.Close()
      End Using
      vResult = mvNS.ProcessBulkEMail(vList.XMLParameterString, vBuffer)
      Return New ParameterList(vResult)
    Else
      Return DataHelperDirect.ProcessBulkEMail(pFileName, vList, pIgnoreProcessJob)
    End If
  End Function

  Public Shared Function GetBranchFromPostCode(ByVal pList As ParameterList) As ParameterList
    Dim vResult As String = ""
    If mvUseWebService Then
      vResult = mvNS.GetBranchFromPostCode(pList.XMLParameterString)
      Return New ParameterList(vResult)
    Else
      Return DataHelperDirect.GetBranchFromPostCode(pList)
    End If
  End Function
  Public Shared Function GetTransactionData(ByVal pType As CareNetServices.XMLTransactionDataSelectionTypes, ByVal pList As ParameterList) As String
    Dim vResult As String = ""
    If mvUseWebService Then
      Return mvNS.SelectTransactionData(pType, pList.XMLParameterString)
    Else
      Return DataHelperDirect.GetTransactionData(pType, pList)
    End If
  End Function

  Public Shared Function FillGrid(ByVal pResult As String, ByVal pGrd As BaseDataList, Optional ByVal pDataRestriction As String = "", Optional ByVal pEditPageNumber As Integer = 0, Optional ByVal pDisplayEditColumn As Boolean = False, Optional ByVal pCommandNameForEditColumn As String = "Edit", Optional ByVal pEditColumnRestriction As String = "") As Integer
    Dim vStringReader As New System.IO.StringReader(pResult)
    Dim vDataSet As New DataSet
    Dim vRowCount As Integer
    Dim vCommandNameForEditColumn As String = ""
    Dim vGrd As DataGrid = Nothing
    'Dim vList As DataList = Nothing
    vDataSet.ReadXml(vStringReader, XmlReadMode.Auto)
    If vDataSet.Tables.Count > 0 Then
      If vDataSet.Tables(0).Columns.Contains("ErrorMessage") Then
        With vDataSet.Tables(0).Rows(0)
          Throw New CareException(.Item("ErrorMessage").ToString, CLng(.Item("ErrorNumber").ToString), .Item("Source").ToString, .Item("Module").ToString, .Item("Method").ToString)
        End With
      End If
      If pEditPageNumber > 0 Then
        If vDataSet.Tables.Contains("DataRow") AndAlso vDataSet.Tables("DataRow").Columns.Contains("ActionNumber") Then
          'If we have an ActionNumber columns then create a new EditColumn and populate it with the ActionNumber (this will become the edit column)
          'Dim vCol As New DataColumn("EditColumn")
          vDataSet.Tables("DataRow").Columns.Add("EditColumn")
          For Each vRow As DataRow In vDataSet.Tables("DataRow").Rows
            vRow("EditColumn") = vRow("ActionNumber").ToString
          Next
          'Now add the row to the Columns table for the new EditColumn, otherwise the column does not get displayed
          If vDataSet.Tables.Contains("Column") Then
            Dim vTable As DataTable = vDataSet.Tables("Column")
            Dim vRow As DataRow = vTable.NewRow
            vRow("Name") = "EditColumn"
            vRow("Visible") = "Y"
            vTable.Rows.InsertAt(vRow, 0)
          End If
        End If
      End If
      If vDataSet.Tables.Contains("DataRow") AndAlso _
          (vDataSet.Tables("DataRow").Columns.Contains("SessionNumber") OrElse pGrd.ID = "RelatedContactData" OrElse pGrd.ID = "PaymentPlans") Then
        vDataSet.Tables("DataRow").Columns.Add("CheckColumn")
        For Each vRow As DataRow In vDataSet.Tables("DataRow").Rows
          vRow("CheckColumn") = vRow("CheckColumn").ToString
        Next
        'Now add the row to the Columns table for the new CheckColumn, otherwise the column does not get displayed
        If vDataSet.Tables.Contains("Column") Then
          Dim vTable As DataTable = vDataSet.Tables("Column")
          Dim vRow As DataRow = vTable.NewRow
          vRow("Name") = "CheckColumn"
          vRow("Visible") = "Y"
          vTable.Rows.InsertAt(vRow, 0)
        End If
      End If
      If vDataSet.Tables.Contains("Column") Then
        Dim vTable As DataTable = vDataSet.Tables("Column")
        Dim vGridCols As New Dictionary(Of String, DataGridColumn)
        'If vGrd IsNot Nothing AndAlso vGrd.Columns.Count = 0 Then
        Dim vDetail As Boolean
        Dim vRow As DataRow
        Dim vHeading As String = ""
        Dim vAttr1 As String = ""
        Dim vName As String
        Dim vVisible As Boolean
        Dim vTemplate As Boolean
        Dim vRequireCommandForEditColumn As Boolean
        Dim vEditColumnValue As String = ""
        Dim vCommandArgColumn As String = String.Empty

        If pDisplayEditColumn AndAlso vDataSet.Tables.Contains("DataRow") Then
          vDataSet.Tables("DataRow").Columns.Add("EditColumn")

          If vDataSet.Tables("DataRow").Columns.Contains("AddressNumber") Then
            For Each vRowAddress As DataRow In vDataSet.Tables("DataRow").Rows
              vRowAddress("EditColumn") = vRowAddress("AddressNumber").ToString
            Next
            vRequireCommandForEditColumn = True
            If pGrd.ID = "PayerData" OrElse pGrd.ID = "DuplicateOrganisations" OrElse pCommandNameForEditColumn.Length > 0 Then
              vCommandNameForEditColumn = pCommandNameForEditColumn
            Else
              vCommandNameForEditColumn = "Edit"
            End If
          End If

          Select Case pGrd.ID
            Case "ContactCPDPoints", "ContactCPDCycle", "ContactCPDObjectives", "UserOrganisationData"
              vRequireCommandForEditColumn = True
              vCommandNameForEditColumn = pCommandNameForEditColumn
            Case "ContactData"
              vRequireCommandForEditColumn = True
              vCommandNameForEditColumn = pCommandNameForEditColumn
              vCommandArgColumn = "ContactNumber"
            Case "PayPlanData"
              vRequireCommandForEditColumn = True
              vCommandNameForEditColumn = pCommandNameForEditColumn
              vCommandArgColumn = "PaymentPlanNumber"
          End Select

          'Now add the row to the Columns table for the new EditColumn, otherwise the column does not get displayed
          If vDataSet.Tables.Contains("Column") Then
            Dim vTableAddress As DataTable = vDataSet.Tables("Column")
            Dim vRowEdit As DataRow = vTable.NewRow
            vRowEdit("Name") = "EditColumn"
            vRowEdit("Visible") = "Y"
            vTable.Rows.InsertAt(vRowEdit, 0)
          End If
        End If

        For Each vRow In vTable.Rows
          vName = vRow.Item("Name").ToString
          vVisible = ((vRow.Item("Visible").ToString = "Y") Or (vRow.Item("Visible").ToString = "")) AndAlso vDetail = False
          vTemplate = False
          If pGrd.ID = "UserOrganisationData" Then
            If vName = "organisation_number" Then vVisible = False
          End If
          If vName = "DisplayTitle" Then
            pGrd.ToolTip = vRow.Item("Value").ToString
          ElseIf vName = "RowCount" Then
            vRowCount = IntegerValue(vRow.Item("Value").ToString)
          ElseIf vName = "MaintenanceDesc" Then
            ' Do Nothing
          ElseIf vName = "DataSelection" Then
            ' Do Nothing
          Else
            If vName = "DetailItems" Then
              vDetail = True
              vVisible = False
            End If
            If vAttr1.Length > 0 Then
              Dim vTCol As New TemplateColumn
              vTCol.HeaderText = vHeading & "<br>" & vRow.Item("Heading").ToString
              vTCol.ItemTemplate = New TwoAttributeTemplate(vAttr1, vName)
              vTCol.Visible = vVisible
              'vGrd.Columns.Add(vTCol)
              vGridCols.Add(vName, vTCol)
              vAttr1 = ""
            Else
              vHeading = vRow.Item("Heading").ToString
              If vHeading.EndsWith("+") Then
                vHeading = vHeading.TrimEnd("+"c)
                vAttr1 = vName
              Else
                If vVisible AndAlso vRow.Item("DataType").ToString = "Memo" Then
                  Dim vTCol As New TemplateColumn
                  vTCol.HeaderText = vHeading
                  vTCol.ItemTemplate = New MemoTemplate(vName)
                  vTCol.Visible = vVisible
                  'vGrd.Columns.Add(vTCol)
                  vGridCols.Add(vName, vTCol)
                ElseIf vVisible AndAlso vName = "WebURL" Then
                  Dim vTCol As New TemplateColumn
                  vTCol.HeaderText = vHeading
                  vTCol.ItemTemplate = New URLOrEMailTemplate(vName)
                  vTCol.Visible = vVisible
                  'vGrd.Columns.Add(vTCol)
                  vGridCols.Add(vName, vTCol)
                ElseIf vName = "EditColumn" AndAlso (pEditPageNumber > 0 Or pDisplayEditColumn) Then
                  Dim vTCol As New TemplateColumn
                  vTCol.HeaderText = ""     'Do not display a header
                  vTCol.ItemTemplate = New EditTemplate(vName, pEditPageNumber, vRequireCommandForEditColumn, vCommandNameForEditColumn, vCommandArgColumn, pEditColumnRestriction)
                  vTCol.Visible = True      'Always make this visible
                  'vGrd.Columns.Add(vTCol)
                  vGridCols.Add(vName, vTCol)
                ElseIf vVisible AndAlso (vName = "CheckColumn" OrElse vName = "PayCheck") Then
                  Dim vTCol As New TemplateColumn
                  vTCol.HeaderText = vHeading
                  vTCol.ItemTemplate = New CheckBoxTemplate(vName)
                  vTCol.Visible = vVisible
                  'vGrd.Columns.Add(vTCol)
                  vGridCols.Add(vName, vTCol)
                Else
                  Dim vTCol As New TemplateColumn
                  vTCol.HeaderText = vHeading
                  vTCol.ItemTemplate = New DisplayTemplate(vName)
                  vTCol.SortExpression = vName
                  Select Case vRow.Item("DataType").ToString
                    Case "Integer", "Long", "Numeric"
                      vTCol.ItemStyle.HorizontalAlign = HorizontalAlign.Right
                  End Select
                  vTCol.Visible = vVisible
                  'vGrd.Columns.Add(vBCol)
                  vGridCols.Add(vName, vTCol)
                End If
              End If
            End If
          End If
        Next

        If TypeOf pGrd Is DataGrid Then
          vGrd = CType(pGrd, DataGrid)
          If vGrd.Columns.Count = 0 Then
            For Each vCol As DataGridColumn In vGridCols.Values
              vGrd.Columns.Add(vCol)
            Next
          End If                  'End if the Grid has zero columns
        Else
          Dim vCols As New StringBuilder
          Dim vHeaders As New StringBuilder
          For Each vCol As String In vGridCols.Keys
            If vGridCols(vCol).Visible Then
              If vCols.Length > 0 Then
                vCols.Append(",")
                vHeaders.Append(",")
              End If
              vCols.Append(vCol)
              vHeaders.Append(vGridCols(vCol).HeaderText)
            End If
          Next
          Dim vList As DataList = CType(pGrd, DataList)
          'vList.HeaderTemplate = New AlternateDisplayFormatTemplate(ListItemType.Header)
          vList.ItemTemplate = New AlternateDisplayFormatTemplate(ListItemType.Item, vCols.ToString, vHeaders.ToString)
          'vList.FooterTemplate = New AlternateDisplayFormatTemplate(ListItemType.Footer)
          vList.Visible = True
        End If

        vDataSet.Tables.Remove(vTable)      'Remove the columns table
      End If       'Has columns table
      If vDataSet.Tables.Count > 0 Then
        If pDataRestriction.Length > 0 Then
          If pDataRestriction.StartsWith("PAGE") Then
            pGrd.DataSource = vDataSet
            pGrd.DataBind()
            Return vRowCount
          Else
            Dim vView As DataView = SetDataViewRowFilter(vDataSet, pDataRestriction)
            pGrd.DataSource = vView
            pGrd.DataBind()
            Return vView.Count
          End If
        Else
          pGrd.DataSource = vDataSet
          pGrd.DataBind()
          Return vDataSet.Tables(0).Rows.Count
        End If
      Else
        Dim vDT As New DataTable
        Dim vDV As New DataView(vDT)
        pGrd.DataSource = vDV
        pGrd.DataBind()
      End If
    End If
  End Function

  Public Shared Function GetDataTable(ByVal pResult As String) As DataTable
    Dim vStream As System.IO.StringReader = New System.IO.StringReader(pResult)
    Dim vDataSet As New DataSet
    Dim vTable As DataTable = Nothing
    vDataSet.ReadXml(vStream, XmlReadMode.Auto)
    If vDataSet.Tables.Count > 0 Then
      If vDataSet.Tables(0).Columns.Contains("ErrorMessage") Then
        Throw New CareException(vDataSet.Tables(0).Rows(0).Item("ErrorMessage").ToString)
      Else
        vTable = vDataSet.Tables(0)
      End If
    End If
    Return vTable
  End Function

  Public Shared Function AccountNoVerify(ByVal pSortCode As String, ByVal pAccountNo As String) As ParameterList
    Dim vList As New ParameterList(HttpContext.Current)
    vList("SortCode") = pSortCode
    vList("AccountNumber") = pAccountNo
    If mvUseWebService Then
      Return New ParameterList(mvNS.AccountNoVerify(vList.XMLParameterString))
    Else
      Return DataHelperDirect.AccountNoVerify(pSortCode, pAccountNo)
    End If
  End Function

  Public Shared Function SetDataViewRowFilter(ByVal pDataSet As DataSet, ByVal pRowFilter As String) As DataView
    Dim vView As DataView = Nothing
    Try
      vView = pDataSet.Tables(0).DefaultView
      vView.RowFilter = pRowFilter
    Catch vEval As EvaluateException
      RaiseError(DataAccessErrors.daeInvalidContactDataSelectionFilter, pRowFilter)
    Catch vEx As Exception
      PreserveStackTrace(vEx)
      Throw vEx
    End Try
    Return vView
  End Function

  Public Shared Function GetPagedFinderData(ByVal pType As CareNetServices.XMLDataFinderTypes, ByVal pGrd As BaseDataList, ByVal pRequest As HttpRequest, ByVal pPlaceHolder As PlaceHolder, ByVal pList As ParameterList, Optional ByVal pPageSize As Integer = 10, Optional ByVal pEditPageNumber As Integer = 0, Optional ByVal pDisplayEditColumn As Boolean = False, Optional ByVal pQueryString As String = "") As Integer
    Dim vNext As New HyperLink
    Dim vPrev As New HyperLink
    Dim vPageNumber As Integer

    If pRequest.QueryString("PAGE") IsNot Nothing Then vPageNumber = IntegerValue(pRequest.QueryString("PAGE"))
    Dim vUrl As String = pRequest.Url.GetLeftPart(UriPartial.Path) 'Get the URL without the query string params
    'Remove the PAGE param but preserve the other querystring params 
    Dim vParams As NameValueCollection = HttpUtility.ParseQueryString(pRequest.QueryString.ToString)
    If vParams.Item("PAGE") IsNot Nothing Then vParams.Remove("PAGE")
    If vParams.Item("Product") IsNot Nothing Then vParams.Remove("Product")
    If vParams.Item("Event") IsNot Nothing Then vParams.Remove("Event")
    If vParams.Item("Document") IsNot Nothing Then vParams.Remove("Document")
    vUrl = vUrl & "?" & vParams.ToString

    vNext.Text = "Next "
    vNext.CssClass = "TableFooter"
    vPrev.Text = "Previous "
    vPrev.CssClass = "TableFooter"

    Dim vRestriction As String = ""
    If vPageNumber > 0 Then
      vRestriction = "PAGE," & vPageNumber & "," & pPageSize
    Else
      'If pGrd.AllowCustomPaging = False Then vRestriction = "PAGE,0," & pPageSize
      vRestriction = "PAGE,0," & pPageSize
    End If

    pList("StartRow") = (vPageNumber * pPageSize) 'Rows start from 0
    pList("NumberOfRows") = pPageSize
    If vUrl.Contains(pQueryString) Then pQueryString = ""
    Dim vResult As String = DataHelper.FindData(pType, pList)
    Dim vCount As Integer = DataHelper.FillGrid(vResult, pGrd, vRestriction, pEditPageNumber, pDisplayEditColumn)
    If vCount > pPageSize Then
      pPlaceHolder.Controls.Clear()
      If vPageNumber > 0 Then
        vPrev.NavigateUrl = vUrl & "&PAGE=" & vPageNumber - 1 & pQueryString
        pPlaceHolder.Controls.Add(vPrev)
      End If
      Dim vNumbers As Integer = vCount \ pPageSize            'Round down
      If vNumbers * pPageSize < vCount Then vNumbers += 1
      Dim vIndex As Integer
      Dim vHL As HyperLink
      For vIndex = 1 To vNumbers
        vHL = New HyperLink
        vHL.Text = vIndex.ToString & " "
        vHL.NavigateUrl = vUrl & "&PAGE=" & vIndex - 1 & pQueryString
        vHL.CssClass = "TableFooter"
        If vIndex - 1 = vPageNumber Then vHL.Font.Bold = True
        pPlaceHolder.Controls.Add(vHL)
      Next
      If vCount > ((vPageNumber + 1) * pPageSize) Then
        vNext.NavigateUrl = vUrl & "&PAGE=" & vPageNumber + 1 & pQueryString
        pPlaceHolder.Controls.Add(vNext)
      End If
    Else
      pPlaceHolder.Controls.Clear()
    End If
    Return vCount
  End Function
  Public Shared Function AddEventDelegate(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Dim vResult As String = String.Empty
      vResult = mvNS.AddEventDelegate(pList.XMLParameterString)
      Dim vReturnList As New ParameterList(vResult)
      Return vReturnList
    Else
      Return DataHelperDirect.AddEventDelegate(pList)
    End If
  End Function
  Public Shared Function DeleteEventDelegate(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Dim vResult As String = String.Empty
      vResult = mvNS.DeleteEventDelegate(pList.XMLParameterString)
      Dim vReturnList As New ParameterList(vResult)
      Return vReturnList
    Else
      Return DataHelperDirect.DeleteEventDelegate(pList)
    End If
  End Function
  Public Shared Function AddRegisteredUser(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Dim vResult As String = String.Empty
      vResult = mvNS.AddRegisteredUser(pList.XMLParameterString)
      Dim vReturnList As New ParameterList(vResult)
      Return vReturnList
    Else
      Return DataHelperDirect.AddRegisteredUser(pList)
    End If
  End Function

  Public Shared Function AddErrorLog(ByVal pList As ParameterList) As DataTable
    If mvUseWebService Then
      Dim vResult As DataTable
      vResult = GetDataTable(mvNS.AddErrorLog(pList.XMLParameterString))
      Return vResult
    Else
      Return DataHelperDirect.AddErrorLog(pList)
    End If
  End Function

  Public Shared Function UpdateContactSurveyResponse(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Dim vResult As String
      vResult = mvNS.UpdateContactSurveyResponses((pList.XMLParameterString)) 'GetDataTable(mvNS.AddErrorLog(pList.XMLParameterString))
      Dim vReturnList As New ParameterList(vResult)
      Return vReturnList
    Else
      Return DataHelperDirect.UpdateContactSurveyResponse(pList)
    End If
  End Function

  Public Shared Function UpdateEventDelegate(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Dim vResult As String = String.Empty
      vResult = mvNS.UpdateEventDelegate(pList.XMLParameterString)
      Dim vReturnList As New ParameterList(vResult)
      Return vReturnList
    Else
      Return DataHelperDirect.UpdateEventDelegate(pList)
    End If
  End Function

  Shared Sub New()
    If GetCustomConfigItem("CustomConfiguration/UserWebServices/value") <> "N" Then
      mvUseWebService = True
    Else
      mvUseWebService = False
    End If
  End Sub
  ''' <summary>
  ''' This method will make a webservice call to get the possible Membership start dates based on the fixed_renewal_m config. 
  ''' </summary>
  ''' <param name="pList">Will have current date as parameter</param>
  ''' <returns>List of possible start dates as Parameter List</returns>
  ''' <remarks>fixed_renewal_m config should be set for this method to return possilb start dates</remarks>
  Public Shared Function GetPaymentPlanStartDate(ByVal pList As ParameterList) As ParameterList
    If mvUseWebService Then
      Dim vResult As String = String.Empty
      vResult = mvNS.GetPaymentPlanStartDate(pList.XMLParameterString)
      Return New ParameterList(vResult)
    Else
      Return DataHelperDirect.GetPaymentPlanStartDate(pList)
    End If
  End Function
End Class


Public Class NumberInfo
  Private mvIdentifier As String
  Private mvDeviceCode As String
  Private mvID As Integer
  Private mvMaxLen As Integer
  Private mvNumber As String = ""
  Private mvDeviceDefault As Boolean = False
  Private mvDefault As Boolean = False
  Private mvMail As Boolean = False
  Private mvPreferredMethod As Boolean = False

  Public Sub New(ByVal pIdentifier As String, ByVal pDevice As String, ByVal pMaxLen As Integer)
    mvIdentifier = pIdentifier
    mvDeviceCode = pDevice
    mvMaxLen = pMaxLen
  End Sub

  Public Property CommunicationNumber() As Integer
    Get
      Return mvID
    End Get
    Set(ByVal Value As Integer)
      mvID = Value
    End Set
  End Property

  Public ReadOnly Property DeviceCode() As String
    Get
      Return mvDeviceCode
    End Get
  End Property

  Public Property IsDefault() As Boolean
    Get
      Return mvDefault
    End Get
    Set(ByVal pValue As Boolean)
      mvDefault = pValue
    End Set
  End Property

  Public Property DeviceDefault() As Boolean
    Get
      Return mvDeviceDefault
    End Get
    Set(ByVal pValue As Boolean)
      mvDeviceDefault = pValue
    End Set
  End Property

  Public ReadOnly Property Identifier() As String
    Get
      Return mvIdentifier
    End Get
  End Property

  Public Property Mail() As Boolean
    Get
      Return mvMail
    End Get
    Set(ByVal pValue As Boolean)
      mvMail = pValue
    End Set
  End Property

  Public ReadOnly Property MaxLength() As Integer
    Get
      Return mvMaxLen
    End Get
  End Property

  Public Property Number() As String
    Get
      Return mvNumber
    End Get
    Set(ByVal pValue As String)
      mvNumber = pValue
    End Set
  End Property

  Public Property PreferredMethod() As Boolean
    Get
      Return mvPreferredMethod
    End Get
    Set(ByVal pValue As Boolean)
      mvPreferredMethod = pValue
    End Set
  End Property
End Class
