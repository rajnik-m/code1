Imports CARE.XMLAccess
Imports CARE.Access

Public Class DataHelperDirect

  Public Shared Function GetReportFile(ByVal pList As ParameterList) As String
    Dim vXMLClass As New LookupData
    Return vXMLClass.GetReportFile(pList.XMLParameterString)
  End Function

  Public Shared Function AddItem(ByVal pType As CareNetServices.XMLMaintenanceControlTypes, ByVal pList As ParameterList) As ParameterList
    Dim vCA As New XMLAddData
    Dim vResult As String = ""
    Select Case pType
      Case CareNetServices.XMLMaintenanceControlTypes.xmctActivities
        Dim vCA1 As New XMLAddData
        vResult = vCA1.AddActivity(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctAddresses
        Dim vCA1 As New XMLAddData
        vResult = vCA1.AddAddress(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctContact
        Dim vCA1 As New XMLAddData
        vResult = vCA1.AddContact(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctDocument
        vResult = vCA.AddCommunicationsLog(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctReference
        vResult = vCA.AddExternalReference(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctFundraisingEvents
        Dim vFD As New FundraisingData
        vResult = vFD.AddContactFundraisingEvent(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctGiftAidDeclarations
        vResult = vCA.AddGiftAidDeclaration(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctLink
        vResult = vCA.AddLink(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctNumber
        Dim vCA1 As New XMLAddData
        vResult = vCA1.AddCommunicationsNumber(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctSuppression
        Dim vCA1 As New XMLAddData
        vResult = vCA1.AddSuppression(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctContactAccounts
        vResult = vCA.AddContactAccount(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctOrganisation
        vResult = vCA.AddOrganisation(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctPosition
        vResult = vCA.AddPosition(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctPosition
        vResult = vCA.AddPosition(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctContactSurveys
        Dim vXMLHelper As New XMLHelper
        vResult = vXMLHelper.AddRecord(pList.XMLParameterString, New ContactSurvey(Nothing))
      Case CareNetServices.XMLMaintenanceControlTypes.xmctCPDCycles
        Dim vCPDData As New CPDData
        vResult = vCPDData.AddContactCPDCycle(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctCPDPoints
        Dim vCPDData As New CPDData
        vResult = vCPDData.AddContactCPDPoints(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctCPDObjectives
        Dim vXMLHelper As New XMLHelper
        vResult = vXMLHelper.AddRecord(pList.XMLParameterString, New ContactCpdObjective(Nothing))
    End Select
    Return New ParameterList(vResult)
  End Function

  Public Shared Function AddWebItem(ByVal pType As CareNetServices.XMLWebDataSelectionTypes, ByVal pList As ParameterList) As ParameterList
    Dim vWD As New WebData
    Dim vResult As String = ""
    Select Case pType
      Case CareNetServices.XMLWebDataSelectionTypes.wstPage
        vResult = vWD.AddWebPage(pList.XMLParameterString)
      Case CareNetServices.XMLWebDataSelectionTypes.wstPageItem
        vResult = vWD.AddWebPageItem(pList.XMLParameterString)
    End Select
    Return New ParameterList(vResult)
  End Function

  Public Shared Function UpdateWebItem(ByVal pType As CareNetServices.XMLWebDataSelectionTypes, ByVal pList As ParameterList) As ParameterList
    Dim vWD As New WebData
    Dim vResult As String = ""
    Select Case pType
      Case CareNetServices.XMLWebDataSelectionTypes.wstPage
        vResult = vWD.UpdateWebPage(pList.XMLParameterString)
      Case CareNetServices.XMLWebDataSelectionTypes.wstPageItem
        vResult = vWD.UpdateWebPageItem(pList.XMLParameterString)
    End Select
    Return New ParameterList(vResult)
  End Function

  Public Shared Function UpdateProvisionalTransaction(ByVal pList As ParameterList) As ParameterList
    Dim vCA As New XMLUpdateData
    Dim vResult As String = ""
    vResult = vCA.UpdateProvisionalTransaction(pList.XMLParameterString)
    Return New ParameterList(vResult)
  End Function

  Public Shared Function UpdateItem(ByVal pType As CareNetServices.XMLMaintenanceControlTypes, ByVal pList As ParameterList) As ParameterList
    Dim vCA As New XMLUpdateData
    Dim vResult As String = ""
    Select Case pType
      Case CareNetServices.XMLMaintenanceControlTypes.xmctAction
        vResult = vCA.UpdateAction(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctActivities
        Dim vCA1 As New XMLUpdateData
        vResult = vCA1.UpdateActivity(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctAddresses
        Dim vCA1 As New XMLUpdateData
        pList("CarePortal") = "Y"
        vResult = vCA1.UpdateAddress(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctContact
        Dim vCA1 As New XMLUpdateData
        vResult = vCA1.UpdateContact(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctFundraisingEvents
        Dim vFD As New FundraisingData
        vFD.UpdateContactFundraisingEvent(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctLink
        vResult = vCA.UpdateLink(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctNumber
        Dim vCA1 As New XMLUpdateData
        vResult = vCA1.UpdateCommunications(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctContactAccounts
        vResult = vCA.UpdateContactAccount(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctReference
        vResult = vCA.UpdateExternalReference(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctSuppression
        Dim vCA1 As New XMLUpdateData
        vResult = vCA1.UpdateSuppression(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctEventDelegateActivity
        Dim vCA1 As New XMLEvents
        vResult = vCA1.UpdateDelegateActivity(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctPosition
        Dim vXmlUpdatedata As New XMLUpdateData
        vResult = vXmlUpdatedata.UpdatePosition(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctContactSurveys
        Dim vXMLHelper As New XMLHelper
        vResult = vXMLHelper.UpdateRecord(pList.XMLParameterString, New CARE.Access.ContactSurvey(Nothing))
      Case CareNetServices.XMLMaintenanceControlTypes.xmctCPDCycles
        Dim vCPDData As New CPDData()
        vResult = vCPDData.UpdateContactCPDCycle(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctCPDPoints
        Dim vCPDData As New CPDData()
        vResult = vCPDData.UpdateContactCPDPoints(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctCPDObjectives
        Dim vXMLHelper As New XMLHelper
        vResult = vXMLHelper.UpdateRecord(pList.XMLParameterString, New CARE.Access.ContactCpdObjective(Nothing))
    End Select
    Return New ParameterList(vResult)
  End Function

  Public Shared Function UpdateWebDocument(ByVal pList As ParameterList) As ParameterList
    Dim vCA As New XMLUpdateData
    Dim vResult As String = ""
    vResult = vCA.UpdateWebDocument(pList.XMLParameterString)
    Return New ParameterList(vResult)
  End Function

  Public Shared Function DeleteItem(ByVal pType As CareNetServices.XMLTransactionDataSelectionTypes, ByVal pList As ParameterList) As ParameterList
    Dim vCA As New XMLDeleteData
    Dim vResult As String = ""
    Select Case pType
      Case CareNetServices.XMLTransactionDataSelectionTypes.xtdtTransactionAnalysis
        vResult = vCA.DeleteProvisionalTransaction(pList.XMLParameterString)
    End Select
    Return New ParameterList(vResult)
  End Function

  Public Shared Function DeleteItem(ByVal pType As CareNetServices.XMLMaintenanceControlTypes, ByVal pList As ParameterList) As ParameterList
    Dim vCA As New XMLDeleteData
    Dim vResult As String = ""
    Select Case pType
      Case CareNetServices.XMLMaintenanceControlTypes.xmctActionLink
        vResult = vCA.DeleteActionLink(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctActionTopic
        vResult = vCA.DeleteActionSubject(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctActivities
        Dim vCA1 As New XMLDeleteData
        vResult = vCA1.DeleteActivity(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctLink
        vResult = vCA.DeleteLink(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctNumber
        Dim vCA1 As New XMLDeleteData
        vResult = vCA1.DeleteCommunications(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctReference
        vResult = vCA.DeleteExternalReference(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctSuppression
        Dim vCA1 As New XMLDeleteData
        vResult = vCA1.DeleteSuppression(pList.XMLParameterString)
      Case CareNetServices.XMLMaintenanceControlTypes.xmctEventDelegateActivity
        Dim vCA1 As New XMLEvents
        vResult = vCA1.DeleteDelegateActivity(pList.XMLParameterString)
    End Select
    Return New ParameterList(vResult)
  End Function

  Public Shared Function GetAddressDataTable(ByVal pType As CareNetServices.XMLAddressDataSelectionTypes, ByVal pAddressNumber As Long) As DataTable
    Dim vList As New ParameterList(HttpContext.Current)
    vList("AddressNumber") = pAddressNumber
    Dim vCX As New XMLDataSelection
    Return GetDataTable(vCX.SelectAddressData(CType(pType, XMLDataSelection.XMLAddressDataSelectionTypes), vList.XMLParameterString))
  End Function

  Public Shared Function GetMembershipData(ByVal pType As CareNetServices.XMLMembershipDataSelectionTypes, ByVal pMembershipNumber As Integer, Optional ByVal pList As ParameterList = Nothing) As DataTable
    Dim vCX As New SelectData
    If pList Is Nothing Then pList = New ParameterList(HttpContext.Current)
    pList("MembershipNumber") = pMembershipNumber
    Dim vResult As String = vCX.SelectMembershipData(CType(pType, SelectData.XMLMembershipDataSelectionTypes), pList.XMLParameterString)
    Return GetDataTable(vResult)
  End Function

  Public Shared Function SelectContactData(ByVal pType As CareNetServices.XMLContactDataSelectionTypes, ByVal pList As ParameterList) As String
    Dim vSelectData As New SelectData
    Return vSelectData.SelectContactData(CType(pType, SelectData.XMLContactDataSelectionTypes), pList.XMLParameterString)
  End Function

  Public Shared Function SelectEventData(ByVal pType As CareNetServices.XMLEventDataSelectionTypes, ByVal pList As ParameterList) As String
    Dim vEventData As New EventData
    Return vEventData.SelectEventData(CType(pType, EventData.XMLEventDataSelectionTypes), pList.XMLParameterString)
  End Function

  Public Shared Function SelectExamData(ByVal pType As ExamsAccess.XMLExamDataSelectionTypes, ByVal pList As ParameterList) As String
    Dim vExamData As New ExamData
    Return vExamData.SelectExamData(CType(pType, ExamData.XMLExamDataSelectionTypes), pList.XMLParameterString)
  End Function

  Public Shared Function SelectFundraisingEventData(ByVal pType As CareNetServices.XMLFundraisingEventDataSelectionTypes, ByVal pList As ParameterList) As String
    Dim vSD As New SelectData
    Return vSD.SelectFundraisingEventData(CType(pType, SelectData.XMLFundraisingEventDataSelectionTypes), pList.XMLParameterString)
  End Function

  Public Shared Function SelectTableData(ByVal pType As CareNetServices.XMLTableDataSelectionTypes, ByVal pList As ParameterList) As DataTable
    Dim vSelectData As New SelectData
    Return Utilities.GetDataTable(vSelectData.SelectTableData(CType(pType, SelectData.XMLTableDataSelectionTypes), pList.XMLParameterString), True)
  End Function

  Public Shared Function SelectTableDataString(ByVal pType As CareNetServices.XMLTableDataSelectionTypes, ByVal pList As ParameterList) As String
    Dim vSelectData As New SelectData
    Return vSelectData.SelectTableData(CType(pType, SelectData.XMLTableDataSelectionTypes), pList.XMLParameterString)
  End Function

  Public Shared Function GetEventDataTable(ByVal pType As CareNetServices.XMLEventDataSelectionTypes, ByVal pList As ParameterList) As DataTable
    Dim vCX As New EventData
    Return GetDataTable(vCX.SelectEventData(CType(pType, EventData.XMLEventDataSelectionTypes), pList.XMLParameterString))
  End Function

  Public Shared Function SelectWebDataTable(ByVal pType As CareNetServices.XMLWebDataSelectionTypes, ByVal pList As ParameterList) As DataTable
    Dim vWD As New WebData
    Return GetDataTable(vWD.SelectWebData(CType(pType, WebData.XMLWebDataSelectionTypes), pList.XMLParameterString))
  End Function

  Public Shared Function AddEventBooking(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New XMLAddData
    Return New ParameterList(vAD.AddEventBooking(pList.XMLParameterString))
  End Function

  Public Shared Function AddExamBooking(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New ExamData
    Return New ParameterList(vAD.AddExamBooking(pList.XMLParameterString))
  End Function

  Public Shared Function CalculateExamBookingPrice(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New ExamData
    Return New ParameterList(vAD.CalculateExamBookingPrice(pList.XMLParameterString))
  End Function

  Public Shared Function AddAction(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New XMLAddData
    Return New ParameterList(vAD.AddAction(pList.XMLParameterString))
  End Function

  Public Shared Function AddActionLink(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New XMLAddData
    Return New ParameterList(vAD.AddActionLink(pList.XMLParameterString))
  End Function

  Public Shared Function AddActionSubject(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New XMLAddData
    Return New ParameterList(vAD.AddActionSubject(pList.XMLParameterString))
  End Function

  Public Shared Function AddActivity(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New XMLAddData
    Return New ParameterList(vAD.AddActivity(pList.XMLParameterString))
  End Function

  Public Shared Function AddEvent(ByVal pList As ParameterList) As ParameterList
    Dim vEV As New XMLEvents
    Return New ParameterList(vEV.AddEvent(pList.XMLParameterString))
  End Function

  Public Shared Function AddEventBookingOption(ByVal pList As ParameterList) As ParameterList
    Dim vEV As New XMLEvents
    Return New ParameterList(vEV.AddEventBookingOption(pList.XMLParameterString))
  End Function

  Public Shared Function UpdateEvent(ByVal pList As ParameterList) As ParameterList
    Dim vEV As New XMLEvents
    Return New ParameterList(vEV.UpdateEvent(pList.XMLParameterString))
  End Function

  Public Shared Function AddDirectDebit(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New XMLPaymentPlans
    Return New ParameterList(vAD.AddAutoPaymentMethod(PaymentPlan.ppAutoPayMethods.ppAPMDD, pList.XMLParameterString))
  End Function

  Public Shared Function AddLink(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New XMLAddData
    Return New ParameterList(vAD.AddLink(pList.XMLParameterString))
  End Function

  Public Shared Function AddSuppresion(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New XMLAddData
    Return New ParameterList(vAD.AddSuppression(pList.XMLParameterString))
  End Function

  Public Shared Function AddMember(ByVal pList As ParameterList) As ParameterList
    Dim vPP As New XMLPaymentPlans
    Return New ParameterList(vPP.AddMembership(pList.XMLParameterString))
  End Function

  Public Shared Function AddPaymentPlan(ByVal pType As CareNetServices.ppType, ByVal pList As ParameterList) As ParameterList
    Dim vPP As New XMLPaymentPlans
    Return New ParameterList(vPP.AddPaymentPlan(CType(pType, CDBEnvironment.ppType), pList.XMLParameterString))
  End Function

  Public Shared Function AddInvoicePayment(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New XMLAddData
    Return New ParameterList(vAD.AddInvoicePayment(pList.XMLParameterString))
  End Function

  Public Shared Function AddPaymentPlanPayment(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New XMLAddData
    Return New ParameterList(vAD.AddPaymentPlanPayment(pList.XMLParameterString))
  End Function

  Public Shared Function UpdatePaymentPlan(ByVal pType As CareNetServices.XMLPaymentPlanUpdateTypes, ByVal pList As ParameterList) As ParameterList
    Dim vPP As New XMLPaymentPlans
    Return New ParameterList(vPP.UpdatePaymentPlan(CType(pType, XMLPaymentPlans.XMLPaymentPlanUpdateTypes), pList.XMLParameterString))
  End Function

  Public Shared Function AddProductSale(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New XMLAddData
    Return New ParameterList(vAD.AddProductSale(pList.XMLParameterString))
  End Function

  Public Shared Function ConfirmCardSaleTransaction(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New XMLAddData
    Return New ParameterList(vAD.ConfirmCardSaleTransaction(pList.XMLParameterString))
  End Function

  Public Shared Function ConfirmCreditSaleTransaction(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New XMLAddData
    Return New ParameterList(vAD.ConfirmCreditSaleTransaction(pList.XMLParameterString))
  End Function

  Public Shared Function ConfirmCashSaleTransaction(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New XMLAddData
    Return New ParameterList(vAD.ConfirmCashSaleTransaction(pList.XMLParameterString))
  End Function

  Public Shared Function ConfirmCreditAndCardSaleTransaction(ByVal pList As ParameterList) As ParameterList
    Dim vAD As New XMLAddData
    Return New ParameterList(vAD.ConfirmCreditAndCardSaleTransaction(pList.XMLParameterString))
  End Function

  Public Shared Sub FillCombo(ByVal pType As CareNetServices.XMLLookupDataTypes, ByVal pCombo As DropDownList, Optional ByVal pAddBlankRow As Boolean = False, Optional ByVal pList As ParameterList = Nothing)
    Dim vTable As DataTable
    If pList Is Nothing Then pList = New ParameterList(HttpContext.Current)
    Dim vCX As New LookupData
    vTable = GetDataTable((vCX.SelectLookupData(CType(pType, LookupData.XMLLookupDataTypes), pList.XMLParameterString)))
    If vTable IsNot Nothing Then
      If pAddBlankRow Then vTable.Rows.InsertAt(vTable.NewRow(), 0)
      pCombo.DataSource = vTable
      pCombo.DataBind()
    End If
  End Sub

  Public Shared Sub FillList(ByVal pType As CareNetServices.XMLLookupDataTypes, ByVal pListBox As ListBox, Optional ByVal pAddBlankRow As Boolean = False, Optional ByVal pList As ParameterList = Nothing, Optional ByVal pMultiSelect As Boolean = False)
    Dim vTable As DataTable
    If pList Is Nothing Then pList = New ParameterList(HttpContext.Current)
    Dim vCX As New LookupData
    vTable = GetDataTable((vCX.SelectLookupData(CType(pType, LookupData.XMLLookupDataTypes), pList.XMLParameterString)))
    If vTable IsNot Nothing Then
      If pAddBlankRow Then vTable.Rows.InsertAt(vTable.NewRow(), 0)
      pListBox.DataTextField = "ActivityValueDesc"
      pListBox.DataValueField = "ActivityValue"
      pListBox.DataSource = vTable
      pListBox.DataBind()
      If pMultiSelect Then pListBox.SelectionMode = ListSelectionMode.Multiple
    End If
  End Sub

  Public Shared Sub FillList(ByVal pType As CareNetServices.XMLLookupDataTypes, ByVal pListBox As ListBox, ByVal pDataAndTextField As ParameterList, Optional ByVal pAddBlankRow As Boolean = False, Optional ByVal pList As ParameterList = Nothing, Optional ByVal pMultiSelect As Boolean = False)
    Dim vTable As DataTable
    If pList Is Nothing Then pList = New ParameterList(HttpContext.Current)
    Dim vCX As New LookupData
    vTable = GetDataTable((vCX.SelectLookupData(CType(pType, LookupData.XMLLookupDataTypes), pList.XMLParameterString)))
    If vTable IsNot Nothing Then
      If pAddBlankRow Then vTable.Rows.InsertAt(vTable.NewRow(), 0)
      pListBox.DataTextField = pDataAndTextField("TextField").ToString
      pListBox.DataValueField = pDataAndTextField("ValueField").ToString
      pListBox.DataSource = vTable
      pListBox.DataBind()
      If pMultiSelect Then pListBox.SelectionMode = ListSelectionMode.Multiple
    End If
  End Sub

  Public Shared Sub FillComboWithRestriction(ByVal pType As CareNetServices.XMLLookupDataTypes, ByVal pCombo As DropDownList, ByVal pAddBlankRow As Boolean, ByVal pList As ParameterList, ByVal pRestriction As String)
    Dim vTable As DataTable
    If pList Is Nothing Then pList = New ParameterList(HttpContext.Current)
    Dim vCX As New LookupData
    vTable = GetDataTable((vCX.SelectLookupData(CType(pType, LookupData.XMLLookupDataTypes), pList.XMLParameterString)))
    If vTable IsNot Nothing Then
      If pAddBlankRow Then vTable.Rows.InsertAt(vTable.NewRow(), 0)
      If pRestriction.Length > 0 Then vTable.DefaultView.RowFilter = pRestriction
      pCombo.DataSource = vTable
      pCombo.DataBind()
    End If
  End Sub

  Public Shared Function FindData(ByVal pType As CareNetServices.XMLDataFinderTypes, ByVal pList As ParameterList) As String
    Dim vCF As New FindData
    Return vCF.FindData(CType(pType, FindData.XMLDataFinderTypes), pList.XMLParameterString)
  End Function

  Public Shared Function FindDataTable(ByVal pType As CareNetServices.XMLDataFinderTypes, ByVal pList As ParameterList) As DataTable
    Dim vCF As New FindData
    Return GetDataTable(vCF.FindData(CType(pType, FindData.XMLDataFinderTypes), pList.XMLParameterString))
  End Function

  Public Shared Function GetLookupData(ByVal pType As CareNetServices.XMLLookupDataTypes, Optional ByVal pList As ParameterList = Nothing, Optional ByVal pAddBlankRow As Boolean = False) As DataTable
    If pList Is Nothing Then pList = New ParameterList(HttpContext.Current)
    If pType <> CareNetServices.XMLLookupDataTypes.xldtMerchantDetails Then pList("SmartClient") = "Y"
    Dim mvCS As New LookupData
    Return GetDataTable(mvCS.SelectLookupData(CType(pType, LookupData.XMLLookupDataTypes), pList.XMLParameterString))
  End Function

  Public Shared Function GetMemberBalance(ByVal pList As ParameterList) As ParameterList
    Dim mvCS As New XMLPaymentPlans
    Return New ParameterList(mvCS.GetMemberBalance(pList.XMLParameterString))
  End Function

  Public Shared Function GetNearest(ByVal pList As ParameterList, ByVal pGRd As DataGrid) As Integer
    Dim vResult As String = ""
    Dim vCX As New XMLLookupData
    vResult = vCX.GetNearestOrganisation(pList.XMLParameterString)
    Return DataHelper.FillGrid(vResult, pGRd)
  End Function

  Public Shared Function GetNextPaymentData(ByVal pList As ParameterList) As DataTable
    Dim vCX As New XMLPaymentPlans
    Return GetDataTable(vCX.GetNextPaymentData(pList.XMLParameterString))
  End Function

  Public Shared Function GetWebControls(ByVal pPageType As CareNetServices.WebControlTypes, ByVal pList As ParameterList) As DataTable
    Dim vCX As New WebData
    Return GetDataTable(vCX.GetWebControls(CType(pPageType, WebPageUserControl.WebControlTypes), pList.XMLParameterString))
  End Function

  Public Shared Function GetWebInfo(ByVal pList As ParameterList) As DataTable
    Dim vCX As New WebData
    Return GetDataTable(vCX.GetWebInfo(pList.XMLParameterString))
  End Function

  Public Shared Function GetWebMenus(ByVal pList As ParameterList) As DataTable
    Dim vCX As New WebData
    Return GetDataTable(vCX.GetWebMenus(pList.XMLParameterString))
  End Function

  Public Shared Function GetWebPageInfo(ByVal pList As ParameterList) As DataTable
    Dim vCX As New WebData
    Return GetDataTable(vCX.GetWebPageInfo(pList.XMLParameterString))
  End Function

  Public Shared Function GetWebPageItems(ByVal pList As ParameterList) As DataTable
    Dim vCX As New WebData
    Return GetDataTable(vCX.GetWebPageItems(pList.XMLParameterString))
  End Function

  Public Shared Function Login(ByVal pList As ParameterList) As ParameterList
    Dim vCX As New CARE.XMLAccess.Login
    Return New ParameterList(vCX.Login(pList.XMLParameterString))
  End Function

  Public Shared Function LoginRegisteredUser(ByVal pList As ParameterList) As ParameterList
    Dim vCX As New XMLLogin
    Return New ParameterList(vCX.LoginRegisteredUser(pList.XMLParameterString))
  End Function

  Public Shared Function MovePosition(ByVal pList As ParameterList) As ParameterList
    Dim vCX As New XMLUpdateData
    Return New ParameterList(vCX.MovePosition(pList.XMLParameterString))
  End Function

  Public Shared Function UpdateRegisteredUser(ByVal pList As ParameterList) As ParameterList
    Dim vCX As New XMLLogin
    Return New ParameterList(vCX.UpdateRegisteredUser(pList.XMLParameterString))
  End Function
  Public Shared Function AddCreditCustomer(ByVal pList As ParameterList) As ParameterList
    Dim vCX As New XMLAddData
    Return New ParameterList(vCX.AddCreditCustomer(pList.XMLParameterString))
  End Function
  Public Shared Function AddDelegateActivity(ByVal pList As ParameterList) As ParameterList
    Dim vCX As New XMLEvents
    Return New ParameterList(vCX.AddDelegateActivity(pList.XMLParameterString))
  End Function

  Public Shared Function UpdateDirectoryPreferences(ByVal pList As ParameterList) As ParameterList
    Dim vCX As New UpdateData
    Return New ParameterList(vCX.UpdateDirectoryPreferences(pList.XMLParameterString))
  End Function

  Public Shared Function ProcessBulkEMail(ByVal pFileName As String, ByVal vList As ParameterList, Optional ByVal pIgnoreProcessJob As Boolean = False) As ParameterList
    Dim vResult As String = ""
    Dim vCX As New MailingData
    vList("IgnoreProcessJob") = BooleanString(pIgnoreProcessJob)
    vResult = vCX.ProcessBulkEMail(vList.XMLParameterString, pFileName)
    Return New ParameterList(vResult)
  End Function

  Public Shared Function GetBranchFromPostCode(ByVal pList As ParameterList) As ParameterList
    Dim vResult As String = ""
    Dim vXMLClass As New XMLLookupData
    vResult = vXMLClass.GetBranchFromPostCode(pList.XMLParameterString)
    Return New ParameterList(vResult)
  End Function

  Public Shared Function GetTransactionData(ByVal pType As CareNetServices.XMLTransactionDataSelectionTypes, ByVal pList As ParameterList) As String
    Dim vResult As String = ""
    Dim vCX As New XMLDataSelection
    vResult = vCX.SelectTransactionData(CType(pType, XMLDataSelection.XMLTransactionDataSelectionTypes), pList.XMLParameterString)
    Return vResult
  End Function

  Public Shared Function AccountNoVerify(ByVal pSortCode As String, ByVal pAccountNo As String) As ParameterList
    Dim vList As New ParameterList(HttpContext.Current)
    vList("SortCode") = pSortCode
    vList("AccountNumber") = pAccountNo
    Dim vLookup As New XMLLookupData
    Return New ParameterList(vLookup.AccountNoVerify(vList.XMLParameterString))
  End Function

  Public Shared Function AddEventDelegate(ByVal pList As ParameterList) As ParameterList
    Dim vResult As String = String.Empty
    Dim vEV As New XMLEvents
    vResult = vEV.AddEventDelegate(pList.XMLParameterString)
    Dim vReturnList As New ParameterList(vResult)
    Return vReturnList
  End Function
  Public Shared Function DeleteEventDelegate(ByVal pList As ParameterList) As ParameterList
    Dim vResult As String = String.Empty
    Dim vEV As New XMLEvents
    vResult = vEV.DeleteEventDelegate(pList.XMLParameterString)
    Dim vReturnList As New ParameterList(vResult)
    Return vReturnList
  End Function
  Public Shared Function AddRegisteredUser(ByVal pList As ParameterList) As ParameterList
    Dim vResult As String = String.Empty
    Dim vLo As New XMLLogin
    vResult = vLo.AddRegisteredUser(pList.XMLParameterString)
    Dim vReturnList As New ParameterList(vResult)
    Return vReturnList
  End Function

  Public Shared Function AddErrorLog(ByVal pList As ParameterList) As DataTable
    Dim vResult As DataTable
    Dim vXh As New XMLHelper
    vResult = GetDataTable((vXh.AddRecord(pList.XMLParameterString, New ErrorLog(Nothing))))
    Return vResult
  End Function

  Public Shared Function UpdateContactSurveyResponse(ByVal pList As ParameterList) As ParameterList
    Dim vResult As String
    Dim vXMLUpdateData As New XMLUpdateData
    vResult = vXMLUpdateData.UpdateContactSurveyResponse(pList.XMLParameterString)
    Dim vReturnList As New ParameterList(vResult)
    Return vReturnList

  End Function
  Public Shared Function UpdateEventDelegate(ByVal pList As ParameterList) As ParameterList
    Dim vResult As String = String.Empty
    Dim vEV As New XMLEvents
    vResult = vEV.UpdateEventDelegate(pList.XMLParameterString)
    Dim vReturnList As New ParameterList(vResult)
    Return vReturnList
  End Function

  Public Shared Function GetPaymentPlanStartDate(ByVal pList As ParameterList) As ParameterList
    Dim vResult As String = String.Empty
    Dim vPP As New XMLPaymentPlans
    vResult = vPP.GetPaymentPlanStartDate(pList.XMLParameterString)
    Return New ParameterList(vResult)
  End Function
End Class
