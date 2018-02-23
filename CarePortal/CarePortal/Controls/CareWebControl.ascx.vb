Public Interface ICareChildWebControl
  Sub SubmitChild(ByVal pList As ParameterList)
End Interface

Public Interface ICareParentWebControl
  Sub ProcessChildControls(ByVal pList As ParameterList)
End Interface

Public Interface IMultiViewWebControl
  'Make sure the following is applied to the class implementing this Interface 
  '1. HandleMultiViewDisplay() is called in Page.Load after initialising the controls
  '2. The class (.ascx control) does not refer to tblDataEntry or tblDisplayData etc. when finding a control
  '3. GridHyperlink case is added to the ProcessButtonClickEvent (type of control is LinkButton)
  '4. GridHyperLinkVisibility should return True if New/Search button is visible to the user
  '5. mvMultiView.SetActiveView(mvView2) is called on getting an error/warning which is to be displayed to the user (in a Label control)
  Function GridHyperLinkVisibility() As Boolean
End Interface

Partial Public Class CareWebControl
  Inherits System.Web.UI.UserControl

#Region "Enums and Module Level Variables"

  Public Enum ContentPosition
    cpHeader
    cpFooter
    cpLeft
    cpCenter
    cpRight
  End Enum

  Protected Enum PostBackTriggerEventTypes
    None
    SelectedIndexChanged
    TextChanged
    ButtonClick
  End Enum

  Public Enum DateType
    ValidFrom
    ValidTo
  End Enum

  Public Enum TNSResult
    NONE = -1
    SUCCESSFUL = 0
    SESSION_EXPIRED = 2
    INVALID_FIELD_VALUES = 3
  End Enum

  Public Const TNSHOSTED As String = "TNSHOSTED"


  Private mvPosition As ContentPosition = ContentPosition.cpCenter
  Private mvInitialParameters As ParameterList
  Private mvDefaultParameters As ParameterList
  Private mvRequiredInitialParameters As String
  Private mvRequiredDefaultParameters As String
  Private mvWebPageItemNumber As Integer
  Private mvWebPageNumber As Integer
  Private mvWebPageItemName As String = ""
  Private mvSubmitItemUrl As String = ""
  Private mvParentGroup As String = ""
  Private mvGroupName As String = ""
  Private mvHTML As String = ""
  Private mvURL As String = ""
  Private mvStyle As String = ""
  Private mvNumberOfRows As Integer
  Private mvTableControl As HtmlTable
  Private mvGridControlTable As HtmlTable 'Only used when mvSupportsMultiView is True
  Private mvFocusControl As Control
  Private mvDateControls As List(Of String)
  Private mvUseNewContact As Boolean
  Private mvSupportsOnlineCCAuthorisation As Boolean

  Protected mvSubmitItemNumber As Integer
  Protected mvControlType As CareNetServices.WebControlTypes
  Protected mvNeedsAuthentication As Boolean
  Protected mvContactEntryHidden As Boolean
  Protected mvCenterControl As Boolean
  Protected mvNeedsParent As Boolean
  Protected mvPageCareControls As List(Of CareWebControl)
  Protected mvDependantControls As List(Of ICareChildWebControl)
  Protected mvRadioButtons As List(Of RadioButton)
  Protected mvCheckBoxes As List(Of CheckBox)
  Protected mvHandlesActivities As Boolean
  Protected mvHandlesSuppressions As Boolean
  Protected mvHandlesExtReferences As Boolean
  Protected mvHandlesLinks As Boolean
  Protected mvHideHistoricLinks As Boolean
  Protected mvHandlesBankAccounts As Boolean
  Protected mvUsesHiddenContactNumber As Boolean
  Protected mvHiddenFields As String = ""
  Protected mvCommNumbers(5) As NumberInfo
  Protected mvLoginPageNumber As Integer
  Protected mvUpdateDetailsPageNumber As Integer
  Protected mvEventDelegateNumber As String
  Protected mvDontClearChild As Boolean

  Protected mvSupportsMultiView As Boolean  'Used to determine if only grid should be displayed first and clicking edit should display the edit fields
  Protected mvMultiView As New WebControls.MultiView  'Only used when mvSupportsMultiView is True
  Protected mvView1 As New WebControls.View 'Only used when mvSupportsMultiView is True
  Protected mvView2 As New WebControls.View 'Only used when mvSupportsMultiView is True

  Protected mvTnsControlName As IList(Of String) = New List(Of String)(New String() {"gatewayCardNumber",
                                                                                     "gatewayCardScheme",
                                                                                     "gatewayCardExpiryDateMonth",
                                                                                     "gatewayCardExpiryDateYear",
                                                                                     "gatewayCardSecurityCode"}).AsReadOnly

  Private mvTempEmailValidator As EmailAddressValidator = Nothing

  Public Enum AuthorisationService As Integer
    None = 0
    <AuthorisationServiceAttribute("SCXLVPCSCP")>
    SecureCXL = 1
    <AuthorisationServiceAttribute("PROTX")>
    ProtX = 2
    <AuthorisationServiceAttribute("TNSHOSTED")>
    TnsHosted = 3
    <AuthorisationServiceAttribute("SAGEPAYHOSTED")>
    SagePayHosted = 4
    <AuthorisationServiceAttribute("CSXL210FE")>
    CommsXL = 5
  End Enum

  Protected Property ServiceType As AuthorisationService = AuthorisationService.None

#End Region

#Region "Web Control Properties"

  Public Property PageCareControls() As List(Of CareWebControl)
    Get
      Return mvPageCareControls
    End Get
    Set(ByVal pValue As List(Of CareWebControl))
      mvPageCareControls = pValue
    End Set
  End Property
  Public Property DependantControls() As List(Of ICareChildWebControl)
    Get
      Return mvDependantControls
    End Get
    Set(ByVal pValue As List(Of ICareChildWebControl))
      mvDependantControls = pValue
    End Set
  End Property
  Public ReadOnly Property HandlesActivities() As Boolean
    Get
      Return mvHandlesActivities
    End Get
  End Property
  Public ReadOnly Property HandlesSuppressions() As Boolean
    Get
      Return mvHandlesSuppressions
    End Get
  End Property
  Public ReadOnly Property HandlesExtReferences() As Boolean
    Get
      Return mvHandlesExtReferences
    End Get
  End Property
  Public ReadOnly Property HandlesLinks() As Boolean
    Get
      Return mvHandlesLinks
    End Get
  End Property
  Public ReadOnly Property HandlesBankAccounts() As Boolean
    Get
      Return mvHandlesBankAccounts
    End Get
  End Property
  Public Property UseNewContact() As Boolean
    Get
      Return mvUseNewContact
    End Get
    Set(ByVal pValue As Boolean)
      mvUseNewContact = pValue
    End Set
  End Property
  Public Property SubmitItemUrl() As String
    Get
      Return mvSubmitItemUrl
    End Get
    Set(ByVal pValue As String)
      mvSubmitItemUrl = pValue
    End Set
  End Property
  Public Property ItemStyle() As String
    Get
      Return mvStyle
    End Get
    Set(ByVal pValue As String)
      mvStyle = pValue
    End Set
  End Property
  Public Property NumberOfRows() As Integer
    Get
      Return mvNumberOfRows
    End Get
    Set(ByVal pValue As Integer)
      mvNumberOfRows = pValue
    End Set
  End Property
  Public Property SubmitItemNumber() As Integer
    Get
      Return mvSubmitItemNumber
    End Get
    Set(ByVal pValue As Integer)
      mvSubmitItemNumber = pValue
    End Set
  End Property
  Public Property WebPageNumber() As Integer
    Get
      Return mvWebPageNumber
    End Get
    Set(ByVal pValue As Integer)
      mvWebPageNumber = pValue
    End Set
  End Property
  Public Property WebPageItemNumber() As Integer
    Get
      Return mvWebPageItemNumber
    End Get
    Set(ByVal pValue As Integer)
      mvWebPageItemNumber = pValue
    End Set
  End Property
  Public Property WebPageItemName() As String
    Get
      Return mvWebPageItemName
    End Get
    Set(ByVal pValue As String)
      mvWebPageItemName = pValue
    End Set
  End Property
  Public Property ParentGroup() As String
    Get
      Return mvParentGroup
    End Get
    Set(ByVal pValue As String)
      mvParentGroup = pValue
    End Set
  End Property
  Public Property GroupName() As String
    Get
      Return mvGroupName
    End Get
    Set(ByVal pValue As String)
      mvGroupName = pValue
    End Set
  End Property
  Public Property InitialParameters() As ParameterList
    Get
      Return mvInitialParameters
    End Get
    Set(ByVal pValue As ParameterList)
      mvInitialParameters = pValue
      If mvInitialParameters.OptionalValue("UseNewContact") = "Y" Then mvUseNewContact = True
    End Set
  End Property
  Public Property DefaultParameters() As ParameterList
    Get
      Return mvDefaultParameters
    End Get
    Set(ByVal pValue As ParameterList)
      mvDefaultParameters = pValue
    End Set
  End Property
  Public Property RequiredInitialParameters() As String
    Get
      Return mvRequiredInitialParameters
    End Get
    Set(ByVal pValue As String)
      mvRequiredInitialParameters = pValue
    End Set
  End Property
  Public Property RequiredDefaultParameters() As String
    Get
      Return mvRequiredDefaultParameters
    End Get
    Set(ByVal pValue As String)
      mvRequiredDefaultParameters = pValue
    End Set
  End Property
  Public Property Position() As ContentPosition
    Get
      Return mvPosition
    End Get
    Set(ByVal Value As ContentPosition)
      mvPosition = Value
    End Set
  End Property
  Public Property HTML() As String
    Get
      Return mvHTML
    End Get
    Set(ByVal pValue As String)
      mvHTML = pValue
    End Set
  End Property
  Public ReadOnly Property FocusControl() As Control
    Get
      Return mvFocusControl
    End Get
  End Property
  Public ReadOnly Property NeedsAuthentication() As Boolean
    Get
      Return mvNeedsAuthentication And Not mvUseNewContact
    End Get
  End Property
  Public ReadOnly Property NeedsParent() As Boolean
    Get
      Return mvNeedsParent
    End Get
  End Property
  Public ReadOnly Property CenterControl() As Boolean
    Get
      Return mvCenterControl
    End Get
  End Property
  Protected Property URL() As String
    Get
      Return mvURL
    End Get
    Set(ByVal pValue As String)
      mvURL = pValue
    End Set
  End Property
  Public Property LoginPageNumber() As Integer
    Get
      Return mvLoginPageNumber
    End Get
    Set(ByVal pValue As Integer)
      mvLoginPageNumber = pValue
    End Set
  End Property
  Public Property UpdateDetailsPageNumber() As Integer
    Get
      Return mvUpdateDetailsPageNumber
    End Get
    Set(ByVal pValue As Integer)
      mvUpdateDetailsPageNumber = pValue
    End Set
  End Property
  Public Property EventDelegateNumber() As String
    Get
      Return mvEventDelegateNumber
    End Get
    Set(ByVal pValue As String)
      mvEventDelegateNumber = pValue
    End Set
  End Property

  Public Property DontClearChild() As Boolean
    Get
      Return mvDontClearChild
    End Get
    Set(ByVal pValue As Boolean)
      mvDontClearChild = pValue
    End Set
  End Property

  Public ReadOnly Property IsBackOfficeUser As Boolean
    Get
      Dim vDepartment As String = ""
      If HttpContext.Current.User.Identity.IsAuthenticated Then
        If TypeOf (Page.User.Identity) Is System.Security.Principal.WindowsIdentity Then
          vDepartment = Session("UserDepartment").ToString
        Else
          Dim vIdentity As FormsIdentity = CType(Page.User.Identity, FormsIdentity)
          If vIdentity.Ticket.UserData.Length > 0 Then
            Dim vItems As String() = vIdentity.Ticket.UserData.Split("|"c)
            If vItems.Length > 3 Then vDepartment = vItems(3)
          End If
        End If
      End If
      Return Not String.IsNullOrWhiteSpace(vDepartment)
    End Get
  End Property

  Friend ReadOnly Property HideHistoricLinks() As Boolean
    Get
      Return (mvHandlesLinks AndAlso mvHideHistoricLinks)
    End Get
  End Property

#End Region

#Region "Initialisation and setup of control"

  Protected Sub InitialiseControls(ByVal pType As CareNetServices.WebControlTypes, ByVal pHTMLTable As HtmlTable)
    InitialiseControls(pType, pHTMLTable, "", "")
  End Sub

  Protected Sub InitialiseControls(ByVal pType As CareNetServices.WebControlTypes, ByVal pHTMLTable As HtmlTable, ByVal pMandatoryFields As String)
    InitialiseControls(pType, pHTMLTable, pMandatoryFields, "")
  End Sub

  Protected Sub InitialiseControls(ByVal pType As CareNetServices.WebControlTypes, ByVal pHTMLTable As HtmlTable, ByVal pMandatoryFields As String, ByVal pNonMandatoryFields As String)
    Try
      mvControlType = pType
      mvTableControl = pHTMLTable
      mvSupportsMultiView = InWebPageDesigner() = False AndAlso TypeOf Me Is IMultiViewWebControl AndAlso (Not (InitialParameters.ContainsKey("ShowGridWithFields")) OrElse BooleanValue(InitialParameters("ShowGridWithFields").ToString) = False)
      If ItemStyle.Length > 0 Then pHTMLTable.Attributes("Class") = ItemStyle
      ValidateParameters(pType)
      Dim vList As New ParameterList(HttpContext.Current)
      vList("WebPageItemNumber") = mvWebPageItemNumber.ToString
      Dim vTable As DataTable = DataHelper.GetWebControls(pType, vList)
      If mvSupportsMultiView AndAlso MultiViewGridOnTop() Then
        'Add Cancel button at the end of the table. Currently all modules having grid on top have a button control New so copy this row to add Cancel button.
        If vTable.Select("ControlType = 'btn' AND ParameterName = 'New'").Length > 0 Then
          Dim vNewButtonRow As DataRow = vTable.NewRow
          Dim vRow As DataRow = vTable.Select("ControlType = 'btn' AND ParameterName = 'New'")(0)
          vNewButtonRow.ItemArray = vRow.ItemArray
          'Now set the parameters
          vNewButtonRow("ControlCaption") = "Cancel"
          vNewButtonRow("ParameterName") = "Cancel"
          vNewButtonRow("Visible") = "Y"
          vNewButtonRow("ReadOnlyItem") = "N"
          vTable.Rows.InsertAt(vNewButtonRow, IntegerValue(vRow("SequenceNumber").ToString))
        End If
      End If
      SetCCMandatoryParameters(pMandatoryFields)
      If Not pNonMandatoryFields.Contains("SecurityCode") AndAlso
         DefaultParameters.ContainsKey("CV2Optional") AndAlso
         BooleanValue(DefaultParameters("CV2Optional").ToString) Then
        pNonMandatoryFields = If(String.IsNullOrWhiteSpace(pNonMandatoryFields), "SecurityCode", pNonMandatoryFields & ",SecurityCode")
      End If
      If pMandatoryFields.Length > 0 Then SetMandatory(vTable, pMandatoryFields, True)
      If pNonMandatoryFields.Length > 0 Then SetMandatory(vTable, pNonMandatoryFields, False)
      SetSystemDefaults(vTable, pMandatoryFields)
      AddWebControls(pType, vTable, pHTMLTable)

      If mvDateControls IsNot Nothing Then
        For Each vParameterName As String In mvDateControls
          Dim vTextBox As Control = FindControl(vParameterName)
          Dim vButton As Control = FindControl("cmdFind" & vParameterName)
          If vButton IsNot Nothing AndAlso vTextBox IsNot Nothing Then
            DirectCast(vButton, HtmlInputButton).Attributes("OnClick") = String.Format("javascript:PopupPicker(  '{0}',this)", vTextBox.ClientID)
          End If
        Next
      End If
      If mvUsesHiddenContactNumber Then AddHiddenField("HiddenContactNumber")
      If mvHiddenFields.Length > 0 Then
        For Each vName As String In mvHiddenFields.Split(","c)
          AddHiddenField(vName)
        Next
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Function GetClass() As String
    If ItemStyle.Length > 0 Then
      Return ItemStyle
    Else
      Dim vClass As String = ""
      Select Case Position
        Case ContentPosition.cpHeader
          vClass = "HeaderItem"
        Case ContentPosition.cpFooter
          vClass = "FooterItem"
        Case ContentPosition.cpLeft
          vClass = "LeftItem"
        Case ContentPosition.cpCenter
          vClass = "CenterItem"
        Case ContentPosition.cpRight
          vClass = "RightItem"
      End Select
      Return vClass
    End If
  End Function

  Protected Sub ValidateParameters(ByVal pType As CareNetServices.WebControlTypes)
    Dim vItems() As String
    If mvRequiredInitialParameters.Length > 0 Then
      vItems = mvRequiredInitialParameters.Split(","c)
      Dim vCheck As Boolean
      For Each vItem As String In vItems
        vCheck = True
        Select Case vItem
          Case "Activity"
            If pType = CareNetServices.WebControlTypes.wctAddLink Then vCheck = False
          Case "TitleLookupGroup", "RelationshipLookupGroup", "TopicLookupGroup", "Topic", "SubTopic", _
               "DefaultRelationship", "UserLookupGroup", "DataFilter", "Device1", "Device2", "Device3", _
               "HideHistoricalAddresses", "DeviceLookupGroup", "SalesGroup", "SecondaryGroup", "ProductCategory", _
               "ProductSalePage", "BookingPageNumber", "BookingPage", "DelegateSelectionPage", "DirectDebitSalePage", "CreditCardSalePage", "NewPositionPage", _
               "SurveyEntryPage", "SessionList", "EventNumber", "BookingNumber", "DirectoryName", "DirectoryDetailsPageNumber", "MaximumRecords", _
               "ContactType", "ItemsPerPage", "DisplayFormat", "HyperlinkFormat", "Activity1", "Activity2", "Activity3", "Activity4", "Activity5", "Activity6", _
               "ProductUpdatePage", "AccessView", "PaymentService", "BookingDocument", "DelegateDocument", "PayerDocument", "ManageRelatedOrganisationPage", _
               "UpdateView", "AccessView", "EditRelatedContactPage", "MoveRelatedContactPage", "SendEmailPage", "MailMergePage", "DataExportPage", "SetDefaultPage", _
               "NewContactPageNumber", "Position", "PositionFunction", "PositionSeniority", "CPDPointPage", "CPDObjectivePage", "AllowEditing", "NewOrganisationPageNumber", _
               "PrintReceiptPage", "PaymentMethod", "BatchCategory", "CommunicationUsage1", "CommunicationUsage2", "CommunicationUsage3", "CommunicationUsage4", _
               "Device4", "Device5", "Device6", "Device7", "Device8", "AddressUsage", "ShowGridWithFields", "GridHyperlinkText", "Subject", "SkillLevel", _
               "ExamSessionCode", "ExamCentreCode", "OrganisationGroup"
            vCheck = False
          Case "LookupGroup"
            If pType = CareNetServices.WebControlTypes.wctSelectMembershipTypes Then vCheck = False
          Case "ContactNumber"
            If pType = CareNetServices.WebControlTypes.wctViewDirectoryDetails Then vCheck = False
          Case "OptionNumber"
            If pType = CareNetServices.WebControlTypes.wctSelectOptionSessions OrElse pType = CareNetServices.WebControlTypes.wctBookEvent Then vCheck = False
          Case "Product"
            If pType = CareNetServices.WebControlTypes.wctProductPurchase Then vCheck = False
          Case "Rate"
            If pType = CareNetServices.WebControlTypes.wctBookEvent OrElse pType = CareNetServices.WebControlTypes.wctProductPurchase Then vCheck = False
          Case "Topic", "DownloadPage"
            If pType = CareNetServices.WebControlTypes.wctDownloadSelection Then vCheck = False
          Case "HyperlinkText"
            If pType = CareNetServices.WebControlTypes.wctUpdatePosition Then vCheck = False
          Case "HyperlinkText1", "HyperlinkText2"
            Select Case pType
              Case CareNetServices.WebControlTypes.wctUpdateAddress, CareNetServices.WebControlTypes.wctUpdatePhoneNumber, CareNetServices.WebControlTypes.wctUpdateEmailAddress
                vCheck = False
            End Select
        End Select
        If vCheck AndAlso Not mvInitialParameters.ContainsKey(vItem) Then
          'If the mandatory item does not exist and it is just a CheckBox then add it to the InitialParameters collection (otherwise client has to manually update each Module)
          Select Case vItem
            Case "HideHistoricalPositions", "CheckBoxMandatory", "CreateInvoice"
              mvInitialParameters.Add(vItem, "N")
            Case "OnlineCCAuthorisation"
              Select Case pType
                Case CareNetServices.WebControlTypes.wctAddMemberCC, CareNetServices.WebControlTypes.wctBookEventCC, CareNetServices.WebControlTypes.wctMakeDonationCC, _
                     CareNetServices.WebControlTypes.wctProductPurchaseCC, CareNetServices.WebControlTypes.wctRapidProductPurchase, CareNetServices.WebControlTypes.wctRenewMembershipCC, _
                     CareNetServices.WebControlTypes.wctRenewSubscriptionCC, CareNetServices.WebControlTypes.wctAddMemberCI
                  mvInitialParameters.Add(vItem, "N")
              End Select
            Case "PaymentService"
              Select Case pType
                Case CareNetServices.WebControlTypes.wctProcessPayment
                  mvInitialParameters.Add(vItem, "")
              End Select
            Case "MembershipFor"
              Select Case pType
                Case CareNetServices.WebControlTypes.wctAddMemberCC, CareNetServices.WebControlTypes.wctPayMultiplePaymentPlans, _
                     CareNetServices.WebControlTypes.wctAddMemberCI
                  mvInitialParameters.Add(vItem, "U")
              End Select
          End Select
          If mvInitialParameters.ContainsKey(vItem) = False Then Throw New CareException(String.Format("The Initial Parameters for this User Control do not contain all the mandatory items ({0})", mvRequiredInitialParameters))
        End If
      Next
    End If
    If mvRequiredDefaultParameters.Length > 0 Then
      vItems = mvRequiredDefaultParameters.Split(","c)
      Dim vCheck As Boolean
      For Each vItem As String In vItems
        vCheck = True
        Select Case vItem
          Case "DataSource"
            If pType = CareNetServices.WebControlTypes.wctAddExternalReference Then vCheck = False
          Case "EMailAddress"
            If pType = CareNetServices.WebControlTypes.wctMakeDonationCC OrElse _
               pType = CareNetServices.WebControlTypes.wctRegister Then vCheck = False
          Case "SendEMail", "TestEMailAddress", "DistributionCode", "Status", "ReturnToRegisterPage"
            vCheck = False 'Non mandatory
          Case "StandardDocument", "ConfirmationPage", "Name"
            If pType = CareNetServices.WebControlTypes.wctRegister OrElse pType = CareNetServices.WebControlTypes.wctMakeDonationCC Then vCheck = False
          Case "DonationAmount", "DonationProduct", "DonationRate", "PartSource"
            If pType = CareNetServices.WebControlTypes.wctRapidProductPurchase Then vCheck = False
          Case "Source"
            If pType = CareNetServices.WebControlTypes.wctRapidProductPurchase OrElse _
               pType = CareNetServices.WebControlTypes.wctAddSuppression Then vCheck = False
          Case "NoteMandatory", "SetValidFromDate", "SetValidToDate", "DisplayFirstRecord", "AllowMaintenance", "ShowActivitiesSuppressionsThatExists", "UseCurrentDate", _
               "ShowActivitiesThatExists", "RegisterPageNumber", "RegisterAtOrgPageNumber", _
               "Activity1", "ItemSelectType1", "Activity2", "ItemSelectType2", "Activity3", "ItemSelectType3", _
               "Activity4", "ItemSelectType4", "Activity5", "ItemSelectType5", "Activity6", "ItemSelectType6", _
               "BatchCategory", "CreditCategory", "DeDuplicateOrganisationPage", "ResetPasswordPage"
            vCheck = False
          Case Else
            'Check
        End Select
        If vCheck AndAlso Not mvDefaultParameters.ContainsKey(vItem) Then
          'If the mandatory item does not exist and it is just a CheckBox then add it to the DefaultParameters collection (otherwise client has to manually update each Module)
          Select Case vItem
            Case "AllowMaintenance", "DisplayFirstRecord"
              If pType = CareNetServices.WebControlTypes.wctAddExternalReference Then mvDefaultParameters.Add(vItem, "N")
            Case "DefaultTopicSubTopic"
              If pType = CareNetServices.WebControlTypes.wctAddCommunicationNote Then mvDefaultParameters.Add(vItem, "N")
            Case "SetHistoric"
              Select Case pType
                Case CareNetServices.WebControlTypes.wctAddContact, CareNetServices.WebControlTypes.wctAddRelatedContact, CareNetServices.WebControlTypes.wctAddSuppression, _
                     CareNetServices.WebControlTypes.wctAddCategory, CareNetServices.WebControlTypes.wctAddCategoryCheckboxes, CareNetServices.WebControlTypes.wctAddCategoryNotes, _
                     CareNetServices.WebControlTypes.wctAddCategoryOptions, CareNetServices.WebControlTypes.wctAddCategoryValue
                  mvDefaultParameters.Add(vItem, "N")
              End Select
            Case "CreateCreditCustomer"
              mvDefaultParameters.Add(vItem, "N")
            Case "DataSource"
              If pType = CareNetServices.WebControlTypes.wctRapidProductPurchase Then mvDefaultParameters.Add(vItem, "")
            Case "UseSourceFromEvent"
              If pType = CareNetServices.WebControlTypes.wctBookEvent OrElse pType = CareNetServices.WebControlTypes.wctBookEventCC Then
                mvDefaultParameters.Add(vItem, "N")
              End If
            Case "HideHistoricLinks"
              If pType = CareNetServices.WebControlTypes.wctAddRelatedContact Then mvDefaultParameters.Add(vItem, "N")
          End Select
          If mvDefaultParameters.ContainsKey(vItem) = False Then Throw New CareException(String.Format("The Default Parameters for this User Control do not contain all the mandatory items ({0})", mvRequiredDefaultParameters))
        End If
      Next
    End If
  End Sub

  Protected Sub SetMandatory(ByVal pTable As DataTable, ByVal pParameters As String, ByVal pMandatory As Boolean)
    If pTable Is Nothing Then Return
    Dim vValue As String = "N"
    If pMandatory Then vValue = "Y"
    Dim vItems() As String = pParameters.Split(","c)
    'Dim vHasMandatoryItem As Boolean = pTable.Columns.Contains("MandatoryItem")
    For Each vRow As DataRow In pTable.Rows
      Dim vParameter As String = vRow("ParameterName").ToString
      For Each vItem As String In vItems
        If vItem = vParameter Then
          vRow("NullsInvalid") = vValue
          'If vHasMandatoryItem Then vRow("MandatoryItem") = vValue
        End If
      Next
    Next
  End Sub

  Private Sub SetSystemDefaults(ByVal pTable As DataTable, ByVal pParameters As String)
    If pTable Is Nothing Then Return
    Dim vContinue As Boolean
    Dim vSetMandatory As Boolean

    'Check if a change is required
    If SupportsOnlineCCAuthorisation Then
      vSetMandatory = pParameters.Contains("SecurityCode")
      vContinue = True
    Else
      Select Case mvControlType
        Case CareNetServices.WebControlTypes.wctAddCommunicationNote
          vContinue = True
      End Select
    End If

    If vContinue Then
      If SupportsOnlineCCAuthorisation Then
        For Each vRow As DataRow In pTable.Rows
          If vRow("ParameterName").ToString = "SecurityCode" Then
            Dim vColumns() As DataColumn = {pTable.Columns("ParameterName")}
            pTable.PrimaryKey = vColumns
            Dim vCCRow As DataRow = pTable.Rows.Find("CreditCardNumber")
            If vSetMandatory Then         'InitialParameters and config 'fp_cc_security_code_mandatory' is set
              If vCCRow("MandatoryItem").ToString = "Y" Then
                vRow("MandatoryItem") = "Y"
              End If
              If vCCRow("Visible").ToString = "Y" AndAlso vRow("MandatoryItem").ToString = "Y" Then
                vRow("Visible") = "Y"
              End If
            Else
              'Always hide
              vRow("Visible") = "N"
            End If
          End If
        Next
        If InitialParameters.ContainsKey("OnlineCCAuthorisation") AndAlso BooleanValue(InitialParameters("OnlineCCAuthorisation").ToString) Then
          Dim vRowColl() As DataRow = pTable.Select("ParameterName = 'AuthorisationCode'")
          If vRowColl.Length > 0 AndAlso DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = "SCXLVPCSCP" Then
            vRowColl(0)("ReadOnlyItem") = "Y"
            vRowColl(0)("MandatoryItem") = "N"
          End If
        End If
      Else
        Select Case mvControlType
          Case CareNetServices.WebControlTypes.wctAddCommunicationNote
            'Always whow SubTopic when it is Hidden and Topic is Visible
            Dim vRowColl() As DataRow = pTable.Select("ParameterName = 'SubTopic'")
            If vRowColl.Length > 0 Then
              Dim vRow As DataRow = vRowColl(0)
              If pTable.Select("ParameterName = 'Topic'")(0).Item("Visible").ToString = "Y" AndAlso vRow("Visible").ToString = "N" Then
                vRow("Visible") = "Y"
              End If
            End If
        End Select
      End If
    End If
  End Sub

  Private Sub SetCCMandatoryParameters(ByRef pParameters As String)
    If SupportsOnlineCCAuthorisation Then
      If InitialParameters.Contains("OnlineCCAuthorisation") AndAlso BooleanValue(InitialParameters("OnlineCCAuthorisation").ToString) _
         AndAlso DataHelper.ConfigurationValueOption(DataHelper.ConfigurationValues.fp_cc_security_code_mandatory) Then
        If pParameters.Length > 0 Then pParameters = pParameters & ","
        pParameters = pParameters & "SecurityCode"
      End If
    End If
  End Sub

  Protected Property SupportsOnlineCCAuthorisation() As Boolean
    Get
      Return mvSupportsOnlineCCAuthorisation
    End Get
    Set(ByVal pValue As Boolean)
      mvSupportsOnlineCCAuthorisation = pValue
    End Set
  End Property

  Protected Sub AddWebControls(ByVal pPageType As CareNetServices.WebControlTypes, ByVal pTable As DataTable, ByVal pHTMLTable As HtmlTable)
    If pTable IsNot Nothing Then
      Dim vHideContactEntry As Boolean
      Select Case pPageType
        Case CareNetServices.WebControlTypes.wctAddContact, _
             CareNetServices.WebControlTypes.wctAddRelatedContact, _
             CareNetServices.WebControlTypes.wctFindFundraiser, _
             CareNetServices.WebControlTypes.wctSelectEventDelegates, _
             CareNetServices.WebControlTypes.wctRegisterMember, _
             CareNetServices.WebControlTypes.wctRegisterCorporateMember, _
             CareNetServices.WebControlTypes.wctRegisterCombined, _
             CareNetServices.WebControlTypes.wctSearchOrganisation, _
             CareNetServices.WebControlTypes.wctSearchDirectory, _
             CareNetServices.WebControlTypes.wctSearchContact, _
             CareNetServices.WebControlTypes.wctViewDirectoryDetails
          'Don't hide contact fields
        Case Else
          'Hide the contact entry fields if this is a public page and we are not authenticated
          vHideContactEntry = Not mvNeedsAuthentication AndAlso HttpContext.Current.User.Identity.IsAuthenticated
      End Select

      Dim vHTMLTable As HtmlTable = pHTMLTable
      Dim vUseQAS As Nullable(Of Boolean)

      If mvSupportsMultiView Then
        'As the user can customise the sequence of the grid, make sure you always process the grid control first to put it in mvView1.
        Dim vSource As DataRow = pTable.Select("ControlType = 'dgr'")(0)
        Dim vTable As DataTable = pTable.DefaultView.ToTable
        vTable.Rows.Clear()
        Dim vTarget As DataRow = vTable.NewRow
        vTarget.ItemArray = vSource.ItemArray
        vTable.Rows.Add(vTarget)
        'vTable should only have one row containing grid control. The GridHyperlink will automatically be added
        AddControls(vTable, vHTMLTable, vHideContactEntry, pPageType, vUseQAS, True)
      End If

      AddControls(pTable, vHTMLTable, vHideContactEntry, pPageType, vUseQAS, False)

      If mvSupportsMultiView Then
        'vHTMLTable contains all Edit/Search fields with Buttons
        mvView2.Controls.Add(vHTMLTable)
        'If the grid is displayed first then check the visibility of the grid hyperlink. The search modules would always display the link.
        If MultiViewGridOnTop() Then FindControlByName(mvGridControlTable, "GridHyperlink").Visible = CType(Me, IMultiViewWebControl).GridHyperLinkVisibility
      End If

      If vUseQAS.HasValue AndAlso vUseQAS.Value = True Then
        'Add a hidden field for the Paf Status
        Dim vHTMLRow As New HtmlTableRow
        'First Add the label
        Dim vHTMLCell As New HtmlTableCell
        vHTMLCell.Attributes("Class") = "DataMessage"
        vHTMLCell.ColSpan = 2
        'Add a label to the cell
        Dim vLabel As New Label
        vLabel.ID = "PafStatus"
        vLabel.CssClass = "DataMessage"
        vLabel.Visible = False
        vHTMLCell.Controls.Add(vLabel)    'Add it to the cell
        AddUpdatePanel(vHTMLCell)
        vHTMLRow.Cells.Add(vHTMLCell)
        'Now add the row to the table
        vHTMLTable.Rows.Add(vHTMLRow)
      End If

      'Add any event handler and triggers after loading all controls.
      Select Case mvControlType
        Case CareNetServices.WebControlTypes.wctUpdateAddress
          AddAsyncPostBackTrigger("Country,Postcode,BuildingNumber,HouseName,Address,Town,County,Branch,ValidFrom,ValidTo", "ContactAddress,New", PostBackTriggerEventTypes.None)
          AddAsyncPostBackTrigger("Default", "ContactAddress", PostBackTriggerEventTypes.None)

          'On Change of Postcode, default Branch should be selected
          AddTextChangedHandler("Postcode")
          AddAsyncPostBackTrigger("Branch", "Postcode", PostBackTriggerEventTypes.TextChanged)
          AddAsyncPostBackTrigger("ContactAddress", "Default,Save,New", PostBackTriggerEventTypes.ButtonClick)


        Case CareNetServices.WebControlTypes.wctUpdateEmailAddress
          AddAsyncPostBackTrigger("AddressNumber,Device,Number,ExDirectory,Mail,ValidFrom,DeviceDefault,ValidTo,PreferredMethod,Notes", "EmailAddress,New", PostBackTriggerEventTypes.None)
          AddAsyncPostBackTrigger("EmailAddress", "Save,New", PostBackTriggerEventTypes.ButtonClick)

        Case CareNetServices.WebControlTypes.wctUpdatePhoneNumber
          AddAsyncPostBackTrigger("AddressNumber,Device,DiallingCode,STDCode,Number,ExDirectory,Extension,ValidFrom,DeviceDefault,ValidTo,PreferredMethod,Notes", "TelephoneData,New", PostBackTriggerEventTypes.None)
          AddAsyncPostBackTrigger("Default", "TelephoneData", PostBackTriggerEventTypes.None)
          AddAsyncPostBackTrigger("TelephoneData", "Default,Save,New", PostBackTriggerEventTypes.ButtonClick)

        Case CareNetServices.WebControlTypes.wctUpdatePosition
          AddAsyncPostBackTrigger("Name,Address,Position,Location,PositionFunction,PositionSeniority,Started,Finished,Mail,PageError", "OrganisationPositionData", PostBackTriggerEventTypes.None)
          AddAsyncPostBackTrigger("OrganisationPositionData,PageError", "Save", PostBackTriggerEventTypes.ButtonClick)

        Case CareNetServices.WebControlTypes.wctUpdateCpdPoints
          AddAsyncPostBackTrigger("ContactCpdPeriodNumber,CpdCategoryType,CpdCategory,PointsDate,CpdPoints,CpdPoints2,WebPublish,CpdItemType,CpdOutcome,EvidenceSeen,Notes,PageError", "ContactCPDPoints,New", PostBackTriggerEventTypes.None)
          AddAsyncPostBackTrigger("ContactCPDPoints", "Save", PostBackTriggerEventTypes.ButtonClick)
          AddAsyncPostBackTrigger("CpdCategory", "CpdCategoryType", PostBackTriggerEventTypes.SelectedIndexChanged)
          AddAsyncPostBackTrigger("CpdPoints,CpdPoints2,PageError,WarningMessage", "CpdCategory,CpdCategoryType", PostBackTriggerEventTypes.SelectedIndexChanged)
          AddAsyncPostBackTrigger("PageError,WarningMessage", "Save", PostBackTriggerEventTypes.ButtonClick)
        Case CareNetServices.WebControlTypes.wctCPDCycle
          AddAsyncPostBackTrigger("EndYear,EndMonth", "StartYear", PostBackTriggerEventTypes.TextChanged)
          AddAsyncPostBackTrigger("CpdCycleType,StartMonth,StartYear,EndMonth,EndYear,CpdCycleStatus,PageError,WarningMessage", "ContactCPDCycle,New", PostBackTriggerEventTypes.None)
          AddAsyncPostBackTrigger("ContactCPDCycle,PageError", "Save", PostBackTriggerEventTypes.ButtonClick)
          AddAsyncPostBackTrigger("StartMonth,EndMonth,CpdCycleStatus", "CpdCycleType", PostBackTriggerEventTypes.None)
        Case CareNetServices.WebControlTypes.wctUpdateCpdObjectives
          AddAsyncPostBackTrigger("ContactCpdPeriodNumber,CpdCategoryType,CpdCategory,CpdObjectiveDesc,LongDescription,CompletionDate,TargetDate,Notes", "ContactCPDObjectives,New", PostBackTriggerEventTypes.None)
          AddAsyncPostBackTrigger("ContactCPDObjectives", "Save,New", PostBackTriggerEventTypes.ButtonClick)
          AddAsyncPostBackTrigger("CpdCategory", "CpdCategoryType", PostBackTriggerEventTypes.SelectedIndexChanged)
          AddAsyncPostBackTrigger("ContactCpdPeriodNumber,CpdCategoryType,CpdCategory", "Save", PostBackTriggerEventTypes.ButtonClick)
        Case CareNetServices.WebControlTypes.wctAddRelatedContact
          AddAsyncPostBackTrigger("RelationshipStddLinkatus", "Relationship", PostBackTriggerEventTypes.SelectedIndexChanged)
        Case CareNetServices.WebControlTypes.wctRelatedContacts
          AddAsyncPostBackTrigger("RelatedContactData,WarningMessage1,WarningMessage2", "SetDefault,SendEmail,MailMerge,DataExport", PostBackTriggerEventTypes.ButtonClick)
        Case CareNetServices.WebControlTypes.wctSearchContact
          AddAsyncPostBackTrigger("ContactData,WarningMessage1,WarningMessage2,WarningMessage3,NewContact", "Search", PostBackTriggerEventTypes.None)
        Case CareNetServices.WebControlTypes.wctBookEvent
          'On Change of AdultQuantity/ChildQuantity, Quantity, Total Amount field should be updated
          AddAsyncPostBackTrigger("Quantity,TotalAmount,WarningMessage3,WarningMessage4,WarningMessage5,ChildQuantity,AdultQuantity", "AdultQuantity,ChildQuantity", PostBackTriggerEventTypes.TextChanged)
          AddAsyncPostBackTrigger("Quantity,TotalAmount,WarningMessage3,WarningMessage4,WarningMessage5,AdultQuantity", "ChildQuantity", PostBackTriggerEventTypes.TextChanged)
          AddAsyncPostBackTrigger("Quantity,TotalAmount,WarningMessage3,WarningMessage4,WarningMessage5,ChildQuantity", "AdultQuantity", PostBackTriggerEventTypes.TextChanged)
          AddAsyncPostBackTrigger("TotalAmount,WarningMessage3,WarningMessage4,WarningMessage5", "Quantity", PostBackTriggerEventTypes.TextChanged)
        Case CareNetServices.WebControlTypes.wctAddMemberCC, CareNetServices.WebControlTypes.wctProcessPayment
          AddComboBoxCheckedChangedHandler(TryCast(FindControlByName(mvTableControl, "CreateToken"), CheckBox))
      End Select
      AddCustomValidator(vHTMLTable)
    End If
  End Sub

  Private Sub AddControls(ByRef pTable As DataTable, ByRef vHtmlTable As HtmlTable, ByVal vHideContactEntry As Boolean, ByVal pPageType As CareNetServices.WebControlTypes, ByRef vUseQAS As Nullable(Of Boolean), ByVal pAddGridForMultiView As Boolean)
    'This should be called from within AddWebControls only
    Dim vFirstOption As Boolean = True
    Dim vHTMLRows(mvNumberOfRows) As HtmlTableRow
    Dim vCurrentRow As Integer
    Dim vNumberOfColumns As Integer
    Dim vPasswordParameter As String = "Password"
    For Each vRow As DataRow In pTable.Rows
      Dim vType As String = vRow("ControlType").ToString
      Dim vVisible As Boolean = vRow.Item("Visible").ToString = "Y"
      Dim vParameterName As String = vRow.Item("ParameterName").ToString
      Dim vHelpText As String = vRow.Item("HelpText").ToString

      Select Case vParameterName
        Case "Title", "Forenames", "Surname", "EMailAddress", "DirectNumber", "MobileNumber", _
             "Address", "Town", "County", "Postcode", "Country", "ConfirmEMailAddress", "LabelNameFormatCode"
          If vHideContactEntry Then
            vVisible = False
            mvContactEntryHidden = True
          End If

          Select Case mvControlType
            Case CareNetServices.WebControlTypes.wctAddMemberCC, CareNetServices.WebControlTypes.wctAddMemberCI
              'If the control is configured to create membership on behalf of an organisation then hide contact entry controls
              If InitialParameters.ContainsKey("MembershipFor") AndAlso InitialParameters("MembershipFor").ToString.ToUpper = "O" OrElse ParentGroup.Length > 0 Then
                vVisible = False
                mvContactEntryHidden = True
              End If
            Case CareNetServices.WebControlTypes.wctProcessPayment
              If InitialParameters.OptionalValue("PaymentType").ToUpper = "CS" Then
                vVisible = False
                mvContactEntryHidden = True
              End If
            Case CareNetServices.WebControlTypes.wctAddMemberDD, CareNetServices.WebControlTypes.wctAddMemberCS
              If ParentGroup.Length > 0 Then
                vVisible = False
                mvContactEntryHidden = True
              End If
            Case Else
              '
          End Select

          If vParameterName = "LabelNameFormatCode" Then
            If Not DataHelper.ConfigurationOption(DataHelper.ConfigurationOptions.multiple_label_name_formats, False) Then
              vVisible = False
            End If
          End If

        Case "UserName"
          If pPageType = CareNetServices.WebControlTypes.wctRegister AndAlso InitialParameters.OptionalValue("EmailAddressIsUserName") = "Y" Then
            vVisible = False
          End If
        Case "PostcoderPostcode", "PostcoderAddress"
          If vHideContactEntry Then
            vVisible = False
          Else
            If Not vUseQAS.HasValue Then vUseQAS = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.qas_pro_web_url).Length > 0
            If vUseQAS.Value = False Then vVisible = False

            Select Case mvControlType
              Case CareNetServices.WebControlTypes.wctAddMemberCC, CareNetServices.WebControlTypes.wctAddMemberCI
                'If the control is configured to create membership on behalf of an organisation then hide contact entry controls
                If InitialParameters.ContainsKey("MembershipFor") AndAlso InitialParameters("MembershipFor").ToString.ToUpper = "O" OrElse ParentGroup.Length > 0 Then
                  vVisible = False
                  mvContactEntryHidden = True
                End If
              Case CareNetServices.WebControlTypes.wctProcessPayment
                If InitialParameters.OptionalValue("PaymentType").ToUpper = "CS" Then
                  vVisible = False
                  mvContactEntryHidden = True
                End If
              Case CareNetServices.WebControlTypes.wctAddMemberDD
                If ParentGroup.Length > 0 Then
                  vVisible = False
                  mvContactEntryHidden = True
                End If
              Case Else
                '
            End Select
          End If
        Case "CreditCardType", "CreditCardNumber", "CardExpiryDate", "IssueNumber", "CardStartDate", "SecurityCode"
          If mvControlType = CareNetServices.WebControlTypes.wctProcessPayment AndAlso InitialParameters.OptionalValue("PaymentType").ToUpper = "CS" Then
            vVisible = False
            mvContactEntryHidden = True
          End If
          If mvControlType = CareNetServices.WebControlTypes.wctAddMemberCC OrElse mvControlType = CareNetServices.WebControlTypes.wctAddMemberCI AndAlso _
            DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type).ToUpper = "PROTX" AndAlso _
            vParameterName = "CreditCardType" Then
            vVisible = True
          End If

          If (vParameterName = "IssueNumber" OrElse vParameterName = "CardStartDate") AndAlso
            ((mvControlType = CareNetServices.WebControlTypes.wctProcessPayment AndAlso InitialParameters.OptionalValue("PaymentService").Trim.ToUpper = TNSHOSTED) OrElse _
             (mvControlType = CareNetServices.WebControlTypes.wctAddMemberCC AndAlso InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" AndAlso _
              DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED)) Then
            vVisible = False
          End If
        Case "Organisation", "Payer"
          If mvControlType = CareNetServices.WebControlTypes.wctAddMemberCC OrElse mvControlType = CareNetServices.WebControlTypes.wctAddMemberCI Then
            If InitialParameters.ContainsKey("MembershipFor") AndAlso mvParentGroup.Length = 0 Then
              If InitialParameters("MembershipFor").ToString.ToUpper = "O" Then vVisible = True Else vVisible = False
            Else
              vVisible = False
            End If
          End If
        Case "StartDateList"
          If mvControlType = CareNetServices.WebControlTypes.wctSelectMembershipTypes AndAlso ParentGroup.Length > 0 Then
            If DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fixed_renewal_M).Length > 0 Then
              vVisible = True
            End If
          End If
        Case "GiftAid"
          If (mvControlType = CareNetServices.WebControlTypes.wctAddMemberCS OrElse mvControlType = CareNetServices.WebControlTypes.wctAddMemberCC OrElse mvControlType = CareNetServices.WebControlTypes.wctAddMemberDD OrElse mvControlType = CareNetServices.WebControlTypes.wctAddMemberCI) AndAlso ParentGroup.Length > 0 AndAlso ParentGroup = "SelectedOrganisation" Then
            vVisible = False
          End If
        Case "gatewayCardExpiryDateYear"
          If (mvControlType = CareNetServices.WebControlTypes.wctAddMemberCC AndAlso ((DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) <> TNSHOSTED) OrElse InitialParameters.OptionalValue("OnlineCCAuthorisation") <> "Y")) OrElse _
          ((mvControlType = CareNetServices.WebControlTypes.wctProcessPayment AndAlso InitialParameters.OptionalValue("PaymentService").Trim.ToUpper <> TNSHOSTED)) Then
            vVisible = False
          End If

      End Select

      If vVisible Then
        Select Case vType.Substring(0, 3)
          Case "pnl"
            Dim vControl As Panel = CType(FindControlByName(Me, "pnl" & vParameterName), Panel)
            If vControl IsNot Nothing Then
              Dim vTable As HtmlTable = CType(FindControlByName(Me, "tbl" & vParameterName), HtmlTable)
              If vTable IsNot Nothing Then vHtmlTable = vTable
            End If
          Case "dgr"
            'Only add grid when MultiView support is not ON or its ON but we are just dealing with grid control to be displayed in mvView1
            If Not mvSupportsMultiView OrElse pAddGridForMultiView Then
              Dim vRowFormat As Boolean = True
              'Check if the display format is set to rows or columnar
              If InitialParameters.ContainsKey("DisplayFormat") AndAlso InitialParameters("DisplayFormat").ToString <> "0" Then vRowFormat = False
              If vRowFormat Then
                Dim vHTMLRow As New HtmlTableRow
                Dim vHTMLCell As New HtmlTableCell
                Dim vDGR As New DataGrid
                With vDGR
                  .AutoGenerateColumns = False
                  .CssClass = "Table"
                  .AllowSorting = False
                  .GridLines = GridLines.None
                  .SelectedItemStyle.CssClass = "TableSelectedData"
                  .AlternatingItemStyle.CssClass = "TableAlternateData"
                  .ItemStyle.CssClass = "TableData"
                  .HeaderStyle.CssClass = "TableHeader"
                  .FooterStyle.CssClass = "TableFooter"
                  .PagerStyle.HorizontalAlign = HorizontalAlign.Center
                  .PagerStyle.ForeColor = Drawing.Color.Black
                  .PagerStyle.BackColor = Drawing.Color.Transparent
                  .PagerStyle.Mode = PagerMode.NumericPages
                  .ID = vParameterName
                End With
                Select Case mvControlType
                  Case CareNetServices.WebControlTypes.wctUpdateAddress, _
                          CareNetServices.WebControlTypes.wctUpdatePhoneNumber, _
                          CareNetServices.WebControlTypes.wctUpdateEmailAddress, _
                          CareNetServices.WebControlTypes.wctUpdatePosition, _
                          CareNetServices.WebControlTypes.wctCPDCycle, _
                          CareNetServices.WebControlTypes.wctUpdateCpdPoints, _
                          CareNetServices.WebControlTypes.wctUpdateCpdObjectives, _
                          CareNetServices.WebControlTypes.wctSearchContact, _
                          CareNetServices.WebControlTypes.wctPayerSelection, _
                          CareNetServices.WebControlTypes.wctDeDupOrgForRegistration, _
                          CareNetServices.WebControlTypes.wctSetUserOrganisation, _
                          CareNetServices.WebControlTypes.wctSelecttPayPlanForDD
                    AddHandler vDGR.ItemCommand, AddressOf DataGridItemClickedHandler
                End Select
                vHTMLCell.Controls.Add(vDGR)
                vHTMLCell.ColSpan = 2
                vHTMLRow.Cells.Add(vHTMLCell)
                If pAddGridForMultiView Then
                  AddMultiViewGrid(vHtmlTable, vHTMLRow)
                Else
                  vHtmlTable.Rows.Add(vHTMLRow)
                End If

              Else
                'Alternate display - columnar format
                Dim vHTMLRow As New HtmlTableRow
                Dim vHTMLCell As New HtmlTableCell
                Dim vDataList As New DataList
                With vDataList
                  .ItemStyle.CssClass = "TableData"
                  .ItemStyle.VerticalAlign = VerticalAlign.Top
                  .SelectedItemStyle.CssClass = "TableSelectedData"
                  .RepeatColumns = IntegerValue(InitialParameters.OptionalValue("DisplayFormat").ToString)
                  .RepeatDirection = RepeatDirection.Horizontal
                  .RepeatLayout = RepeatLayout.Table
                  .ID = vParameterName
                  .CellSpacing = 5
                  .CellPadding = 5
                End With
                Select Case mvControlType
                  Case CareNetServices.WebControlTypes.wctSelectMembershipTypes, _
                       CareNetServices.WebControlTypes.wctSelectEvents, _
                       CareNetServices.WebControlTypes.wctSelectProducts, _
                       CareNetServices.WebControlTypes.wctSelectBookingOptions, _
                       CareNetServices.WebControlTypes.wctSearchDirectory, _
                       CareNetServices.WebControlTypes.wctViewTransaction, _
                       CareNetServices.WebControlTypes.wctDownloadSelection, _
                       CareNetServices.WebControlTypes.wctRelatedOrganisations, _
                       CareNetServices.WebControlTypes.wctSelectExams
                    AddHandler vDataList.ItemDataBound, AddressOf DataListItemDataBoundHandler
                End Select
                vHTMLCell.Controls.Add(vDataList)
                vHTMLCell.ColSpan = 2
                vHTMLRow.Cells.Add(vHTMLCell)
                If pAddGridForMultiView Then
                  AddMultiViewGrid(vHtmlTable, vHTMLRow)
                Else
                  vHtmlTable.Rows.Add(vHTMLRow)
                End If
              End If
            End If

          Case "cbo"
            Dim vHTMLRow As New HtmlTableRow
            'First Add the label
            Dim vHTMLCell As New HtmlTableCell
            vHTMLCell.InnerHtml = vRow("ControlCaption").ToString
            vHTMLCell.Attributes("Class") = "DataEntryLabel"
            vHTMLRow.Cells.Add(vHTMLCell)

            vHTMLCell = New HtmlTableCell
            If vNumberOfColumns > 0 Then
              vHTMLCell.ColSpan = (vNumberOfColumns * 2) - 1
            End If
            Dim vDDL As New DropDownList
            AddFocusScript(vDDL.Attributes)
            vDDL.ID = vParameterName
            If mvFocusControl Is Nothing Then mvFocusControl = vDDL
            vDDL.CssClass = "DataEntryItem"
            vDDL.Width = New Unit(CInt(vRow("ControlWidth")) / 100, UnitType.Em)
            Select Case vParameterName
              Case "Actioner"
                vDDL.DataTextField = "FullName"
                vDDL.DataValueField = "ContactNumber"
                Dim vList As New ParameterList(HttpContext.Current)
                vList("Active") = "Y"
                If InitialParameters.ContainsKey("UserLookupGroup") Then vList("LookupGroup") = InitialParameters("UserLookupGroup").ToString
                DataHelper.FillComboWithRestriction(CareNetServices.XMLLookupDataTypes.xldtUsers, vDDL, True, vList, "LogName is null OR ContactNumber <>''")
              Case "ContactNumber1"
                If pPageType = CareNetServices.WebControlTypes.wctAddLink Then
                  Dim vList As New ParameterList(HttpContext.Current)
                  Select Case InitialParameters("LinkType").ToString
                    Case "C"
                      'Contact with Activity
                      vList("Activity") = InitialParameters("Activity").ToString
                      vList("RestrictNonHistoricActivity") = "Y"
                      Dim vTable As DataTable = DataHelper.FindDataTable(CareNetServices.XMLDataFinderTypes.xdftContacts, vList)
                      If vTable IsNot Nothing Then
                        vTable.Rows.InsertAt(vTable.NewRow, 0)
                        vDDL.DataTextField = "LabelName"
                        vDDL.DataValueField = "ContactNumber"
                        vDDL.DataSource = vTable
                        vDDL.DataBind()
                      End If
                    Case "O"
                      'Organisation with Activity
                      vList("Activity") = InitialParameters("Activity").ToString
                      vList("RestrictNonHistoricActivity") = "Y"
                      Dim vTable As DataTable = DataHelper.FindDataTable(CareNetServices.XMLDataFinderTypes.xdftOrganisations, vList)
                      If vTable IsNot Nothing Then
                        vTable.Rows.InsertAt(vTable.NewRow(), 0)
                        vDDL.DataTextField = "Name"
                        vDDL.DataValueField = "OrganisationNumber"
                        vDDL.DataSource = vTable
                        vDDL.DataBind()
                      End If
                    Case Else   '"U"
                      'CARE User
                      vDDL.DataTextField = "FullName"
                      vDDL.DataValueField = "ContactNumber"
                      If InitialParameters.ContainsKey("UserLookupGroup") Then vList("LookupGroup") = InitialParameters("UserLookupGroup").ToString
                      DataHelper.FillComboWithRestriction(CareNetServices.XMLLookupDataTypes.xldtUsers, vDDL, True, vList, "LogName is null OR ContactNumber <>''")
                  End Select
                End If
              Case "DataSource"
                vDDL.DataTextField = "DataSourceDesc"
                vDDL.DataValueField = "DataSource"
                If mvControlType = CareNetServices.WebControlTypes.wctRapidProductPurchase Then
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtDataSources, vDDL, True)
                  If Not IsPostBack AndAlso DefaultParameters.ContainsKey("DataSource") AndAlso DefaultParameters("DataSource").ToString.Length > 0 Then
                    vDDL.SelectedValue = DefaultParameters("DataSource").ToString
                  End If
                  AddSelectedIndexChangedHandler(vDDL)
                Else
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtDataSources, vDDL, False)
                End If
              Case "Direction"
                vDDL.DataTextField = "LookupDesc"
                vDDL.DataValueField = "LookupCode"
                Dim vList As New ParameterList(HttpContext.Current)
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtDocumentDirections, vDDL, True, vList)
              Case "EventNumber"
                vDDL.DataTextField = "EventDesc"
                vDDL.DataValueField = "EventNumber"
                Dim vList As New ParameterList(HttpContext.Current)
                If mvInitialParameters.ContainsKey("Topic") AndAlso mvInitialParameters.ContainsKey("SubTopic") Then
                  vList("Topic") = mvInitialParameters("Topic")
                  vList("SubTopic") = mvInitialParameters("SubTopic")
                End If
                Dim vTable As DataTable = DataHelper.FindDataTable(CareNetServices.XMLDataFinderTypes.xdftEvents, vList)
                If vTable IsNot Nothing Then
                  vTable.Rows.InsertAt(vTable.NewRow(), 0)
                  vDDL.DataSource = vTable
                  vDDL.DataBind()
                End If
              Case "Sex"
                vDDL.DataTextField = "LookupDesc"
                vDDL.DataValueField = "LookupCode"
                Dim vList As New ParameterList(HttpContext.Current)
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtSexes, vDDL, True, vList)
              Case "ClaimDay"
                vDDL.DataTextField = "ClaimDay"
                vDDL.DataValueField = "ClaimDay"
                If DataHelper.ControlValue(DataHelper.ControlValues.auto_pay_claim_date_method) = "D" Then
                  If DataHelper.ConfigurationOption(DataHelper.ConfigurationOptions.fp_default_blank_claim_day, False) Then
                    DataHelper.FillComboWithRestriction(CareNetServices.XMLLookupDataTypes.xldtBankAccountClaimDays, vDDL, True, Nothing, "(BankAccount = '" & DefaultParameters("BankAccount").ToString & "' AND ClaimType = 'DD') OR (BankAccount IS NULL)")
                  Else
                    DataHelper.FillComboWithRestriction(CareNetServices.XMLLookupDataTypes.xldtBankAccountClaimDays, vDDL, False, Nothing, "BankAccount = '" & DefaultParameters("BankAccount").ToString & "' AND ClaimType = 'DD'")
                  End If
                Else
                  vDDL.Enabled = False
                End If
              Case "LabelNameFormatCode"
                vDDL.DataTextField = "LabelNameFormatCodeDesc"
                vDDL.DataValueField = "LabelNameFormatCode"
              Case "BranchName"
                vDDL.DataTextField = "BranchName"
                vDDL.DataValueField = "BranchName"
                If GetDropDownValue("Bank").Length > 0 Then
                  Dim vList As New ParameterList(HttpContext.Current)
                  vList("Bank") = GetDropDownValue("Bank")
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtBanks, vDDL, True, vList)
                End If
              Case "Bank"
                vDDL.DataTextField = "Bank"
                vDDL.DataValueField = "Bank"
                Dim vList As New ParameterList(HttpContext.Current)
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtBankNames, vDDL, True, vList)
              Case "STDCode"
                vDDL.DataTextField = "STDCodeDesc"
                vDDL.DataValueField = "STDCode"
                Dim vList As New ParameterList(HttpContext.Current)
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtSTDCodes, vDDL, True, vList)
              Case "DiallingCode"
                vDDL.DataTextField = "DiallingCodeDesc"
                vDDL.DataValueField = "DiallingCode"
                Dim vList As New ParameterList(HttpContext.Current)
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtDiallingCodes, vDDL, True, vList)
              Case "Device"
                vDDL.DataTextField = "DeviceDesc"
                vDDL.DataValueField = "Device"
                Dim vList As New ParameterList(HttpContext.Current)
                If InitialParameters.ContainsKey("DeviceLookupGroup") Then vList("LookupGroup") = InitialParameters("DeviceLookupGroup")
                AddSelectedIndexChangedHandler(vDDL)
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtDevices, vDDL, True, vList)
              Case "CommunicationUsage"
                vDDL.DataTextField = "CommunicationUsageDesc"
                vDDL.DataValueField = "CommunicationUsage"
                Dim vList As New ParameterList(HttpContext.Current)
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtCommunicationUsages, vDDL, True, vList)
              Case "AddressNumber"
                Dim vList As New ParameterList(HttpContext.Current)
                vList("ContactNumber") = UserContactNumber()
                Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddresses, vList)
                If vTable IsNot Nothing Then
                  vTable.Rows.InsertAt(vTable.NewRow, 0)
                  vDDL.DataTextField = "AddressLine"
                  vDDL.DataValueField = "AddressNumber"
                  vDDL.DataSource = vTable
                  vDDL.DataBind()
                End If
              Case "CreditCardType"
                vDDL.DataTextField = "CreditCardTypeDesc"
                vDDL.DataValueField = "CreditCardType"
                Dim vList As New ParameterList(HttpContext.Current)
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtCreditCardTypes, vDDL, True, vList)
              Case "PositionFunction"
                vDDL.DataTextField = "PositionFunctionDesc"
                vDDL.DataValueField = "PositionFunction"
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtPositionFunctions, vDDL, True)
              Case "PositionSeniority"
                vDDL.DataTextField = "PositionSeniorityDesc"
                vDDL.DataValueField = "PositionSeniority"
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtPositionSeniorities, vDDL, True)
              Case "CpdCycleType"
                Dim vCycleTypeList As New ParameterList(HttpContext.Current)
                vCycleTypeList("ForPortal") = "Y"
                vDDL.DataTextField = "CPDCycleTypeDesc"
                vDDL.DataValueField = "CPDCycleType"
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtCPDCycleTypes, vDDL, True, vCycleTypeList)
              Case "CpdCycleStatus"
                Dim vCycleStatusList As New ParameterList(HttpContext.Current)
                vCycleStatusList("ForPortal") = "Y"
                vDDL.DataTextField = "CPDCycleStatusDesc"
                vDDL.DataValueField = "CPDCycleStatus"
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtCPDCycleStatuses, vDDL, True, vCycleStatusList)
              Case "CpdCategory"
                Dim vCpdCategory As New ParameterList(HttpContext.Current)
                vDDL.DataTextField = "CpdCategoryDesc"
                vDDL.DataValueField = "CpdCategory"
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtCPDCategories, vDDL, True, vCpdCategory)
              Case "CpdCategoryType"
                Dim vCatrgoryTypeList As New ParameterList(HttpContext.Current)
                vCatrgoryTypeList("ForPortal") = "Y"
                vDDL.DataTextField = "CpdCategoryTypeDesc"
                vDDL.DataValueField = "CpdCategoryType"
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtCPDCategoryTypes, vDDL, True, vCatrgoryTypeList)
              Case "CpdItemType"
                vDDL.DataValueField = "CpdItemType"
                vDDL.DataTextField = "CpdItemTypeDesc"
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtCPDItemTypes, vDDL, True)
              Case "WebPublish"
                vDDL.DataValueField = "LookupCode"
                vDDL.DataTextField = "LookupDesc"
                DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtWebPublishFlags, vDDL, True)
            End Select
            vDDL.Visible = True
            vHTMLCell.Controls.Add(vDDL)      'Add it to the cell
            CheckRequiredValidator(vRow, vParameterName, False, vHTMLCell, pPageType)
            CheckReadOnlyField(vHTMLCell, vParameterName, vRow)
            Select Case vParameterName
              Case "PostcoderAddress"
                vDDL.Items.Add("Enter postcode then choose address from list")
                AddTBUpdatePanel(vHTMLCell, TryCast(FindControlByName(mvTableControl, "PostcoderPostcode"), TextBox))
              Case "Bank", "BranchName"
                If mvControlType = CareNetServices.WebControlTypes.wctAddBankAccount Then
                  AddUpdatePanel(vHTMLCell)
                  AddSelectedIndexChangedHandler(vDDL)
                End If
              Case "LabelNameFormatCode"
                'This will add the SelectedIndexChanged handler for Title. Make sure not to add it again when using ajax for contact names
                AddDDLUpdatePanel(vHTMLCell, TryCast(FindControlByName(mvTableControl, "Title"), DropDownList))
              Case "CpdCycleType"
                If mvControlType = CareNetServices.WebControlTypes.wctCPDCycle Then
                  AddSelectedIndexChangedHandler(vDDL)
                End If
              Case "CpdCategoryType", "CpdCategory"
                If mvControlType = CareNetServices.WebControlTypes.wctUpdateCpdPoints OrElse mvControlType = CareNetServices.WebControlTypes.wctUpdateCpdObjectives Then
                  AddSelectedIndexChangedHandler(vDDL)
                End If
              Case "StartDateList"
                If mvControlType = CareNetServices.WebControlTypes.wctSelectMembershipTypes Then
                  AddSelectedIndexChangedHandler(vDDL)
                  Dim vList As List(Of Date) = GetMembershipStartDate()
                  If vList.Count > 0 Then
                    For Each vDate As Date In vList
                      vDDL.Items.Add(vDate.ToShortDateString)
                    Next
                    If vList.Count > 0 Then
                      Dim vSelectedDate As Date = GetNearestDate(vList)
                      If vSelectedDate <> Nothing Then
                        vDDL.SelectedValue = vSelectedDate.ToShortDateString
                      End If
                    End If
                  End If
                End If
            End Select
            vHTMLRow.Cells.Add(vHTMLCell)
            'Now add the row to the table
            vHtmlTable.Rows.Add(vHTMLRow)

          Case "txt", "rdo", "dtp"
            Dim vReadOnly As Boolean = (vType = "rdo")
            Dim vHTMLRow As New HtmlTableRow
            Dim vControlID As String
            'First Add the label
            Dim vHTMLCell As New HtmlTableCell
            Dim vFieldType As FieldTypes
            vHTMLCell.InnerHtml = vRow("ControlCaption").ToString
            vHTMLCell.Attributes("Class") = "DataEntryLabel"
            vHTMLRow.Cells.Add(vHTMLCell)
            vControlID = vParameterName

            vHTMLCell = New HtmlTableCell
            If vNumberOfColumns > 0 Then
              vHTMLCell.ColSpan = (vNumberOfColumns * 2) - 1
            End If
            vFieldType = GetFieldType(vRow("Type").ToString)

            Dim vControl As Control = Nothing

            Dim vUseTextBox As Boolean = False
            If vRow("ParameterName").ToString = "StatusReason" Then
              vUseTextBox = True
            End If

            If vReadOnly = False AndAlso vRow("ValidationTable").ToString.Length > 0 AndAlso vUseTextBox = False Then
              Dim vDDL As New DropDownList
              AddFocusScript(vDDL.Attributes)
              vDDL.ID = vParameterName
              If mvFocusControl Is Nothing Then mvFocusControl = vDDL
              vDDL.CssClass = "DataEntryItem"
              vDDL.Width = New Unit(CInt(vRow("ControlWidth")) / 100, UnitType.Em)
              Select Case vRow("ParameterName").ToString
                Case "ActionPriority"
                  vDDL.DataTextField = "ActionPriorityDesc"
                  vDDL.DataValueField = "ActionPriority"
                  Dim vList As New ParameterList(HttpContext.Current)
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtActionPriorities, vDDL, True, vList)
                Case "ActivityValue"
                  vDDL.DataTextField = "ActivityValueDesc"
                  vDDL.DataValueField = "ActivityValue"
                  Dim vList As New ParameterList(HttpContext.Current)
                  vList("Activity") = InitialParameters("Activity")
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtActivityValues, vDDL, True, vList)
                Case "Branch"
                  vDDL.DataTextField = "Name"
                  vDDL.DataValueField = "Branch"
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtBranches, vDDL, True)
                Case "Country"
                  vDDL.DataTextField = "CountryDesc"
                  vDDL.DataValueField = "Country"
                  Select Case pPageType
                    Case CareNetServices.WebControlTypes.wctSearchDirectory
                      DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtCountries, vDDL, True)
                      SelectListItem(vDDL, "")
                    Case Else
                      DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtCountries, vDDL)
                      SelectListItem(vDDL, "UK")
                      AddSelectedIndexChangedHandler(vDDL)
                  End Select
                Case "DistributionCode"
                  vDDL.DataTextField = "DistributionCodeDesc"
                  vDDL.DataValueField = "DistributionCode"
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtDistributionCodes, vDDL, True)
                Case "DocumentClass"
                  vDDL.DataTextField = "DocumentClassDesc"
                  vDDL.DataValueField = "DocumentClass"
                  Dim vList As New ParameterList(HttpContext.Current)
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtDocumentClasses, vDDL, True, vList)
                Case "EventGroup"
                  vDDL.DataTextField = "EventGroupDesc"
                  vDDL.DataValueField = "EventGroup"
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtEventGroups, vDDL, True)
                Case "Organiser"
                  vDDL.DataTextField = "OrganiserDesc"
                  vDDL.DataValueField = "Organiser"
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtEventOrganisers, vDDL, True)
                Case "Relationship"
                  vDDL.DataTextField = "RelationshipDesc"
                  vDDL.DataValueField = "Relationship"
                  Dim vList As New ParameterList(HttpContext.Current)
                  vList("FromContactGroup") = "CON"
                  vList("ToContactGroup") = "CON"
                  If InitialParameters.ContainsKey("RelationshipLookupGroup") Then vList("LookupGroup") = InitialParameters("RelationshipLookupGroup")
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtRelationships, vDDL, True, vList)
                  If mvControlType = CareNetServices.WebControlTypes.wctAddRelatedContact Then
                    AddSelectedIndexChangedHandler(vDDL)
                  End If
                Case "SkillLevel"
                  vDDL.DataTextField = "SkillLevelDesc"
                  vDDL.DataValueField = "SkillLevel"
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtEventSkillLevels, vDDL, True)
                Case "Status"
                  vDDL.DataTextField = "StatusDesc"
                  vDDL.DataValueField = "Status"
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtStatuses, vDDL, True)
                Case "SubTopic"
                  vDDL.DataTextField = "SubTopicDesc"
                  vDDL.DataValueField = "SubTopic"
                  If FindControlByName(mvTableControl, "Topic") Is Nothing Then
                    Dim vList As New ParameterList(HttpContext.Current)
                    If InitialParameters.ContainsKey("Topic") Then
                      vList("Topic") = InitialParameters("Topic")
                    Else
                      vList("Topic") = DefaultParameters("Topic")   'Used in AddCommunicationNote
                    End If
                    DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtSubTopics, vDDL, True, vList)
                  End If
                Case "Topic"
                  vDDL.DataTextField = "TopicDesc"
                  vDDL.DataValueField = "Topic"
                  Dim vList As New ParameterList(HttpContext.Current)
                  If InitialParameters.ContainsKey("TopicLookupGroup") Then vList("LookupGroup") = InitialParameters("TopicLookupGroup")
                  If InitialParameters.ContainsKey("Topic") Then vList("Topic") = InitialParameters("Topic")
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtTopics, vDDL, True, vList)
                Case "Title"
                  vDDL.DataTextField = "Title"
                  vDDL.DataValueField = "Title"
                  Dim vList As New ParameterList(HttpContext.Current)
                  If InitialParameters.ContainsKey("TitleLookupGroup") Then vList("LookupGroup") = InitialParameters("TitleLookupGroup")
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtTitles, vDDL, True, vList)
                Case "Venue"
                  vDDL.DataTextField = "VenueDesc"
                  vDDL.DataValueField = "Venue"
                  DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtEventVenues, vDDL, True)
              End Select
              If vDDL IsNot Nothing Then
                vDDL.Visible = True
                vHTMLCell.Controls.Add(vDDL)      'Add it to the cell
              End If
            Else
              Dim vTextBox As New TextBox
              vControl = vTextBox
              AddFocusScript(vTextBox.Attributes)
              If vType.StartsWith("txt_") Then
                vParameterName += vType.Substring(3)
                vTextBox.ID = vParameterName
              Else
                vTextBox.ID = vParameterName
              End If
              vControlID = vTextBox.ID
              If mvFocusControl Is Nothing AndAlso vReadOnly = False Then mvFocusControl = vTextBox
              vTextBox.CssClass = "DataEntryItem"
              Dim vEntryLength As Integer = CInt(vRow("EntryLength"))
              If vEntryLength > 0 Then
                Dim vDevice As String = ""
                Select Case vParameterName
                  Case "DirectNumber"
                    vDevice = DataHelper.ControlValue(DataHelper.ControlValues.direct_device)
                  Case "SwitchboardNumber"
                    vDevice = DataHelper.ControlValue(DataHelper.ControlValues.switchboard_device)
                  Case "MobileNumber"
                    vDevice = DataHelper.ControlValue(DataHelper.ControlValues.mobile_device)
                  Case "FaxNumber"
                    vDevice = DataHelper.ControlValue(DataHelper.ControlValues.fax_device)
                  Case "EMailAddress"
                    vDevice = DataHelper.ControlValue(DataHelper.ControlValues.email_device)
                  Case "WebAddress"
                    vDevice = DataHelper.ControlValue(DataHelper.ControlValues.web_device)
                  Case "AdditionalNumber1"
                    vDevice = mvInitialParameters.OptionalValue("Device1")
                  Case "AdditionalNumber2"
                    vDevice = mvInitialParameters.OptionalValue("Device2")
                  Case "AdditionalNumber3"
                    vDevice = mvInitialParameters.OptionalValue("Device3")
                  Case "Quantity", "AdultQuantity", "ChildQuantity"
                    If mvControlType = CareNetServices.WebControlTypes.wctBookEventCC Then
                      vEntryLength = 2
                    End If
                  Case "MemberNumber"
                    If mvControlType = CareNetServices.WebControlTypes.wctLogin Then
                      vEntryLength = 9
                    End If
                End Select
                If vDevice.Length > 0 Then vEntryLength = DataHelper.DeviceMaxLength(vDevice)
                vTextBox.MaxLength = vEntryLength
              Else
                vTextBox.TextMode = TextBoxMode.MultiLine
                vTextBox.Wrap = True
                vTextBox.Height = New Unit((CInt(vRow("ControlHeight")) / 300) + 1, UnitType.Em)
              End If
              vTextBox.Width = New Unit(CInt(vRow("ControlWidth")) / 100, UnitType.Em)
              vTextBox.Visible = True
              'CheckReadOnlyField(vHTMLCell, vParameterName, vRow)

              Select Case vParameterName
                Case "Address", "Town", "County", "Postcode"
                  If Not vUseQAS.HasValue Then vUseQAS = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.qas_pro_web_url).Length > 0
                  If vVisible Then
                    If vUseQAS.Value = True Then AddTextChangedHandler(vTextBox)
                    vHTMLCell.Controls.Add(vTextBox)    'Add it to the cell
                  End If
                Case "ContactNumber"
                  vTextBox.AutoPostBack = True
                  AddHandler vTextBox.TextChanged, AddressOf ContactNumberChangedHandler
                  Dim vTB As New Label
                  vTB.ID = vParameterName & "_Desc"
                  vTB.CssClass = "DataEntryViewItem"
                  vTB.Visible = True
                  vTB.Width = New Unit(20, UnitType.Em)

                  If pPageType = CareNetServices.WebControlTypes.wctContactSelection OrElse pPageType = CareNetServices.WebControlTypes.wctRapidProductPurchase Then
                    Dim vPanel As New Panel
                    vPanel.Controls.Add(vTextBox)
                    vPanel.Controls.Add(vTB)
                    vHTMLCell.Controls.Add(vPanel)    'Add it to the cell
                  Else
                    Dim vPanel As New UpdatePanel
                    vPanel.ContentTemplateContainer.Controls.Add(vTextBox)
                    vPanel.ContentTemplateContainer.Controls.Add(vTB)
                    vHTMLCell.Controls.Add(vPanel)    'Add it to the cell
                  End If
                Case "ExternalReference"
                  If pPageType = CareNetServices.WebControlTypes.wctContactSelectionExternalRef OrElse _
                    pPageType = CareNetServices.WebControlTypes.wctRapidProductPurchase Then
                    vTextBox.AutoPostBack = True
                    AddHandler vTextBox.TextChanged, AddressOf ExternalReferenceChangedHandler
                    Dim vTB As New Label
                    vTB.ID = vParameterName & "_Desc"
                    vTB.CssClass = "DataEntryViewItem"
                    vTB.Visible = True
                    vTB.Width = New Unit(20, UnitType.Em)

                    Dim vPanel As New Panel
                    vPanel.Controls.Add(vTextBox)
                    vPanel.Controls.Add(vTB)
                    vHTMLCell.Controls.Add(vPanel)    'Add it to the cell                    
                  Else
                    vHTMLCell.Controls.Add(vTextBox)    'Add it to the cell
                  End If

                Case "FriendlyUrl"
                  vTextBox.AutoPostBack = True
                  AddHandler vTextBox.TextChanged, AddressOf FriendlyUrlChangedHandler
                  vHTMLCell.Controls.Add(vTextBox)    'Add it to the cell
                Case "StartYear" ', "EndYear"
                  vHTMLCell.Controls.Add(vTextBox)    'Add it to the cell
                  If pPageType = CareNetServices.WebControlTypes.wctCPDCycle Then
                    vTextBox.MaxLength = 4
                    AddTextChangedHandler(vTextBox)
                    AddRangeValidator(vHTMLCell, vTextBox.ID, 2000, 3000, "Start Year must be between 2000 to 3000")
                  End If
                Case "EndYear"
                  vHTMLCell.Controls.Add(vTextBox)    'Add it to the cell
                  If pPageType = CareNetServices.WebControlTypes.wctCPDCycle Then
                    vTextBox.MaxLength = 4
                    AddRangeValidator(vHTMLCell, vTextBox.ID, 2000, 3000, "End Year must be between 2000 to 3000")
                  End If
                Case "AdultQuantity", "ChildQuantity"
                  AddTextChangedHandler(vTextBox)
                  vHTMLCell.Controls.Add(vTextBox)
                Case Else
                  vHTMLCell.Controls.Add(vTextBox)    'Add it to the cell
              End Select

              Dim vDateItem As Boolean = vType.StartsWith("dtp")
              Select Case vTextBox.ID
                Case "AccountNumber"
                  If pPageType = CareNetServices.WebControlTypes.wctAddBankAccount Then
                    AddRegExValidator(vHTMLCell, vParameterName, RegularExpressionTypes.retEncryptedAccountNumber)
                  Else
                    AddRegExValidator(vHTMLCell, vParameterName, RegularExpressionTypes.retAccountNumber)
                  End If
                  Dim vBankValidation As String = ""
                  If DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.afd_software).Length > 0 Then
                    vBankValidation = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.afd_software).ToLower
                  Else
                    vBankValidation = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.albacs_verify).ToLower
                  End If
                  Select Case vBankValidation
                    Case "error", "warn", "both", "bank"
                      AddCustomValidator(vHTMLCell, vParameterName, "", vParameterName)
                  End Select
                Case "CpdPoints", "CpdPoints2"
                  If vReadOnly = False Then
                    If DataHelper.ConfigurationOption(DataHelper.ConfigurationOptions.cpd_points_allow_numeric) Then
                      AddDoubleValidator(vHTMLCell, vParameterName)
                    Else
                      AddDataTypeValidator(vHTMLCell, vParameterName, ValidationDataType.Integer)
                    End If
                  End If
                Case "Amount", "DonationAmount", "TargetAmount"
                  If vReadOnly = False Then AddDoubleValidator(vHTMLCell, vParameterName)
                Case "Quantity"
                  If vReadOnly = False AndAlso vFieldType = FieldTypes.cftNumeric Then AddDoubleValidator(vHTMLCell, vParameterName)
                Case "CreditCardNumber"
                  If InitialParameters.OptionalValue("PaymentService").Trim.ToUpper <> TNSHOSTED AndAlso _
                     Not (mvControlType = CareNetServices.WebControlTypes.wctAddMemberCC AndAlso InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" AndAlso _
                          DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED) Then
                    AddRegExValidator(vHTMLCell, vParameterName, RegularExpressionTypes.retCreditCardNumber)
                  End If
                Case "gatewayCardExpiryDateYear"
                  AddRegExValidator(vHTMLCell, vParameterName, RegularExpressionTypes.retExpiryDateYear)
                Case "EMailAddress"
                  AddEmailValidator(vHTMLCell, vParameterName)
                Case "CardExpiryDate"
                  If InitialParameters.OptionalValue("PaymentService").Trim.ToUpper <> TNSHOSTED AndAlso _
                    Not (mvControlType = CareNetServices.WebControlTypes.wctAddMemberCC AndAlso InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" AndAlso _
                                DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED) Then
                    AddRegExValidator(vHTMLCell, vParameterName, RegularExpressionTypes.retExpiryDate)
                    AddCustomValidator(vHTMLCell, vParameterName, "Invalid expiry date", vParameterName)
                  End If
                Case "CardStartDate"
                  If InitialParameters.OptionalValue("PaymentService").Trim.ToUpper <> TNSHOSTED AndAlso _
                    Not (mvControlType = CareNetServices.WebControlTypes.wctAddMemberCC AndAlso InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" AndAlso _
                      DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED) Then
                    AddRegExValidator(vHTMLCell, vParameterName, RegularExpressionTypes.retStartDate)
                    AddCustomValidator(vHTMLCell, vParameterName, "Invalid start date", vParameterName)
                  End If
                Case "SecurityCode"
                  If InitialParameters.OptionalValue("PaymentService").Trim.ToUpper <> TNSHOSTED AndAlso _
                    Not (mvControlType = CareNetServices.WebControlTypes.wctAddMemberCC AndAlso InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" AndAlso _
                      DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED) Then
                    AddRegExValidator(vHTMLCell, vParameterName, RegularExpressionTypes.retSecurityCode)
                  End If

                Case "ConfirmEMailAddress"
                  AddCompareValidator(vHTMLCell, vParameterName, "EMailAddress", ValidationCompareOperator.Equal)
                Case "Forenames"
                  vTextBox.Attributes("onChange") = "javascript:CapitaliseWords(this)"
                Case "IssueNumber"
                  If InitialParameters.OptionalValue("PaymentService").Trim.ToUpper <> TNSHOSTED AndAlso _
                          Not (mvControlType = CareNetServices.WebControlTypes.wctAddMemberCC AndAlso InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" AndAlso _
                                      DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED) Then
                    AddRegExValidator(vHTMLCell, vParameterName, RegularExpressionTypes.retIssueNumber)
                  End If

                Case "Password", "OldPassword", "NewPassword"
                  vTextBox.TextMode = TextBoxMode.Password
                  vPasswordParameter = vParameterName
                  If mvControlType <> CareNetServices.WebControlTypes.wctLogin Then
                    If DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.portal_password_complexity) = "C" OrElse IntegerValue(DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.portal_password_min_length)) > 1 Then
                      AddRegExValidator(vHTMLCell, vParameterName, RegularExpressionTypes.retPassword)
                    End If
                  End If
                Case "ConfirmPassword"
                  vTextBox.TextMode = TextBoxMode.Password
                  AddCompareValidator(vHTMLCell, vParameterName, vPasswordParameter, ValidationCompareOperator.Equal)
                Case "Postcode", "Town", "PostcoderPostcode"
                  vTextBox.Attributes("onChange") = "javascript:UpperCaseField(this)"
                Case "SortCode"
                  vTextBox.MaxLength += 2
                  If pPageType = CareNetServices.WebControlTypes.wctAddBankAccount Then
                    AddUpdatePanel(vHTMLCell)
                    AddTextChangedHandler(vTextBox)
                    AddRegExValidator(vHTMLCell, vParameterName, RegularExpressionTypes.retEncryptedSortCode)
                  Else
                    AddRegExValidator(vHTMLCell, vParameterName, RegularExpressionTypes.retSortCode)
                  End If
                Case "ValidFrom", "ValidTo", "DateOfBirth", "Dated", "TargetDate", "ActivityDate", "ScheduledOn", "CompletedOn", "StatusDate"
                  vDateItem = True
                Case "Started", "Finished"
                  vDateItem = True
                Case "SecurityAnswer"
                  AddCustomValidator(vHTMLCell, vParameterName, "Question is mandatory for answer", vParameterName)
                Case "SecurityQuestion"
                  AddCustomValidator(vHTMLCell, vParameterName, "Answer is mandatory for question", vParameterName)
                Case "Number"
                  If pPageType = CareNetServices.WebControlTypes.wctUpdateEmailAddress Then
                    AddEmailValidator(vHTMLCell, vParameterName)
                  Else
                    AddTelephoneNumberValidator(vHTMLCell, vParameterName)
                  End If
                Case "DirectNumber", "SwitchboardNumber", "MobileNumber", "FaxNumber"
                  AddTelephoneNumberValidator(vHTMLCell, vParameterName)
              End Select
              If vDateItem Then
                Dim vButton As New HtmlInputButton
                vButton.ID = "cmdFind" & vParameterName
                vButton.Attributes("value") = "..."
                vButton.Attributes("class") = "Button"
                vButton.Attributes("style") = "width:2em;"
                vButton.CausesValidation = False
                If mvDateControls Is Nothing Then mvDateControls = New List(Of String)
                mvDateControls.Add(vParameterName)
                vHTMLCell.Controls.Add(vButton)
                CheckReadOnlyField(vHTMLCell, vButton.ID, vRow)
                AddDateValidator(vHTMLCell, vParameterName)
                Select Case vParameterName
                  Case "ValidFrom"
                    If IsTableItemVisible(pTable, "ValidTo") Then
                      AddDateCompareValidator(vHTMLCell, vParameterName, "ValidTo", ValidationCompareOperator.LessThanEqual)
                    End If
                  Case "ValidTo"
                    If IsTableItemVisible(pTable, "ValidFrom") Then
                      AddDateCompareValidator(vHTMLCell, vParameterName, "ValidFrom", ValidationCompareOperator.GreaterThanEqual)
                    End If
                  Case "DateOfBirth"
                    If IsTableItemVisible(pTable, "Today") Then
                      AddDateCompareValidator(vHTMLCell, vParameterName, "Today", ValidationCompareOperator.LessThanEqual)
                    End If
                  Case "CompletedOn"
                    If IsTableItemVisible(pTable, "ScheduledOn") Then
                      AddDateCompareValidator(vHTMLCell, vParameterName, "ScheduledOn", ValidationCompareOperator.GreaterThanEqual)
                    End If
                  Case "ScheduledOn"
                    If IsTableItemVisible(pTable, "CompletedOn") Then
                      AddDateCompareValidator(vHTMLCell, vParameterName, "CompletedOn", ValidationCompareOperator.LessThanEqual)
                    End If
                  Case "Started"
                    If IsTableItemVisible(pTable, "Finished") Then
                      AddDateCompareValidator(vHTMLCell, vParameterName, "Finished", ValidationCompareOperator.LessThanEqual)
                    End If
                  Case "Finished"
                    If IsTableItemVisible(pTable, "Started") Then
                      AddDateCompareValidator(vHTMLCell, vParameterName, "Started", ValidationCompareOperator.GreaterThanEqual)
                    End If
                End Select
              End If
              If vType = "rdo" Then vRow("ReadOnlyItem") = "Y"
            End If
            CheckRequiredValidator(vRow, vParameterName, vReadOnly, vHTMLCell, pPageType)
            CheckReadOnlyField(vHTMLCell, vControlID, vRow)

            Select Case vParameterName
              Case "Address", "County", "Postcode", "Country", "Town"
                If vUseQAS.HasValue AndAlso vUseQAS.Value = True Then AddDDLUpdatePanel(vHTMLCell, TryCast(FindControlByName(mvTableControl, "PostcoderAddress"), DropDownList))
                If vParameterName = "Town" Then AddCustomValidator(vHTMLCell, vParameterName, "Required Field", vParameterName)
              Case "SubTopic"
                AddDDLUpdatePanel(vHTMLCell, TryCast(FindControlByName(mvTableControl, "Topic"), DropDownList))
              Case "TargetDate"
                AddDDLUpdatePanel(vHTMLCell, TryCast(FindControlByName(mvTableControl, "EventNumber"), DropDownList))
              Case "FundraisingDescription"
                AddDDLUpdatePanel(vHTMLCell, TryCast(FindControlByName(mvTableControl, "EventNumber"), DropDownList))
              Case "TotalAmount"
                If mvControlType = CareNetServices.WebControlTypes.wctBookEventCC Or mvControlType = CareNetServices.WebControlTypes.wctProductPurchase _
                  Or mvControlType = CareNetServices.WebControlTypes.wctBookEvent Or mvControlType = CareNetServices.WebControlTypes.wctProductPurchaseCC Then
                  AddTBUpdatePanel(vHTMLCell, TryCast(FindControlByName(mvTableControl, "Quantity"), TextBox))
                End If
              Case "Quantity"
                If mvControlType = CareNetServices.WebControlTypes.wctBookEventCC Or mvControlType = CareNetServices.WebControlTypes.wctProductPurchase _
                  Or mvControlType = CareNetServices.WebControlTypes.wctProductPurchaseCC Then
                  AddTBUpdatePanel(vHTMLCell, TryCast(FindControlByName(mvTableControl, "Amount"), TextBox))
                End If
              Case "Number"
                AddUpdatePanel(vHTMLCell)
                AddTBUpdatePanel(vHTMLCell, DirectCast(vControl, TextBox))
            End Select
            vHTMLRow.Cells.Add(vHTMLCell)
            If vHelpText.Length > 0 Then
              Dim vHelpCell As New HtmlTableCell
              vHelpCell.InnerHtml = vHelpText
              vHelpCell.Attributes("Class") = "DataEntryHelp"
              vHTMLRow.Cells.Add(vHelpCell)
            End If
            'Now add the row to the table
            vHtmlTable.Rows.Add(vHTMLRow)

          Case "btn"
            Dim vHTMLRow As HtmlTableRow = Nothing
            Dim vHTMLCell As HtmlTableCell = Nothing

            For Each vRowHtml As HtmlTableRow In vHtmlTable.Rows
              If vRowHtml.ID = "trBtn_" & vRow("ControlTop").ToString Then
                vHTMLRow = vRowHtml
                For Each vCellHtml As HtmlTableCell In vHTMLRow.Cells
                  If vCellHtml.ID = "tcBtn_" & vRow("ControlTop").ToString Then
                    vHTMLCell = vCellHtml
                    Exit For
                  End If
                Next
                Exit For
              End If
            Next
            If vHTMLRow Is Nothing OrElse vHTMLCell Is Nothing Then
              vHTMLRow = New HtmlTableRow
              vHTMLRow.ID = "trBtn_" & vRow("ControlTop").ToString
              vHTMLCell = New HtmlTableCell
              If Not mvCenterControl Then vHTMLRow.Cells.Add(vHTMLCell)
              vHTMLCell = New HtmlTableCell
              vHTMLCell.ID = "tcBtn_" & vRow("ControlTop").ToString
            End If

            Dim vButton As New Button
            AddFocusScript(vButton.Attributes)
            vButton.ID = vParameterName
            vButton.Text = vRow("ControlCaption").ToString
            vButton.CssClass = "Button"
            Select Case mvControlType
              Case CareNetServices.WebControlTypes.wctSelectEventDelegates
                If vParameterName = "AddDelegate" Then
                  vButton.ValidationGroup = "SelectEventDelegates"
                End If
              Case CareNetServices.WebControlTypes.wctPayMultiplePaymentPlans
                'BR19328
                If vButton.Text = "Submit" Then
                  vButton.ID = "Submit"
                End If
            End Select

            If vParameterName = "PrintButton" Then
              vButton.Attributes("onClick") = "javascript:SendPagetoPrinter()"
            End If

            AddHandler vButton.Click, AddressOf ButtonClickHandler

            ' Set CausesValidation to False for "New" Button
            If vButton.ID = "New" OrElse vButton.ID = "Cancel" Then
              vButton.CausesValidation = False
            End If
            vHTMLCell.Controls.Add(vButton)
            CheckReadOnlyField(vHTMLCell, vParameterName, vRow)
            If mvCenterControl Then
              vHtmlTable.Align = "Center"
              vHTMLCell.Align = "Center"
            End If
            vHTMLRow.Cells.Add(vHTMLCell)

            ' Append space after each buttons
            Dim vSpaceForButtons As New LiteralControl()
            vSpaceForButtons.Text = "&nbsp;&nbsp;&nbsp;"
            vHTMLCell.Controls.Add(vSpaceForButtons)

            'Now add the row to the table
            vHtmlTable.Rows.Add(vHTMLRow)

          Case "fil"
            Dim vHTMLRow As New HtmlTableRow
            'First Add the label
            Dim vHTMLCell As New HtmlTableCell
            vHTMLCell.InnerHtml = vRow("ControlCaption").ToString
            vHTMLCell.Attributes("Class") = "DataEntryLabel"
            vHTMLRow.Cells.Add(vHTMLCell)

            Dim vFile As New HtmlInputFile
            vFile.ID = vParameterName
            vFile.Attributes("Accept") = "image/*"
            vFile.MaxLength = 254
            vFile.Attributes("Class") = "DataEntryItem"
            vHTMLCell = New HtmlTableCell
            vHTMLCell.Controls.Add(vFile)
            vHTMLRow.Cells.Add(vHTMLCell)
            'Now add the row to the table
            vHtmlTable.Rows.Add(vHTMLRow)

          Case "img"
            Dim vHTMLRow As New HtmlTableRow
            'First Add the label
            Dim vHTMLCell As New HtmlTableCell
            vHTMLCell.InnerHtml = vRow("ControlCaption").ToString
            vHTMLCell.Attributes("Class") = "DataEntryLabel"
            vHTMLRow.Cells.Add(vHTMLCell)

            '<div style="position:absolute;">
            '<div style="position:absolute; top: 0px ;left: 0px; width:200px; height:20px "><img alt="meter" src="meter.jpg" width="200px", height="20px" /></div>
            '<div style="background:red; position:absolute; top:2px; left:2px; width:20px; overflow:hidden; height:14px;"></div>
            '	</div>
            Dim vDiv As New HtmlGenericControl("div")
            vDiv.Style.Add("position", "absolute")

            Dim vDivImage As New HtmlGenericControl("div")
            vDivImage.Style.Add("position", "absolute")
            vDivImage.Style.Add("top", "0px")
            vDivImage.Style.Add("left", "0px")
            vDivImage.Style.Add("width", "150px")
            vDivImage.Style.Add("height", "20px")
            vDiv.Controls.Add(vDivImage)

            Dim vFile As New Image
            vFile.ID = vParameterName
            vFile.ImageUrl = "~/images/meter.jpg"
            vFile.Width = New Unit(200)
            vFile.Height = New Unit(20)
            vFile.Attributes("Class") = "DataEntryItem"
            vDivImage.Controls.Add(vFile)

            Dim vDivProgress As New HtmlGenericControl("div")
            vDivProgress.Style.Add("background", "red")
            vDivProgress.Style.Add("position", "absolute")
            vDivProgress.Style.Add("top", "2px")
            vDivProgress.Style.Add("left", "2px")
            vDivProgress.Style.Add("width", "20px")
            vDivProgress.Style.Add("height", "14px")
            vDivProgress.Style.Add("overflow", "hidden")
            vDivProgress.ID = "DivProgress"
            vDiv.Controls.Add(vDivProgress)

            vHTMLCell = New HtmlTableCell
            vHTMLCell.Controls.Add(vDiv)
            vHTMLRow.Cells.Add(vHTMLCell)
            vHTMLRow.Height = "22"
            'Now add the row to the table
            vHtmlTable.Rows.Add(vHTMLRow)
          Case "opt"
            Dim vHTMLRow As HtmlTableRow
            If mvNumberOfRows > 0 Then
              If vCurrentRow >= mvNumberOfRows Then
                vCurrentRow = 0
                vNumberOfColumns += 1
              End If
              If vHTMLRows(vCurrentRow) Is Nothing Then
                For vIndex As Integer = 0 To mvNumberOfRows - 1
                  vHTMLRows(vIndex) = New HtmlTableRow
                Next
                vNumberOfColumns = 1
              End If
              vHTMLRow = vHTMLRows(vCurrentRow)
              vCurrentRow += 1
            Else
              vHTMLRow = New HtmlTableRow
            End If
            'First Add the label
            Dim vHTMLCell As New HtmlTableCell
            vHTMLCell.InnerHtml = vRow("ControlCaption").ToString
            vHTMLCell.Attributes("Class") = "DataEntryLabel"
            vHTMLRow.Cells.Add(vHTMLCell)

            vHTMLCell = New HtmlTableCell
            Dim vRadioButton As New RadioButton
            AddFocusScript(vRadioButton.InputAttributes)
            vRadioButton.ID = vParameterName.TrimEnd("0123456789".ToCharArray) & vType.Substring(3)
            vRadioButton.CssClass = "DataEntryCheckBox"
            vRadioButton.Width = New Unit(CInt(vRow("ControlWidth")) / 100, UnitType.Em)
            vRadioButton.Visible = True
            vRadioButton.GroupName = vParameterName.TrimEnd("0123456789".ToCharArray) & mvWebPageItemNumber
            If vFirstOption Then
              vRadioButton.Checked = True
              mvRadioButtons = New List(Of RadioButton)
            End If
            vFirstOption = False
            mvRadioButtons.Add(vRadioButton)
            vHTMLCell.Controls.Add(vRadioButton)    'Add it to the cell
            CheckReadOnlyField(vHTMLCell, vParameterName.TrimEnd("0123456789".ToCharArray) & vType.Substring(3), vRow)
            vHTMLRow.Cells.Add(vHTMLCell)
            If mvNumberOfRows > 0 Then
              For vIndex As Integer = 0 To mvNumberOfRows - 1
                vHtmlTable.Rows.Add(vHTMLRows(vIndex))
              Next
            Else
              'Now add the row to the table
              vHtmlTable.Rows.Add(vHTMLRow)
            End If

          Case "chk"
            Dim vHTMLRow As HtmlTableRow
            If mvNumberOfRows > 0 Then
              If vCurrentRow >= mvNumberOfRows Then
                vCurrentRow = 0
                vNumberOfColumns += 1
              End If
              If vHTMLRows(vCurrentRow) Is Nothing Then
                For vIndex As Integer = 0 To mvNumberOfRows - 1
                  vHTMLRows(vIndex) = New HtmlTableRow
                Next
                vNumberOfColumns = 1
              End If
              vHTMLRow = vHTMLRows(vCurrentRow)
              vCurrentRow += 1
            Else
              vHTMLRow = New HtmlTableRow
            End If
            'First Add the label
            Dim vHTMLCell As New HtmlTableCell
            vHTMLCell.InnerHtml = vRow("ControlCaption").ToString
            vHTMLCell.Attributes("Class") = "DataEntryLabel"
            vHTMLRow.Cells.Add(vHTMLCell)

            vHTMLCell = New HtmlTableCell
            Dim vCheckBox As New CheckBox
            AddFocusScript(vCheckBox.InputAttributes)
            vCheckBox.ID = vParameterName.TrimEnd("0123456789".ToCharArray) & vType.Substring(3)
            vCheckBox.CssClass = "DataEntryCheckBox"
            vCheckBox.Width = New Unit(CInt(vRow("ControlWidth")) / 100, UnitType.Em)
            vCheckBox.Visible = True
            If mvCheckBoxes Is Nothing Then mvCheckBoxes = New List(Of CheckBox)
            mvCheckBoxes.Add(vCheckBox)

            Select Case pPageType
              Case CType(CareNetServices.WebControlTypes.wctAddCategoryCheckboxes, CareNetServices.WebControlTypes)
                If mvInitialParameters.ContainsKey("CheckBoxMandatory") AndAlso mvInitialParameters("CheckBoxMandatory").ToString = "Y" _
                   AndAlso vRow("MandatoryItem").ToString = "Y" Then
                  AddCustomValidator(vHTMLCell, vCheckBox.ID)
                End If
            End Select
            vHTMLCell.Controls.Add(vCheckBox)    'Add it to the cell
            CheckReadOnlyField(vHTMLCell, vParameterName.TrimEnd("0123456789".ToCharArray) & vType.Substring(3), vRow)


            vHTMLRow.Cells.Add(vHTMLCell)
            If mvNumberOfRows > 0 Then
              For vIndex As Integer = 0 To mvNumberOfRows - 1
                vHtmlTable.Rows.Add(vHTMLRows(vIndex))
              Next
            Else
              'Now add the row to the table
              vHtmlTable.Rows.Add(vHTMLRow)
            End If

          Case "lbl"
            Dim vHTMLRow As New HtmlTableRow
            'First Add the label
            Dim vHTMLCell As New HtmlTableCell
            vHTMLCell.Attributes("Class") = "DataMessage"
            vHTMLCell.ColSpan = 2
            'Add a label to the cell
            If vParameterName = "MailTo" Then
              Dim vHyperLink As New HyperLink
              vHyperLink.ID = vParameterName
              vHyperLink.CssClass = "DataMessage"
              vHyperLink.Visible = True
              vHyperLink.NavigateUrl = "MailTo:?subject=Sponsor Me&body=Please look at this page {0}"
              vHyperLink.Text = vRow("ControlCaption").ToString
              vHyperLink.Enabled = False
              vHTMLCell.Controls.Add(vHyperLink) 'Add it to the cell
            Else
              Dim vLabel As New Label
              vLabel.ID = vRow("ParameterName").ToString
              vLabel.CssClass = "DataMessage"
              vLabel.Visible = True
              vLabel.Text = vRow("ControlCaption").ToString
              vHTMLCell.Controls.Add(vLabel)    'Add it to the cell
              If vParameterName = "PageError" OrElse vParameterName = "ValidationError" OrElse vParameterName = "NoCriteriaError" Then vLabel.CssClass = "PageError"
              If vParameterName.StartsWith("WarningMessage") Then vLabel.CssClass = "WarningMessage"
            End If
            vHTMLRow.Cells.Add(vHTMLCell)
            'Now add the row to the table
            vHtmlTable.Rows.Add(vHTMLRow)

          Case "lst"
            Dim vHTMLRow As New HtmlTableRow
            'First Add the label
            Dim vHTMLCell As New HtmlTableCell
            vHTMLCell.InnerHtml = vRow("ControlCaption").ToString
            vHTMLCell.Attributes("Class") = "DataEntryLabel"
            vHTMLRow.Cells.Add(vHTMLCell)
            vHTMLCell = New HtmlTableCell
            If vNumberOfColumns > 0 Then
              vHTMLCell.ColSpan = (vNumberOfColumns * 2) - 1
            End If
            Dim vListBox As New ListBox
            AddFocusScript(vListBox.Attributes)
            vListBox.ID = vParameterName
            If mvFocusControl Is Nothing Then mvFocusControl = vListBox
            vListBox.CssClass = "DataEntryItem"
            vListBox.Width = New Unit(CInt(vRow("ControlWidth")) / 100, UnitType.Em)
            vListBox.Height = New Unit((CInt(vRow("ControlHeight")) / 300) + 1, UnitType.Em)
            If mvControlType <> CareNetServices.WebControlTypes.wctProcessPayment AndAlso mvControlType <> CareNetServices.WebControlTypes.wctAddMemberCC Then vListBox.Items.Insert(0, "<None>")
            vHTMLCell.Controls.Add(vListBox)
            vHTMLRow.Cells.Add(vHTMLCell)
            CheckReadOnlyField(vHTMLCell, vParameterName, vRow)
            vHtmlTable.Rows.Add(vHTMLRow)
        End Select
      End If
    Next
  End Sub

  Protected Overridable Function MultiViewGridOnTop() As Boolean
    'For all modules where displayed should be displayed first, this should return True unlike any Search modules where False should be returned
    Return True
  End Function

  Private Sub AddFocusScript(ByVal pAttributes As AttributeCollection)
    pAttributes.Add("onFocus", "document.getElementById('LastControl').value=this.id")
  End Sub

  Private Sub AddUpdatePanel(ByVal pHTMLCell As HtmlTableCell)
    'This will add a non conditional UpdatePanel i.e. Triggers are not required
    Dim vPanel As New UpdatePanel
    While pHTMLCell.Controls.Count > 0
      vPanel.ContentTemplateContainer.Controls.Add(pHTMLCell.Controls(0))
    End While
    pHTMLCell.Controls.Add(vPanel)      'Add it to the cell
  End Sub

  'Private Sub AddPostBackTrigger(ByVal pControl As Control, ByVal SourceName As String, ByVal pTiggerEventType As PostBackTriggerEventTypes)
  '  'Call this routine after loading all controls to mvTableControls to make sure that Source control exists 
  '  If pControl IsNot Nothing AndAlso pControl.Parent IsNot Nothing AndAlso pControl.Parent.Parent IsNot Nothing Then
  '    Dim vPanel As UpdatePanel = TryCast(pControl.Parent.Parent, UpdatePanel)
  '    If vPanel IsNot Nothing Then
  '      Dim vTrigger As New AsyncPostBackTrigger
  '      vTrigger.ControlID = SourceName
  '      vTrigger.EventName = [Enum].GetName(GetType(PostBackTriggerEventTypes), pTiggerEventType)
  '      vPanel.Triggers.Add(vTrigger)
  '    End If
  '  End If
  'End Sub
  Private Sub AddSelectedIndexChangedHandler(ByVal pSourceControlNames As String)
    For Each vControlID As String In pSourceControlNames.Split(","c)
      AddSelectedIndexChangedHandler(TryCast(FindControlByName(mvTableControl, vControlID), DropDownList))
    Next
  End Sub
  Private Sub AddSelectedIndexChangedHandler(ByVal pSourceDDL As DropDownList)
    If pSourceDDL IsNot Nothing Then
      pSourceDDL.AutoPostBack = True
      AddHandler pSourceDDL.SelectedIndexChanged, AddressOf DropDownListSelectedIndexChangedHandler
    End If
  End Sub
  Protected Sub AddTextChangedHandler(ByVal pSourceControlNames As String)
    For Each vControlID As String In pSourceControlNames.Split(","c)
      AddTextChangedHandler(TryCast(FindControlByName(mvTableControl, vControlID), TextBox))
    Next
  End Sub

  Protected Sub AddCheckBoxCheckedHandler(ByVal pSourceControlName As String)
    If TryCast(FindControlByName(mvTableControl, pSourceControlName), CheckBox) IsNot Nothing Then
      AddHandler TryCast(FindControlByName(mvTableControl, pSourceControlName), CheckBox).CheckedChanged, AddressOf CheckBoxChecked
    End If

  End Sub
  Private Sub AddTextChangedHandler(ByVal pSourceTB As TextBox)
    If pSourceTB IsNot Nothing Then
      pSourceTB.AutoPostBack = True
      AddHandler pSourceTB.TextChanged, AddressOf TextBoxTextChangedHandler
    End If
  End Sub

  Protected Sub AddHandlersAndTriggers(ByVal pHTMLTable As HtmlTable)
    AddTextChangedHandler("Forenames,Surname,Initials,Honorifics")
    'If LabelNameFormatCode exists then we have already added the SelectedIndexChanged handler for Title. 
    'In such case, only add the LabelNameFormatCode handler for UpdateContact module as Label Name control is not available in Add Contact or Add Related Contact modules
    Dim vUseLabelNameFormatCode As Boolean = FindControlByName(mvTableControl, "LabelNameFormatCode") IsNot Nothing
    AddSelectedIndexChangedHandler("Sex" & If(vUseLabelNameFormatCode, If(mvControlType = CareNetServices.WebControlTypes.wctUpdateContact, ",LabelNameFormatCode", ""), ",Title"))

    'On Change of Forename, Initials/Preferredforename and HiddenOldForename fields should be updated
    AddAsyncPostBackTrigger("Initials,PreferredForename,HiddenOldForename", "Forenames", PostBackTriggerEventTypes.TextChanged)

    'On Change of Surname, HiddenSurname, HiddenSurname2 and Surname (for capitalisation) fields should be updated
    AddAsyncPostBackTrigger("HiddenSurname,HiddenSurname2,Surname", "Surname", PostBackTriggerEventTypes.TextChanged)

    'On Change of Forename/Surname/Title/Sex, Salutation field should be updated
    Dim vSalutation As String = "Salutation"
    If FindControlByName(mvTableControl, "Salutation") Is Nothing Then vSalutation = "HiddenSalutation"
    AddAsyncPostBackTrigger(vSalutation, "Forenames,Surname", PostBackTriggerEventTypes.TextChanged)
    AddAsyncPostBackTrigger(vSalutation, "Sex,Title", PostBackTriggerEventTypes.SelectedIndexChanged)

    AddAsyncPostBackTrigger("Title", "Title", PostBackTriggerEventTypes.SelectedIndexChanged)
    AddAsyncPostBackTrigger("Forenames", "Forenames", PostBackTriggerEventTypes.TextChanged)
    AddAsyncPostBackTrigger("Sex", "Sex", PostBackTriggerEventTypes.SelectedIndexChanged)
    AddAsyncPostBackTrigger("Relationship", "Relationship", PostBackTriggerEventTypes.SelectedIndexChanged)

    'On Change of Forename/Surname/Initials/Title/Honorifics/PreferredForename, Label Name field should be updated
    AddAsyncPostBackTrigger("LabelName", "Forenames,Surname,Initials,Honorifics,PreferredForename", PostBackTriggerEventTypes.TextChanged)
    AddAsyncPostBackTrigger("LabelName", "Title" & If(vUseLabelNameFormatCode AndAlso mvControlType = CareNetServices.WebControlTypes.wctUpdateContact, ",LabelNameFormatCode", ""), PostBackTriggerEventTypes.SelectedIndexChanged)
  End Sub

  Private Sub AddDDLUpdatePanel(ByVal pHTMLCell As HtmlTableCell, ByVal pSourceDDL As DropDownList)
    If pSourceDDL IsNot Nothing Then
      pSourceDDL.AutoPostBack = True
      AddHandler pSourceDDL.SelectedIndexChanged, AddressOf DropDownListSelectedIndexChangedHandler
      Dim vPanel As New UpdatePanel
      vPanel.UpdateMode = UpdatePanelUpdateMode.Conditional
      While pHTMLCell.Controls.Count > 0
        vPanel.ContentTemplateContainer.Controls.Add(pHTMLCell.Controls(0))
      End While
      Dim vTrigger As New AsyncPostBackTrigger
      vTrigger.ControlID = pSourceDDL.ID
      vTrigger.EventName = "SelectedIndexChanged"
      vPanel.Triggers.Add(vTrigger)
      pHTMLCell.Controls.Add(vPanel)      'Add it to the cell
    End If
  End Sub

  Private Sub AddTBUpdatePanel(ByVal pHTMLCell As HtmlTableCell, ByVal pSourceTB As TextBox)
    If pSourceTB IsNot Nothing Then
      pSourceTB.AutoPostBack = True
      AddHandler pSourceTB.TextChanged, AddressOf TextBoxTextChangedHandler
      Dim vPanel As New UpdatePanel
      vPanel.UpdateMode = UpdatePanelUpdateMode.Conditional
      While pHTMLCell.Controls.Count > 0
        vPanel.ContentTemplateContainer.Controls.Add(pHTMLCell.Controls(0))
      End While
      Dim vTrigger As New AsyncPostBackTrigger
      vTrigger.ControlID = pSourceTB.ID
      vTrigger.EventName = "TextChanged"
      vPanel.Triggers.Add(vTrigger)
      pHTMLCell.Controls.Add(vPanel)      'Add it to the cell
    End If
  End Sub

  Private Sub AddComboBoxCheckedChangedHandler(ByVal pCheckBox As CheckBox)
    If pCheckBox IsNot Nothing Then
      pCheckBox.AutoPostBack = True
      AddHandler pCheckBox.CheckedChanged, AddressOf CheckBoxChecked
    End If
  End Sub

  Protected Sub AddAsyncPostBackTrigger(ByVal pControlNames As String, ByVal pSourceNames As String, ByVal pTiggerEventType As PostBackTriggerEventTypes)
    'Call this routine after loading all controls to mvTableControls to make sure that Source control exists 
    Dim vSourceControls As New CollectionList(Of Control)
    Dim vControl As Control = Nothing
    For Each vName As String In pSourceNames.Split(","c)
      vControl = FindControlByName(mvTableControl, vName)
      If vControl IsNot Nothing Then
        vSourceControls.Add(vName, vControl)
      ElseIf mvSupportsMultiView Then 'If a control is not found in mvTableControl then check if it exists in mvGridControlTable
        vControl = FindControlByName(mvGridControlTable, vName)
        If vControl IsNot Nothing Then vSourceControls.Add(vName, vControl)
      End If
    Next
    vControl = Nothing
    For Each vName As String In pControlNames.Split(","c)
      vControl = FindControlByName(mvTableControl, vName)
      If vControl Is Nothing AndAlso mvSupportsMultiView Then vControl = FindControlByName(mvGridControlTable, vName) 'If a control is not found in mvTableControl then check if it exists in mvGridControlTable
      If vControl IsNot Nothing AndAlso vControl.Parent IsNot Nothing AndAlso vControl.Parent.Parent IsNot Nothing Then
        For Each vSource As Control In vSourceControls
          Dim vUpdatePanel As UpdatePanel = TryCast(vControl.Parent.Parent, UpdatePanel)
          Dim vTrigger As New AsyncPostBackTrigger
          If vUpdatePanel Is Nothing Then
            Dim vHtmlCell As New HtmlTableCell
            vHtmlCell = TryCast(vControl.Parent, HtmlTableCell)
            If vHtmlCell IsNot Nothing Then
              Dim vPanel As New UpdatePanel
              vPanel.UpdateMode = UpdatePanelUpdateMode.Conditional
              While vHtmlCell.Controls.Count > 0
                vPanel.ContentTemplateContainer.Controls.Add(vHtmlCell.Controls(0))
              End While

              vTrigger.ControlID = vSource.ID
              If pTiggerEventType <> PostBackTriggerEventTypes.None Then
                Select Case pTiggerEventType
                  Case PostBackTriggerEventTypes.ButtonClick
                    vTrigger.EventName = "Click"
                  Case Else
                    vTrigger.EventName = [Enum].GetName(GetType(PostBackTriggerEventTypes), pTiggerEventType)
                End Select
              End If
              vPanel.Triggers.Add(vTrigger)
              vHtmlCell.Controls.Add(vPanel)
            End If
          Else
            Dim vIsTriggerAlreadyPresent As Boolean = False
            For vIndex As Integer = 0 To vUpdatePanel.Triggers.Count - 1
              If TryCast(vUpdatePanel.Triggers(vIndex), AsyncPostBackTrigger).ControlID = vSource.ID Then
                vIsTriggerAlreadyPresent = True
              End If
            Next
            If Not vIsTriggerAlreadyPresent Then
              vTrigger.ControlID = vSource.ID
              If pTiggerEventType <> PostBackTriggerEventTypes.None Then
                Select Case pTiggerEventType
                  Case PostBackTriggerEventTypes.ButtonClick
                    vTrigger.EventName = "Click"
                  Case Else
                    vTrigger.EventName = [Enum].GetName(GetType(PostBackTriggerEventTypes), pTiggerEventType)
                End Select
              End If
              vUpdatePanel.Triggers.Add(vTrigger)
            End If
          End If
        Next
      End If
    Next
  End Sub

  Private Sub CheckReadOnlyField(ByVal pControl As Control, ByVal pName As String, ByVal pRow As DataRow)
    Dim vFoundControl As Control = FindControlByName(pControl, pName)
    Dim vViewStateParameterName As String = pName & "~ReadOnly" ' Store the original readonly value of the control
    If vFoundControl IsNot Nothing Then
      Dim vReadonly As Boolean
      If pRow.Table.Columns.Contains("ReadonlyItem") AndAlso pRow("ReadonlyItem").ToString.Length > 0 Then
        vReadonly = pRow("ReadonlyItem").ToString = "Y"
      End If

      If TypeOf (vFoundControl) Is TextBox Then
        DirectCast(vFoundControl, TextBox).ReadOnly = vReadonly
        If vReadonly Then
          Dim vControl As TextBox = DirectCast(vFoundControl, TextBox)
          vControl.Attributes.Add("readonly", "readonly")
          If vControl.ID.EndsWith("Amount") OrElse vControl.ID.EndsWith("Total") OrElse vControl.ID.EndsWith("Balance") Then
            vControl.CssClass = "ReadOnlyNumber"
          Else
            vControl.CssClass = "ReadOnly"
          End If
          ViewState(vViewStateParameterName) = True
        End If
      ElseIf TypeOf (vFoundControl) Is RadioButton Then
        DirectCast(vFoundControl, RadioButton).Enabled = Not vReadonly

        If vReadonly Then
          ViewState(vViewStateParameterName) = True
        End If
      ElseIf TypeOf (vFoundControl) Is CheckBox Then
        DirectCast(vFoundControl, CheckBox).Enabled = Not vReadonly
        If vReadonly Then
          ViewState(vViewStateParameterName) = True
        End If
      ElseIf TypeOf (vFoundControl) Is DropDownList Then
        DirectCast(vFoundControl, DropDownList).Enabled = Not vReadonly
        If vReadonly Then
          DirectCast(vFoundControl, DropDownList).CssClass = "ReadOnly"
          ViewState(vViewStateParameterName) = True
        End If
      ElseIf TypeOf (vFoundControl) Is Button Then
        DirectCast(vFoundControl, Button).Enabled = Not vReadonly
        If vReadonly Then
          ViewState(vViewStateParameterName) = True
        End If
      ElseIf TypeOf (vFoundControl) Is HtmlInputButton Then
        DirectCast(vFoundControl, HtmlInputButton).Disabled = vReadonly
        If vReadonly Then
          ViewState(vViewStateParameterName) = True
        End If
      ElseIf TypeOf (vFoundControl) Is HyperLink Then
        DirectCast(vFoundControl, HyperLink).Enabled = Not vReadonly
        If vReadonly Then
          DirectCast(vFoundControl, HyperLink).CssClass = "ReadOnly"
          ViewState(vViewStateParameterName) = True
        End If
      ElseIf TypeOf (vFoundControl) Is ListBox Then
        DirectCast(vFoundControl, ListBox).Enabled = Not vReadonly
        If vReadonly Then
          DirectCast(vFoundControl, ListBox).CssClass = "ListboxReadonly"
          ViewState(vViewStateParameterName) = True
        End If
      End If
    End If
  End Sub

  Private Sub AddMandatoryStyle(ByVal pControl As Control, ByVal pName As String)
    Dim vFoundControl As WebControl = TryCast(FindControlByName(pControl, pName), WebControl)
    If vFoundControl IsNot Nothing Then
      vFoundControl.CssClass = "DataEntryItemMandatory"
    End If
  End Sub
  Private Sub CheckRequiredValidator(ByVal pRow As DataRow, ByVal pParameterName As String, ByVal pReadOnly As Boolean, ByVal pHTMLCell As HtmlTableCell, ByVal pType As CareNetServices.WebControlTypes)
    Dim vMandatory As Boolean
    Dim vValidationGroup As String = ""
    If pRow.Table.Columns.Contains("MandatoryItem") AndAlso pRow("MandatoryItem").ToString.Length > 0 Then
      vMandatory = pRow("MandatoryItem").ToString = "Y"
    Else
      vMandatory = pRow("NullsInvalid").ToString = "Y"
    End If
    Select Case pParameterName
      Case "OldPassword", "NewPassword"
        vMandatory = True
      Case "Status"
        If Not vMandatory AndAlso pType = CareNetServices.WebControlTypes.wctAddContact Then
          vMandatory = DataHelper.ConfigurationValueOption(DataHelper.ConfigurationValues.cd_status_mandatory)
        End If
      Case "ClaimDay"
        If DataHelper.ControlValue(DataHelper.ControlValues.auto_pay_claim_date_method) = "D" Then
          If DataHelper.ConfigurationOption(DataHelper.ConfigurationOptions.fp_default_blank_claim_day, False) Then
            vMandatory = True
          End If
        End If
      Case "EMailAddress"
        If pType = CareNetServices.WebControlTypes.wctAddOrganisation Then
          vMandatory = False
        End If
      Case "SwitchboardNumber"
        If pType = CareNetServices.WebControlTypes.wctAddOrganisation OrElse pType = CareNetServices.WebControlTypes.wctUpdateOrganisation Then
          vMandatory = False
        End If
      Case "WebAddress"
        If pType = CareNetServices.WebControlTypes.wctAddOrganisation OrElse pType = CareNetServices.WebControlTypes.wctUpdateOrganisation Then
          vMandatory = False
        End If
      Case "FaxNumber"
        If pType = CareNetServices.WebControlTypes.wctUpdateOrganisation Then vMandatory = False
      Case "WebPublish", "CpdPoints2"
        If pType = CareNetServices.WebControlTypes.wctUpdateCpdPoints Then vMandatory = True
      Case "CreditCardType"
        If pType = CareNetServices.WebControlTypes.wctAddMemberCC OrElse pType = CareNetServices.WebControlTypes.wctAddMemberCI OrElse _
          InitialParameters.OptionalValue("PaymentService").Trim.ToUpper = TNSHOSTED Then vMandatory = False
      Case "gatewayCardExpiryDateYear"
        If (pType = CareNetServices.WebControlTypes.wctAddMemberCC OrElse pType = CareNetServices.WebControlTypes.wctAddMemberCI AndAlso _
           DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED) OrElse InitialParameters.OptionalValue("PaymentService").Trim.ToUpper = TNSHOSTED Then vMandatory = True
    End Select
    Select Case pType
      'Case CareNetServices.WebControlTypes.wctCPDCycle
      '  Select Case pParameterName
      '    Case "EndDate", "EndMonth", "EndYear", "StartDate", "StartMonth", "StartYear"
      '      vMandatory = False
      '  End Select
      Case CareNetServices.WebControlTypes.wctSelectEventDelegates
        Select Case pParameterName
          Case "Surname", "EmailAddress"
            vValidationGroup = "SelectEventDelegates"
        End Select
    End Select
    If vMandatory Then AddMandatoryStyle(pHTMLCell, pParameterName)
    If vMandatory AndAlso Not pReadOnly Then AddRequiredValidator(pHTMLCell, pParameterName, vValidationGroup)
  End Sub

  Protected Sub AddBoundColumn(ByVal pGrd As DataGrid, ByVal pDataField As String, ByVal pHeaderText As String)
    Dim vTCol As New TemplateColumn
    vTCol.ItemTemplate = New DisplayTemplate(pDataField)
    vTCol.HeaderText = pHeaderText
    pGrd.Columns.Add(vTCol)
  End Sub

  Protected Sub AddMemoColumn(ByVal pGrd As DataGrid, ByVal pDataField As String, ByVal pHeaderText As String)
    Dim vTCol As New TemplateColumn
    vTCol.HeaderText = pHeaderText
    vTCol.ItemTemplate = New MemoTemplate(pDataField)
    pGrd.Columns.Add(vTCol)
  End Sub

  Protected Sub AddHiddenField(ByVal pName As String)
    If Me.ViewState.Item(pName) Is Nothing Then Me.ViewState.Add(pName, "")
  End Sub

#End Region

#Region "Event Handlers"

  Protected Overridable Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Page.Validate()
    Dim vButton As WebControls.WebControl = TryCast(sender, Button)
    If vButton Is Nothing Then vButton = TryCast(sender, LinkButton)
    Select Case mvControlType
      Case CareNetServices.WebControlTypes.wctUpdateAddress, _
              CareNetServices.WebControlTypes.wctUpdatePhoneNumber, _
              CareNetServices.WebControlTypes.wctUpdateEmailAddress, _
              CareNetServices.WebControlTypes.wctUpdatePosition, _
              CareNetServices.WebControlTypes.wctUpdateCpdPoints, _
              CareNetServices.WebControlTypes.wctCPDCycle, _
              CareNetServices.WebControlTypes.wctUpdateCpdObjectives, _
              CareNetServices.WebControlTypes.wctSearchContact, _
              CareNetServices.WebControlTypes.wctSearchOrganisation
        Dim vCanProcess As Boolean
        If (vButton.ID = "New" OrElse vButton.ID = "Cancel" OrElse vButton.ID = "GridHyperlink") Then
          vCanProcess = True
        Else
          If IsValid() Then
            vCanProcess = True
          End If
        End If

        If vCanProcess Then
          Try
            ProcessButtonClickEvent(vButton.ID)
          Catch vEx As ThreadAbortException
            Throw vEx
          Catch vException As Exception
            ProcessError(vException)
          End Try
        End If
      Case Else
        If IsValid() Then
          Try
            ProcessSubmit()
            GoToSubmitPage()
          Catch vEx As ThreadAbortException
            Throw vEx
          Catch vException As Exception
            ProcessError(vException)
          End Try
        End If
    End Select
  End Sub

  Public Sub SetAuthenticationRequired(ByVal pControlName As String)
    Select Case pControlName.ToUpper
      Case "ADDMEMBERCC", "ADDMEMBERCI"
        If InitialParameters IsNot Nothing AndAlso InitialParameters.ContainsKey("MembershipFor") Then
          If InitialParameters("MembershipFor").ToString.ToUpper = "U" Then mvNeedsAuthentication = False
        Else
          mvNeedsAuthentication = False
        End If
      Case "PROCESSPAYMENT"
        'Process Payment should not redirect to login page if this module is 
        'called from Trader
        If (Request.QueryString("ContactNumber") IsNot Nothing AndAlso IntegerValue(Request.QueryString("ContactNumber")) > 0) OrElse _
          (Request.QueryString("Trader") IsNot Nothing AndAlso Request.QueryString("Trader").ToString = "Y") OrElse _
          (Session("Trader") IsNot Nothing AndAlso Session("Trader").ToString = "Y") Then _
          mvNeedsAuthentication = False
      Case Else
        '
    End Select
  End Sub

  Public Overridable Sub ProcessSubmit()
    'implementation in inherited controls
  End Sub

  Public Overridable Sub ProcessButtonClickEvent(ByVal pValues As Object)
    'implementation in inherited controls
  End Sub

  Public Overridable Sub HandleDataGridEdit(ByVal e As DataGridCommandEventArgs)
    'implementation in inherited controls
  End Sub

  Public Overridable Sub HandleDataListItemDataBound(ByVal e As DataListItemEventArgs)
    'implementation in inherited controls
  End Sub

  Private Sub DataListItemDataBoundHandler(ByVal sender As Object, ByVal e As DataListItemEventArgs)
    'The ItemDataBound event is raised only for items, alternating items, and selected items. 
    'Use this event if the customization depends on the data. eg. Creating hyperlinks to other pages
    Try
      Select Case mvControlType
        Case CareNetServices.WebControlTypes.wctSelectMembershipTypes, _
             CareNetServices.WebControlTypes.wctSelectEvents, _
             CareNetServices.WebControlTypes.wctSelectProducts, _
             CareNetServices.WebControlTypes.wctSelectBookingOptions, _
             CareNetServices.WebControlTypes.wctSearchDirectory, _
             CareNetServices.WebControlTypes.wctViewTransaction, _
             CareNetServices.WebControlTypes.wctDownloadSelection, _
             CareNetServices.WebControlTypes.wctRelatedOrganisations, _
             CareNetServices.WebControlTypes.wctSelectExams
          HandleDataListItemDataBound(e)
      End Select
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Sub DataGridItemClickedHandler(ByVal sender As Object, ByVal e As DataGridCommandEventArgs)
    Try
      'Always raise the event for following control types as the command name is configurable
      Select Case mvControlType
        Case CareNetServices.WebControlTypes.wctUpdateAddress, _
             CareNetServices.WebControlTypes.wctUpdatePhoneNumber, _
             CareNetServices.WebControlTypes.wctUpdateEmailAddress, _
             CareNetServices.WebControlTypes.wctSearchContact, _
             CareNetServices.WebControlTypes.wctUpdateCpdPoints, _
             CareNetServices.WebControlTypes.wctUpdateCpdObjectives, _
             CareNetServices.WebControlTypes.wctCPDCycle, _
             CareNetServices.WebControlTypes.wctUpdatePosition, _
             CareNetServices.WebControlTypes.wctPayerSelection, _
             CareNetServices.WebControlTypes.wctDeDupOrgForRegistration, _
             CareNetServices.WebControlTypes.wctSetUserOrganisation, _
             CareNetServices.WebControlTypes.wctSelecttPayPlanForDD
          HandleDataGridEdit(e)
      End Select

      ' Highlight Selected Row of Datagrid
      Dim vDGR As DataGrid = DirectCast(sender, DataGrid)
      vDGR.SelectedIndex = e.Item.ItemIndex
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Sub ContactNumberChangedHandler(ByVal sender As Object, ByVal e As EventArgs)
    ContactNumberChanged(DirectCast(sender, TextBox), IntegerValue(DirectCast(sender, TextBox).Text))
  End Sub

  Private Sub ExternalReferenceChangedHandler(ByVal sender As Object, ByVal e As EventArgs)
    ExternalReferenceChanged(DirectCast(sender, TextBox), DirectCast(sender, TextBox).Text)
  End Sub

  Private Sub DateButtonClick(ByVal sender As Object, ByVal e As EventArgs)
    Dim vButton As Button = DirectCast(sender, Button)
    Dim vBaseID As String = vButton.ID.Replace("_Button", "_Calendar")
    Dim vCalendar As Calendar = DirectCast(FindControl(vBaseID), Calendar)
    If vCalendar IsNot Nothing Then
      vCalendar.Visible = Not vCalendar.Visible
    End If
  End Sub
  ''' <summary>
  ''' This method is implemented in Membership Type Selection module when user changes the membershisp start date
  ''' </summary>
  ''' <remarks></remarks>
  Public Overridable Sub MembershipStartDateChangeHandler()
    'Handled in Membership Type selection Module 
  End Sub


  Protected Sub RaiseDropDownListSelectedIndexChanged(ByVal pDLL As DropDownList)
    DropDownListSelectedIndexChangedHandler(pDLL, New System.EventArgs)
  End Sub

  Private Sub DropDownListSelectedIndexChangedHandler(ByVal sender As Object, ByVal e As EventArgs)
    Try
      Dim vDDL As DropDownList = DirectCast(sender, DropDownList)
      Select Case vDDL.ID
        Case "Bank"
          Dim vDDLBranchName As DropDownList = DirectCast(FindControlByName(Me, "BranchName"), DropDownList)
          If vDDL.SelectedValue.Length > 0 Then
            Dim vList As New ParameterList(HttpContext.Current)
            vList("Bank") = vDDL.SelectedItem.Value
            DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtBanks, vDDLBranchName, True, vList)
          Else
            vDDLBranchName.DataSource = Nothing
            vDDLBranchName.Items.Clear()
          End If
          vDDLBranchName.SelectedIndex = -1
          SetTextBoxText("SortCode", "")
        Case "BranchName"
          If vDDL.SelectedIndex > 0 AndAlso GetDropDownValue("Bank").Length > 0 Then
            Dim vTable As DataTable = TryCast(vDDL.DataSource, DataTable)
            If vTable Is Nothing Then
              Dim vList As New ParameterList(HttpContext.Current)
              vList("Bank") = GetDropDownValue("Bank")
              vTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtBanks, vList)
            End If
            vTable.DefaultView.RowFilter = "BranchName = '" & vDDL.SelectedValue & "' AND Bank = '" & GetDropDownValue("Bank") & "'"
            Dim vRow As DataRow = vTable.DefaultView.ToTable.Rows(0)
            If vRow("Bank").ToString.Length > 0 Then SetTextBoxText("SortCode", vRow("SortCode").ToString)
          Else
            SetTextBoxText("SortCode", "")
          End If
        Case "PostcoderAddress"
          If vDDL IsNot Nothing AndAlso vDDL.SelectedIndex > 0 Then
            Try
              If Len(DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.afd_everywhere_server)) > 0 And _
                (Len(DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.afd_software)) = 0 Or DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.afd_software) = "POSTCODE" Or DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.afd_software) = "BOTH") Then
                Dim vAFD As New AFDInterface
                If vAFD.DoSearch("", vDDL.SelectedItem.Value).Count > 0 Then
                  'The town field is part of the content of the update panel 
                  Dim vTownTextBox As TextBox = TryCast(FindControlByName(Me, "Town"), TextBox)
                  If vTownTextBox IsNot Nothing Then vTownTextBox.Text = vAFD.Town
                  Dim vAddressTextBox As TextBox = TryCast(FindControlByName(Me, "Address"), TextBox)
                  If vAddressTextBox IsNot Nothing Then
                    vAddressTextBox.Text = vAFD.Address
                  Else
                    vAddressTextBox.Text = String.Empty
                  End If
                  Dim vCountyTextBox As TextBox = TryCast(FindControlByName(Me, "County"), TextBox)
                  If vCountyTextBox IsNot Nothing Then
                    vCountyTextBox.Text = vAFD.County
                  Else
                    vCountyTextBox.Text = String.Empty
                  End If
                  Dim vPostcodeTextBox As TextBox = TryCast(FindControlByName(Me, "Postcode"), TextBox)
                  If vPostcodeTextBox IsNot Nothing Then
                    vPostcodeTextBox.Text = vAFD.Postcode
                  End If
                End If
              Else
                Dim vQAS As New QASInterface
                Dim vAddress As String = vQAS.GetAddress(vDDL.SelectedItem.Value)
                'The town field is part of the content of the update panel 
                Dim vTownTextBox As TextBox = TryCast(FindControlByName(Me, "Town"), TextBox)
                If vTownTextBox IsNot Nothing Then vTownTextBox.Text = vQAS.Town
                Dim vAddressTextBox As TextBox = TryCast(FindControlByName(Me, "Address"), TextBox)
                If vAddressTextBox IsNot Nothing Then
                  vAddressTextBox.Text = vAddress
                  'If TypeOf (vAddressTextBox.Parent) Is UpdatePanel Then DirectCast(vAddressTextBox.Parent, UpdatePanel).Update()
                End If
                Dim vCountyTextBox As TextBox = TryCast(FindControlByName(Me, "County"), TextBox)
                If vCountyTextBox IsNot Nothing Then
                  vCountyTextBox.Text = vQAS.County
                  'If TypeOf (vCountyTextBox.Parent) Is UpdatePanel Then DirectCast(vCountyTextBox.Parent, UpdatePanel).Update()
                End If
                Dim vPostcodeTextBox As TextBox = TryCast(FindControlByName(Me, "Postcode"), TextBox)
                If vPostcodeTextBox IsNot Nothing Then
                  vPostcodeTextBox.Text = vQAS.Postcode
                  'If TypeOf (vPostcodeTextBox.Parent) Is UpdatePanel Then DirectCast(vPostcodeTextBox.Parent, UpdatePanel).Update()
                End If
              End If
              Dim vCountryDDL As DropDownList = TryCast(FindControlByName(Me, "Country"), DropDownList)
              If vCountryDDL IsNot Nothing Then
                SelectListItem(vCountryDDL, "UK")
                'If TypeOf (vCountryDDL.Parent) Is UpdatePanel Then DirectCast(vCountryDDL.Parent, UpdatePanel).Update()
              End If
              SetPafStatus("VB")

            Catch vEx As Exception
              vDDL.Items.Clear()
              Dim vItems(1) As ListItem
              vItems(0) = New ListItem("Cannot select address at this time...", "")
              vItems(1) = New ListItem(vEx.Message, "")
              vDDL.Items.AddRange(vItems)
            End Try
          End If

        Case "Topic"
          Dim vDDLSubTopic As DropDownList = TryCast(FindControlByName(Me, "SubTopic"), DropDownList)
          If vDDLSubTopic IsNot Nothing Then
            Dim vList As New ParameterList(HttpContext.Current)
            If vDDL.SelectedIndex > 0 Then
              vList("Topic") = vDDL.SelectedItem.Value
              DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtSubTopics, vDDLSubTopic, True, vList)
            Else
              vDDLSubTopic.Items.Clear()
            End If
          End If
        Case "EventNumber"
          Dim vDesc As TextBox = TryCast(FindControlByName(Me, "FundraisingDescription"), TextBox)
          Dim vDate As TextBox = TryCast(FindControlByName(Me, "TargetDate"), TextBox)
          If vDesc IsNot Nothing Then
            If vDDL.SelectedIndex > 0 Then
              If vDate IsNot Nothing Then
                Dim vTable As DataTable = DirectCast(vDDL.DataSource, DataTable)
                For Each vRow As DataRow In vTable.Rows
                  If vRow("EventNumber").ToString = vDDL.SelectedValue Then
                    vDate.Text = vRow("StartDate").ToString
                  End If
                Next
                SetControlEnabled("TargetDate", False)
              End If
              vDesc.Text = vDDL.SelectedItem.Text
              SetControlEnabled("FundraisingDescription", False)
            Else
              vDesc.Text = ""
              SetControlEnabled("FundraisingDescription", True)
              If vDate IsNot Nothing Then
                vDate.Text = ""
                SetControlEnabled("TargetDate", False)
              End If
            End If
          End If
        Case "Title"
          UpdateSalutation(vDDL)
          Dim vDropDown As DropDownList = TryCast(FindControlByName(Me, "LabelNameFormatCode"), DropDownList)
          If vDropDown Is Nothing Then
            UpdateLabelName(vDDL)
          Else
            UpdateLabelNameFormatCodes(vDropDown)
            UpdateLabelName(vDDL)
          End If
        Case "LabelNameFormatCode"
          UpdateLabelName(vDDL)
        Case "Sex"
          UpdateSalutation(vDDL)
        Case "Country"
          If vDDL IsNot Nothing AndAlso vDDL.SelectedIndex > 0 Then
            If DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.qas_pro_web_url).Length > 0 Then SetPafStatus("PA")
          End If
        Case "CpdCycleType"
          Dim vCycleTypeList As New ParameterList(HttpContext.Current)
          vCycleTypeList("ForPortal") = "Y"
          Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCPDCycleTypes, vCycleTypeList)
          Dim vRowColl() As DataRow = vTable.Select("CpdCycleType = '" & vDDL.SelectedValue & "'")
          If vRowColl.Length > 0 Then
            Dim vRow As DataRow = vRowColl(0)
            'Dim vUseFlexibleCycles As Boolean = False
            'Dim vGotDateFields As Boolean = HasCPDCycleStartAndEndDates()
            'If vRow.Item("StartMonth").ToString.Length = 0 Then vUseFlexibleCycles = True
            'If vUseFlexibleCycles = False Then
            SetTextBoxText("StartMonth", MonthName(IntegerValue(vRow.Item("StartMonth").ToString)))
            SetTextBoxText("EndMonth", MonthName(IntegerValue(vRow.Item("EndMonth").ToString)))
            '  If vGotDateFields Then
            '    SetTextBoxText("StartDate", "")
            '    SetTextBoxText("EndDate", "")
            '    SetControlEnabled("StartDate", False)
            '    SetControlEnabled("EndDate", False)
            '    CType(FindControl("StartDate"), TextBox).Enabled = False
            '    CType(FindControl("EndDate"), TextBox).Enabled = False
            '  End If
            'Else
            '  SetTextBoxText("StartMonth", "")
            '  SetTextBoxText("EndMonth", "")
            '  SetTextBoxText("EndYear", "")
            '  'SetTextBoxText("StartYear", "")
            '  CType(FindControl("StartYear"), TextBox).Text = String.Empty
            'End If
            SetControlEnabled("StartMonth", False)
            SetControlEnabled("EndMonth", False)
            SetControlEnabled("EndYear", False)
            'SetControlEnabled("StartYear", (Not vUseFlexibleCycles))
            If Request.QueryString("CN") IsNot Nothing AndAlso Request.QueryString("CN").Length > 0 Then
              Dim vTemp As Date = Date.Parse(vRow.Item("StartDate").ToString)
              SetTextBoxText("StartYear", vTemp.Year.ToString)
              vTemp = Date.Parse(vRow.Item("EndDate").ToString)
              SetTextBoxText("EndYear", vTemp.Year.ToString)
              DirectCast(Me.FindControl("CpdCycleType"), DropDownList).Enabled = False
              'Dim vStartDate As Date = Date.Parse(vRow.Item("StartDate").ToString)
              'Dim vEndDate As Date = Date.Parse(vRow.Item("EndDate").ToString)
              'If vUseFlexibleCycles Then
              '  If vGotDateFields Then
              '    SetTextBoxText("StartDate", vStartDate.ToString(CAREDateFormat))
              '    SetTextBoxText("EndDate", vEndDate.ToString(CAREDateFormat))
              '  End If
              'Else
              '  SetTextBoxText("StartYear", vStartDate.Year.ToString)
              '  SetTextBoxText("EndYear", vEndDate.Year.ToString)
              'End If
              'DirectCast(Me.FindControl("CpdCycleType"), DropDownList).Enabled = False
            Else
              If IntegerValue(vRow.Item("DefaultDuration").ToString) > 0 AndAlso IntegerValue(GetTextBoxText("StartYear")) > 0 Then
                ' We have a default duration - need to set controls accordingly
                Dim vDate As New Date(IntegerValue(GetTextBoxText("StartYear")), IntegerValue(vRow("StartMonth").ToString), 1)
                vDate = vDate.AddYears(IntegerValue(vRow.Item("DefaultDuration").ToString)).AddDays(-1)
                SetTextBoxText("EndYear", vDate.Year.ToString)
                SetControlEnabled("EndYear", False)
              End If
              'If IntegerValue(vRow.Item("DefaultDuration").ToString) > 0 Then
              '  ' We have a default duration - need to set controls accordingly
              '  Dim vStartDate As Nullable(Of Date)
              '  If vUseFlexibleCycles = False AndAlso IntegerValue(GetTextBoxText("StartYear")) > 0 Then
              '    vStartDate = New Date(IntegerValue(GetTextBoxText("StartYear")), IntegerValue(vRow("StartMonth").ToString), 1)
              '  ElseIf vUseFlexibleCycles = True AndAlso vGotDateFields = True AndAlso GetTextBoxText("StartDate").Length > 0 Then
              '    vStartDate = Date.Parse(GetTextBoxText("StartDate"))
              '  End If
              '  If vStartDate.HasValue Then
              '    Dim vEndDate As Date = vStartDate.Value.AddYears(IntegerValue(vRow.Item("DefaultDuration").ToString)).AddDays(-1)
              '    If vUseFlexibleCycles Then
              '      'SetTextBoxText("EndDate", vEndDate.ToString(CAREDateFormat))
              '      CType(FindControl("EndDate"), TextBox).Text = vEndDate.ToString(CAREDateFormat)
              '    Else
              '      SetTextBoxText("EndYear", vEndDate.Year.ToString)
              '      SetControlEnabled("EndYear", False)
              '      CType(FindControl("EndYear"), TextBox).Enabled = False
              '    End If
              '  End If
              '  'Dim vDate As New Date(IntegerValue(GetTextBoxText("StartYear")), IntegerValue(vRow("StartMonth").ToString), 1)
              '  'vDate = vDate.AddYears(IntegerValue(vRow.Item("DefaultDuration").ToString)).AddDays(-1)
              '  'SetTextBoxText("EndYear", vDate.Year.ToString)
              'End If
            End If
            Dim vCycleStatusList As DropDownList = DirectCast(Me.FindControl("CpdCycleStatus"), DropDownList)
            If Not vCycleStatusList Is Nothing Then
              Dim vList As ParameterList = New ParameterList(HttpContext.Current)
              vList("CPDCycleType") = vDDL.SelectedValue
              vList("ForPortal") = "Y"
              vDDL.DataTextField = "CPDCycleStatusDesc"
              vDDL.DataValueField = "CPDCycleStatus"
              vCycleStatusList.Items.Clear()
              DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtCPDCycleStatuses, vCycleStatusList, True, vList)
            End If
          End If
        Case "CpdCategoryType"
          Dim vCpdCategory As New ParameterList(HttpContext.Current)
          Dim vRestriction As String
          Dim vCpdCategoryDropDown As DropDownList
          vCpdCategory("FromCPDPoints") = "Y"
          vCpdCategory("CpdCategoryType") = vDDL.SelectedValue
          vCpdCategoryDropDown = DirectCast(Me.FindControl("CpdCategory"), DropDownList)
          vCpdCategoryDropDown.DataTextField = "CpdCategoryDesc"
          vCpdCategoryDropDown.DataValueField = "CpdCategory"
          If mvControlType = CareNetServices.WebControlTypes.wctUpdateCpdPoints Then
            vRestriction = "Approved ='Y'"
          Else
            vRestriction = ""
          End If
          vCpdCategoryDropDown.Items.Clear()
          DataHelper.FillComboWithRestriction(CareNetServices.XMLLookupDataTypes.xldtCPDCategories, vCpdCategoryDropDown, False, vCpdCategory, vRestriction)
          If mvControlType = CareNetServices.WebControlTypes.wctUpdateCpdPoints Then
            Dim vTable As DataTable
            vTable = CType(vCpdCategoryDropDown.DataSource, DataTable)
            Dim vRowColl() As DataRow = vTable.Select("CpdCategory = '" & vCpdCategoryDropDown.SelectedValue & "'")
            Dim vRow As DataRow = Nothing
            If vRowColl.Length > 0 Then vRow = vRowColl(0)
            If vRow IsNot Nothing Then
              SetLabelText("PageError", "")
              If vRow.Item("CPDPoints").ToString <> "" Then
                SetTextBoxText("CpdPoints", vRow.Item("CPDPoints").ToString)
                If vRow.Item("PointsOverride").ToString = "N" Then
                  SetControlEnabled("CpdPoints", False)
                  SetControlEnabled("CpdPoints2", False)
                Else
                  SetControlEnabled("CpdPoints", True)
                  SetControlEnabled("CpdPoints2", True)
                End If
              Else
                SetControlEnabled("CpdPoints", True)
                SetControlEnabled("CpdPoints2", True)
              End If
              If vRow.Item("DateMandatory").ToString = "Y" Then
                ViewState("DateMandatory") = "Y"
              Else
                ViewState("DateMandatory") = "N"
              End If
            End If
          End If
        Case "CpdCategory"
          If mvControlType = CareNetServices.WebControlTypes.wctUpdateCpdPoints Then
            Dim vTable As DataTable
            vTable = CType(vDDL.DataSource, DataTable)
            Dim vRowColl() As DataRow = vTable.Select("CpdCategory = '" & vDDL.SelectedValue & "'")
            Dim vRow As DataRow = vRowColl(0)
            If vRow IsNot Nothing Then
              SetLabelText("PageError", "")
              If vRow.Item("CPDPoints").ToString <> "" Then
                SetTextBoxText("CpdPoints", vRow.Item("CPDPoints").ToString)
                If vRow.Item("PointsOverride").ToString = "N" Then
                  SetControlEnabled("CpdPoints", False)
                  SetControlEnabled("CpdPoints2", False)
                Else
                  SetControlEnabled("CpdPoints", True)
                  SetControlEnabled("CpdPoints2", True)
                End If
              Else
                SetControlEnabled("CpdPoints", True)
                SetControlEnabled("CpdPoints2", True)
              End If
              If vRow.Item("DateMandatory").ToString = "Y" Then
                ViewState("DateMandatory") = "Y"
              Else
                ViewState("DateMandatory") = "N"
              End If
            End If
          End If
        Case "Relationship"
          If Me.FindControl("RelationshipStatus") IsNot Nothing Then
            Dim vRelationship As New ParameterList(HttpContext.Current)
            Dim vRelationshipStatusDropDown As DropDownList
            vRelationship("Relationship") = vDDL.SelectedValue
            vRelationshipStatusDropDown = DirectCast(Me.FindControl("RelationshipStatus"), DropDownList)
            vRelationshipStatusDropDown.DataTextField = "RelationshipStatusDesc"
            vRelationshipStatusDropDown.DataValueField = "RelationshipStatus"
            vRelationshipStatusDropDown.Items.Clear()
            DataHelper.FillComboWithRestriction(CareNetServices.XMLLookupDataTypes.xldtRelationshipStatuses, vRelationshipStatusDropDown, True, vRelationship, "Relationship Is Null OR Relationship = '" & vDDL.SelectedValue & "'")
          End If
        Case "StartDateList"
          MembershipStartDateChangeHandler()
        Case "DataSource"
          Dim vExtContactTextBox As TextBox = TryCast(FindControlByName(Me, "ExternalReference"), TextBox)
          If vExtContactTextBox IsNot Nothing AndAlso vExtContactTextBox.Text.Length > 0 Then ExternalReferenceChanged(vExtContactTextBox, vExtContactTextBox.Text)
        Case "Device"
          If Me.FindControl("CommunicationUsage") IsNot Nothing Then
            Dim vCUDDL As DropDownList = DirectCast(Me.FindControl("CommunicationUsage"), DropDownList)
            Dim vList As New ParameterList(HttpContext.Current)
            If vDDL.SelectedValue.Length > 0 Then vList("Device") = vDDL.SelectedValue
            DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtCommunicationUsages, vCUDDL, True, vList)
          End If
      End Select
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  Private Sub TextBoxTextChangedHandler(ByVal sender As Object, ByVal e As EventArgs)
    Dim vTB As TextBox = DirectCast(sender, TextBox)
    Try
      Select Case vTB.ID
        Case "AdultQuantity", "ChildQuantity"
          DirectCast(Me, BookEvent).ValidateQuantity(vTB)
          SetTotalAmount(IntegerValue(GetTextBoxText("Quantity")))
        Case "Quantity"
          If mvControlType = CareNetServices.WebControlTypes.wctBookEventCC OrElse _
             mvControlType = CareNetServices.WebControlTypes.wctProductPurchase OrElse _
             mvControlType = CareNetServices.WebControlTypes.wctBookEvent OrElse _
             mvControlType = CareNetServices.WebControlTypes.wctProductPurchaseCC Then
            If vTB.Text.StartsWith("-") OrElse vTB.Text.StartsWith(".") Then vTB.Text = vTB.Text.Remove(0, 1)
            If vTB.Text.EndsWith("-") OrElse vTB.Text.EndsWith(".") Then vTB.Text = vTB.Text.Remove(vTB.Text.Length - 1, 1)
            If IntegerValue(vTB.Text.ToString) > 0 Then
              If mvControlType = CareNetServices.WebControlTypes.wctProductPurchase OrElse mvControlType = CareNetServices.WebControlTypes.wctProductPurchaseCC Then
                SetTotalAmount(IntegerValue(vTB.Text.ToString))
              ElseIf mvControlType = CareNetServices.WebControlTypes.wctBookEvent Then
                If Not DirectCast(Me, BookEvent).ValidateQuantity(vTB) Then Exit Sub
                SetTotalAmount(IntegerValue(vTB.Text.ToString))
              Else
                SetTextBoxText("TotalAmount", FixTwoPlaces((IntegerValue(vTB.Text.ToString) * DoubleValue(GetTextBoxText("Amount"))).ToString).ToString("0.00"))
              End If
            End If
          End If
        Case "PostcoderPostcode"
          Dim vDDL As DropDownList = TryCast(FindControlByName(Me, "PostcoderAddress"), DropDownList)
          If vDDL IsNot Nothing Then
            vDDL.Items.Clear()
            Try
              If vTB.Text.Length = 0 Then
                vDDL.Items.Add("Enter postcode then choose address from list")
              Else
                If Len(DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.afd_everywhere_server)) > 0 And _
                (Len(DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.afd_software)) = 0 Or DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.afd_software) = "POSTCODE" Or DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.afd_software) = "BOTH") Then
                  Dim vAFD As New AFDInterface
                  Dim vAddresses As ListItemCollection = vAFD.DoSearch(vTB.Text)
                  Dim vItems(vAddresses.Count) As ListItem
                  vItems(0) = New ListItem("Select an address...", "")
                  vAddresses.CopyTo(vItems, 1)
                  vDDL.Items.AddRange(vItems)
                Else
                  Dim vQAS As New QASInterface
                  Dim vAddresses As ListItemCollection = vQAS.DoSearch(vTB.Text)
                  Dim vItems(vAddresses.Count) As ListItem
                  vItems(0) = New ListItem("Select an address...", "")
                  vAddresses.CopyTo(vItems, 1)
                  vDDL.Items.AddRange(vItems)
                End If
              End If
            Catch vEx As Exception
              Dim vItems(1) As ListItem
              vItems(0) = New ListItem("Cannot find addresses at this time...", "")
              vItems(1) = New ListItem(vEx.Message, "")
              vDDL.Items.AddRange(vItems)
              SetPafStatus("NV")
            End Try
          End If
        Case "SortCode"
          If mvControlType = CareNetServices.WebControlTypes.wctAddBankAccount Then
            Dim vClearControls As Boolean = True
            If vTB.Text.Length > 0 Then
              Dim vList As New ParameterList(HttpContext.Current)
              vList.Add("SortCode", vTB.Text.Replace("-", ""))
              Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtBanks, vList)
              If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 AndAlso vTable.Rows(0)("Bank").ToString.Length > 0 AndAlso vTable.Rows(0)("BranchName").ToString.Length > 0 Then
                SetDropDownText("Bank", vTable.Rows(0)("Bank").ToString, True)
                SetDropDownText("BranchName", vTable.Rows(0)("BranchName").ToString, True)
                vClearControls = False
              End If
            End If
            If vClearControls Then
              SetDropDownText("Bank", "")
              SetDropDownText("BranchName", "")
              Dim vDDLBranchName As DropDownList = DirectCast(FindControlByName(Me, "BranchName"), DropDownList)
              vDDLBranchName.DataSource = Nothing
              vDDLBranchName.Items.Clear()
            End If
          End If
        Case "Forenames"
          If mvControlType = CareNetServices.WebControlTypes.wctUpdateContact Then UpdateInitials(True, vTB.ID, GetTextBoxText(vTB.ID))
          UpdateSalutation(vTB)
          UpdateLabelName(vTB)
          UpdatePreferredForename("PreferredForename", GetTextBoxText("Forenames"))
        Case "Surname"
          If Not CapitalisationChangedOnly("HiddenSurname2", "Surname") Then
            'Note: HiddenOldSurname should remain unchanged as this should be the original value (it will be set once salutation updated)
            Capitalise(vTB, CapitaliseOptions.caoSurname)
            SetHiddenText("HiddenSurname2", vTB.Text)
          End If
          If DataHelper.ConfigurationOption(DataHelper.ConfigurationOptions.use_ajax_for_contact_names, False) Then
            UpdateSalutation(vTB)
            UpdateLabelName(vTB)
          End If
        Case "Initials", "Honorifics"
          UpdateLabelName(vTB)
        Case "Postcode"
          If mvControlType = CareNetServices.WebControlTypes.wctUpdateAddress Then
            Dim vDefaultBranch As String = GetDefaultBranch(GetTextBoxText("Postcode"))
            If Not String.IsNullOrEmpty(vDefaultBranch) Then
              Dim vDDLBranch As DropDownList = CType(FindControl("Branch"), DropDownList)
              If Not vDDLBranch Is Nothing Then
                If Not vDDLBranch.Items.FindByValue(vDefaultBranch) Is Nothing Then
                  vDDLBranch.SelectedIndex = -1
                  vDDLBranch.Items.FindByValue(vDefaultBranch).Selected = True
                End If
              End If
            End If
          End If
          SetPafStatus("")
        Case "Address", "Town", "County", "Country"
          If DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.qas_pro_web_url).Length > 0 Then SetPafStatus("PA")
        Case "StartYear"    ', "StartDate"
          If mvControlType = CareNetServices.WebControlTypes.wctCPDCycle Then
            'Dim vUseStartYear As Boolean = False

            'If vTB.ID.Equals("StartYear", StringComparison.InvariantCultureIgnoreCase) Then
            '  vUseStartYear = True
            Dim vControl As RangeValidator = CType(vTB.FindControl("rnvStartYear"), RangeValidator)
            vControl.Validate()
            If Not vControl.IsValid Then Exit Select
            'End If

            'SetCPDEndYearOrDate(vUseStartYear)
            Dim vDDL As DropDownList = TryCast(FindControlByName(Me, "CpdCycleType"), DropDownList)
            Dim vCycleTypeList As New ParameterList(HttpContext.Current)
            vCycleTypeList("ForPortal") = "Y"
            Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCPDCycleTypes, vCycleTypeList)
            Dim vRowColl() As DataRow = vTable.Select("CpdCycleType = '" & vDDL.Items(vDDL.SelectedIndex).Value & "'")
            If vDDL.SelectedValue.ToString.Length > 0 Then
              If vRowColl.Length > 0 Then
                Dim vRow As DataRow = vRowColl(0)
                If vRow IsNot Nothing AndAlso IntegerValue(vRow.Item("DefaultDuration").ToString) > 0 Then
                  ' We have a default duration - need to set controls accordingly
                  Dim vDate As New Date(IntegerValue(GetTextBoxText("StartYear")), IntegerValue(vRow("StartMonth").ToString), 1)
                  vDate = vDate.AddYears(IntegerValue(vRow.Item("DefaultDuration").ToString)).AddDays(-1)
                  TryCast(FindControlByName(Me, "EndYear"), TextBox).Text = vDate.Year.ToString
                  SetControlEnabled("EndYear", False)
                Else
                  ' No default duration - let the user decide
                  SetControlEnabled("EndYear", True)
                End If
              End If
            Else
              SetControlEnabled("EndYear", False)
            End If
          End If
      End Select

    Catch vEx As Exception
      If vEx.Message.StartsWith("Conversion from string") AndAlso vEx.Message.EndsWith(" to type 'Integer' is not valid.") Then
        ' invalid Quantity string type inputted
      Else
        ProcessError(vEx)
      End If
    End Try
  End Sub

  Protected Sub SetTotalAmount(ByVal pQuantity As Integer)
    Dim vPercentage As Double = DoubleValue(GetHiddenText("HiddenPercentage"))
    Dim vAmount As Double = DoubleValue(GetHiddenText("HiddenCurrentPrice"))
    Dim vVatRate As Double
    vAmount = vAmount * pQuantity
    If GetHiddenText("HiddenVatExclusive") = "Y" Then
      vVatRate = FixTwoPlaces((vAmount * (vPercentage / 100)).ToString)
      vAmount = vAmount + vVatRate
    End If
    SetTextBoxText("TotalAmount", vAmount.ToString("0.00"))
  End Sub

  Private Sub Capitalise(ByVal pControl As Control, ByVal pCapitaliseOptions As CapitaliseOptions)
    Try
      SetTextBoxText(pControl.ID, CapitaliseWords(pControl.ID, pCapitaliseOptions))
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  Private Enum CapitaliseOptions
    caoNone
    caoItalianAnd = 1
    caoSurname = 2
  End Enum


  Private Function CapitaliseWords(ByVal pParameterName As String, ByVal pCapitaliseOptions As CapitaliseOptions) As String
    Dim vPos As Integer
    Dim vResult As New StringBuilder
    Dim vResultString As String
    Dim vLastChar As Char
    Dim vThisChar As Char
    Dim vIndex As Integer
    Dim vLen As Integer
    Dim vString As String

    vString = GetTextBoxText(pParameterName)
    vString = vString.Trim().ToLower
    If vString.Length = 0 Then Return vString

    vLen = vString.Length
    vLastChar = " "c
    For vIndex = 0 To vString.Length - 1
      vThisChar = vString.Chars(vIndex)
      Select Case vLastChar
        Case "'"c                                  'An apostrophe
          vThisChar = Char.ToUpper(vThisChar)
          If vThisChar = "S"c Then                 'Could be a plural like Fred's
            If vIndex = vString.Length - 1 OrElse vString.Chars(vIndex + 1) = " "c Then
              vThisChar = "s"c                     'If so then make it lower case
            End If
          End If
          vResult.Append(vThisChar)
        Case " "c, "-"c, "."c, ","c, "("c, "/"c, "\"c, ChrW(10)
          vResult.Append(Char.ToUpper(vThisChar))
        Case Else
          vResult.Append(vThisChar)
      End Select
      vLastChar = vThisChar
    Next
    vResultString = vResult.ToString

    'Change the letter following Mc or Mac to upper case e.g. McDonald
    'If the Mc or Mac is not at the start of the string then it must have a preceeding space e.g. Raith McDonald
    If (pCapitaliseOptions And CapitaliseOptions.caoSurname) = CapitaliseOptions.caoSurname Then
      vPos = vResultString.IndexOf("Mc")
      If (vPos + 2 < vResultString.Length) AndAlso ((vPos = 0) Or ((vPos > 0) AndAlso (vResultString.Chars(vPos - 1) = " "c))) Then
        vResult.Chars(vPos + 2) = Char.ToUpper(vResult.Chars(vPos + 2))
      End If

      vPos = vResultString.IndexOf("Mac")
      If (vPos + 3 < vResultString.Length) AndAlso ((vPos = 0) Or ((vPos > 0) AndAlso (vResultString.Chars(vPos - 1) = " "c))) Then
        vResult.Chars(vPos + 3) = Char.ToUpper(vResult.Chars(vPos + 3))
      End If
    End If

    'If any of the following words are found (space on each side) then lowercase them
    Dim vLCaseWords() As String = {" And ", " Und ", " Et ", " Of ", " On ", " To "}
    For Each vWord As String In vLCaseWords
      vPos = vResultString.IndexOf(vWord)
      If vPos > 0 Then vResult.Chars(vPos + 1) = Char.ToLower(vResult.Chars(vPos + 1))
    Next

    'Change Of at the end to of
    If vResultString.EndsWith(" Of") Then vResult.Chars(vResultString.Length - 2) = "o"c

    'If dealing with Italian And then set to lower case e.g. Donna e Gabacci
    If (pCapitaliseOptions And CapitaliseOptions.caoItalianAnd) = CapitaliseOptions.caoItalianAnd Then
      vPos = vResultString.IndexOf(" E ")
      If vPos > 0 Then vResult.Chars(vPos + 1) = Char.ToLower(vResult.Chars(vPos + 1))
    End If

    'If we find " The " and what preceeds it is not a number then set it to lower case
    vPos = vResultString.IndexOf(" The ")
    If vPos > 0 AndAlso Char.IsDigit(vResult.Chars(vPos - 1)) = False Then vResult.Chars(vPos + 1) = "t"c

    'If any of the folowing words are found either at the end of a line or stand-alone then upper case them
    Dim vUCaseWords() As String = {" Plc", " Ag"}
    For Each vWord As String In vUCaseWords
      vPos = vResultString.IndexOf(vWord)
      If vPos > 0 Then
        Dim vChar As Char = " "c
        If vPos + vWord.Length < vResultString.Length Then vChar = vResult.Chars(vPos + vWord.Length)
        If vChar = " "c Or vChar = ChrW(10) Then
          For vIndex = 2 To vWord.Length - 1
            vResult.Chars(vPos + vIndex) = Char.ToUpper(vResult.Chars(vPos + vIndex))
          Next
        End If
      End If
    Next
    Return vResult.ToString
  End Function

  Private Function SpacePadInitials(ByVal pString As String) As String
    Dim vInitials As String = ""
    Dim vWords() As String
    Dim vItems() As String
    Dim vItems2() As String
    Dim vIndex As Integer
    Dim vIndex2 As Integer
    Dim vIndex3 As Integer
    Dim vInitial As String
    Dim vAddStyle As Boolean
    Dim vStyle As String

    vStyle = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.initials_format)
    If vStyle.Length > 2 Then
      vStyle = vStyle.Substring(1, vStyle.Length - 2)
    ElseIf vStyle.Length = 2 Then
      vStyle = ""
    Else
      vStyle = " "
    End If
    If pString.Length > 0 Then
      vWords = pString.Split(" "c)
      vAddStyle = False
      For vIndex = 0 To vWords.Length - 1
        vItems = vWords(vIndex).Split("-"c)
        For vIndex2 = 0 To vItems.Length - 1
          vItems2 = vItems(vIndex2).Split("."c)
          For vIndex3 = 0 To vItems2.Length - 1
            If vItems2(vIndex3).Length > 0 Then
              Select Case vItems2(vIndex3).ToLower
                Case "+", "et", "und", "and"
                  vInitial = " + "
                  vAddStyle = False
                Case "&"
                  vInitial = "&"
                  If vAddStyle Then
                    vInitials &= vStyle
                    If vStyle.IndexOf(" ") >= 0 Then vInitial = "& "
                    vAddStyle = False
                  End If
                Case Else
                  vInitial = vItems2(vIndex3).Substring(0, 1).ToUpper
                  If vAddStyle Then
                    vInitials = vInitials & vStyle
                  Else
                    vAddStyle = True
                  End If
              End Select
              vInitials = vInitials & vInitial
            End If
          Next
        Next
      Next
      If vInitials.Length > 0 Then
        If Not vInitials.EndsWith(vStyle) Then vInitials &= vStyle
      End If
      ' BR11961 - Need to ensure that the initials are truncated to the permitted 7 characters
      If vInitials.Length > 7 Then vInitials = vInitials.Substring(0, 7)
    End If
    Return vInitials.Trim
  End Function

  Private Function GetLabelName(ByVal pJunior As Boolean, ByVal pParameterName As String) As String
    Dim vFormat As String = ""
    Dim vLabelName As New StringBuilder
    Dim vItem As String
    Dim vJoint As Boolean
    Dim vTitle1 As String = ""
    Dim vTitle2 As String = ""
    Dim vInitials1 As String = ""
    Dim vInitials2 As String = ""
    Dim vSurname1 As String = ""
    Dim vSurname2 As String = ""
    Dim vTitle As String
    Dim vInitials As String
    Dim vSurname As String
    Dim vHonorifics As String
    Dim vForenames As String
    Dim vPreferred As String
    Dim vItems() As String
    Dim vAddSpace As Boolean
    Dim vPrefixHonorifics As String = ""
    Dim vSurnamePrefix As String = ""
    Dim vHonorifics1 As String = ""
    Dim vHonorifics2 As String = ""
    Dim vForenames1 As String = ""
    Dim vForenames2 As String = ""
    Dim vLabelNameFormatCode As String = ""
    Dim vLabelNameFormat As DataTable = Nothing

    vTitle = GetDropDownValue(ContactFieldName("Title", pParameterName))
    vInitials = GetTextBoxText(ContactFieldName("Initials", pParameterName))
    vSurname = GetTextBoxText(ContactFieldName("Surname", pParameterName))
    vHonorifics = GetTextBoxText(ContactFieldName("Honorifics", pParameterName))
    vForenames = GetTextBoxText(ContactFieldName("Forenames", pParameterName))
    vPreferred = GetTextBoxText(ContactFieldName("PreferredForename", pParameterName))
    vLabelNameFormatCode = GetDropDownValue(ContactFieldName("LabelNameFormatCode", pParameterName))

    Dim vTitlesList As New ParameterList(HttpContext.Current)
    Dim vTitles As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtTitles, vTitlesList)
    Dim vTitleLabel As String = ""

    If pJunior Then
      vFormat = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.jnr_label_name_format)
      If vFormat.Length = 0 Then vFormat = "forenames surname honorifics"
    Else
      If vTitle.IndexOf(" & ") >= 0 Then
        vJoint = True
      ElseIf vInitials.IndexOf(" & ") >= 0 Then
        vJoint = True
      ElseIf vSurname.IndexOf(" & ") >= 0 Then
        vJoint = True
      End If
      If Not vJoint Then
        If vTitle.Length > 0 AndAlso vLabelNameFormatCode.Length = 0 Then
          Dim vDDL As DropDownList = TryCast(FindControlByName(mvTableControl, "Title"), DropDownList)
          If vDDL IsNot Nothing Then
            Dim vTable As DataTable = TryCast(vDDL.DataSource, DataTable)
            If vTable IsNot Nothing And vTable.Columns.Contains("LabelName") Then
              vFormat = vTable.Select("Title = '" & vTitle & "'")(0)("LabelName").ToString
            End If
          End If
        ElseIf vTitle.Length > 0 AndAlso vLabelNameFormatCode.Length > 0 Then
          'get format from new table
          Dim vList As New ParameterList(HttpContext.Current)
          vList("LabelNameFormatCode") = vLabelNameFormatCode
          vLabelNameFormat = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtLabelNameFormatCodes, vList)
          If vLabelNameFormat IsNot Nothing AndAlso vLabelNameFormat.Columns.Contains("LabelNameFormat") Then
            If vLabelNameFormat.Select("Title = '" & vTitle & "'").Length > 0 Then vFormat = vLabelNameFormat.Select("Title = '" & vTitle & "'")(0)("LabelNameFormat").ToString
          End If
        End If
        If vFormat.Length = 0 Then vFormat = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.label_name_format)
        If vFormat.Length = 0 Then vFormat = "title initials surname honorifics"
      End If
    End If

    vItems = vFormat.Split(" "c)
    Dim vItemsList As New ArrayListEx(vFormat.Replace(" ", ","))

    If vJoint Then
      vItem = ContactInfo.JointItem(vTitle, vTitle, vTitle1, vTitle2, False)
      vItem = ContactInfo.JointItem(vTitle, vInitials, vInitials1, vInitials2, False)
      vItem = ContactInfo.JointItem(vTitle, vSurname, vSurname1, vSurname2, False)
      Dim vList As New ArrayListEx
      vList.Add(vTitle1)
      vList.Add(vInitials1)
      If vSurname2.Length > 0 Then vList.Add(vSurname1)
      vList.Add("&")
      vList.Add(vTitle2)
      vList.Add(vInitials2)
      If vSurname2.Length > 0 Then
        vList.Add(vSurname2)
      Else
        vList.Add(vSurname1)
      End If
      vLabelName.Append(vList.SSNonBlankList)
    Else
      For Each vItem In vItems
        Dim vContinue As Boolean = True

        Select Case vItem
          Case "title1", "title2", "initials1", "initials2", "forenames1", "forenames2", "surname1", "surname2", "honorifics1", "honorifics2"
            If vLabelNameFormatCode.Length = 0 Then
              Dim vDDL As DropDownList = TryCast(FindControlByName(mvTableControl, "Title"), DropDownList)
              If vDDL IsNot Nothing Then
                Dim vTable As DataTable = TryCast(vDDL.DataSource, DataTable)
                If vTable IsNot Nothing AndAlso vTable.Columns.Contains("JointTitle") Then
                  If BooleanValue(vTable.Select("Title = '" & vTitle & "'")(0)("JointTitle").ToString) = False Then
                    vLabelName = vLabelName.Append(vItem & " ")
                    vContinue = False
                  End If
                End If
              End If
            Else
              If vLabelNameFormat IsNot Nothing AndAlso vLabelNameFormat.Columns.Contains("JointTitle") Then
                If BooleanValue(vLabelNameFormat.Select("LabelNameFormatCode = '" & vLabelNameFormatCode & "'")(0)("JointTitle").ToString) = False Then
                  vLabelName = vLabelName.Append(vItem & " ")
                  vContinue = False
                End If
              End If
            End If
          Case "title", "forenames", "initials", "surname", "honorifics", "preferred_forename", "prefix_honorifics"
            If vLabelNameFormatCode.Length = 0 Then
              Dim vDDL As DropDownList = TryCast(FindControlByName(mvTableControl, "Title"), DropDownList)
              If vDDL IsNot Nothing Then
                Dim vTable As DataTable = TryCast(vDDL.DataSource, DataTable)
                If vTable IsNot Nothing AndAlso vTable.Columns.Contains("JointTitle") Then
                  Dim vDataRows() As DataRow = vTable.Select("Title = '" & vTitle & "'")
                  If vDataRows.Length > 0 AndAlso BooleanValue(vDataRows(0)("JointTitle").ToString) = True Then
                    vLabelName = vLabelName.Append(vItem & " ")
                    vContinue = False
                  End If
                End If
              End If
            Else
              If vLabelNameFormat IsNot Nothing AndAlso vLabelNameFormat.Columns.Contains("JointTitle") Then
                If BooleanValue(vLabelNameFormat.Select("LabelNameFormatCode = '" & vLabelNameFormatCode & "'")(0)("JointTitle").ToString) = True Then
                  vLabelName = vLabelName.Append(vItem & " ")
                  vContinue = False
                End If
              End If
            End If
        End Select
        If vContinue Then
          If vAddSpace Then vLabelName.Append(" ")
          vAddSpace = True
          Select Case vItem
            Case "title1"
              vItem = ContactInfo.JointItem(vTitle, vTitle, vTitle1, vTitle2, False)
              If vTitle1.Length > 0 Then vLabelName = vLabelName.Append(vTitle1)
            Case "title2"
              vItem = ContactInfo.JointItem(vTitle, vTitle, vTitle1, vTitle2, False)
              If vTitle2.Length > 0 Then vLabelName = vLabelName.Append(vTitle2)
            Case "title"
              If vTitle.Length > 0 Then vLabelName.Append(vTitle) Else vAddSpace = False
            Case "initials1"
              vItem = ContactInfo.JointItem(vTitle, vInitials, vInitials1, vInitials2, False)
              If vInitials1.Length > 0 Then vLabelName = vLabelName.Append(vInitials1)
            Case "initials2"
              vItem = ContactInfo.JointItem(vTitle, vInitials, vInitials1, vInitials2, False)
              If vInitials2.Length > 0 Then vLabelName = vLabelName.Append(vInitials2)
            Case "initials"
              If vInitials.Length > 0 Then vLabelName.Append(vInitials) Else vAddSpace = False
            Case "surname1"
              vItem = ContactInfo.JointItem(vTitle, vSurname, vSurname1, vSurname2, False)
              If vSurname1.Length > 0 Then vLabelName = vLabelName.Append(vSurname1)
            Case "surname2"
              vItem = ContactInfo.JointItem(vTitle, vSurname, vSurname1, vSurname2, False)
              If vSurname2.Length > 0 Then vLabelName = vLabelName.Append(vSurname2)
            Case "surname"
              If vSurnamePrefix.Length > 0 Then
                If (vInitials.Length > 0 AndAlso vItemsList.Contains("initials")) OrElse (vForenames.Length > 0 AndAlso vItemsList.Contains("forenames")) Then
                  vLabelName.Append(vSurname)
                Else
                  vLabelName.Append(vSurname.Substring(0, 1).ToUpper)
                  vLabelName.Append(vSurname.Substring(1))
                End If
              ElseIf vSurname.Length > 0 Then
                vLabelName.Append(vSurname)
              Else
                vAddSpace = False
              End If
            Case "honorifics1"
              vItem = ContactInfo.JointItem(vTitle, vHonorifics, vHonorifics1, vHonorifics2, False)
              If vHonorifics1.Length > 0 Then vLabelName = vLabelName.Append(vHonorifics1)
            Case "honorifics2"
              vItem = ContactInfo.JointItem(vTitle, vHonorifics, vHonorifics1, vHonorifics2, False)
              If vHonorifics2.Length > 0 Then vLabelName = vLabelName.Append(vHonorifics2)
            Case "honorifics"
              If vHonorifics.Length > 0 Then vLabelName.Append(vHonorifics) Else vAddSpace = False
            Case "forenames1"
              vItem = ContactInfo.JointItem(vTitle, vForenames, vForenames1, vForenames2, False)
              If vForenames1.Length > 0 Then vLabelName = vLabelName.Append(vForenames1)
            Case "forenames2"
              vItem = ContactInfo.JointItem(vTitle, vForenames, vForenames1, vForenames2, False)
              If vForenames2.Length > 0 Then vLabelName = vLabelName.Append(vForenames2)
            Case "forenames"
              If vForenames.Length > 0 Then vLabelName.Append(vForenames) Else vAddSpace = False
            Case "preferred_forename"
              If vPreferred.Length > 0 Then vLabelName.Append(vPreferred) Else vAddSpace = False
            Case "prefix_honorifics"
              If vPrefixHonorifics.Length > 0 Then vLabelName.Append(vPrefixHonorifics) Else vAddSpace = False
            Case Else
              If vItem.Trim.Length > 0 Then vLabelName.Append(vItem.Trim) Else vAddSpace = False
          End Select
        End If
      Next
    End If
    Return vLabelName.ToString.Trim
  End Function
  Private Function IsJointContact() As Boolean
    Dim vTitle As String = GetDropDownValue("Title")
    Dim vRow As DataRow = ContactInfo.JointTitleRow(vTitle)
    If vRow IsNot Nothing Then
      IsJointContact = vRow("JointTitle").ToString = "Y"
    Else
      If DataHelper.ConfigurationOption(DataHelper.ConfigurationOptions.cd_joint_contact_support) Then
        If vTitle.IndexOf("&") > 0 Then
          IsJointContact = True
        ElseIf GetTextBoxText("Initials").IndexOf("&") > 0 Then
          Return True
        ElseIf GetTextBoxText("Forenames").IndexOf("&") > 0 Then
          Return True
        ElseIf GetTextBoxText("Surname").IndexOf("&") > 0 Then
          Return True
        End If
      End If
    End If
  End Function

  Private Function GetSalutation(ByVal pParameterName As String) As String
    Dim vSurname As String
    Dim vForenames() As String
    Dim vNewValue As String

    If mvControlType <> CareNetServices.WebControlTypes.wctUpdateContact AndAlso IsJointContact() Then
      Return ContactInfo.JointSalutation(GetDropDownValue("Title"), GetTextBoxText("Forenames"), GetDropDownValue("Surname"))
    Else
      Select Case pParameterName
        Case "Title"
          vNewValue = GetDefaultSalutationOnTitle(GetDropDownValue("Title"))
          vForenames = GetTextBoxText("Forenames").Split(" "c)
          If vForenames(0).Length > 0 AndAlso vNewValue.Contains("forename") Then vNewValue = vNewValue.Replace("forename", vForenames(0))
          vSurname = GetTextBoxText("Surname")
          If vSurname.Length > 0 AndAlso vNewValue.Contains("surname") Then vNewValue = vNewValue.Replace("surname", vSurname)
          Return vNewValue
        Case "Forenames"
          Dim vTextBox As TextBox = TryCast(FindControl("Salutation"), TextBox)
          If vTextBox IsNot Nothing Then
            vNewValue = GetTextBoxText("Salutation")
          Else
            vNewValue = GetHiddenText("HiddenSalutation")
          End If
          If vNewValue.Length > 0 Then
            'Me.ViewState.Item("HiddenOldForename") = "Lisa"
            Dim vOldForename As String = GetHiddenText("HiddenOldForename").Split(" "c)(0)

            Dim vNewForename As String = GetTextBoxText("Forenames").Split(" "c)(0)
            Dim vTemplate As String = GetSalutationFormat(pParameterName)
            If vTemplate.Contains("forename") Then
              If vOldForename.Length = 0 Then vOldForename = "forename"
              If vNewForename.Length = 0 Then vNewForename = "forename"
              'find index of forename in the existing salutation
              Dim v1st As Integer = GetSalutationNameIndex(vNewValue, vOldForename)
              'if forename has been found, and it's not at the start of the string, check template to see if it should be at the start & if it is, check if it is at the start (or if salutation has been manually changed)
              If v1st > 0 AndAlso vTemplate.IndexOf("forename") = 0 AndAlso vNewValue.StartsWith(vOldForename) Then v1st = -1
              If v1st >= 0 Then
                'if surname is in salutation template, and forename appears after surname, check to see if there's another instance of the forename (surname and forename could be the same)
                If vTemplate.Contains("surname") AndAlso vTemplate.IndexOf("forename") > vTemplate.IndexOf("surname") Then
                  Dim v2nd As Integer = GetSalutationNameIndex(vNewValue.Substring(v1st + vOldForename.Length), vOldForename)
                  If v2nd >= 0 Then
                    v1st += v2nd + vOldForename.Length
                  End If
                  'if forename hasn't been found again, check if it's there without spaces either side of it (forename could have been modified in salutation) in which case it shouldn't be changed
                  If v2nd = -1 AndAlso vNewValue.Substring(v1st + vOldForename.Length).IndexOf(vOldForename) >= 0 Then Return vNewValue
                End If
                vNewValue = vNewValue.Remove(v1st, vOldForename.Length)
                vNewValue = vNewValue.Insert(v1st, vNewForename)
              End If
            End If
          End If
          Return vNewValue
        Case "Surname"
          Dim vTextBox As TextBox = TryCast(FindControl("Salutation"), TextBox)
          If vTextBox IsNot Nothing Then
            vNewValue = GetTextBoxText("Salutation")
          Else
            vNewValue = GetHiddenText("HiddenSalutation")
          End If
          If vNewValue.Length > 0 Then
            Dim vOldSurname As String = GetHiddenText("HiddenOldSurname")
            Dim vNewSurname As String = GetTextBoxText("Surname")
            Dim vTemplate As String = GetSalutationFormat(pParameterName)
            If vTemplate.Contains("surname") Then
              If vOldSurname.Length = 0 Then vOldSurname = "surname"
              If vNewSurname.Length = 0 Then vNewSurname = "surname"
              'find index of surname in the existing salutation
              Dim v1st As Integer = GetSalutationNameIndex(vNewValue, vOldSurname)
              'if surname has been found, and it's not at the start of the string, check template to see if it should be at the start & if it is, check if it is at the start (or if salutation has been manually changed)
              If v1st > 0 AndAlso vTemplate.IndexOf("surname") = 0 AndAlso vNewValue.StartsWith(vOldSurname) Then v1st = -1
              If v1st >= 0 Then
                'if forename is in salutation template, and surname appears after surname, check to see if there's another instance of the surname (surname and forename could be the same)
                If vTemplate.Contains("forename") AndAlso vTemplate.IndexOf("forename") < vTemplate.IndexOf("surname") Then
                  Dim v2nd As Integer = GetSalutationNameIndex(vNewValue.Substring(v1st + vOldSurname.Length), vOldSurname)
                  If v2nd >= 0 Then
                    v1st = v2nd + v1st + vOldSurname.Length
                  End If
                  'if surname hasn't been found again, check if it's there without spaces either side of it (surname could have been modified in salutation) in which case it shouldn't be changed
                  If v2nd = -1 AndAlso vNewValue.Substring(v1st + vOldSurname.Length).IndexOf(vOldSurname) >= 0 Then Return vNewValue
                End If
                vNewValue = vNewValue.Remove(v1st, vOldSurname.Length)
                vNewValue = vNewValue.Insert(v1st, vNewSurname)
              End If
            End If
          End If
          SetHiddenText("HiddenSalutation", vNewValue)
          Return vNewValue
        Case "Sex"
          Dim vTextBox As TextBox = TryCast(FindControl("Salutation"), TextBox)
          If vTextBox Is Nothing Then
            vNewValue = GetHiddenText("HiddenSalutation")
          Else
            vNewValue = vTextBox.Text
          End If
          Dim vDDL As DropDownList = TryCast(FindControlByName(mvTableControl, "Title"), DropDownList)
          If vNewValue.Length = 0 OrElse _
          ((vDDL IsNot Nothing AndAlso vDDL.SelectedItem.Value.ToString.Length = 0) OrElse (vDDL Is Nothing AndAlso GetHiddenText("HiddenTitle").Length = 0)) Then
            vNewValue = GetGenderSalutationFormat(pParameterName)
            vSurname = GetTextBoxText(ContactFieldName("Surname", pParameterName))
            If vSurname.Length > 0 Then
              vSurname = vSurname.Substring(0, 1).ToUpper & vSurname.Substring(1)
              vNewValue = vNewValue.Replace("surname", vSurname)
            End If
          End If
          Return vNewValue
        Case Else
          Return ""
      End Select
    End If
  End Function

  Private Function GetSalutationNameIndex(ByVal pSalutation As String, ByVal pName As String) As Integer
    'find index of the whole word.  If it's at the start or end then it wont have spaces on both sides of it.
    Dim vValueAmended As String = " " & pName & " "
    Dim vIndex As Integer = pSalutation.IndexOf(vValueAmended)
    If vIndex < 0 Then
      vValueAmended = pName & " "
      If pSalutation.StartsWith(vValueAmended) Then vIndex = 0
      If vIndex < 0 Then
        vValueAmended = " " & pName
        If pSalutation.EndsWith(vValueAmended) Then vIndex = pSalutation.LastIndexOf(vValueAmended)
        If vIndex < 0 AndAlso pSalutation = pName Then
          vValueAmended = pName
          vIndex = 0
        End If
      End If
    End If
    If vIndex >= 0 AndAlso vValueAmended.StartsWith(" ") Then vIndex += 1
    Return vIndex
  End Function

  Private Function GetSalutationFormat(ByVal pParameterName As String) As String
    If GetDropDownValue(ContactFieldName("Title", pParameterName)).Length > 0 Then
      Return GetDefaultSalutationOnTitle(GetDropDownValue(ContactFieldName("Title", pParameterName)))
    ElseIf GetDropDownValue(ContactFieldName("Sex", pParameterName)).Length > 0 Then
      Return GetGenderSalutationFormat(pParameterName).ToString
    End If
    Return ""
  End Function

  Private Function GetGenderSalutationFormat(ByVal pParameterName As String) As String
    Dim vSalutationTemplate As String = ""
    If GetDropDownValue(ContactFieldName("Sex", pParameterName)) = "F" Then
      vSalutationTemplate = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.default_female_salutation)
    ElseIf GetDropDownValue(ContactFieldName("Sex", pParameterName)) = "M" Then
      vSalutationTemplate = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.default_male_salutation)
    End If
    If vSalutationTemplate = "" Then vSalutationTemplate = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.default_salutation)
    Return vSalutationTemplate
  End Function

  Private Function GetDefaultSalutationOnTitle(ByVal pTitle As String) As String
    Dim vSalutation As String = String.Empty
    Try
      Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtTitles)
      If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
        vTable.DefaultView.RowFilter = "Title = '" & pTitle & "'"
        If vTable.DefaultView.ToTable.Rows.Count > 0 Then vSalutation = vTable.DefaultView.ToTable.Rows(0)("Salutation").ToString
      End If
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
    Return vSalutation
  End Function

  Private Function ContactFieldName(ByVal pControlID As String, ByVal pParameterName As String) As String
    Return String.Format("{0}{1}", pControlID, FieldSuffix(pParameterName))
  End Function

  Private Function FieldSuffix(ByVal pParameterName As String) As String
    If pParameterName.EndsWith("1") Then
      Return "1"
    ElseIf pParameterName.EndsWith("2") Then
      Return "2"
    Else
      Return ""
    End If
  End Function

  Private Sub UpdateSalutation(ByVal pControl As Control)
    Try
      Dim vSalutation As String = GetSalutation(pControl.ID)
      SetTextBoxText("Salutation", vSalutation)
      SetHiddenText("HiddenSalutation", vSalutation)
      SetHiddenText("HiddenOldSurname", GetTextBoxText("Surname"))
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  Private Sub UpdateLabelName(ByVal pControl As Control)
    Try
      SetTextBoxText("LabelName", GetLabelName(False, pControl.ID))
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  Private Sub UpdateLabelNameFormatCodes(ByVal pLabelNameFormatDDL As DropDownList)
    Try
      If pLabelNameFormatDDL IsNot Nothing Then
        pLabelNameFormatDDL.Items.Clear()
        Dim vValue As String = GetDropDownValue("Title")
        Dim vList As New ParameterList(HttpContext.Current)
        vList("Title") = vValue
        DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtLabelNameFormatCodes, pLabelNameFormatDDL, True, vList)
      End If
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  Private Sub UpdateInitials(ByVal pResetExisitingInitials As Boolean, ByVal pParameterName As String, ByVal pValue As String)
    Try
      If pResetExisitingInitials OrElse GetTextBoxText(ContactFieldName("Initials", pParameterName)).Length = 0 Then
        SetTextBoxText(ContactFieldName("Initials", pParameterName), SpacePadInitials(pValue))
      End If
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  Private Sub UpdatePreferredForename(ByVal pParameterName As String, ByVal pValue As String)
    Try
      If GetTextBoxText(ContactFieldName("PreferredForename", pParameterName)).Length = 0 Then
        Dim vControl As Control = FindControlByName(mvTableControl, pParameterName)
        If vControl IsNot Nothing Then
          Dim vTextBox As TextBox = TryCast(vControl, TextBox)
          Dim vPreferredNameLength As Integer = vTextBox.MaxLength
          SetTextBoxText(ContactFieldName("PreferredForename", pParameterName), TruncateString(FirstWord(pValue), vPreferredNameLength))
        Else
          'get Hidenpreferredforename
          Dim vHiddenPreferredForename As String = GetHiddenText("HiddenPreferredForename")
          If vHiddenPreferredForename.ToLower = FirstWord(GetHiddenText("HiddenOldForename").ToLower) Then
            SetHiddenText("HiddenPreferredForename", TruncateString(FirstWord(GetTextBoxText("Forenames")), 30))
          End If
        End If
        SetHiddenText("HiddenOldForename", GetTextBoxText("Forenames"))  'Always set this. If not changing the forenames 2nd time will not update preferredforename
      ElseIf GetTextBoxText(ContactFieldName("PreferredForename", pParameterName)).Length > 0 Then
        If ValueChanged("HiddenOldForename", "Forenames") Then
          Dim vControl As Control = FindControlByName(mvTableControl, pParameterName)
          If vControl IsNot Nothing Then
            Dim vTextBox As TextBox = TryCast(vControl, TextBox)
            'If the preferred forename is set to the FirstWord of the old forenames then change preferred forename to first word of new forename
            If GetTextBoxText(ContactFieldName("PreferredForename", pParameterName)).ToLower = FirstWord(GetHiddenText("HiddenOldForename").ToLower) Then
              Dim vPreferredNameLength As Integer = vTextBox.MaxLength
              SetTextBoxText(ContactFieldName("PreferredForename", pParameterName), TruncateString(FirstWord(pValue), vPreferredNameLength))
            End If
          End If
          SetHiddenText("HiddenOldForename", GetTextBoxText("Forenames"))
        End If
      End If
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  Private Function GetDefaultBranch(ByVal pPostcode As String) As String
    If Not String.IsNullOrEmpty(pPostcode) Then
      Dim vList As New ParameterList(HttpContext.Current)
      vList("Postcode") = pPostcode
      Dim vPostcodeBranch As ParameterList = DataHelper.GetBranchFromPostCode(vList)
      Return vPostcodeBranch("Branch").ToString
    Else
      Return ""
    End If
  End Function

  Private Function TruncateString(ByVal pString As String, ByVal pLength As Integer) As String
    If pString.Length > pLength Then
      Return pString.Substring(0, pLength)
    Else
      Return pString
    End If
  End Function

  Private Function FirstWord(ByVal pString As String) As String
    Dim vPos As Integer = pString.IndexOf(" "c)
    If vPos >= 0 Then
      Return pString.Substring(0, vPos).Trim
    Else
      Return pString.Trim
    End If
  End Function

  Private Sub FriendlyUrlChangedHandler(ByVal sender As Object, ByVal e As EventArgs)
    FriendlyUrlChanged(DirectCast(sender, TextBox), DirectCast(sender, TextBox).Text)
  End Sub

#End Region

  Protected Function FindGroupControl(ByVal pGroupName As String) As CareWebControl
    For Each vControl As CareWebControl In PageCareControls
      If vControl.GroupName = pGroupName Then
        Return vControl
      End If
    Next
    Return Nothing
  End Function

  Public Overridable Sub ClearControls()
    ClearControls(False, Nothing)
  End Sub

  Public Overridable Sub ClearControls(ByVal pClearLabels As Boolean)
    ClearControls(pClearLabels, Nothing)
  End Sub
  Public Overridable Sub ClearControls(ByVal pClearLabels As Boolean, ByVal pErrorLabel As Label)
    If mvTableControl IsNot Nothing Then ClearChildControls(mvTableControl, True, pClearLabels, pErrorLabel)
    If Me.ViewState.Item("HiddenContactNumber") IsNot Nothing Then Me.ViewState.Item("HiddenContactNumber") = ""
  End Sub

  Public Sub ClearChildControls(ByVal pControl As Control, ByRef pFirstControl As Boolean)
    If pControl IsNot Nothing Then ClearChildControls(pControl, pFirstControl, False, Nothing)
  End Sub

  Public Sub ClearChildControls(ByVal pControl As Control, ByRef pFirstControl As Boolean, ByVal pClearLabels As Boolean, ByVal pErrorLabel As Label)
    For Each vControl As Control In pControl.Controls
      If vControl.Controls.Count > 0 Then
        ClearChildControls(vControl, pFirstControl, pClearLabels, pErrorLabel)
      Else
        If TypeOf (vControl) Is TextBox Then
          DirectCast(vControl, TextBox).Text = ""
        ElseIf TypeOf (vControl) Is RadioButton Then
          If pFirstControl Then
            DirectCast(vControl, RadioButton).Checked = True
            pFirstControl = False
          Else
            DirectCast(vControl, RadioButton).Checked = False
          End If
        ElseIf TypeOf (vControl) Is CheckBox Then
          DirectCast(vControl, CheckBox).Checked = False
        ElseIf TypeOf (vControl) Is DropDownList Then
          DirectCast(vControl, DropDownList).SelectedIndex = -1
        ElseIf TypeOf (vControl) Is HiddenField Then
          DirectCast(vControl, HiddenField).Value = ""
        ElseIf pClearLabels AndAlso TypeOf (vControl) Is Label Then
          Dim vClear As Boolean = True
          If pErrorLabel IsNot Nothing Then vClear = DirectCast(vControl, Label).Equals(pErrorLabel) = False
          If vClear Then DirectCast(vControl, ITextControl).Text = ""
        End If
      End If
    Next
  End Sub

  Protected Sub ContactNumberChanged(ByVal pTextBox As TextBox, ByVal pContactNumber As Integer)
    Try
      Dim vTable As DataTable = Nothing
      Dim vRow As DataRow = Nothing

      Dim vLabel As Label = TryCast(Me.FindControl(pTextBox.ID & "_Desc"), Label)
      Try
        If mvControlType <> CareNetServices.WebControlTypes.wctRapidProductPurchase Then
          For Each vCareWebControl As CareWebControl In Me.PageCareControls
            vCareWebControl.ClearControls(True)
          Next
        End If
        'Clear Externalreference and Data source field if using contact number to search contact
        If mvControlType = CareNetServices.WebControlTypes.wctRapidProductPurchase Then
          If Me.FindControlByName(mvTableControl, "ExternalReference") IsNot Nothing AndAlso Me.IsControlVisible("ExternalReference") Then
            SetTextBoxText("ExternalReference", "")
            SetLabelText("ExternalReference_Desc", "")
          End If
          If Me.FindControlByName(mvTableControl, "DataSource") IsNot Nothing AndAlso Me.IsControlVisible("DataSource") Then TryCast(FindControlByName(mvTableControl, "DataSource"), DropDownList).SelectedIndex = 0
        End If

        Session("CurrentContactNumber") = 0
        Session("CurrentAddressNumber") = 0
        If vLabel IsNot Nothing Then vLabel.Text = ""
        If pContactNumber > 0 Then
          pTextBox.Text = pContactNumber.ToString
          vTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, pContactNumber)
          vRow = DataHelper.GetRowFromDataTable(vTable)
        End If
      Catch vEx As Exception
        If vLabel IsNot Nothing Then vLabel.Text = String.Format("Contact Number {0} not found", pContactNumber)
        If mvControlType <> CareNetServices.WebControlTypes.wctRapidProductPurchase Then
          For Each vCareWebControl As CareWebControl In Me.PageCareControls
            vCareWebControl.ClearControls(True, vLabel)
          Next
        End If
      End Try

      If vRow IsNot Nothing Then
        If vRow("OwnershipAccessLevel").ToString = "W" Then
          If vLabel IsNot Nothing Then vLabel.Text = vRow("ContactName").ToString
          If mvControlType = CareNetServices.WebControlTypes.wctContactSelection Then
            Session("CurrentContactNumber") = IntegerValue(vRow("ContactNumber").ToString)
            Session("CurrentAddressNumber") = IntegerValue(vRow("AddressNumber").ToString)
            ProcessContactSelection(vTable)
          End If
        Else
          If vLabel IsNot Nothing Then vLabel.Text = String.Format("Invalid Ownership Access to Contact {0}", pContactNumber)
        End If
      Else
        If mvControlType = CareNetServices.WebControlTypes.wctRapidProductPurchase Then
          If vLabel IsNot Nothing Then vLabel.Text = String.Format("Contact Number {0} not found", pContactNumber)
        End If
      End If
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  Protected Sub ExternalReferenceChanged(ByVal pTextBox As TextBox, ByVal pExternalReference As String)
    Try
      Dim vTable As DataTable = Nothing
      Dim vRow As DataRow = Nothing

      Dim vLabel As Label = TryCast(Me.FindControl(pTextBox.ID & "_Desc"), Label)
      Try
        If mvControlType <> CareNetServices.WebControlTypes.wctRapidProductPurchase Then
          For Each vCareWebControl As CareWebControl In Me.PageCareControls
            vCareWebControl.ClearControls(True)
          Next
        End If
        Session("CurrentContactNumber") = 0
        Session("CurrentAddressNumber") = 0
        If vLabel IsNot Nothing Then vLabel.Text = ""
        If pExternalReference.Length > 0 Then
          pTextBox.Text = pExternalReference
          Dim vList As New ParameterList(HttpContext.Current)
          If mvControlType = CareNetServices.WebControlTypes.wctRapidProductPurchase Then
            If Me.FindControlByName(mvTableControl, "DataSource") IsNot Nothing AndAlso Me.IsControlVisible("DataSource") Then
              vList("DataSource") = TryCast(Me.FindControlByName(mvTableControl, "DataSource"), DropDownList).SelectedValue
            ElseIf DefaultParameters.ContainsKey("DataSource") Then
              vList("DataSource") = DefaultParameters("DataSource")
            Else
              'This should not find any records as we always need a datasource to find an external contact
              vList("DataSource") = ""
            End If
          Else
            vList("DataSource") = InitialParameters("DataSource").ToString
          End If
          vList("ExternalReference") = pExternalReference
          vTable = DataHelper.FindDataTable(CareNetServices.XMLDataFinderTypes.xdftContacts, vList)
          vRow = DataHelper.GetRowFromDataTable(vTable)
        End If
      Catch vEx As Exception
        If vLabel IsNot Nothing Then vLabel.Text = "Contact Number not found"
        'We do not want to clear all the controls while doing rapid product purchase
        If mvControlType = CareNetServices.WebControlTypes.wctRapidProductPurchase Then
          If Me.FindControlByName(mvTableControl, "ContactNumber") IsNot Nothing AndAlso Me.IsControlVisible("ContactNumber") Then
            Dim vContactNumberTextBox As TextBox = TryCast(Me.FindControlByName(mvTableControl, "ContactNumber"), TextBox)
            vContactNumberTextBox.Text = ""
            TryCast(Me.FindControl(vContactNumberTextBox.ID & "_Desc"), ITextControl).Text = ""
          End If
        Else
          For Each vCareWebControl As CareWebControl In Me.PageCareControls
            vCareWebControl.ClearControls(True, vLabel)
          Next
        End If
      End Try
      If vRow IsNot Nothing Then
        Dim vContactNumber As Integer = IntegerValue(vRow("ContactNumber").ToString)
        If vRow("OwnershipAccessLevel").ToString = "W" Then
          If vLabel IsNot Nothing Then vLabel.Text = vRow("ContactName").ToString
          If mvControlType = CareNetServices.WebControlTypes.wctContactSelectionExternalRef Then
            Session("CurrentContactNumber") = IntegerValue(vRow("ContactNumber").ToString)
            Session("CurrentAddressNumber") = IntegerValue(vRow("AddressNumber").ToString)
            vTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vContactNumber)
            ProcessContactSelection(vTable)
          End If
          If mvControlType = CareNetServices.WebControlTypes.wctRapidProductPurchase Then
            If Me.FindControlByName(mvTableControl, "ContactNumber") IsNot Nothing AndAlso Me.IsControlVisible("ContactNumber") Then
              Dim vContactNumberTextBox As TextBox = TryCast(Me.FindControlByName(mvTableControl, "ContactNumber"), TextBox)
              vContactNumberTextBox.Text = vRow("ContactNumber").ToString
              TryCast(Me.FindControl(vContactNumberTextBox.ID & "_Desc"), ITextControl).Text = vRow("ContactName").ToString
            End If
          End If
        Else
          If vLabel IsNot Nothing Then vLabel.Text = String.Format("Invalid Ownership Access to Contact {0}", vContactNumber)
        End If
      Else
        If vLabel IsNot Nothing Then vLabel.Text = "Contact Number not found"
        'We do not want to clear all the controls while doing rapid product purchase
        'Only clear contact number and contact name as it will keep the selected datasource 
        If mvControlType = CareNetServices.WebControlTypes.wctRapidProductPurchase Then
          If Me.FindControlByName(mvTableControl, "ContactNumber") IsNot Nothing AndAlso Me.IsControlVisible("ContactNumber") Then
            Dim vContactNumberTextBox As TextBox = TryCast(Me.FindControlByName(mvTableControl, "ContactNumber"), TextBox)
            vContactNumberTextBox.Text = ""
            TryCast(Me.FindControl(vContactNumberTextBox.ID & "_Desc"), ITextControl).Text = ""
          End If
        Else
          For Each vCareWebControl As CareWebControl In Me.PageCareControls
            vCareWebControl.ClearControls(True, vLabel)
          Next
        End If
      End If
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  Protected Overridable Sub FriendlyUrlChanged(ByVal pTextBox As TextBox, ByVal pValue As String)
    'Implementation in inherited controls
  End Sub

  Protected Overridable Sub SetDefaultDates()
    Select Case mvControlType
      Case CareNetServices.WebControlTypes.wctAddCategory, _
          CareNetServices.WebControlTypes.wctAddCategoryCheckboxes, _
          CareNetServices.WebControlTypes.wctAddCategoryNotes, _
          CareNetServices.WebControlTypes.wctAddCategoryOptions, _
          CareNetServices.WebControlTypes.wctAddCategoryValue, _
          CareNetServices.WebControlTypes.wctAddSuppression

        Dim vContactNumber As Integer = GetHiddenContactNumber()
        If vContactNumber = 0 Then
          If Not IsPostBack Then
            If Not DefaultParameters("SetValidFromDate") Is Nothing Then
              If DefaultParameters("SetValidFromDate").ToString.Length > 0 Then
                SetTextBoxText("ValidFrom", DefaultParameters("SetValidFromDate").ToString)
              End If
            End If

            If Not DefaultParameters("SetValidToDate") Is Nothing Then
              If DefaultParameters("SetValidToDate").ToString.Length > 0 Then
                SetTextBoxText("ValidTo", DefaultParameters("SetValidToDate").ToString)
              End If
            End If
          End If
        End If
    End Select
  End Sub
  Public Overridable Function SetDate(ByVal pDateType As DateType) As String
    Dim vDate As String = String.Empty
    Try
      Select Case mvControlType
        Case CareNetServices.WebControlTypes.wctAddCategory, _
              CareNetServices.WebControlTypes.wctAddCategoryCheckboxes, _
              CareNetServices.WebControlTypes.wctAddCategoryNotes, _
              CareNetServices.WebControlTypes.wctAddCategoryOptions, _
              CareNetServices.WebControlTypes.wctAddCategoryValue, _
              CareNetServices.WebControlTypes.wctAddSuppression, _
              CareNetServices.WebControlTypes.wctAddLink
          Select Case pDateType
            Case DateType.ValidFrom
              Dim vValidFrom As TextBox = TryCast(Me.FindControl("ValidFrom"), TextBox)
              If Not vValidFrom Is Nothing Then
                vDate = vValidFrom.Text
              End If

              If vDate.Length = 0 Then
                If Not DefaultParameters("SetValidFromDate") Is Nothing AndAlso DefaultParameters("SetValidFromDate").ToString.Length > 0 Then
                  vDate = DefaultParameters("SetValidFromDate").ToString
                Else
                  vDate = Today.ToShortDateString
                End If
              End If
            Case DateType.ValidTo
              Dim vValidTo As TextBox = TryCast(Me.FindControl("ValidTo"), TextBox)
              If Not vValidTo Is Nothing Then
                vDate = vValidTo.Text
              End If

              If vDate.Length = 0 Then
                If Not DefaultParameters("SetValidToDate") Is Nothing AndAlso DefaultParameters("SetValidToDate").ToString.Length > 0 Then
                  vDate = DefaultParameters("SetValidToDate").ToString
                Else
                  vDate = Today.AddYears(100).ToShortDateString
                End If
              End If
          End Select
      End Select

    Catch ex As Exception
      Throw ex
    End Try

    If vDate.Length = 0 Then
      Select Case pDateType
        Case DateType.ValidFrom
          vDate = Today.ToShortDateString
        Case DateType.ValidTo
          vDate = Today.AddYears(100).ToShortDateString
      End Select
    End If

    Return vDate
  End Function

  Friend Function SetCategoryValidToDate(ByVal pDateType As DateType, ByVal pValidFrom As String, ByVal pActivity As String, ByVal pActivityValue As String) As String
    Dim vDate As String = String.Empty
    Try
      Dim vValidTo As TextBox = TryCast(Me.FindControl("ValidTo"), TextBox) 'valid to has a value entered
      If Not vValidTo Is Nothing Then
        vDate = vValidTo.Text
      End If
      If vDate.Length = 0 Then
        If Not DefaultParameters("SetValidToDate") Is Nothing AndAlso DefaultParameters("SetValidToDate").ToString.Length > 0 Then 'default set on module
          vDate = DefaultParameters("SetValidToDate").ToString
        Else 'no value set so default the value (use duration if exists, otherwise use default of 100 years)
          vDate = pValidFrom
          'check for activity value duration
          Dim vList As New ParameterList(HttpContext.Current)
          vList("Activity") = pActivity
          vList("ActivityValue") = pActivityValue
          Dim vDataTable As New DataTable
          vDataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtActivityValues, vList)
          If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then
            Dim vDataRow As DataRow = vDataTable.Rows(0)
            If vDataRow("DurationDays").ToString.Length > 0 OrElse vDataRow("DurationMonths").ToString.Length > 0 Then
              Dim vTemp As Date = Date.Parse(vDate)
              If vDataRow("DurationDays").ToString.Length > 0 Then vTemp = vTemp.AddDays(IntegerValue(vDataRow("DurationDays").ToString))
              If vDataRow("DurationMonths").ToString.Length > 0 Then vTemp = vTemp.AddMonths(IntegerValue(vDataRow("DurationMonths").ToString))
              vDate = vTemp.ToShortDateString
            End If
          End If
          'if duration not set on activity value check for activity duration
          If vDate = pValidFrom Then
            vList.Remove("ActivityValue")
            vDataTable = New DataTable
            vDataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtActivities, vList)
            If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then
              Dim vDataRow As DataRow = vDataTable.Rows(0)
              If vDataRow("DurationDays").ToString.Length > 0 OrElse vDataRow("DurationMonths").ToString.Length > 0 Then
                Dim vTemp As Date = Date.Parse(vDate)
                If vDataRow("DurationDays").ToString.Length > 0 Then vTemp = vTemp.AddDays(IntegerValue(vDataRow("DurationDays").ToString))
                If vDataRow("DurationMonths").ToString.Length > 0 Then vTemp = vTemp.AddMonths(IntegerValue(vDataRow("DurationMonths").ToString))
                vDate = vTemp.ToShortDateString
              End If
            End If
          End If
          'no duration has been set, so use the default of 100 yrs after valid from
          If vDate = pValidFrom Then vDate = Today.AddYears(100).ToShortDateString
        End If
      End If
    Catch ex As Exception
      Throw ex
    End Try
    Return vDate
  End Function


  Public Overridable Sub ProcessContactSelection(ByVal pTable As DataTable)
    Dim vRow As DataRow = DataHelper.GetRowFromDataTable(pTable)
    Dim vContactNumber As Integer = IntegerValue(vRow("ContactNumber").ToString)
    Dim vAddressNumber As Integer = IntegerValue(vRow("AddressNumber").ToString)

    Select Case mvControlType
      Case CareNetServices.WebControlTypes.wctAddContact, CareNetServices.WebControlTypes.wctAddRelatedContact, _
            CareNetServices.WebControlTypes.wctUpdateContact
        SetHiddenText("HiddenContactNumber", vContactNumber.ToString)
        SetDropDownText("Title", vRow("Title").ToString, True)  'Make sure to set this first as this will raise an event to update salutation and label name
        SetDropDownText("LabelNameFormatCode", vRow("LabelNameFormatCode").ToString)
        SetTextBoxText("Salutation", vRow("Salutation").ToString)
        SetHiddenText("HiddenSalutation", vRow("Salutation").ToString)
        SetTextBoxText("PreferredForename", vRow("PreferredForename").ToString)
        SetHiddenText("HiddenPreferredForename", vRow("PreferredForename").ToString)
        SetHiddenText("HiddenOldForename", vRow("Forenames").ToString)
        SetHiddenText("HiddenTitle", vRow("Title").ToString)
        SetTextBoxText("Forenames", vRow("Forenames").ToString)
        SetTextBoxText("Surname", vRow("Surname").ToString)
        SetTextBoxText("DateOfBirth", vRow("DateOfBirth").ToString)
        SetHiddenText("HiddenSurnamePrefix", vRow("SurnamePrefix").ToString)
        SetHiddenText("HiddenSurname", vRow("Surname").ToString)
        SetHiddenText("HiddenOldSurname", vRow("Surname").ToString)
        SetHiddenText("HiddenSurname2", vRow("Surname").ToString)
        SetHiddenText("HiddenInitials", vRow("Initials").ToString)
        SetHiddenText("HiddenLabelName", vRow("LabelName").ToString)
        SetHiddenText("HiddenSex", vRow("Sex").ToString)
        SetDropDownText("Sex", vRow("Sex").ToString)

        If mvControlType = CareNetServices.WebControlTypes.wctUpdateContact Then
          SetTextBoxText("Initials", vRow("Initials").ToString)
          SetTextBoxText("Honorifics", vRow("Honorifics").ToString)
          SetTextBoxText("NiNumber", vRow("NiNumber").ToString)
          SetTextBoxText("LabelName", vRow("LabelName").ToString)
        End If

        SetHiddenText("HiddenAddressNumber", vAddressNumber.ToString)

        If mvControlType = CareNetServices.WebControlTypes.wctAddRelatedContact AndAlso vAddressNumber.ToString = Session("CurrentAddressNumber").ToString Then
          'This is a related contact module and the address is the same as the current contact
          'So don't set the address fields
        ElseIf Not mvControlType = CareNetServices.WebControlTypes.wctUpdateContact Then
          SetTextBoxText("Address", vRow("Address").ToString)
          SetTextBoxText("Town", vRow("Town").ToString)
          SetTextBoxText("County", vRow("County").ToString)
          SetTextBoxText("Postcode", vRow("Postcode").ToString)
          SetDropDownText("Country", vRow("CountryCode").ToString)
        End If

        If FindControlByName(mvTableControl, "Status") IsNot Nothing Then
          SetDropDownText("Status", vRow("Status").ToString)
          SetTextBoxText("StatusDate", vRow("StatusDate").ToString)
          SetTextBoxText("StatusReason", vRow("StatusReason").ToString)
        End If
        SetContactCommNumbersInfo(vContactNumber, True)
    End Select

    If GroupName.Length > 0 Then
      For Each vCareWebControl As CareWebControl In PageCareControls
        If vCareWebControl.ParentGroup = GroupName AndAlso _
           (vCareWebControl.mvControlType = CareNetServices.WebControlTypes.wctAddContact Or _
           vCareWebControl.mvControlType = CareNetServices.WebControlTypes.wctDisplayContactData Or _
           vCareWebControl.mvControlType = CareNetServices.WebControlTypes.wctContactSelection Or _
           vCareWebControl.mvControlType = CareNetServices.WebControlTypes.wctUpdateContact) Then
          vCareWebControl.ProcessContactSelection(pTable)
        End If
      Next

      Dim vActivityTable As DataTable = Nothing
      Dim vActivitiesRead As Boolean
      For Each vCareWebControl As CareWebControl In PageCareControls
        If vCareWebControl.ParentGroup = GroupName AndAlso vCareWebControl.HandlesActivities Then
          If vActivitiesRead = False Then
            If vCareWebControl.ParentGroup = "DelegateActivities" AndAlso Not String.IsNullOrEmpty(mvEventDelegateNumber) Then
              Dim vList As New ParameterList(HttpContext.Current)
              vList("EventDelegateNumber") = mvEventDelegateNumber
              vActivityTable = DataHelper.GetEventDataTable(CareNetServices.XMLEventDataSelectionTypes.xedtEventCategories, vList)
            End If
            If vActivityTable Is Nothing OrElse vActivityTable.Rows.Count = 0 Then
              vActivityTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCategories, vContactNumber)
            End If
          End If
          vActivitiesRead = True
          If vActivityTable IsNot Nothing Then
            vActivityTable.DefaultView.RowFilter = "Status <> 'Historic'"
            vCareWebControl.ProcessActivitySelection(vActivityTable.DefaultView.ToTable())
          Else
            vCareWebControl.DisplayActivitySuppressionModule(False)
          End If
          vCareWebControl.SetHiddenText("HiddenContactNumber", vContactNumber.ToString)
        End If
      Next
      Dim vSuppressionTable As DataTable = Nothing
      Dim vSuppressionsRead As Boolean
      For Each vCareWebControl As CareWebControl In PageCareControls
        If vCareWebControl.ParentGroup = GroupName AndAlso vCareWebControl.HandlesSuppressions Then
          If vSuppressionsRead = False Then vSuppressionTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactSuppressions, vContactNumber)
          vSuppressionsRead = True
          vCareWebControl.SetHiddenText("HiddenContactNumber", vContactNumber.ToString)
          vCareWebControl.ProcessSuppressionSelection(vSuppressionTable)
        End If
      Next

      Dim vExtReferenceTable As DataTable = Nothing
      Dim vExtReferenceRead As Boolean
      For Each vCareWebControl As CareWebControl In PageCareControls
        If vCareWebControl.ParentGroup = GroupName AndAlso vCareWebControl.HandlesExtReferences Then
          If vExtReferenceRead = False Then vExtReferenceTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactExternalReferences, vContactNumber)
          vExtReferenceRead = True
          vCareWebControl.SetHiddenText("HiddenContactNumber", vContactNumber.ToString)
          vCareWebControl.ProcessExtReferenceSelection(vExtReferenceTable)
        End If
      Next

      Dim vNeedsLinks As Boolean
      For Each vCareWebControl As CareWebControl In PageCareControls
        If vCareWebControl.ParentGroup = GroupName AndAlso vCareWebControl.HandlesLinks Then
          vNeedsLinks = True
          Exit For
        End If
      Next
      Dim vBankAccountTable As DataTable = Nothing
      Dim vBankAccountRead As Boolean
      For Each vCareWebControl As CareWebControl In PageCareControls
        If vCareWebControl.ParentGroup = GroupName AndAlso vCareWebControl.HandlesBankAccounts Then
          If vBankAccountRead = False Then vBankAccountTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactBankAccounts, vContactNumber)
          vBankAccountRead = True
          If vBankAccountTable IsNot Nothing Then
            vBankAccountTable.DefaultView.RowFilter = "HistoryOnly <> 'Yes'"
            vCareWebControl.ProcessBankAccountSelection(vBankAccountTable.DefaultView.ToTable())
          End If
          vCareWebControl.SetHiddenText("HiddenContactNumber", vContactNumber.ToString)
        End If
      Next
      If vNeedsLinks Then
        Dim vList As New ParameterList(HttpContext.Current)
        vList("RestrictNonHistoricLinks") = "Y"
        Dim vLinksTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactLinksFrom, vContactNumber, vList)
        Dim vRelatedContactCount As Integer = 0
        If vLinksTable IsNot Nothing AndAlso vLinksTable.Rows.Count > 0 Then
          Dim vRowProcessed As Boolean = False
          Dim vRelatedContactNumber As Integer = 0
          Dim vRelationship As String = ""
          For Each vDataRow As DataRow In vLinksTable.Rows
            vRowProcessed = False
            For Each vCareWebControl As CareWebControl In PageCareControls
              vRelationship = ""
              vRelatedContactNumber = 0
              If vCareWebControl.ParentGroup = GroupName AndAlso vCareWebControl.HandlesLinks Then
                vRelationship = String.Empty
                If (vCareWebControl.InitialParameters.ContainsKey("DefaultRelationship") = True AndAlso vCareWebControl.InitialParameters("DefaultRelationship").ToString.Length > 0) Then
                  'ADD RELATED CONTACT
                  'DefaultRelationship is set so find a Row that contains this Relationship and if one is found, find some other Row for that Contact with a different Relationship
                  vRelationship = vCareWebControl.InitialParameters("DefaultRelationship").ToString
                ElseIf (vCareWebControl.DefaultParameters.ContainsKey("Relationship") = True AndAlso vCareWebControl.DefaultParameters("Relationship").ToString.Length > 0) Then
                  'ADD LINK
                  'Relationship is set so find a row that contains this Relationship
                  vRelationship = vCareWebControl.DefaultParameters("Relationship").ToString
                End If
                Debug.WriteLine(vDataRow("RelationshipCode").ToString)
                If vRelationship.Length > 0 Then
                  Dim vContactNumber2 As Integer = 0
                  Dim vLinkTable As DataTable = Nothing
                  If vCareWebControl.HideHistoricLinks Then
                    'Relationship is set so find a row that contains this Relationship
                    If CanDisplayRelationshipLink(vDataRow, vRelationship) Then
                      'vValidFrom <= Today AND vValidTo >= Today AND vValidTo <> vAmendedOn
                      If (vCareWebControl.InitialParameters.ContainsKey("DefaultRelationship") = True AndAlso vCareWebControl.InitialParameters("DefaultRelationship").ToString.Length > 0) Then
                        'DefaultRelationship is set so find a Row that contains this Relationship and if one is found, find some other Row for that Contact with a different Relationship
                        vRelatedContactNumber = IntegerValue(vDataRow("ContactNumber").ToString)
                        For Each vRow2 As DataRow In vLinksTable.Rows
                          If vContactNumber2 = 0 AndAlso ((IntegerValue(vRow2("ContactNumber").ToString) = vRelatedContactNumber) _
                          AndAlso (String.Equals(vRow2("RelationshipCode").ToString, vRelationship, StringComparison.CurrentCultureIgnoreCase) = False)) Then
                            If CanDisplayRelationshipLink(vRow2, vRow2("RelationshipCode").ToString) Then
                              vCareWebControl.ProcessLinkSelection(vRow2)
                              vContactNumber2 = IntegerValue(vRow2("ContactNumber").ToString)
                              vLinkTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vContactNumber2)
                              vCareWebControl.ProcessContactSelection(vLinkTable)
                            End If
                          End If
                        Next
                        If vContactNumber2 = 0 Then
                          'No other Link found so use this one
                          vContactNumber2 = IntegerValue(vDataRow("ContactNumber").ToString)
                          vLinkTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vContactNumber2)
                          vCareWebControl.ProcessContactSelection(vLinkTable)
                        End If
                      Else
                        vCareWebControl.ProcessLinkSelection(vDataRow)
                        vContactNumber2 = IntegerValue(vDataRow("ContactNumber").ToString)
                        vLinkTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vContactNumber2)
                        vCareWebControl.ProcessContactSelection(vLinkTable)
                        vCareWebControl.SetHiddenText("HiddenContactNumber", vDataRow("ContactNumber").ToString)
                      End If
                      vRowProcessed = True
                    End If
                  Else
                    'Should be just ADD RELATED CONTACT
                    If String.Equals(vDataRow("RelationshipCode").ToString, vRelationship, StringComparison.CurrentCultureIgnoreCase) Then
                      vRelatedContactNumber = IntegerValue(vDataRow("ContactNumber").ToString)
                      For Each vRow2 As DataRow In vLinksTable.Rows
                        If (IntegerValue(vRow2("ContactNumber").ToString) = vRelatedContactNumber) AndAlso (vRow2("RelationshipCode").ToString <> vRelationship) Then
                          vCareWebControl.ProcessLinkSelection(vRow2)
                          vContactNumber2 = IntegerValue(vRow2("ContactNumber").ToString)
                          vLinkTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vContactNumber2)
                          vCareWebControl.ProcessContactSelection(vLinkTable)
                          Exit For
                        End If
                      Next
                      If vContactNumber2 = 0 Then
                        'No other Link found so use this one
                        vContactNumber2 = IntegerValue(vDataRow("ContactNumber").ToString)
                        vLinkTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vContactNumber2)
                        vCareWebControl.ProcessContactSelection(vLinkTable)
                      End If
                      vRowProcessed = True
                    End If
                  End If
                End If
                If vRowProcessed = False AndAlso vRelationship.Length = 0 Then
                  'Used for ADD RELATED CONTACT 
                  'No defaulting of Relationship so just process this row
                  If vCareWebControl.GetHiddenText("OldRelationship").ToString.Length = 0 Then
                    'Make sure we have not already set data on this control
                    If ((vCareWebControl.HideHistoricLinks = False) _
                    OrElse (vCareWebControl.HideHistoricLinks = True AndAlso CanDisplayRelationshipLink(vDataRow, "") = True)) Then
                      vCareWebControl.ProcessLinkSelection(vDataRow)
                      Dim vContactNumber2 As Integer = IntegerValue(vDataRow("ContactNumber").ToString)
                      Dim vLinkTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vContactNumber2)
                      vCareWebControl.ProcessContactSelection(vLinkTable)
                      vRowProcessed = True
                    End If
                  End If
                End If
              End If
              If vRowProcessed = True Then Exit For
            Next
          Next
        End If
      End If
    End If
  End Sub

  Private Function CanDisplayRelationshipLink(ByVal pDataRow As DataRow, ByVal pCurrentRelationship As String) As Boolean
    Dim vCanDisplay As Boolean = False

    Dim vValidFrom As Date
    Dim vHasValidFrom As Boolean = Date.TryParse(pDataRow("ValidFrom").ToString, vValidFrom)
    Dim vValidTo As Date
    Dim vHasValidTo As Boolean = Date.TryParse(pDataRow("ValidTo").ToString, vValidTo)
    Dim vAmendedOn As Date = Date.Parse(pDataRow("AmendedOn").ToString)

    If (String.Equals(pDataRow("RelationshipCode").ToString, pCurrentRelationship, StringComparison.CurrentCultureIgnoreCase) = True OrElse pCurrentRelationship.Length = 0) _
    AndAlso (vHasValidFrom = False OrElse Date.Compare(vValidFrom, Date.Today) <= 0) _
    AndAlso (vHasValidTo = False OrElse Date.Compare(vValidTo, Date.Today) >= 0) _
    AndAlso (vHasValidTo = False OrElse Date.Compare(vValidTo, vAmendedOn) <> 0) Then
      vCanDisplay = True
    End If

    Return vCanDisplay
  End Function

  Protected Sub SetContactCommNumbersInfo(ByVal pContactNumber As Integer, ByVal pSelection As Boolean)
    If FindControlByName(mvTableControl, "EMailAddress") IsNot Nothing OrElse _
       FindControlByName(mvTableControl, "DirectNumber") IsNot Nothing OrElse _
       FindControlByName(mvTableControl, "MobileNumber") IsNot Nothing OrElse _
       FindControlByName(mvTableControl, "AdditionalNumber1") IsNot Nothing OrElse _
       FindControlByName(mvTableControl, "AdditionalNumber2") IsNot Nothing OrElse _
       FindControlByName(mvTableControl, "AdditionalNumber3") IsNot Nothing Then
      mvCommNumbers(0) = New NumberInfo("DirectNumber", DataHelper.ControlValue(DataHelper.ControlValues.direct_device), 20)
      mvCommNumbers(1) = New NumberInfo("MobileNumber", DataHelper.ControlValue(DataHelper.ControlValues.mobile_device), 20)
      mvCommNumbers(2) = New NumberInfo("EMailAddress", DataHelper.ControlValue(DataHelper.ControlValues.email_device), 128)
      mvCommNumbers(3) = New NumberInfo("AdditionalNumber1", mvInitialParameters.OptionalValue("Device1"), 128)
      mvCommNumbers(4) = New NumberInfo("AdditionalNumber2", mvInitialParameters.OptionalValue("Device2"), 128)
      mvCommNumbers(5) = New NumberInfo("AdditionalNumber3", mvInitialParameters.OptionalValue("Device3"), 128)
      'When adding the following controls for a contact make sure to put it devices textlookupbox filter in CDBNETCL.frmControlParameters
      'mvCommNumbers(6) = New NumberInfo("SwitchboardNumber", DataHelper.ControlValue(DataHelper.ControlValues.switchboard_device), 20)
      'mvCommNumbers(7) = New NumberInfo("FaxNumber", DataHelper.ControlValue(DataHelper.ControlValues.fax_device), 20)
      'mvCommNumbers(8) = New NumberInfo("WebAddress", DataHelper.ControlValue(DataHelper.ControlValues.web_device), 60)

      Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers, pContactNumber)
      If Not vTable Is Nothing Then

        For Each vCommsRow As DataRow In vTable.Rows
          Dim vValidFrom As Date
          Dim vHasValidFrom As Boolean = Date.TryParse(vCommsRow("ValidFrom").ToString, vValidFrom)
          Dim vValidTo As Date
          Dim vHasValidTo As Boolean = Date.TryParse(vCommsRow("ValidTo").ToString, vValidTo)
          Dim vAmendedOn As Date = Date.Parse(vCommsRow("AmendedOn").ToString)
          If (vHasValidFrom = False OrElse vValidFrom <= Date.Today) AndAlso _
              (vHasValidTo = False OrElse vValidTo >= Date.Today) AndAlso _
              (vHasValidTo = False OrElse Not (vValidTo = vAmendedOn)) Then

            Dim vDeviceCode As String = vCommsRow.Item("DeviceCode").ToString
            For Each vNumber As NumberInfo In mvCommNumbers
              If vNumber IsNot Nothing AndAlso vDeviceCode = vNumber.DeviceCode Then
                Dim vExt As String = vCommsRow.Item("Extension").ToString
                Dim vPhoneNumber As String = vCommsRow.Item("PhoneNumber").ToString
                If vExt.Length > 0 AndAlso vPhoneNumber.EndsWith(vExt) Then
                  Dim vPos As Integer = InStr(vPhoneNumber, " Ext ")
                  If vPos > 0 Then vPhoneNumber = vPhoneNumber.Substring(0, vPos - 1)
                End If
                If pSelection Then
                  SetTextBoxText(vNumber.Identifier, vPhoneNumber)
                  'SetHiddenText("Old" & vNumber.Identifier, vPhoneNumber)
                End If
                vNumber.CommunicationNumber = IntegerValue(vCommsRow.Item("CommunicationNumber").ToString)
                vNumber.Number = vPhoneNumber
                vNumber.DeviceDefault = BooleanValue(vCommsRow.Item("DeviceDefault").ToString)
                vNumber.IsDefault = BooleanValue(vCommsRow.Item("Default").ToString)
                vNumber.Mail = BooleanValue(vCommsRow.Item("Mail").ToString)
                vNumber.PreferredMethod = BooleanValue(vCommsRow.Item("PreferredMethod").ToString)
                Exit For
              End If
            Next
          End If
        Next
      End If
    End If
  End Sub

  Public Overridable Sub ProcessActivitySelection(ByVal pTable As DataTable)
    'These are all overriden in the inherited classes
  End Sub
  Public Overridable Sub DisplayActivitySuppressionModule(ByVal pValue As Boolean)
    'These are all overriden in the inherited classes
  End Sub
  Public Overridable Sub ProcessSuppressionSelection(ByVal pTable As DataTable)
    'These are all overriden in the inherited classes
  End Sub
  Public Overridable Sub ProcessExtReferenceSelection(ByVal pTable As DataTable)
    'These are all overriden in the inherited classes
  End Sub
  Public Overridable Sub ProcessLinkSelection(ByVal pRow As DataRow)
    'These are all overriden in the inherited classes
  End Sub
  Public Overridable Sub ProcessBankAccountSelection(ByVal pTable As DataTable)
    'These are all overriden in the inherited classes
  End Sub

#Region "Add Methods for Validation and Server valdiation"

  Protected Overridable Sub AddCustomValidator(ByVal pHTMLTable As HtmlTable)

  End Sub

  Protected Overridable Sub AddCustomValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String)

  End Sub

  Protected Sub AddCustomValidatorToControl(ByVal pParameterName As String, ByVal pID As String, ByVal pErrorMessage As String)
    Dim vControl As Control = FindControlByName(mvTableControl, pParameterName)
    If vControl IsNot Nothing Then
      AddCustomValidator(DirectCast(vControl.Parent, HtmlTableCell), pID, pErrorMessage)
    End If
  End Sub

  Protected Sub AddCustomValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String, ByVal pErrorMessage As String)
    AddCustomValidator(pHTMLCell, pID, pErrorMessage, Nothing)
  End Sub

  Protected Sub AddCustomValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String, ByVal pErrorMessage As String, ByVal pControlID As String)
    Dim vCFV As New CustomValidator
    With vCFV
      .CssClass = "DataValidator"
      .ID = "cfv" & pID
      If Not String.IsNullOrEmpty(pControlID) Then .ControlToValidate = pID
      .Display = ValidatorDisplay.Dynamic
      .ErrorMessage = pErrorMessage
      .SetFocusOnError = True
    End With

    Select Case pControlID
      Case "CardStartDate"
        If InitialParameters.OptionalValue("PaymentService").Trim.ToUpper <> "TNSHOSTED" AndAlso _
          Not (mvControlType = CareNetServices.WebControlTypes.wctAddMemberCC AndAlso InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" AndAlso _
                          DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED) Then
          AddHandler vCFV.ServerValidate, AddressOf ServerValidateStartDate
        End If

      Case "CardExpiryDate"
        If InitialParameters.OptionalValue("PaymentService").Trim.ToUpper <> "TNSHOSTED" AndAlso _
          Not (mvControlType = CareNetServices.WebControlTypes.wctAddMemberCC AndAlso InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" AndAlso _
                          DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED) Then _
          AddHandler vCFV.ServerValidate, AddressOf ServerValidateExpiryDate
      Case "AccountNumber"
        AddHandler vCFV.ServerValidate, AddressOf ServerValidateAccountNumber
      Case "SecurityQuestion", "SecurityAnswer"
        AddHandler vCFV.ServerValidate, AddressOf ServerValidateSecurityQuestion
      Case "Town"
        vCFV.ValidateEmptyText = True
        AddHandler vCFV.ServerValidate, AddressOf ServerValidateTown
      Case Else
        AddHandler vCFV.ServerValidate, AddressOf ServerValidate
    End Select
    pHTMLCell.Controls.Add(vCFV)
  End Sub

  'Protected Sub AddMultiFieldValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String, ByVal pControls As String, ByVal pCondition As MultiFieldValidator.Conditions)
  '  Dim vMFV As New MultiFieldValidator
  '  With vMFV
  '    .CssClass = "DataValidator"
  '    .ID = "mfv" & pID
  '    .Display = ValidatorDisplay.Dynamic
  '    .ErrorMessage = "Mutiple Fields must be completed"
  '    .Condition = pCondition
  '    .ControlsToValidate = pControls
  '  End With
  '  pHTMLCell.Controls.Add(vMFV)
  'End Sub
  Protected Sub AddRangeValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String, ByVal pMinimumValue As Integer, ByVal pMaximumValue As Integer, Optional ByVal pErrorMessage As String = "", Optional ByVal pUseInteger As Boolean = False)
    Dim vRNV As New RangeValidator
    With vRNV
      .CssClass = "DataValidator"
      .ID = "rnv" & pID
      .ControlToValidate = pID
      .MinimumValue = pMinimumValue.ToString
      .MaximumValue = pMaximumValue.ToString
      If pUseInteger Then .Type = ValidationDataType.Integer
      .Display = ValidatorDisplay.Dynamic
      If pErrorMessage.Length = 0 Then
        .ErrorMessage = "Invalid Value"
      Else
        .ErrorMessage = pErrorMessage
      End If
      .SetFocusOnError = True
    End With
    pHTMLCell.Controls.Add(vRNV)
  End Sub
  Private Sub AddCompareValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String, ByVal pCompareID As String, ByVal pOperator As ValidationCompareOperator)
    Dim vCV As New CompareValidator
    With vCV
      .CssClass = "DataValidator"
      .ID = "cfv" & pID
      .ControlToValidate = pID
      .ControlToCompare = pCompareID
      .Operator = pOperator
      .Display = ValidatorDisplay.Dynamic
      .ErrorMessage = "Invalid Value"
      .Operator = pOperator
      .SetFocusOnError = True
    End With
    pHTMLCell.Controls.Add(vCV)
  End Sub

  Private Sub AddValueCompareValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String, ByVal pOperator As ValidationCompareOperator, ByVal pValue As String)
    Dim vCV As New CompareValidator
    With vCV
      .CssClass = "DataValidator"
      .ID = "cfv" & pID
      .ControlToValidate = pID
      .Operator = pOperator
      .Type = ValidationDataType.Date
      .ValueToCompare = pValue
      .Display = ValidatorDisplay.Dynamic
      Dim vOperator As String = "="
      Select Case pOperator
        Case ValidationCompareOperator.GreaterThan
          vOperator = ">"
        Case ValidationCompareOperator.GreaterThanEqual
          vOperator = ">="
        Case ValidationCompareOperator.LessThan
          vOperator = "<"
        Case ValidationCompareOperator.LessThanEqual
          vOperator = "<="
        Case ValidationCompareOperator.NotEqual
          vOperator = "Not equal to"
      End Select
      .ErrorMessage = String.Format("Value must be {0} {1}", vOperator, pValue)
      .Operator = pOperator
      .SetFocusOnError = True
    End With
    pHTMLCell.Controls.Add(vCV)
  End Sub

  Private Sub AddDateCompareValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String, ByVal pCompareID As String, ByVal pOperator As ValidationCompareOperator)
    Dim vCV As New CompareValidator
    With vCV
      .CssClass = "DataValidator"
      .ID = "cfv" & pID
      .Type = ValidationDataType.Date
      .ControlToValidate = pID
      If pCompareID = "Today" Then
        .ValueToCompare = Date.Today.ToShortDateString
        If pOperator = ValidationCompareOperator.LessThan Or pOperator = ValidationCompareOperator.LessThanEqual Then
          .ErrorMessage = "Date Cannot be in the future"
        Else
          .ErrorMessage = "Invalid Date"
        End If
      Else
        .ControlToCompare = pCompareID
        .ErrorMessage = "Invalid Date Range"
      End If
      .Display = ValidatorDisplay.Dynamic
      .Operator = pOperator
      .Type = ValidationDataType.Date
      .SetFocusOnError = True
    End With
    pHTMLCell.Controls.Add(vCV)
  End Sub

  Private Sub AddDoubleValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String)
    AddDataTypeValidator(pHTMLCell, pID, ValidationDataType.Double)
  End Sub

  Private Sub AddDataTypeValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String, ByVal pType As ValidationDataType)
    Dim vCV As New CompareValidator
    With vCV
      .CssClass = "DataValidator"
      .ID = "afv" & pID
      .ControlToValidate = pID
      .Display = ValidatorDisplay.Dynamic
      .ErrorMessage = "Invalid Value"
      .Operator = ValidationCompareOperator.DataTypeCheck
      .Type = pType
      .SetFocusOnError = True
    End With
    pHTMLCell.Controls.Add(vCV)
  End Sub

  Private Sub AddDateValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String)
    Dim vCV As New CompareValidator
    With vCV
      .CssClass = "DataValidator"
      .ID = "dfv" & pID
      .ControlToValidate = pID
      .Display = ValidatorDisplay.Dynamic
      .ErrorMessage = "Invalid Date"
      .Operator = ValidationCompareOperator.DataTypeCheck
      .Type = ValidationDataType.Date
      .SetFocusOnError = True
    End With
    pHTMLCell.Controls.Add(vCV)
  End Sub

  Protected Sub AddRequiredValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String, Optional ByVal pvalidationGroup As String = "")
    Dim vRFV As New RequiredFieldValidator
    With vRFV
      .CssClass = "DataValidator"
      .ID = "rfv" & pID
      .ControlToValidate = pID
      .Display = ValidatorDisplay.Dynamic
      .ErrorMessage = "Required Field"
      .SetFocusOnError = True
      If pvalidationGroup.Length > 0 Then
        .ValidationGroup = pvalidationGroup
      End If
    End With
    pHTMLCell.Controls.Add(vRFV)
  End Sub


  Public Overridable Sub ServerValidateStartDate(ByVal sender As Object, ByVal e As ServerValidateEventArgs)
    Dim vValid As Boolean = False
    Dim vControl As Control = FindControlByName(mvTableControl, DirectCast(sender, CustomValidator).ControlToValidate)
    If vControl IsNot Nothing Then
      Dim vDateString As String = DirectCast(vControl, TextBox).Text
      If vDateString.Length = 4 Then
        Dim vYear As Integer = IntegerValue(vDateString.Substring(2, 2))
        If vYear > 60 Then
          vYear += 1900
        Else
          vYear += 2000
        End If
        If vYear < Date.Today.Year OrElse (vYear = Date.Today.Year AndAlso IntegerValue(vDateString.Substring(0, 2)) <= Date.Today.Month) Then
          vValid = True
        End If
      End If
    End If
    e.IsValid = vValid
  End Sub

  Public Overridable Sub ServerValidateExpiryDate(ByVal sender As Object, ByVal e As ServerValidateEventArgs)
    Dim vValid As Boolean = False
    Dim vControl As Control = FindControlByName(mvTableControl, DirectCast(sender, CustomValidator).ControlToValidate)
    If vControl IsNot Nothing Then
      Dim vDateString As String = DirectCast(vControl, TextBox).Text
      If vDateString.Length = 4 Then
        Dim vYear As Integer = IntegerValue(vDateString.Substring(2, 2)) + 2000
        If vYear > Date.Today.Year OrElse (vYear = Date.Today.Year AndAlso IntegerValue(vDateString.Substring(0, 2)) >= Date.Today.Month) Then
          vValid = True
        End If
      End If
    End If
    e.IsValid = vValid
  End Sub

  Public Overridable Sub ServerValidateSecurityQuestion(ByVal sender As Object, ByVal e As ServerValidateEventArgs)
    Dim vValid As Boolean = True
    Dim vQuestionTextBox As TextBox = DirectCast(FindControlByName(mvTableControl, "SecurityQuestion"), TextBox)
    Dim vAnswerTextBox As TextBox = DirectCast(FindControlByName(mvTableControl, "SecurityAnswer"), TextBox)
    Dim vCFV As CustomValidator = DirectCast(sender, CustomValidator)
    Try
      If vCFV.ControlToValidate = "SecurityQuestion" AndAlso vQuestionTextBox.Text.Length > 0 AndAlso vAnswerTextBox.Text.Length = 0 Then
        vValid = False
      ElseIf vCFV.ControlToValidate = "SecurityAnswer" AndAlso vQuestionTextBox.Text.Length = 0 AndAlso vAnswerTextBox.Text.Length > 0 Then
        vValid = False
      End If
    Catch vEx As CareException

      Throw vEx
    Finally
    End Try
    e.IsValid = vValid
  End Sub

  Public Overridable Sub ServerValidateAccountNumber(ByVal sender As Object, ByVal e As ServerValidateEventArgs)
    Dim vValid As Boolean = False
    Dim vCFV As CustomValidator = DirectCast(sender, CustomValidator)
    Dim vTextBox As TextBox = DirectCast(FindControlByName(mvTableControl, vCFV.ControlToValidate), TextBox)
    Dim vResult As ParameterList = Nothing
    Try
      Dim vAccountNo As String = vTextBox.Text
      Dim vSortCode As String = GetTextBoxText("SortCode").Replace("-", "")
      If vAccountNo.Contains("*") Then vAccountNo = GetHiddenText("OldAccountNumber")
      If vSortCode.Contains("*") Then vSortCode = GetHiddenText("OldSortCode")
      vResult = DataHelper.AccountNoVerify(vSortCode, vAccountNo)
      If IntegerValue(vResult("VerifyResult").ToString) = DataHelper.AccountNoVerifyResult.avrValid Then vValid = True
    Catch vEx As CareException
      If vEx.ErrorNumber = CareException.ErrorNumbers.enALBACSVerify Then
        vCFV.ErrorMessage = vEx.Message
      Else
        Throw vEx
      End If
    Finally
      If vValid = False Then
        If vResult IsNot Nothing Then vCFV.ErrorMessage = vResult("VerifyResultMessage").ToString
        mvFocusControl = vTextBox
      End If
    End Try
    e.IsValid = vValid
  End Sub

  Public Overridable Sub ServerValidateTown(ByVal sender As Object, ByVal e As ServerValidateEventArgs)
    Dim vValid As Boolean = True
    Dim vTownTextBox As TextBox = DirectCast(FindControlByName(Me, "Town"), TextBox)
    Dim vCountryDropDownList As DropDownList = TryCast(FindControlByName(Me, "Country"), DropDownList)
    Dim vAddress As TextBox = TryCast(FindControlByName(Me, "Address"), TextBox)
    Try
      If vCountryDropDownList IsNot Nothing AndAlso vCountryDropDownList.SelectedValue = "UK" AndAlso vTownTextBox.Text.Length = 0 Then
        vValid = False
      End If
      'when registering as a new user in the register module you should be able to register without entering an address.
      If Not vValid AndAlso vAddress IsNot Nothing AndAlso vAddress.Text.Length = 0 Then
        vValid = True
      End If
    Catch vEx As CareException

      Throw vEx
    Finally
    End Try
    e.IsValid = vValid
  End Sub

  Public Overridable Sub ServerValidate(ByVal sender As Object, ByVal e As ServerValidateEventArgs)
    e.IsValid = True
  End Sub

  Public Sub AddDateTimePicker(ByVal pControlName As String)
    If mvDateControls Is Nothing Then mvDateControls = New List(Of String)
    mvDateControls.Add(pControlName)
  End Sub

  Private Enum RegularExpressionTypes
    retSortCode
    retAccountNumber
    retCreditCardNumber
    retExpiryDate
    retIssueNumber
    retStartDate
    retSecurityCode
    retEncryptedAccountNumber
    retEncryptedSortCode
    retPassword
    retExpiryDateMonth
    retExpiryDateYear
  End Enum

  Private Sub AddRegExValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String, ByVal pType As RegularExpressionTypes)
    Dim vREV As New RegularExpressionValidator
    With vREV
      .CssClass = "DataValidator"
      .ID = "rev" & pID
      .ControlToValidate = pID
      .Display = ValidatorDisplay.Dynamic
      .SetFocusOnError = True
      Select Case pType
        Case RegularExpressionTypes.retAccountNumber
          .ErrorMessage = "Invalid Account Number"
          .ValidationExpression = "^[0-9]{8,16}$"
        Case RegularExpressionTypes.retEncryptedAccountNumber
          .ErrorMessage = "Invalid Account Number"
          .ValidationExpression = "^[0-9]{8,16}|[*]{4}[0-9]{4}|[*]{12}[0-9]{4}$"
        Case RegularExpressionTypes.retCreditCardNumber
          .ErrorMessage = "Invalid Credit Card Number"
          .ValidationExpression = "^(?:4[0-9]{12}(?:[0-9]{3})?|5[1-5][0-9]{14}|6011[0-9]{12}|3(?:0[0-5]|[68][0-9])[0-9]{11}|3[47][0-9]{13}|(5000|4903|5018|5020|5038|6304|6759|6761|6763)\d{8,15}|(6334|6767)\d{15}|(6334|6767)\d{14}|(6334|6767)\d{12}|(6304|6706|6709|6771)[0-9]{12,15})$" 'BR13374 added validation for Maestro and Solo debit cards
        Case RegularExpressionTypes.retExpiryDate
          .ErrorMessage = "Invalid Expiry Date"
          .ValidationExpression = "^[01][0-9][0-3][0-9]$"
        Case RegularExpressionTypes.retIssueNumber
          .ErrorMessage = "Invalid Issue Number"
          .ValidationExpression = "^\d{1,2}$"
        Case RegularExpressionTypes.retSortCode
          .ErrorMessage = "Invalid Sort Code"
          .ValidationExpression = "^[0-9]{2}-?[0-9]{2}-?[0-9]{2}$"
        Case RegularExpressionTypes.retEncryptedSortCode
          .ErrorMessage = "Invalid Sort Code"
          .ValidationExpression = "^[0-9]{2}-?[0-9]{2}-?[0-9]{2}|[0-9]{2}-?[*]{2}-?[*]{2}$"
        Case RegularExpressionTypes.retStartDate
          .ErrorMessage = "Invalid Start Date"
          .ValidationExpression = "^[01][0-9][0-3][0-9]$"
          'Case RegularExpressionTypes.retWebPageUrl
          '.ErrorMessage = "Invalid Url"
          '.ValidationExpression = "[A-Za-z0-9_.]"
        Case RegularExpressionTypes.retSecurityCode
          .ErrorMessage = "Invalid Card Security Code"
          .ValidationExpression = "^\d\d\d\d?"
        Case RegularExpressionTypes.retPassword
          Dim vMinimumLength As Integer
          vMinimumLength = IntegerValue(DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.portal_password_min_length))
          If vMinimumLength = 0 Then vMinimumLength = 1
          .ErrorMessage = "Invalid Password"
          If DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.portal_password_complexity) = "C" Then
            .ValidationExpression = "^(?:(?=.*[a-z])(?:(?=.*[A-Z])(?=.*[\d\W])|(?=.*\W)(?=.*\d))|(?=.*\W)(?=.*[A-Z])(?=.*\d)).{" & vMinimumLength & ",}$"
            .ErrorMessage = "Password does not meet complexity rules"
          Else
            .ValidationExpression = ".{" & vMinimumLength & ",}"
          End If
        Case RegularExpressionTypes.retExpiryDateMonth
          .ErrorMessage = "Invalid Card Expiry Date (month)"
          .ValidationExpression = "^[0-9m]{1,2}$"
        Case RegularExpressionTypes.retExpiryDateYear
          .ErrorMessage = "Invalid Card Expiry Date (Year)"
          .ValidationExpression = "^[0-9y]{4}$"
      End Select
    End With
    pHTMLCell.Controls.Add(vREV)
  End Sub

  Private Sub AddEmailValidator(ByVal pHTMLCell As Control, ByVal pID As String)
    pHTMLCell.Controls.Add(New EmailAddressValidator With {.CssClass = "DataValidator",
                                                           .ID = "rev" & pID,
                                                           .ControlToValidate = pID,
                                                           .Display = ValidatorDisplay.Dynamic,
                                                           .SetFocusOnError = True,
                                                           .ErrorMessage = "Invalid Email Address",
                                                           .ValidateRequestMode = UI.ValidateRequestMode.Enabled
                                                          }
                          )
  End Sub

  Private Sub AddTelephoneNumberValidator(ByVal pHTMLCell As Control, ByVal pID As String)
    pHTMLCell.Controls.Add(New TelephoneNumberValidator With {.CssClass = "DataValidator",
                                                           .ID = "rev" & pID,
                                                           .ControlToValidate = pID,
                                                           .Display = ValidatorDisplay.Dynamic,
                                                           .SetFocusOnError = True,
                                                           .ErrorMessage = "Invalid Telephone Number",
                                                           .ValidateRequestMode = UI.ValidateRequestMode.Enabled
                                                          }
                          )
  End Sub

#End Region

#Region "Methods for getting parameters"

  Protected Sub AddOptionalTextBoxValueWithWildCard(ByVal pList As ParameterList, ByVal pParameterName As String)
    Dim vValue As String = GetTextBoxText(pParameterName)
    If vValue.Length > 0 Then pList(pParameterName) = vValue & "*"
  End Sub

  Protected Sub AddOptionalTextBoxValue(ByVal pList As ParameterList, ByVal pParameterName As String)
    AddOptionalTextBoxValue(pList, pParameterName, False)
  End Sub

  Protected Sub AddOptionalTextBoxValue(ByVal pList As ParameterList, ByVal pParameterName As String, ByVal pCheckControlExists As Boolean)
    Dim vValue As String = GetTextBoxText(pParameterName, pCheckControlExists)
    If pCheckControlExists Then
      pList(pParameterName) = vValue
    Else
      If vValue.Length > 0 Then pList(pParameterName) = vValue
    End If
  End Sub

  Protected Sub AddOptionalDropDownValue(ByVal pList As ParameterList, ByVal pParameterName As String)
    AddOptionalDropDownValue(pList, pParameterName, False)
  End Sub

  Protected Sub AddOptionalDropDownValue(ByVal pList As ParameterList, ByVal pParameterName As String, ByVal pCheckControlExists As Boolean)
    Dim vValue As String = GetDropDownValue(pParameterName, pCheckControlExists)
    If pCheckControlExists Then
      pList(pParameterName) = vValue
    Else
      If vValue.Length > 0 Then pList(pParameterName) = vValue
    End If
  End Sub

  Protected Sub AddDefaultParameters(ByVal pList As ParameterList)
    For Each vItem As DictionaryEntry In DefaultParameters
      pList(vItem.Key) = vItem.Value
    Next
  End Sub

  Protected Sub AddCCParameters(ByVal pList As ParameterList)
    If IsControlVisible("CreditCardNumber") = False AndAlso IsControlVisible("CardExpiryDate") = False AndAlso IsControlVisible("IssueNumber") = False AndAlso _
      IsControlVisible("CardStartDate") = False AndAlso IsControlVisible("SecurityCode") = False Then
      'All controls are hidden so allow CC payment without CC details
      pList("NoClaimRequired") = "Y"
    Else
      AddOptionalTextBoxValue(pList, "CreditCardNumber")

      If mvControlType = CareNetServices.WebControlTypes.wctAddMemberCC AndAlso InitialParameters.OptionalValue("OnlineCCAuthorisation") = "Y" AndAlso _
        DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.fp_cc_authorisation_type) = TNSHOSTED Then

        Dim vExpiryDate As String = String.Empty
        If FindControlByName(Me, "CardExpiryDate") IsNot Nothing AndAlso FindControlByName(Me, "gatewayCardExpiryDateYear") IsNot Nothing Then
          Try
            vExpiryDate = TryCast(Me.FindControl("CardExpiryDate"), TextBox).Text + TryCast(Me.FindControl("gatewayCardExpiryDateYear"), TextBox).Text.Substring(2, 2)
          Catch ex As Exception
            vExpiryDate = ""
          End Try
        End If
        pList.Add("CardExpiryDate", vExpiryDate)
      Else
        AddOptionalTextBoxValue(pList, "CardExpiryDate")
      End If

      AddOptionalTextBoxValue(pList, "IssueNumber")
      AddOptionalTextBoxValue(pList, "CardStartDate")
      AddOptionalDropDownValue(pList, "CreditCardType")
      If IsControlVisible("SecurityCode") Then
        AddOptionalTextBoxValue(pList, "SecurityCode")
        If pList.Contains("SecurityCode") Then pList("GetAuthorisation") = "Y"
      End If
    End If
  End Sub

  Protected Sub AddDDParameters(ByVal pList As ParameterList)
    AddOptionalTextBoxValue(pList, "SortCode")
    AddOptionalTextBoxValue(pList, "AccountNumber")
    AddOptionalTextBoxValue(pList, "AccountName")
    Dim vClaimDay As String = GetClaimDay(pList("BankAccount").ToString)
    If vClaimDay.Length > 0 Then pList.Add("ClaimDay", vClaimDay)
  End Sub

  Protected Sub AddMemberParameters(ByVal pList As ParameterList, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pPaymentMethod As String, ByVal pMembershipType As String, ByVal pMembershipStartDate As String)
    If ParentGroup.Length = 0 Then
      pList("UserID") = pContactNumber
    Else
      pList("UserID") = UserContactNumber()
    End If
    pList("PayerContactNumber") = pContactNumber
    pList("PayerAddressNumber") = pAddressNumber

    If pMembershipStartDate.Length > 0 Then
      pList("StartDate") = pMembershipStartDate
    Else
      pList("StartDate") = Date.Today.ToShortDateString
    End If

    pList("Joined") = Date.Today.ToShortDateString
    pList("PaymentMethod") = pPaymentMethod
    pList("MembershipType") = pMembershipType
    If DefaultParameters.ContainsKey("SetCardExpiry") Then pList("SetCardExpiry") = DefaultParameters("SetCardExpiry").ToString
    pList("MemberFixedAmount") = 0
    AddDefaultParameters(pList)    'BankAccount,PaymentFrequency,Source,ReasonForDespatch
    If (Not InitialParameters.ContainsKey("MembershipFor")) OrElse (InitialParameters("MembershipFor").ToString.ToUpper <> "O") OrElse ParentGroup.Length > 0 Then
      If pList.ContainsKey("Branch") Then
        pList("DefaultBranch") = DefaultParameters("Branch")
        pList.Remove("Branch")
      End If
    End If
    pList("ChangeBranchWithAddress") = "N"
  End Sub

  Protected Sub AddPaymentPlanParameters(ByVal pList As ParameterList, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pPaymentMethod As String)
    pList("UserID") = pContactNumber
    pList("PayerContactNumber") = pContactNumber
    pList("PayerAddressNumber") = pAddressNumber
    pList("StartDate") = Date.Today.ToShortDateString
    pList("PaymentMethod") = pPaymentMethod
    AddDefaultParameters(pList)    'Product,Rate,BankAccount,PaymentFrequency,Source,ReasonForDespatch
  End Sub

  Protected Sub AddUserParameters(ByVal pList As ParameterList)
    If Convert.ToString(pList("UserLogname")).Length = 0 Then
      ' UserID will get set only for registered user
      If Session.Contents.Item("UserContactNumber") IsNot Nothing AndAlso IntegerValue(Session("UserContactNumber").ToString) > 0 Then
        pList("UserID") = IntegerValue(Session("UserContactNumber").ToString)
      End If
    End If
  End Sub

#End Region

#Region "Get and Set Methods for page values"

  Protected Function GetCheckBoxChecked(ByVal pParameterName As String) As Boolean
    Dim vControl As Control = FindControlByName(mvTableControl, pParameterName)
    If vControl IsNot Nothing Then
      Dim vCheckBox As CheckBox = TryCast(vControl, CheckBox)
      If vCheckBox IsNot Nothing AndAlso vCheckBox.Checked Then Return True
    End If
  End Function

  Protected Function GetDropDownValue(ByVal pParameterName As String) As String
    Return GetDropDownValue(pParameterName, False)
  End Function

  Protected Function GetDropDownValue(ByVal pParameterName As String, ByRef pCheckControlExists As Boolean) As String
    Dim vDDL As DropDownList = TryCast(FindControlByName(mvTableControl, pParameterName), DropDownList)
    If vDDL IsNot Nothing Then
      If vDDL.SelectedItem IsNot Nothing Then Return vDDL.SelectedItem.Value.ToString
    Else
      pCheckControlExists = False
    End If
    Return ""
  End Function

  Protected Function GetHiddenText(ByVal pParameterName As String) As String
    If Me.ViewState.Item(pParameterName) IsNot Nothing Then
      Return Me.ViewState.Item(pParameterName).ToString()
    End If
    Return ""
  End Function

  Protected Function GetHiddenAddressNumber() As Integer
    Return IntegerValue(GetHiddenText("HiddenAddressNumber"))
  End Function

  Protected Function GetHiddenContactNumber() As Integer
    Return IntegerValue(GetHiddenText("HiddenContactNumber"))
  End Function

  Protected Function GetTextBoxText(ByVal pParameterName As String) As String
    Return GetTextBoxText(pParameterName, False)
  End Function

  Protected Function GetTextBoxText(ByVal pParameterName As String, ByRef pCheckControlExists As Boolean) As String
    Dim vControl As Control = FindControlByName(mvTableControl, pParameterName)
    If vControl IsNot Nothing Then
      Dim vTextBox As TextBox = TryCast(vControl, TextBox)
      If vTextBox IsNot Nothing Then
        Select Case pParameterName
          Case "SortCode"
            Return vTextBox.Text.Replace("-", "")
          Case Else
            Return vTextBox.Text
        End Select
      End If
    Else
      pCheckControlExists = False
    End If
    Return ""
  End Function

  Protected Function GetLabelText(ByVal pParameterName As String) As String
    Dim vControl As Control = FindControlByName(mvTableControl, pParameterName)
    If vControl IsNot Nothing Then
      Dim vLabel As Label = TryCast(vControl, Label)
      If vLabel IsNot Nothing Then Return vLabel.Text
    End If
    Return ""
  End Function

  Protected Sub SetLookupItem(ByVal pType As CareNetServices.XMLLookupDataTypes, ByVal pParameterName As String)
    SetLookupItem(pType, pParameterName, InitialParameters(pParameterName).ToString)
  End Sub

  Protected Sub SetLookupItem(ByVal pType As CareNetServices.XMLLookupDataTypes, ByVal pParameterName As String, ByVal pParameterValue As String)
    Dim vList As New ParameterList(HttpContext.Current)
    vList(pParameterName) = pParameterValue
    Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetLookupData(pType, vList))
    If vRow IsNot Nothing Then SetTextBoxText(pParameterName, vRow(pParameterName & "Desc").ToString)
  End Sub

  Private Function GetClaimDay(ByVal pBankAccount As String) As String
    Dim vClaimDay As String = ""
    If DataHelper.ControlValue(DataHelper.ControlValues.auto_pay_claim_date_method) = "D" Then
      Dim vControl As Control = FindControlByName(mvTableControl, "ClaimDay")
      If vControl IsNot Nothing Then
        Dim vDDL As DropDownList = TryCast(vControl, DropDownList)
        If vDDL.SelectedItem IsNot Nothing Then vClaimDay = vDDL.SelectedItem.Value
        If vClaimDay.Length = 0 Then
          vClaimDay = vDDL.Items.Item(1).ToString
        End If
      Else
        Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtBankAccountClaimDays)
        If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
          DataHelper.DefaultClaimDay(vTable, pBankAccount)
          If vTable.DefaultView.ToTable.Rows.Count > 0 Then vClaimDay = vTable.DefaultView.ToTable.Rows(0)("ClaimDay").ToString
        End If
      End If
    End If
    Return vClaimDay
  End Function

  Protected Sub SetMemberBalance(ByVal pPaymentMethod As String, ByVal pMembershipType As String, ByVal pUserContactNumber As Integer, ByVal pStartDate As String)
    Dim vList As New ParameterList(HttpContext.Current)
    AddMemberParameters(vList, 1, 1, pPaymentMethod, pMembershipType, pStartDate)
    If vList.ContainsKey("PayerContactNumber") Then
      vList.Remove("PayerContactNumber")
      vList.Remove("PayerAddressNumber")
    End If
    If vList.Contains("BankAccount") AndAlso pPaymentMethod = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.pm_dd) Then
      Dim vClaimDay As String = GetClaimDay(vList("BankAccount").ToString)
      If vClaimDay.Length > 0 Then vList.Add("ClaimDay", vClaimDay)
    End If
    If ParentGroup.Length > 0 Then
      vList("PayerAddressNumber") = GetContactAddress(pUserContactNumber)
      vList("PayerContactNumber") = pUserContactNumber.ToString
    ElseIf UserContactNumber() > 0 Then
      vList("PayerContactNumber") = UserContactNumber()
      vList("PayerAddressNumber") = UserAddressNumber()
    End If
    vList("MemberFixedAmount") = 0
    vList = DataHelper.GetMemberBalance(vList)
    SetTextBoxText("Balance", CDbl(vList("MemberBalance")).ToString("#.00"))
  End Sub

  Protected Sub SetAmountOrBalance(ByVal pParameterName As String)
    SetAmountOrBalance(pParameterName, InitialParameters("Product").ToString, InitialParameters("Rate").ToString)
  End Sub

  Protected Sub SetAmountOrBalance(ByVal pParameterName As String, ByVal pProductCode As String, ByVal pRateCode As String)
    Dim vList As New ParameterList(HttpContext.Current)
    vList("Product") = pProductCode
    vList("Rate") = pRateCode
    Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtRates, vList))
    If vRow IsNot Nothing Then
      Dim vAmount As Double = CDbl(vRow("CurrentPrice"))
      If vAmount > 0 Then
        SetTextBoxText(pParameterName, vAmount.ToString("#.00"))
      End If
    End If
  End Sub

  Protected Sub SetCheckBoxChecked(ByVal pParameterName As String, Optional ByVal pCheck As Boolean = True)
    Dim vControl As Control = FindControlByName(mvTableControl, pParameterName)
    If vControl IsNot Nothing Then
      Dim vCheckBox As CheckBox = TryCast(vControl, CheckBox)
      If vCheckBox IsNot Nothing Then vCheckBox.Checked = pCheck
    End If
  End Sub

  Protected Sub SetRadioButtonChecked(ByVal pParameterName As String)
    Dim vControl As Control = FindControlByName(mvTableControl, pParameterName)
    If vControl IsNot Nothing Then
      Dim vRadioButton As RadioButton = TryCast(vControl, RadioButton)
      If vRadioButton IsNot Nothing Then vRadioButton.Checked = True
    End If
  End Sub

  Protected Sub SetDropDownText(ByVal pParameterName As String, ByVal pValue As String)
    SetDropDownText(pParameterName, pValue, False)
  End Sub
  Protected Sub SetDropDownText(ByVal pParameterName As String, ByVal pValue As String, ByVal pForceEvent As Boolean)
    Dim vControl As Control = FindControlByName(mvTableControl, pParameterName)
    If vControl IsNot Nothing Then
      Dim vDropDown As DropDownList = TryCast(vControl, DropDownList)
      If vDropDown IsNot Nothing AndAlso vDropDown.Items.FindByValue(pValue) IsNot Nothing Then
        vDropDown.SelectedValue = pValue
        If pValue.Length > 0 AndAlso vDropDown.SelectedValue.Length = 0 AndAlso vDropDown.Items.Count = 0 Then
          'BR13665: Setting a value of a combo when the combo has got no items
          vDropDown.Items.Add("")
          vDropDown.Items.Add(pValue)
          vDropDown.SelectedValue = pValue
        End If
        If pForceEvent Then DropDownListSelectedIndexChangedHandler(vDropDown, New System.EventArgs)
      End If
    End If
  End Sub

  Public Sub SetHiddenText(ByVal pParameterName As String, ByVal pValue As String)
    If Me.ViewState.Item(pParameterName) IsNot Nothing Then
      Me.ViewState.Item(pParameterName) = pValue
    End If
  End Sub

  Protected Sub SetTextBoxText(ByVal pParameterName As String, ByVal pValue As String)
    Dim vControl As Control = FindControlByName(mvTableControl, pParameterName)
    If vControl IsNot Nothing Then
      Dim vTextBox As TextBox = TryCast(vControl, TextBox)
      If vTextBox IsNot Nothing Then
        vTextBox.Text = pValue
      Else
        Dim vHiddenField As HiddenField = TryCast(vControl, HiddenField)
        If vHiddenField IsNot Nothing Then
          vHiddenField.Value = pValue
        End If
      End If
    End If
  End Sub

  Protected Sub SetParentParentVisible(ByVal pParameterName As String, ByVal pVisible As Boolean)
    Dim vControl As Control = FindControlByName(Me, pParameterName)
    If vControl IsNot Nothing Then
      vControl.Parent.Parent.Visible = pVisible
    End If
  End Sub

  Protected Sub SetParentVisible(ByVal pParameterName As String, ByVal pVisible As Boolean)
    Dim vControl As Control = FindControlByName(Me, pParameterName)
    If vControl IsNot Nothing Then
      vControl.Parent.Visible = pVisible
    End If
  End Sub

  Protected Sub SetLabelTextFromLabel(ByVal pParameterName As String, ByVal pControlName As String, ByVal pDefaultText As String)
    Dim vControl As Control = FindControlByName(Me, pControlName)
    Dim vText As String
    If vControl IsNot Nothing Then
      vText = DirectCast(vControl, ITextControl).Text
    Else
      vText = pDefaultText
    End If
    SetLabelText(pParameterName, vText)
  End Sub

  Protected Sub SetLabelText(ByVal pParameterName As String, ByVal pValue As String)
    Dim vControl As Control = FindControlByName(mvTableControl, pParameterName)
    If vControl IsNot Nothing Then
      Dim vLabel As Label = TryCast(vControl, Label)
      If vLabel IsNot Nothing Then vLabel.Text = pValue
    End If
  End Sub

  Protected Sub SetErrorLabel(ByVal pValue As String, Optional ByVal pParameterName As String = "PageError")
    Dim vDefault As String = ""
    If pValue.Length > 0 Then vDefault = "Error: "
    SetLabelText(pParameterName, vDefault & pValue)
  End Sub

  Protected Sub SetPafStatus(ByVal pValue As String)
    Dim vControl As Control = FindControlByName(mvTableControl, "PafStatus")
    If vControl IsNot Nothing Then
      Dim vLabel As Label = TryCast(vControl, Label)
      If vLabel IsNot Nothing Then
        vLabel.Text = pValue
        If mvControlType = CareNetServices.WebControlTypes.wctUpdateAddress Then ViewState("PafStatus") = pValue
      End If
    End If
  End Sub

#End Region

  Protected Sub ShowMessageOnlyFromLabel(ByVal pControlName As String, ByVal pDefaultMessage As String)
    Dim vControl As Control = FindControlByName(mvTableControl, pControlName)
    Dim vMessage As String
    If vControl IsNot Nothing Then
      vMessage = DirectCast(vControl, ITextControl).Text
    Else
      vMessage = pDefaultMessage
    End If
    ShowMessageOnly(vMessage, vControl)
  End Sub

  Protected Sub ShowMessageOnly(ByVal pMessageLabel As Control)
    ShowMessageOnly("", pMessageLabel)
  End Sub
  Protected Sub ShowMessageOnly(ByVal pMessage As String)
    ShowMessageOnly(pMessage, Nothing)
  End Sub
  Private Sub ShowMessageOnly(ByVal pMessage As String, ByVal pMessageLabel As Control)
    mvTableControl.Rows.Clear()
    Dim vHTMLRow As New HtmlTableRow
    Dim vHTMLCell As New HtmlTableCell
    vHTMLCell.Attributes("Class") = "DataMessage"
    vHTMLCell.ColSpan = 2
    Dim vLabel As Label = TryCast(pMessageLabel, Label)
    If vLabel Is Nothing Then
      vLabel = New Label
      vLabel.ID = "Message"
      vLabel.CssClass = "DataMessage"
      vLabel.Text = pMessage
    End If
    vLabel.Visible = True
    vHTMLCell.Controls.Add(vLabel)    'Add it to the cell
    vHTMLRow.Cells.Add(vHTMLCell)
    mvTableControl.Rows.Add(vHTMLRow)
  End Sub

  Protected Function FindControlByName(ByVal pControl As Control, ByVal pName As String) As Control
    Dim vControl As Control = Nothing
    For Each vControl In pControl.Controls
      If vControl.ID = pName Then
        Return vControl
      ElseIf vControl.Controls.Count > 0 Then
        Dim vFoundControl As Control = FindControlByName(vControl, pName)
        If vFoundControl IsNot Nothing Then Return vFoundControl
      End If
    Next
    Return Nothing
  End Function

  Protected Function AddGiftAidDeclaration(ByVal pContactNumber As Integer) As ParameterList
    If GetCheckBoxChecked("GiftAid") Then
      Try
        Dim vGADList As New ParameterList(HttpContext.Current)
        vGADList("UserID") = pContactNumber
        vGADList("ContactNumber") = pContactNumber
        vGADList("Source") = DefaultParameters("Source")
        vGADList("DeclarationType") = "A"
        Return DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctGiftAidDeclarations, vGADList)
      Catch vEx As CareException
        If vEx.ErrorNumber <> CareException.ErrorNumbers.enGiftAidDeclarationsOverlap Then Throw vEx
      End Try
    End If
    Return Nothing
  End Function

  Protected Function AddNewContact(Optional ByVal pUseHiddenNumbers As Boolean = False) As ParameterList
    'Create the contact
    If mvContactEntryHidden Then
      Dim vReturnList As New ParameterList(HttpContext.Current)
      If InitialParameters.ContainsKey("MembershipFor") AndAlso InitialParameters("MembershipFor").ToString.ToUpper = "O" AndAlso mvParentGroup.Length = 0 Then
        Dim vOrganisationDDL As DropDownList = TryCast(FindControlByName(Me, "Organisation"), DropDownList)
        If vOrganisationDDL IsNot Nothing Then
          vReturnList("ContactNumber") = vOrganisationDDL.SelectedValue.Split(CChar(","))(0)
          vReturnList("AddressNumber") = vOrganisationDDL.SelectedValue.Split(CChar(","))(1)
        End If
      Else
        If mvParentGroup.Length > 0 AndAlso mvControlType = CareNetServices.WebControlTypes.wctAddMemberDD OrElse _
          mvControlType = CareNetServices.WebControlTypes.wctAddMemberCC OrElse _
          mvControlType = CareNetServices.WebControlTypes.wctAddMemberCS OrElse _
          mvControlType = CareNetServices.WebControlTypes.wctAddMemberCI Then
          vReturnList("ContactNumber") = GetContactNumberFromParentGroup()
          vReturnList("AddressNumber") = GetContactAddress(GetContactNumberFromParentGroup)
        Else
          vReturnList("ContactNumber") = UserContactNumber()
          AddUserParameters(vReturnList)
          vReturnList("AddressNumber") = UserAddressNumber()
        End If
      End If
      If ServiceType = AuthorisationService.SagePayHosted Then
        Cache("ContactNumber") = vReturnList("ContactNumber")
        Cache("AddressNumber") = vReturnList("AddressNumber")
      End If
      Return vReturnList
    Else
      Dim vContactList As New ParameterList
      Dim vReturnList As New ParameterList
      If pUseHiddenNumbers Then
        vReturnList("ContactNumber") = GetHiddenContactNumber()
        AddUserParameters(vReturnList)
        vReturnList("AddressNumber") = GetHiddenAddressNumber()
      Else
        vContactList = GetAddContactParameterList()
        'AddUserParameters(vContactList)
        vReturnList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctContact, vContactList)
      End If
      Session("ContactNumber") = vReturnList("ContactNumber")
      Session("AddressNumber") = vReturnList("AddressNumber")
      If ServiceType = AuthorisationService.SagePayHosted Then
        Cache("ContactNumber") = vReturnList("ContactNumber")
        Cache("AddressNumber") = vReturnList("AddressNumber")
      End If
      Return vReturnList
    End If
  End Function

  Protected Function GetContactAddress(ByVal pCotactNumber As Integer) As String
    Dim vAddress As String = ""
    Dim vParams As New ParameterList(HttpContext.Current)
    vParams("ContactNumber") = pCotactNumber
    Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vParams))
    If vRow IsNot Nothing Then
      vAddress = vRow("AddressNumber").ToString
    End If
    Return vAddress
  End Function

  Protected Function GetAddContactParameterList() As ParameterList
    Return GetAddContactParameterList(True)
  End Function
  Protected Function GetAddContactParameterList(ByVal pGetCommNumbers As Boolean) As ParameterList
    Dim vUpdate As Boolean = GetHiddenContactNumber() > 0
    Dim vContactList As New ParameterList(HttpContext.Current)
    vContactList("CarePortal") = "Y"
    'NFPCARE-522 : Over typing issues
    If Not DataHelper.ConfigurationOption(DataHelper.ConfigurationOptions.use_ajax_for_contact_names, False) Then
      Dim vTriggerField As Control = Nothing
      If ValueChanged("HiddenOldForename", "Forenames") Then 'Trigger field has changed
        vTriggerField = FindControlByName(mvTableControl, "Forenames")
        If vTriggerField IsNot Nothing Then
          'if update field values have not changed then update them using existing rules
          If mvControlType = CareNetServices.WebControlTypes.wctUpdateContact AndAlso (Not ValueChanged("HiddenInitials", "Initials")) Then UpdateInitials(True, vTriggerField.ID, GetTextBoxText(vTriggerField.ID))
          If Not ValueChanged("HiddenSalutation", "Salutation") Then UpdateSalutation(vTriggerField)
          If Not ValueChanged("HiddenLabelName", "LabelName") Then UpdateLabelName(vTriggerField)
          If Not ValueChanged("HiddenPreferredForename", "PreferredForename") Then UpdatePreferredForename("PreferredForename", GetTextBoxText("Forenames"))
        End If
      End If

      If ValueChanged("HiddenSurname", "Surname") Then
        vTriggerField = FindControlByName(mvTableControl, "Surname")
        If Not ValueChanged("HiddenSalutation", "Salutation") Then UpdateSalutation(vTriggerField)
        If Not ValueChanged("HiddenLabelName", "LabelName") Then UpdateLabelName(vTriggerField)
      End If

      If ValueChanged("HiddenTitle", "Title") Then
        vTriggerField = FindControlByName(mvTableControl, "Title")
        If Not ValueChanged("HiddenSalutation", "Salutation") Then UpdateSalutation(vTriggerField)
        If Not ValueChanged("HiddenLabelName", "LabelName") Then UpdateLabelName(vTriggerField)
      End If

      If ValueChanged("HiddenSex", "Sex") Then
        vTriggerField = FindControlByName(mvTableControl, "Sex")
        If Not ValueChanged("HiddenSalutation", "Salutation") Then UpdateSalutation(vTriggerField)
      End If

      If ValueChanged("HiddenInitials", "Initials") Then
        vTriggerField = FindControlByName(mvTableControl, "Initials")
        If Not ValueChanged("HiddenLabelName", "LabelName") Then UpdateLabelName(vTriggerField)
      End If

      If ValueChanged("HiddenHonorifics", "Honorifics") Then
        vTriggerField = FindControlByName(mvTableControl, "Honorifics")
        If Not ValueChanged("HiddenLabelName", "LabelName") Then UpdateLabelName(vTriggerField)
      End If
    End If

    Dim vTitle As String = GetDropDownValue("Title")
    If vTitle.Length > 0 Then vContactList("Title") = vTitle
    If FindControlByName(mvTableControl, "LabelNameFormatCode") IsNot Nothing Then vContactList("LabelNameFormatCode") = GetDropDownValue("LabelNameFormatCode")
    AddOptionalTextBoxValue(vContactList, "Forenames", True)
    AddOptionalTextBoxValue(vContactList, "Surname")
    AddOptionalTextBoxValue(vContactList, "PreferredForename", True)
    AddOptionalTextBoxValue(vContactList, "Salutation", True)
    If mvControlType = CareNetServices.WebControlTypes.wctUpdateContact Then
      AddOptionalTextBoxValue(vContactList, "Initials")
      AddOptionalTextBoxValue(vContactList, "Honorifics")
      AddOptionalTextBoxValue(vContactList, "LabelName")
      AddOptionalTextBoxValue(vContactList, "NiNumber")
    Else
      AddOptionalTextBoxValue(vContactList, "Address")
      AddOptionalTextBoxValue(vContactList, "Town")
      If vContactList.Contains("Address") AndAlso Not vContactList.Contains("Town") Then vContactList.Add("Town", "#")
      AddOptionalTextBoxValue(vContactList, "County", True)
      If FindControlByName(mvTableControl, "Postcode") IsNot Nothing Then
        AddOptionalTextBoxValue(vContactList, "Postcode", True)
      Else
        Dim vValue As String = GetTextBoxText("PostcoderPostcode")
        If vValue.Length > 0 Then vContactList("Postcode") = vValue
      End If
      Dim vPAFStatus As String = GetLabelText("PafStatus")
      If vPAFStatus.Length > 0 Then vContactList("PafStatus") = vPAFStatus
      vContactList("Source") = DefaultParameters("Source")
      Dim vCountry As String = GetDropDownValue("Country")
      If vCountry.Length = 0 AndAlso vUpdate = False Then vCountry = "UK" 'Only add default country when adding a new contact
      If vCountry.Length > 0 Then vContactList("Country") = vCountry

      If pGetCommNumbers Then
        AddOptionalTextBoxValue(vContactList, "EMailAddress", True)
        AddOptionalTextBoxValue(vContactList, "DirectNumber", True)
        AddOptionalTextBoxValue(vContactList, "MobileNumber", True)
      End If
    End If

    AddOptionalTextBoxValue(vContactList, "DateOfBirth", True)
    Dim vSex As String = GetDropDownValue("Sex")
    If vSex.Length > 0 Then vContactList("Sex") = vSex

    AddOptionalTextBoxValue(vContactList, "PositionLocation", True)
    If FindControlByName(mvTableControl, "Position") IsNot Nothing Then
      AddOptionalTextBoxValue(vContactList, "Position", True)
      AddOptionalTextBoxValue(vContactList, "Name", True)
    End If


    If FindControlByName(mvTableControl, "Status") IsNot Nothing AndAlso GetDropDownValue("Status").Length > 0 Then
      vContactList("Status") = GetDropDownValue("Status")
      AddOptionalTextBoxValue(vContactList, "StatusDate", True)
      AddOptionalTextBoxValue(vContactList, "StatusReason", True)
    End If

    If vUpdate Then
      If Not vContactList.Contains("Salutation") Then
        Dim vSalutation As String = GetHiddenText("HiddenSalutation")
        If vSalutation.Length = 0 Then vSalutation = vContactList("Surname").ToString
        vContactList.Add("Salutation", vSalutation)
      Else
        If vContactList("Salutation").ToString.Length = 0 Then vContactList("Salutation") = vContactList("Surname").ToString
      End If
      If Not vContactList.Contains("PreferredForename") Then vContactList.Add("PreferredForename", GetHiddenText("HiddenPreferredForename"))
      'As this is an existing Contact and Source is not a field the user can change/set, we should not be updating it
      If vContactList.ContainsKey("Source") Then vContactList.Remove("Source")
    End If
    Return vContactList
  End Function
  Protected Function GetAddOrganisationParameterList() As ParameterList
    Dim vOrganisationList As New ParameterList(HttpContext.Current)
    vOrganisationList("ContactNumber") = UserContactNumber().ToString
    If DefaultParameters.ContainsKey("Source") Then vOrganisationList("Source") = DefaultParameters("Source").ToString
    AddOptionalTextBoxValue(vOrganisationList, "Name")
    AddOptionalTextBoxValue(vOrganisationList, "FaxNumber")
    AddOptionalTextBoxValue(vOrganisationList, "SwitchboardNumber")
    AddOptionalTextBoxValue(vOrganisationList, "WebAddress")
    AddOptionalTextBoxValue(vOrganisationList, "PostcoderPostcode")
    AddOptionalTextBoxValue(vOrganisationList, "PostcoderAddress")
    AddOptionalTextBoxValue(vOrganisationList, "Address")
    AddOptionalTextBoxValue(vOrganisationList, "Town")
    If vOrganisationList.Contains("Address") AndAlso Not vOrganisationList.Contains("Town") Then vOrganisationList.Add("Town", "#")
    AddOptionalTextBoxValue(vOrganisationList, "County")
    AddOptionalTextBoxValue(vOrganisationList, "BuildingNumber")
    AddOptionalTextBoxValue(vOrganisationList, "HouseName")

    If FindControlByName(mvTableControl, "Postcode") IsNot Nothing Then
      AddOptionalTextBoxValue(vOrganisationList, "Postcode", True)
    Else
      Dim vValue As String = GetTextBoxText("PostcoderPostcode")
      If vValue.Length > 0 Then vOrganisationList("Postcode") = vValue
    End If

    AddOptionalTextBoxValue(vOrganisationList, "EMailAddress")
    vOrganisationList("Country") = GetDropDownValue("Country")

    Dim vPAFStatus As String = GetLabelText("PafStatus")
    If vPAFStatus.Length > 0 Then vOrganisationList("PafStatus") = vPAFStatus

    Return vOrganisationList
  End Function
  Protected Function GetContactFromEMailAddress(ByVal pEMailAddress As String) As ParameterList
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vContactList As New ParameterList(HttpContext.Current)
    vList("EMailAddress") = pEMailAddress
    Dim vTable As DataTable = DataHelper.FindDataTable(CareNetServices.XMLDataFinderTypes.xdftContacts, vList)
    If vTable IsNot Nothing AndAlso vTable.Rows.Count = 1 Then
      'We have found the contact
      vContactList("ContactNumber") = vTable.Rows(0)("ContactNumber").ToString
      vContactList("AddressNumber") = vTable.Rows(0)("AddressNumber").ToString
    End If
    Return vContactList
  End Function

  Protected Sub SetAuthentication(ByVal pList As ParameterList)
    SetAuthentication(pList, False)
  End Sub
  Protected Sub SetAuthentication(ByVal pList As ParameterList, ByVal pCheckSingleSignOn As Boolean)
    Dim vAccessViewNames As String = pList.OptionalValue("AccessViewName")
    Dim vUserData As String = String.Format(DataHelper.Database & "|{0}|{1}|{2}|{3}", pList("ContactNumber"), pList("AddressNumber"), pList("UserDepartment"), vAccessViewNames)
    Dim vUserName As String
    If pList("UserLogname") IsNot Nothing Then
      vUserName = pList("UserLogname").ToString
    Else
      vUserName = "guest"
    End If

    'The number of minutes added on here is the time with no activity after which the ticket will be invaldi
    Dim vTicket As New FormsAuthenticationTicket(1, vUserName, Now, Now.AddMinutes(Session.Timeout), False, vUserData)
    Dim vCookie As New HttpCookie(FormsAuthentication.FormsCookieName)
    vCookie.Value = FormsAuthentication.Encrypt(vTicket)
    Response.Cookies.Add(vCookie)
    If pCheckSingleSignOn Then
      If pList("UserLogname") Is Nothing Then pList("UserLogname") = vUserName 'Just in case
      ProcessSingleSignOn(pList)
    End If
  End Sub

  Protected Sub ProcessSingleSignOn(ByVal pList As ParameterList)
    'Single Sign-On
    Dim vSingleSignOnURL As String = GetCustomConfigItem("CustomConfiguration/SingleSignOnURL", True)
    If vSingleSignOnURL.Length > 0 Then
      Dim vSingleSignOnKey As String = GetCustomConfigItem("CustomConfiguration/SingleSignOnKey")
      If vSingleSignOnKey.Length > 0 Then
        Dim vData As String = If(pList.Contains("ContactNumber"), pList("ContactNumber").ToString, If(pList.Contains("UserLogname"), pList("UserLogname").ToString, "guest"))
        Dim vMD5 As System.Security.Cryptography.MD5
        vMD5 = System.Security.Cryptography.MD5.Create()
        vMD5.ComputeHash(Encoding.UTF8.GetBytes(vSingleSignOnKey & vData))
        Dim vHashValue As New StringBuilder
        For vIndex As Integer = 0 To vMD5.Hash.Length - 1
          vHashValue.Append(vMD5.Hash(vIndex).ToString("x2")) 'Create 32 characters hexadecimal-formatted hash string
        Next
        vMD5.Clear()
        Dim vReturnURL As String = ""
        If Session("UpdateDetailsURL") IsNot Nothing Then vReturnURL = Session("UpdateDetailsURL").ToString
        If vReturnURL.Length = 0 Then vReturnURL = Request.Params("ReturnURL")
        If vReturnURL Is Nothing OrElse vReturnURL.Length = 0 Then
          If Request.QueryString("ReturnURL") IsNot Nothing Then
            vReturnURL = Request.QueryString("ReturnUrl").ToString
          ElseIf SubmitItemUrl.Length > 0 Then
            vReturnURL = SubmitItemUrl
          ElseIf SubmitItemNumber > 0 Then
            vReturnURL = String.Format("{0}?pn={1}", If(String.IsNullOrEmpty(Request.Url.Query), Request.Url.AbsoluteUri, Request.Url.AbsoluteUri.Replace(Request.Url.Query, "")), SubmitItemNumber)
          End If
        End If
        'BR18438 
        'Debug.WriteLine("PRE BR18438 fix: ProcessSingleSignOn vReturnURL: " & vReturnURL)
        vReturnURL = Server.UrlEncode(vReturnURL)
        'BR18438
        'Debug.WriteLine("POST BR18438 fix: ProcessSingleSignOn vReturnURL: " & vReturnURL)
        RedirectViaWhiteList(String.Format(vSingleSignOnURL, vData, vHashValue.ToString, vReturnURL))
      End If
    End If
  End Sub

  Protected Function IsValid() As Boolean
    If Me.Page.IsValid Then
      If Not InWebPageDesigner() Then Return True
    End If
  End Function

  Protected Function InWebPageDesigner() As Boolean
    If Request.QueryString("cwpd") = "Y" Then Return True
  End Function

  Public Function UserOrNewContactNumber() As Integer
    If mvUseNewContact Then
      Return IntegerValue(Session("ContactNumber").ToString)
    Else
      Return UserContactNumber()
    End If
  End Function

  Public Function UserOrNewAddressNumber() As Integer
    If mvUseNewContact Then
      Return IntegerValue(Session("AddressNumber").ToString)
    Else
      Return UserAddressNumber()
    End If
  End Function

  Public Function UserContactNumber() As Integer
    If HttpContext.Current.User.Identity.IsAuthenticated Then
      If TypeOf (Page.User.Identity) Is System.Security.Principal.WindowsIdentity AndAlso Session("UserContactNumber") IsNot Nothing Then
        Return IntegerValue(Session("UserContactNumber").ToString)
      Else
        Dim vIdentity As FormsIdentity = CType(HttpContext.Current.User.Identity, FormsIdentity)
        If vIdentity.Ticket.UserData.Length > 0 Then
          Dim vString As String = vIdentity.Ticket.UserData
          Dim vItems As String() = vString.Split("|"c)
          If vItems.Length > 1 Then UserContactNumber = IntegerValue(vItems(1))
        End If
      End If
    Else
      If Request.QueryString("ucn") IsNot Nothing AndAlso Request.QueryString("ucn").Length > 0 Then Return CInt(Request.QueryString("ucn"))
    End If
  End Function

  Public Function UserAddressNumber() As Integer
    If HttpContext.Current.User.Identity.IsAuthenticated Then
      If TypeOf (Page.User.Identity) Is System.Security.Principal.WindowsIdentity AndAlso Session("UserAddressNumber") IsNot Nothing Then
        Return IntegerValue(Session("UserAddressNumber").ToString)
      Else
        Dim vIdentity As FormsIdentity = CType(HttpContext.Current.User.Identity, FormsIdentity)
        If vIdentity.Ticket.UserData.Length > 0 Then
          Dim vString As String = vIdentity.Ticket.UserData
          Dim vItems As String() = vString.Split("|"c)
          If vItems.Length > 2 Then UserAddressNumber = IntegerValue(vItems(2))
        End If
      End If
    Else
      If Request.QueryString("uan") IsNot Nothing AndAlso Request.QueryString("uan").Length > 0 Then Return CInt(Request.QueryString("uan"))
    End If
  End Function

  Protected Sub SelectListItem(ByVal pList As DropDownList, ByVal pValue As String)
    pList.SelectedIndex = -1
    Dim vItem As ListItem
    For Each vItem In pList.Items
      If vItem.Value = pValue Then
        vItem.Selected = True
        Exit For
      End If
    Next
  End Sub

  Protected Sub SetControlVisible(ByVal pParameterName As String, ByVal pVisible As Boolean)
    Dim vControl As Control = FindControlByName(mvTableControl, pParameterName)
    If vControl IsNot Nothing Then vControl.Visible = pVisible
  End Sub

  Private Function IsControlVisible(ByVal pParameterName As String) As Boolean
    Dim vVisible As Boolean = False
    Dim vControl As Control = FindControlByName(mvTableControl, pParameterName)
    If vControl IsNot Nothing Then vVisible = vControl.Visible
    Return vVisible
  End Function

  Private Function CheckViewStateReadonly(ByVal pParameterName As String) As Boolean
    Dim vViewStateParameterName As String = pParameterName & "~ReadOnly"
    Dim vIsReadonly As Boolean = False
    If ViewState(vViewStateParameterName) IsNot Nothing AndAlso ViewState(vViewStateParameterName).ToString.Length > 0 Then
      vIsReadonly = Convert.ToBoolean(ViewState(vViewStateParameterName))
    End If
    Return vIsReadonly
  End Function

  Protected Sub SetControlEnabled(ByVal pParameterName As String, ByVal pEnabled As Boolean)
    Dim vControl As Control = FindControlByName(mvTableControl, pParameterName)
    If vControl IsNot Nothing Then
      If TypeOf vControl Is CheckBox Then
        If Not CheckViewStateReadonly(pParameterName) Then
          DirectCast(vControl, CheckBox).Enabled = pEnabled
        End If
      ElseIf TypeOf vControl Is DropDownList Then
        If Not CheckViewStateReadonly(pParameterName) Then
          DirectCast(vControl, DropDownList).Enabled = pEnabled
          If pEnabled Then
            If DirectCast(vControl, DropDownList).CssClass = "ReadOnly" Then
              DirectCast(vControl, DropDownList).CssClass = "DataEntryItem"
            End If
          Else
            DirectCast(vControl, DropDownList).CssClass = "ReadOnly"
          End If
        End If
      ElseIf TypeOf vControl Is TextBox Then
        If Not CheckViewStateReadonly(pParameterName) Then
          If Not pEnabled Then
            DirectCast(vControl, TextBox).Attributes.Add("readonly", "readonly") ' .ReadOnly = Not pEnabled
          Else
            DirectCast(vControl, TextBox).Attributes.Remove("readonly") 'Make sure to remove this attribute when enabling a control
          End If
          If pEnabled Then
            If DirectCast(vControl, TextBox).CssClass = "ReadOnly" Then
              If CType(DirectCast(vControl, TextBox).FindControl("rfv" & pParameterName), RequiredFieldValidator) IsNot Nothing Then
                DirectCast(vControl, TextBox).CssClass = "DataEntryItemMandatory"
              Else
                DirectCast(vControl, TextBox).CssClass = "DataEntryItem"
              End If
            End If
          Else
            DirectCast(vControl, TextBox).CssClass = "ReadOnly"
          End If
        End If
      ElseIf TypeOf vControl Is Button Then
        If Not CheckViewStateReadonly(pParameterName) Then
          DirectCast(vControl, Button).Enabled = pEnabled
        End If
      ElseIf TypeOf vControl Is HtmlInputButton Then
        If Not CheckViewStateReadonly(pParameterName) Then
          DirectCast(vControl, HtmlInputButton).Disabled = Not pEnabled
        End If
      ElseIf TypeOf vControl Is HyperLink Then
        If Not CheckViewStateReadonly(pParameterName) Then
          DirectCast(vControl, HyperLink).Enabled = pEnabled
          If pEnabled Then
            If DirectCast(vControl, HyperLink).CssClass = "ReadOnly" Then
              DirectCast(vControl, HyperLink).CssClass = "DataEntryItem"
            End If
          Else
            DirectCast(vControl, TextBox).CssClass = "ReadOnly"
          End If
        End If
      End If
    End If
  End Sub

  Public Sub SendEmail(ByVal pFromAddress As String, ByVal pToAddress As String, ByVal pSubject As String, ByVal pMessageBody As String)
    If pFromAddress.Length > 0 AndAlso pToAddress.Length > 0 Then
      Dim vMessage As New System.Net.Mail.MailMessage(pFromAddress, pToAddress)
      vMessage.Subject = pSubject
      vMessage.BodyEncoding = New System.Text.ASCIIEncoding
      vMessage.Body = pMessageBody
      Dim vSMTPServer As String = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.email_smtp_server)
      If vSMTPServer.Length = 0 Then vSMTPServer = "localhost"
      Dim vClient As System.Net.Mail.SmtpClient = New System.Net.Mail.SmtpClient(vSMTPServer)
      vClient.Send(vMessage)
    End If
  End Sub

  Private Function IsTableItemVisible(ByVal pTable As DataTable, ByVal pParameterName As String) As Boolean
    Dim vVisible As Boolean
    For Each vRow As DataRow In pTable.Rows
      Dim vParameterName As String = vRow.Item("ParameterName").ToString
      If vParameterName = pParameterName Then
        vVisible = vRow.Item("Visible").ToString = "Y"
        Exit For
      End If
    Next
    Return vVisible
  End Function

  Protected Function ValidateContactStatus() As Boolean
    Dim vValid As Boolean = True
    If GetDropDownValue("Status").Length > 0 Then
      Dim vDDL As DropDownList = TryCast(FindControlByName(Me, "Status"), DropDownList)
      If vDDL IsNot Nothing Then
        Dim vDT As DataTable = TryCast(vDDL.DataSource, DataTable)
        If vDT IsNot Nothing Then
          If vDT.Rows(vDDL.SelectedIndex).Item("ReasonRequired").ToString = "Y" Then
            vValid = GetTextBoxText("StatusReason").Length > 0
          End If
        End If
      End If
    End If
    Return vValid
  End Function

  Protected Sub SaveContactCommNumbers(ByVal pNumbers As NumberInfo(), ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer)
    SaveContactCommNumbers(pNumbers, pContactNumber, pAddressNumber, False, True)
  End Sub
  Protected Sub SaveContactCommNumbers(ByVal pNumbers As NumberInfo(), ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pAdditionalNumbersOnly As Boolean, ByVal pDeleteExisting As Boolean)
    Dim vNewNumber As String = ""
    Dim vContinue As Boolean
    Dim vSetHistoric As Boolean = DefaultParameters.Contains("SetHistoric") AndAlso BooleanValue(DefaultParameters("SetHistoric").ToString)
    For Each vNumber As NumberInfo In pNumbers
      If vNumber IsNot Nothing Then
        vContinue = vNumber.DeviceCode.Length > 0
        If pAdditionalNumbersOnly Then vContinue = vNumber.Identifier.StartsWith("AdditionalNumber") AndAlso vContinue
        If vContinue Then
          vNewNumber = GetTextBoxText(vNumber.Identifier)
          Dim vList As New ParameterList(HttpContext.Current)
          vList("ContactNumber") = pContactNumber
          vList("AddressNumber") = pAddressNumber
          vList("Device") = vNumber.DeviceCode
          vList("DiallingCode") = ""
          vList("STDCode") = ""

          Dim vDataChanged As Boolean = False
          Dim vAddNew As Boolean = False
          If vNumber.CommunicationNumber = 0 AndAlso vNewNumber.Length > 0 Then
            'Add New Record
            vAddNew = True
          ElseIf vNumber.CommunicationNumber > 0 AndAlso vNewNumber.Length = 0 Then
            'The record is removed
            vDataChanged = True
          ElseIf vNumber.Number = vNewNumber Then
            'Nothing is changed. Don't add/delete anything
          Else
            'Data is changed. Set the existing as historic and add new record when SetHistoric is set otherwise either update or delete the record
            vDataChanged = True
            vAddNew = vSetHistoric
          End If
          If vDataChanged Then
            vList("CommunicationNumber") = vNumber.CommunicationNumber
            If vNewNumber.Length > 0 OrElse vAddNew Then
              vList("OldContactNumber") = pContactNumber
              If pAdditionalNumbersOnly Then vList.Remove("AddressNumber")
              If vAddNew Then
                'Where SetHistoric set update comm to set as historic
                vList("Number") = vNumber.Number
                vList("ValidTo") = TodaysDate()
                vList("AmendedOn") = TodaysDate()
              Else
                vList("Number") = vNewNumber
              End If
              DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctNumber, vList)
            ElseIf pDeleteExisting Then
              DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctNumber, vList)
            End If
          End If
          If vAddNew Then
            If vDataChanged Then
              'Where previous comm is set to historic create new comm copying flags
              vList("ValidFrom") = TodaysDate()
              If vList.Contains("ValidTo") Then vList.Remove("ValidTo")
              If vList.Contains("CommunicationNumber") Then vList.Remove("CommunicationNumber")
              If vList.Contains("OldContactNumber") Then vList.Remove("OldContactNumber")
              If vNumber.DeviceDefault Then vList("DeviceDefault") = "Y"
              If vNumber.IsDefault Then vList("Default") = "Y"
              If vNumber.Mail Then vList("Mail") = "Y"
              If vNumber.PreferredMethod Then vList("PreferredMethod") = "Y"
            End If
            vList("Number") = vNewNumber
            DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctNumber, vList)
          End If
        End If
      End If
    Next
  End Sub

  Protected Function ValueChanged(ByVal pHiddenParameterName As String, ByVal pParameterName As String) As Boolean
    'Used to compare a control's value with a hidden control's value. Returns FALSE if the control is not displayed
    Dim vControl As Control = FindControlByName(Me, pParameterName)
    If vControl IsNot Nothing Then
      Dim vTextBox As TextBox = TryCast(vControl, TextBox)
      If vTextBox IsNot Nothing Then Return GetHiddenText(pHiddenParameterName) <> vTextBox.Text
      Dim vDropDown As DropDownList = TryCast(vControl, DropDownList)
      If vDropDown IsNot Nothing Then Return GetHiddenText(pHiddenParameterName) <> vDropDown.SelectedItem.Value
    End If
    Return False
  End Function

  Protected Function CapitalisationChangedOnly(ByVal pHiddenParameterName As String, ByVal pParameterName As String) As Boolean
    'checks if only capitals have changed
    Dim vControl As Control = FindControlByName(Me, pParameterName)
    If vControl IsNot Nothing Then
      Dim vTextBox As TextBox = TryCast(vControl, TextBox)
      If vTextBox IsNot Nothing Then
        Dim vOldText As String = GetHiddenText(pHiddenParameterName)
        Dim vNewText As String = vTextBox.Text
        If vOldText.Length > 0 AndAlso (String.Compare(vOldText, vNewText, True) = 0) Then
          Return True
        Else
          Return False
        End If
      End If
    End If
  End Function

  Protected Sub ShowActivityOrSuppression(ByVal pDisplayActivity As Boolean)
    If DefaultParameters.ContainsKey("ShowActivitiesSuppressionsThatExists") AndAlso DefaultParameters("ShowActivitiesSuppressionsThatExists").ToString.Length > 0 Then
      If DefaultParameters("ShowActivitiesSuppressionsThatExists").ToString = "Y" Then
        If Not pDisplayActivity Then
          mvTableControl.Rows.Clear()
        End If
      End If
    End If
  End Sub

  Protected Function GetContactNumberFromParentGroup(Optional ByVal pParentGroup As String = "") As Integer
    Try
      If InWebPageDesigner() Then
        Return UserContactNumber()
      Else
        Dim vParentGroup As String = mvParentGroup
        If pParentGroup.Length > 0 Then
          vParentGroup = pParentGroup
        End If
        If vParentGroup = "SelectedContact" Then
          If Session("SelectedContactNumber") IsNot Nothing AndAlso Session("SelectedContactNumber").ToString.Length > 0 Then
            Return IntegerValue(Session("SelectedContactNumber").ToString)
          Else
            RaiseError(DataAccessErrors.daeSessionValueNotSet, "SelectedContact")
          End If
        ElseIf vParentGroup = "SelectedOrganisation" Then
          If Session("SelectedOrganisationNumber") IsNot Nothing AndAlso Session("SelectedOrganisationNumber").ToString.Length > 0 Then
            Return IntegerValue(Session("SelectedOrganisationNumber").ToString)
          Else
            RaiseError(DataAccessErrors.daeSessionValueNotSet, "SelectedOrganisationNumber")
          End If
        Else
          Return UserContactNumber()
        End If
      End If
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Function

  Protected Function GetShoppingBasketTransaction(ByVal pContactNumber As Integer, ByVal pList As ParameterList) As Boolean
    Dim vList As New ParameterList(HttpContext.Current)
    If pContactNumber = 0 Then
      vList("ContactNumber") = GetContactNumberFromParentGroup()
    Else
      vList("ContactNumber") = pContactNumber
    End If
    vList("FindTransactionType") = "V"
    vList("BatchType") = "CA"
    Dim vTable As DataTable = Nothing
    Dim vRow As DataRow = Nothing
    vTable = DataHelper.FindDataTable(CareNetServices.XMLDataFinderTypes.xdftTransactions, vList)
    If vTable IsNot Nothing Then
      If vTable.Rows.Count > 0 Then
        Dim vMaxBatchNumber As Integer = 0
        For vRowIndex As Integer = 0 To vTable.Rows.Count - 1
          If IntegerValue(vTable.Rows(vRowIndex).Item("BatchNumber").ToString) > vMaxBatchNumber Then
            vMaxBatchNumber = IntegerValue(vTable.Rows(vRowIndex).Item("BatchNumber").ToString)
            vRow = vTable.Rows(vRowIndex)
          End If
        Next
      End If
    End If
    If vRow IsNot Nothing Then
      pList("BatchNumber") = vRow.Item("BatchNumber")
      pList("TransactionNumber") = vRow.Item("TransactionNumber")
      Return True
    Else
      Return False
    End If
  End Function


  Private Function GetMembershipStartDate() As List(Of Date)
    Dim vStartDate As New List(Of Date)
    Dim vStortedDate As SortedSet(Of Date) = New SortedSet(Of Date)
    Dim vPreviousDate As Date = DateAdd(DateInterval.Month, -11, Date.Now.Date)
    Dim vFutureDate As Date = DateAdd(DateInterval.Month, 11, Date.Now.Date)

    Do While DateDiff(DateInterval.Month, vPreviousDate, vFutureDate) >= 0
      vStortedDate.Add(vPreviousDate)
      vPreviousDate = DateAdd(DateInterval.Month, 1, vPreviousDate)
    Loop

    For Each vDate As Date In vStortedDate
      Dim vParam As New ParameterList(HttpContext.Current)
      vParam.Add("Date", vDate)
      Dim vResult As ParameterList = DataHelper.GetPaymentPlanStartDate(vParam)
      For Each vKey As String In vResult.Keys
        If vResult(vKey).ToString.Length > 0 Then
          If Not (vStartDate.Contains(CDate(vResult(vKey)))) AndAlso (DateDiff(DateInterval.Month, CDate(vResult(vKey)), Date.Now.Date) <= 11) AndAlso (DateDiff(DateInterval.Month, CDate(vResult(vKey)), Date.Now.Date) >= -11) Then
            vStartDate.Add(CDate(vResult(vKey)))
          End If
        End If
      Next
    Next
    If vStartDate.Count = 0 Then vStartDate.Add(CDate(Date.Now.ToShortDateString))
    Return vStartDate
  End Function
  ''' <summary>
  ''' Get the nearest date from the List 
  ''' </summary>
  ''' <param name="pDates">Sorted list of dates </param>
  ''' <returns>nearest date</returns>
  ''' <remarks></remarks>
  Private Function GetNearestDate(ByVal pDates As List(Of Date)) As Date
    Dim vPreviousDate As Date = Nothing
    Dim vNextDate As Date = Nothing
    Dim vCurrentDate As Date = CDate(DateTime.Now.ToShortDateString)
    Dim vReturnDate As Date = Nothing

    For Each vDate As Date In pDates
      Select Case Date.Compare(vCurrentDate, vDate)
        Case Is < 0
          vNextDate = vDate
          Exit For
        Case 0
          Return vDate
        Case Is > 0
          vPreviousDate = vDate
      End Select
    Next
    If vNextDate <> Nothing AndAlso vPreviousDate <> Nothing Then
      Dim vFutureDateDiff As Integer = CInt(DateDiff(DateInterval.Day, vCurrentDate, vNextDate))
      Dim vPastDateDiff As Integer = CInt(DateDiff(DateInterval.Day, vPreviousDate, vCurrentDate))
      If vFutureDateDiff < vPastDateDiff Then
        vReturnDate = vNextDate
      ElseIf vFutureDateDiff > vPastDateDiff Then
        vReturnDate = vPreviousDate
      ElseIf vFutureDateDiff = vPastDateDiff Then
        vReturnDate = vPreviousDate
      End If
    ElseIf vNextDate <> Nothing AndAlso vPreviousDate = Nothing Then
      vReturnDate = vNextDate
    ElseIf vPreviousDate <> Nothing And vNextDate = Nothing Then
      vReturnDate = vPreviousDate
    End If
    Return vReturnDate
  End Function


#Region "Submit and Error handling"

  Protected Sub GoToSubmitPage()
    GoToSubmitPage(Nothing)
  End Sub

  Protected Sub GoToSubmitPage(ByVal pParams As String)
    If Session("UpdateDetailsURL") IsNot Nothing Then
      Session("UpdateDetailsURL") = Nothing
      Dim vParams As New ParameterList(HttpContext.Current)
      vParams("OldUserName") = Session("RegisteredUserName")
      vParams("LastUpdatedOn") = DateTime.Now.ToString(CAREDateTimeFormat)
      Dim vResultList As ParameterList = DataHelper.UpdateRegisteredUser(vParams)
    End If
    Dim vRedirectString As String = ""
    If Request.QueryString("ReturnURL") IsNot Nothing Then
      ProcessRedirect(Request.QueryString("ReturnUrl").ToString)
    ElseIf mvSubmitItemUrl.Length > 0 Then
      vRedirectString = mvSubmitItemUrl
      If pParams IsNot Nothing AndAlso pParams.Length > 0 Then
        If Not vRedirectString.Contains("?") Then vRedirectString &= "?"
        vRedirectString &= pParams
      End If
      ProcessRedirect(vRedirectString)
    ElseIf mvSubmitItemNumber > 0 Then
      vRedirectString = String.Format("Default.aspx?pn={0}{1}", mvSubmitItemNumber, pParams)
      ProcessRedirect(vRedirectString)
    End If
  End Sub

  Public Sub SubmitChildControls(ByVal pList As ParameterList)
    If DependantControls IsNot Nothing Then
      For Each vCareWebControl As ICareChildWebControl In DependantControls
        vCareWebControl.SubmitChild(pList)
      Next
    End If
  End Sub

  Protected Sub ProcessError(ByVal pException As Exception)
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vResult As DataTable
    Dim vErrorId As Integer = 0
    Dim vErrorURL As String = "ShowErrors.aspx"
    Dim vRedirectString As String = ""
    LogException(pException)
    Session("LastException") = pException
    Session("LastPageNumber") = Request.QueryString("pn")
    Try
      'Check Exception Type and fetch Error Number and Source if it is a CareException.
      If TypeOf (pException) Is CareException Then
        Dim vEx As CareException = CType(pException, CareException)
        vList("ErrorNumber") = vEx.ErrorNumber
        vList("ErrorSource") = vEx.Source
      Else
        vList("ErrorNumber") = 0
        vList("ErrorSource") = pException.Source
      End If
      vList("WebPageNumber") = Request.QueryString("PN")
      vList("ErrorMessage") = pException.Message
      vList("StackTrace") = pException.StackTrace
      'Record Error in Database.
      vResult = DataHelper.AddErrorLog(vList)
      If vResult.Columns.Contains("ErrorId") AndAlso vResult.Rows(0).Item("ErrorId").ToString.Length > 0 Then
        vErrorId = IntegerValue(vResult.Rows(0).Item("ErrorId").ToString)
        vErrorURL = String.Format(vErrorURL & "?EI={0}", vErrorId)
      End If
      ProcessRedirect(vErrorURL)
    Catch vEx As ThreadAbortException
      'Do nothing since it is expected from ProcessRedirect
    Catch ex As Exception
      ProcessRedirect("ShowErrors.aspx")
    End Try
  End Sub

  Protected Sub ProcessRedirect(ByVal pPage As String)
    'Server.Transfer(pPage)
    ' we cannot call Server.Transfer on an ASP.NET AJAX enabled page
    ' It throws error, thats why used Response.Redirect
    'Response.Redirect(pPage, True)
    RedirectViaWhiteList(pPage)
  End Sub


  Public Function GetDevices() As String
    Dim vDevices As String = ""
    Dim vList As New ParameterList(HttpContext.Current)
    If InitialParameters.ContainsKey("DeviceLookupGroup") Then
      vList("LookupGroup") = InitialParameters("DeviceLookupGroup")
      Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtDevices, vList)
      If vTable IsNot Nothing Then
        For Each vRow As DataRow In vTable.Rows
          vDevices = vDevices & "'" & Convert.ToString(vRow("Device")) & "',"
        Next
        If vDevices.Length > 0 Then
          vDevices = vDevices.Substring(0, vDevices.Length - 1)
        End If
      End If
    End If
    Return vDevices
  End Function

#End Region

#Region "MultiView Support"
  Protected Sub HandleMultiViewDisplay()
    HandleMultiViewDisplay("")
  End Sub
  Protected Sub HandleMultiViewDisplay(ByVal pButtonNamesToShowGrid As String)
    If mvSupportsMultiView Then
      'mvView1 always hold the grid control and mvView2 always hold the edit/search fields with buttons
      Dim vDisplayGrid As Boolean = False
      If IsPostBack Then
        'Find out if the post back is due to a button Click event which should displayed the grid control
        Dim vNames As String = "Save,Cancel"  'Standard button names that should display the grid on Click event
        If pButtonNamesToShowGrid.Length > 0 Then vNames += "," & pButtonNamesToShowGrid
        For Each vParameterName As String In vNames.Split(","c)
          Dim vControl As Control = FindControlByName(Me, vParameterName)
          If vControl IsNot Nothing Then
            Dim vNameID As String = FindControlByName(Me, vParameterName).UniqueID
            'Search through all the Keys in the Request.Form as it will contain the button id which raised the Click event
            For Each vName As String In Request.Form.Keys
              If vName = vNameID Then
                vDisplayGrid = True
                Exit For
              End If
            Next
            If vDisplayGrid Then Exit For
          End If
        Next
      End If
      If MultiViewGridOnTop() Then
        If (Not IsPostBack OrElse vDisplayGrid) Then
          'Grid should be displayed on loading the page first time (no post back) or when a button Click event is raised that should display the grid e.g. Save, Cancel, Default buttons.
          mvMultiView.SetActiveView(mvView1)
        Else
          'Otherwise, display the edit fields with buttons
          mvMultiView.SetActiveView(mvView2)
        End If
      Else
        If vDisplayGrid Then
          'Only display the grid when a button Click event was raised which requires displaying the grid e.g. Search button
          mvMultiView.SetActiveView(mvView1)
        Else
          'Display the search fields when Grid is not displayed
          mvMultiView.SetActiveView(mvView2)
        End If
      End If
    End If
  End Sub

  Private Sub AddMultiViewGrid(ByRef pHtmlTable As HtmlTable, ByVal pHtmlRow As HtmlTableRow)
    'This is used only when MultiView support is ON
    Dim vLinkButton As New LinkButton
    'Add the default text
    vLinkButton.Text = If(InitialParameters.ContainsKey("GridHyperlinkText"), InitialParameters("GridHyperlinkText").ToString, If(MultiViewGridOnTop, "Click here to add a New record...", "Click here to Search again..."))
    vLinkButton.ID = "GridHyperlink"
    vLinkButton.CausesValidation = False  'No validation is required
    AddHandler vLinkButton.Click, AddressOf ButtonClickHandler  'Need to raise an event to perform New/Search
    Dim vLinkCell As New HtmlTableCell
    vLinkCell.Controls.Add(vLinkButton)
    Dim vLinkRow As New HtmlTableRow
    vLinkRow.Cells.Add(vLinkCell)
    pHtmlTable.Rows.Add(vLinkRow) 'Add the grid hyper link row before adding the grid
    pHtmlTable.Rows.Add(pHtmlRow)
    mvView1.ID = "View1"
    mvView1.Controls.Add(pHtmlTable)
    mvMultiView.Views.Add(mvView1)
    mvView2.ID = "View2"
    mvMultiView.Views.Add(mvView2)
    mvMultiView.ActiveViewIndex = -1  'Do not display any View
    Dim vUpdatePanel As New UpdatePanel
    vUpdatePanel.ContentTemplateContainer.Controls.Add(mvMultiView)
    'Add the UpdatePanel to the module's controls as HTMLTable is removed as it was added to mvView1
    Me.Controls.Add(vUpdatePanel)
    mvGridControlTable = pHtmlTable 'Save this table (containing grid hyperlink and grid control)
    pHtmlTable = New HtmlTable  'Reset this table as the rest of the controls should be in a new table
    mvTableControl = pHtmlTable
  End Sub

#End Region

#Region "TNS"

  ''' <summary>
  ''' Remove session values for TNS Hosted 
  ''' </summary>
  ''' <param name="pResultsOnly"></param>
  ''' <remarks></remarks>
  Protected Sub ClearSessionForTnsValues(ByVal pResultsOnly As Boolean)
    If pResultsOnly Then
      If Session("FormErrorContents") IsNot Nothing Then Session.Remove("FormErrorContents")
      If Session("FormErrorCode") IsNot Nothing Then Session.Remove("FormErrorCode")
    Else
      If Session("FormErrorContents") IsNot Nothing Then Session.Remove("FormErrorContents")
      If Session("FormErrorCode") IsNot Nothing Then Session.Remove("FormErrorCode")
      If Session("ReturnUrl") IsNot Nothing Then Session.Remove("ReturnUrl")
      If Session("ActionLink") IsNot Nothing Then Session.Remove("ActionLink")
    End If
  End Sub

  ''' <summary>
  ''' Remove specified key and value from the session
  ''' </summary>
  ''' <param name="pKey"></param>
  ''' <remarks></remarks>
  Protected Sub ClearSessionValue(ByVal pKey As String)
    If Session(pKey) IsNot Nothing Then Session.Remove(pKey)
  End Sub

  ''' <summary>
  ''' Rename control IDs as required by TNS else the card authorisation will
  ''' be rejected   
  ''' </summary>
  ''' <remarks></remarks>
  Protected Sub RenameControlsForTNS(ByVal pHtmlTable As HtmlTable)
    If FindControl("CreditCardNumber") IsNot Nothing Then
      FindControlByName(Me, "CreditCardNumber").ID = "gatewayCardNumber"
      Dim vHTMLCell As HtmlTableCell = GetHtmlCellFromControlId("gatewayCardNumber", pHtmlTable)
      AddRegExValidator(vHTMLCell, "gatewayCardNumber", RegularExpressionTypes.retCreditCardNumber)
      AddRequiredValidator(vHTMLCell, "gatewayCardNumber")
    End If
    If FindControl("CreditCardType") IsNot Nothing Then
      FindControlByName(Me, "CreditCardType").ID = "gatewayCardScheme"
      Dim vHTMLCell As HtmlTableCell = GetHtmlCellFromControlId("gatewayCardScheme", pHtmlTable)
      AddRequiredValidator(vHTMLCell, "gatewayCardScheme")
    End If
    If FindControl("SecurityCode") IsNot Nothing Then
      FindControlByName(Me, "SecurityCode").ID = "gatewayCardSecurityCode"
      Dim vHTMLCell As HtmlTableCell = GetHtmlCellFromControlId("gatewayCardSecurityCode", pHtmlTable)
      AddRegExValidator(vHTMLCell, "gatewayCardSecurityCode", RegularExpressionTypes.retSecurityCode)
      If Not BooleanValue(DefaultParameters("CV2Optional").ToString) Then
        AddRequiredValidator(vHTMLCell, "gatewayCardSecurityCode")
      End If
    End If
    If FindControl("CardExpiryDate") IsNot Nothing Then
      Dim vExpiryDate As TextBox = TryCast(FindControlByName(Me, "CardExpiryDate"), TextBox)
      vExpiryDate.ID = "gatewayCardExpiryDateMonth"
      vExpiryDate.MaxLength = 2
      Dim vHTMLCell As HtmlTableCell = GetHtmlCellFromControlId("gatewayCardExpiryDateMonth", pHtmlTable)
      AddRegExValidator(vHTMLCell, "gatewayCardExpiryDateMonth", RegularExpressionTypes.retExpiryDateMonth)
      AddRequiredValidator(vHTMLCell, "gatewayCardExpiryDateMonth")
    End If
  End Sub

  Protected Function GetResponseValue(ByVal pAttributeName As String) As String
    Return GetResponseValue(pAttributeName, False)
  End Function


  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="pAttributeName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Protected Function GetResponseValue(ByVal pAttributeName As String, ByVal pWithErrorCode As Boolean) As String
    Dim vResponse As String = String.Empty
    If Session("formErrorContents") IsNot Nothing Then
      Dim vFormFieldValue As IDictionary(Of String, String) = TryCast(Session("formErrorContents"), IDictionary(Of String, String))
      Dim vKeyName As String = String.Empty

      'Asp.net decorates the attribute name with some additional information so find
      ' the actual key name to get the value
      For Each vKey As String In vFormFieldValue.Keys
        If vKey.Contains(pAttributeName) Then
          vKeyName = vKey
          Exit For
        End If
      Next

      If vKeyName.Length > 0 Then
        vResponse = vFormFieldValue(vKeyName)
        If vResponse IsNot Nothing AndAlso vResponse.Length > 0 Then
          If pWithErrorCode Then
            Return vResponse
          Else
            If vResponse.Contains("~"c) Then
              If vResponse.Length > vResponse.IndexOf("~"c) + 1 Then
                Return vResponse.Substring(vResponse.IndexOf("~"c) + 1)
              Else
                Return ""
              End If
            Else
              Return vResponse
            End If
          End If

        End If
      Else
        Return vResponse
      End If
    End If
    Return vResponse
  End Function

  ''' <summary>
  ''' Show invalid field error
  ''' </summary>
  ''' <param name="pField"></param>
  ''' <param name="pFieldName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Protected Function ShowTNSErrorDetails(ByVal pField As String, ByVal pFieldName As String) As String
    Dim vErrorCode As String = pField.Substring(0, 1)
    Dim vErrorString As String = String.Empty
    Select Case CInt(vErrorCode)
      Case 1
        vErrorString = "Mandatory Field " + pFieldName
      Case 2
        vErrorString = "Invalid value " + pFieldName
      Case Else
        vErrorString = ""
    End Select
    Return vErrorString
  End Function

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <remarks></remarks>
  Protected Sub DisplayError()
    'Show error messgae to the user
    For Each vControlName As String In mvTnsControlName
      If GetResponseValue(vControlName, True).Contains("~"c) Then SetErrorLabel(ShowTNSErrorDetails(GetResponseValue(vControlName, True), GetDisplayFieldName(vControlName)))
    Next
  End Sub

  Private Function GetDisplayFieldName(ByVal pFieldName As String) As String
    Select Case pFieldName
      Case "gatewayCardNumber"
        Return "Card Number"
      Case "gatewayCardSecurityCode"
        Return "Card Security Code"
      Case "gatewayCardExpiryDateMonth"
        Return "Card Expiry Date (Month)"
      Case "gatewayCardExpiryDateYear"
        Return "Card Expiry Date (Year)"
      Case Else
        Return ""
    End Select
  End Function

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="pTable"></param>
  ''' <param name="pText"></param>
  ''' <param name="pControlID"></param>
  ''' <remarks></remarks>
  Protected Sub AddHiddenField(ByVal pTable As HtmlTable, ByVal pText As String, ByVal pControlID As String)
    Dim vHTMLRow As New HtmlTableRow
    Dim vHTMLCell As New HtmlTableCell
    Dim vHiddenControl As New HiddenField

    vHiddenControl.ID = pControlID
    vHiddenControl.Value = pText
    vHTMLCell.Controls.Add(vHiddenControl)
    vHTMLRow.Cells.Add(vHTMLCell)
    pTable.Rows.Add(vHTMLRow)

  End Sub

  ''' <summary>
  ''' Method will restore the values back from the TNS on the controls so that user do not have to the data entry again
  ''' </summary>
  ''' <remarks></remarks>
  Protected Sub SetFieldValues()

    If FindControlByName(Me, "Title") IsNot Nothing Then TryCast(FindControlByName(Me, "Title"), DropDownList).Text = GetResponseValue("Title")
    If FindControlByName(Me, "Forenames") IsNot Nothing Then TryCast(FindControlByName(Me, "Forenames"), TextBox).Text = GetResponseValue("Forenames")
    If FindControlByName(Me, "Surname") IsNot Nothing Then TryCast(FindControlByName(Me, "Surname"), TextBox).Text = GetResponseValue("Surname")
    If FindControlByName(Me, "EMailAddress") IsNot Nothing Then TryCast(FindControlByName(Me, "EMailAddress"), TextBox).Text = GetResponseValue("EMailAddress")
    If FindControlByName(Me, "ConfirmEMailAddress") IsNot Nothing Then TryCast(FindControlByName(Me, "ConfirmEMailAddress"), TextBox).Text = GetResponseValue("ConfirmEMailAddress")
    If FindControlByName(Me, "PostcoderPostcode") IsNot Nothing Then TryCast(FindControlByName(Me, "PostcoderPostcode"), TextBox).Text = GetResponseValue("PostcoderPostcode")
    If FindControlByName(Me, "PostcoderAddress") IsNot Nothing Then TryCast(FindControlByName(Me, "PostcoderAddress"), TextBox).Text = GetResponseValue("PostcoderAddress")
    If FindControlByName(Me, "Address") IsNot Nothing Then TryCast(FindControlByName(Me, "Address"), TextBox).Text = GetResponseValue("Address")
    If FindControlByName(Me, "Town") IsNot Nothing Then TryCast(FindControlByName(Me, "Town"), TextBox).Text = GetResponseValue("Town")
    If FindControlByName(Me, "County") IsNot Nothing Then TryCast(FindControlByName(Me, "County"), TextBox).Text = GetResponseValue("County")
    If FindControlByName(Me, "Country") IsNot Nothing Then TryCast(FindControlByName(Me, "Country"), TextBox).Text = GetResponseValue("Country")
    If FindControlByName(Me, "CreditCardType") IsNot Nothing Then TryCast(FindControlByName(Me, "CreditCardType"), DropDownList).Text = GetResponseValue("gatewayCardScheme")
    If FindControlByName(Me, "NetAmount") IsNot Nothing Then TryCast(FindControlByName(Me, "NetAmount"), TextBox).Text = GetResponseValue("NetAmount")
    If FindControlByName(Me, "VatAmount") IsNot Nothing Then TryCast(FindControlByName(Me, "VatAmount"), TextBox).Text = GetResponseValue("VatAmount")
    If FindControlByName(Me, "GrossAmount") IsNot Nothing Then TryCast(FindControlByName(Me, "GrossAmount"), TextBox).Text = GetResponseValue("GrossAmount")
    If FindControlByName(Me, "Balance") IsNot Nothing Then TryCast(FindControlByName(Me, "Balance"), TextBox).Text = GetResponseValue("Balance")
    If FindControlByName(Me, "MembershipType") IsNot Nothing Then TryCast(FindControlByName(Me, "MembershipType"), TextBox).Text = GetResponseValue("MembershipType")
    If FindControlByName(Me, "CreditCardNumber") IsNot Nothing Then TryCast(FindControlByName(Me, "CreditCardNumber"), TextBox).Text = GetResponseValue("gatewayCardNumber")
    If FindControlByName(Me, "CardExpiryDate") IsNot Nothing Then TryCast(FindControlByName(Me, "CardExpiryDate"), TextBox).Text = GetResponseValue("gatewayCardExpiryDateMonth")
    If FindControlByName(Me, "gatewayCardExpiryDateYear") IsNot Nothing Then TryCast(FindControlByName(Me, "gatewayCardExpiryDateYear"), TextBox).Text = GetResponseValue("gatewayCardExpiryDateYear")
    If FindControlByName(Me, "IssueNumber") IsNot Nothing Then TryCast(FindControlByName(Me, "IssueNumber"), TextBox).Text = GetResponseValue("IssueNumber")
    If FindControlByName(Me, "CardStartDate") IsNot Nothing Then TryCast(FindControlByName(Me, "CardStartDate"), TextBox).Text = GetResponseValue("CardStartDate")
    If FindControlByName(Me, "SecurityCode") IsNot Nothing Then TryCast(FindControlByName(Me, "SecurityCode"), TextBox).Text = GetResponseValue("gatewayCardSecurityCode")

  End Sub

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="vHtmlTable"></param>
  ''' <remarks></remarks>
  Protected Sub AddReturnLinkForTns(ByRef vHtmlTable As HtmlTable)
    If Me.FindControlByName(Me, "gatewayReturnURL") IsNot Nothing AndAlso Session("ReturnUrl") IsNot Nothing Then TryCast(FindControlByName(Me, "gatewayReturnURL"), TextBox).Text = Session("ReturnUrl").ToString
  End Sub

  ''' <summary>
  ''' Returns the HtmlCell for the specified control ID
  ''' </summary>
  ''' <param name="pControlID">Control ID </param>
  ''' <param name="pTable">HTMLTable which needs searching for html cell</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetHtmlCellFromControlId(ByVal pControlID As String, ByVal pTable As HtmlTable) As HtmlTableCell
    Dim vCell As HtmlTableCell = Nothing
    Dim vFound As Boolean
    For Each vHtmlRow As HtmlTableRow In pTable.Rows
      For Each vHtmlcells As HtmlTableCell In vHtmlRow.Cells
        For Each vControl As Control In vHtmlcells.Controls
          If vControl.ID = pControlID Then
            vFound = True
            vCell = vHtmlcells
            Exit For
          End If
        Next
        If vFound Then Exit For
      Next
      If vFound Then Exit For
    Next
    Return vCell
  End Function


  Private Sub ClearTNSControls()

  End Sub
#End Region

  Public Overridable Sub CheckBoxChecked(sender As Object, e As EventArgs)
    If FindControlByName(Me, "TokenDesc") IsNot Nothing Then
      Dim vCheckBox As CheckBox = TryCast(sender, CheckBox)
      If vCheckBox.Checked Then
        FindControlByName(Me, "TokenDesc").Visible = True
        FindControlByName(Me, "TokenDesc").Parent.Parent.Visible = True
      Else
        FindControlByName(Me, "TokenDesc").Visible = False
        FindControlByName(Me, "TokenDesc").Parent.Parent.Visible = False
      End If
    End If
    If FindControlByName(Me, "TokenList") IsNot Nothing Then
      DirectCast(FindControlByName(Me, "TokenList"), ListBox).ClearSelection()
    End If
  End Sub

#Region "SagePay"
  Protected Sub PopulateListBox()
    Dim vParamList As New ParameterList(HttpContext.Current)
    vParamList("ContactNumber") = UserContactNumber.ToString
    If FindControlByName(Me, "TokenList") IsNot Nothing Then
      FindControlByName(Me, "TokenList").Parent.Parent.Visible = True
      Dim vListBox As ListBox = DirectCast(FindControlByName(Me, "TokenList"), ListBox)
      Dim vDataItem As ParameterList = New ParameterList(HttpContext.Current)
      vDataItem("TextField") = "TokenDesc"
      vDataItem("ValueField") = "TokenId"
      DataHelper.FillList(CareNetServices.XMLLookupDataTypes.xldtContactCreditCards, vListBox, vDataItem, False, vParamList, False)
    End If
  End Sub

#End Region


  Protected Sub HideSagePayControls()
    Dim vControlDictionary As Dictionary(Of String, Boolean) = New Dictionary(Of String, Boolean)
    vControlDictionary.Add("TokenList", False)
    vControlDictionary.Add("CreateToken", False)
    vControlDictionary.Add("TokenDesc", False)

    For Each vItem As KeyValuePair(Of String, Boolean) In vControlDictionary
      SetControlVisible(vItem.Key, vItem.Value)
      If FindControlByName(Me, vItem.Key) IsNot Nothing AndAlso FindControlByName(Me, vItem.Key).Parent IsNot Nothing AndAlso FindControlByName(Me, vItem.Key).Parent.Parent IsNot Nothing Then FindControlByName(Me, vItem.Key).Parent.Parent.Visible = vItem.Value
    Next

  End Sub

  'CSXL210FE|SCXLVPCSCP|PROTX|TNSHOSTED|SAGEPAYHOSTED
  Protected Overridable Function GetAuthorisationType() As AuthorisationService
    '
  End Function

  ' ''' <summary>For CPD Cycles, are the StartDate and EndDate controls visible?</summary>
  ' ''' <returns>True if both controls are visisle, otherwise False</returns>
  'Protected Function HasCPDCycleStartAndEndDates() As Boolean
  '  If IsControlVisible("StartDate") = True AndAlso IsControlVisible("EndDate") = True Then
  '    Return True
  '  Else
  '    Return False
  '  End If
  'End Function

  'Protected Sub SetCPDEndYearOrDate(ByVal pUseStartYear As Boolean)
  '  Dim vDDL As DropDownList = TryCast(FindControlByName(Me, "CpdCycleType"), DropDownList)
  '  Dim vCycleTypeList As New ParameterList(HttpContext.Current)
  '  vCycleTypeList("ForPortal") = "Y"
  '  Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCPDCycleTypes, vCycleTypeList)
  '  Dim vRowColl() As DataRow = vTable.Select("CpdCycleType = '" & vDDL.Items(vDDL.SelectedIndex).Value & "'")
  '  If vDDL.SelectedValue.ToString.Length > 0 Then
  '    If vRowColl.Length > 0 Then
  '      Dim vRow As DataRow = vRowColl(0)
  '      If vRow IsNot Nothing AndAlso IntegerValue(vRow.Item("DefaultDuration").ToString) > 0 Then
  '        ' We have a default duration - need to set controls accordingly
  '        Dim vStartDate As Nullable(Of Date)
  '        If pUseStartYear = True AndAlso IntegerValue(GetTextBoxText("StartYear")) > 0 Then
  '          vStartDate = New Date(IntegerValue(GetTextBoxText("StartYear")), IntegerValue(vRow("StartMonth").ToString), 1)
  '        ElseIf pUseStartYear = False AndAlso GetTextBoxText("StartDate").Length > 0 Then
  '          vStartDate = Date.Parse(GetTextBoxText("StartDate"))
  '        End If
  '        If vStartDate.HasValue Then
  '          Dim vEndDate As Date = vStartDate.Value.AddYears(IntegerValue(vRow.Item("DefaultDuration").ToString)).AddDays(-1)
  '          If pUseStartYear Then
  '            SetTextBoxText("EndYear", vEndDate.Year.ToString)
  '            SetControlEnabled("EndYear", False)
  '          Else
  '            SetTextBoxText("EndDate", vEndDate.ToString(CAREDateFormat))
  '            SetControlEnabled("EndDate", False)
  '          End If
  '        End If
  '        'Dim vDate As New Date(IntegerValue(GetTextBoxText("StartYear")), IntegerValue(vRow("StartMonth").ToString), 1)
  '        'vDate = vDate.AddYears(IntegerValue(vRow.Item("DefaultDuration").ToString)).AddDays(-1)
  '        'TryCast(FindControlByName(Me, "EndYear"), TextBox).Text = vDate.Year.ToString
  '        'SetControlEnabled("EndYear", False)
  '      Else
  '        ' No default duration - let the user decide
  '        If pUseStartYear Then
  '          SetControlEnabled("EndYear", True)
  '        Else
  '          SetControlEnabled("EndDate", True)
  '        End If
  '      End If
  '    End If
  '  Else
  '    If pUseStartYear Then
  '      SetControlEnabled("EndYear", False)
  '    Else
  '      SetControlEnabled("EndDate", False)
  '    End If
  '  End If
  'End Sub

End Class


#Region "URL or EMail Template"

Public Class URLOrEMailTemplate
  Implements ITemplate

  Private mvAttr As String

  Public Sub New(ByVal pAttr As String)
    mvAttr = pAttr
  End Sub

  'must implement following method
  Public Sub InstantiateIn(ByVal pContainer As Control) Implements ITemplate.InstantiateIn
    Dim vLabel As New Label
    AddHandler vLabel.DataBinding, AddressOf OnDataBinding
    pContainer.Controls.Add(vLabel)
  End Sub

  Public Sub OnDataBinding(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vLabel As Label = DirectCast(sender, Label)
    Dim vDGItem As DataGridItem = DirectCast(vLabel.NamingContainer, DataGridItem)
    Dim vDRView As DataRowView = DirectCast(vDGItem.DataItem, DataRowView)
    Dim vResult As String = vDRView(mvAttr).ToString
    Dim vRegEx As Regex

    Dim vPatternEmail As String = "[a-zA-Z_0-9.-]+\@[a-zA-Z_0-9.-]+\.\w+"
    vRegEx = New Regex(vPatternEmail)
    If vRegEx.IsMatch(vResult) Then
      vResult = vRegEx.Replace(vResult, AddressOf MailToMatchEvaluator)
    Else
      Dim vPatternSite As String = "\w*[\://]*\w+\.\w+\.\w+[/\w+]*[.\w+]*"
      vRegEx = New Regex(vPatternSite)
      If vRegEx.IsMatch(vResult) Then
        vResult = vRegEx.Replace(vResult, AddressOf WebSiteMatchEvaluator)
      ElseIf vResult.StartsWith("http:") OrElse vResult.StartsWith("https:") Then
        Dim ub As UriBuilder = New UriBuilder(vResult)
        Dim sb As StringBuilder = New StringBuilder("<a href='")
        sb.Append(ub.ToString())
        sb.Append("'>")
        sb.Append(vResult)
        sb.Append("</a>")
        vResult = sb.ToString()
      End If
    End If
    vLabel.Text = vResult
  End Sub

  Private Function MailToMatchEvaluator(ByVal m As Match) As String
    Dim sb As StringBuilder = New StringBuilder("<a href='mailto:")
    sb.Append(m.Value)
    sb.Append("'>")
    sb.Append(m.Value)
    sb.Append("</a>")
    Return sb.ToString()
  End Function

  Private Function WebSiteMatchEvaluator(ByVal m As Match) As String
    Dim ub As UriBuilder = New UriBuilder(m.Value)
    Dim sb As StringBuilder = New StringBuilder("<a href='")
    sb.Append(ub.ToString())
    sb.Append("'>")
    sb.Append(m.Value)
    sb.Append("</a>")
    Return sb.ToString()
  End Function

End Class

#End Region

#Region "TwoAttributeTemplate"

Public Class TwoAttributeTemplate
  Implements ITemplate

  Private mvAttr1 As String
  Private mvAttr2 As String

  Public Sub New(ByVal pAttr1 As String, ByVal pAttr2 As String)
    mvAttr1 = pAttr1
    mvAttr2 = pAttr2
  End Sub

  'must implement following method
  Public Sub InstantiateIn(ByVal pContainer As Control) Implements ITemplate.InstantiateIn
    Dim vLabel As New Label
    AddHandler vLabel.DataBinding, AddressOf OnDataBinding
    pContainer.Controls.Add(vLabel)
  End Sub

  Public Sub OnDataBinding(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vLabel As Label = DirectCast(sender, Label)
    Dim vDGItem As DataGridItem = DirectCast(vLabel.NamingContainer, DataGridItem)
    Dim vDRView As DataRowView = DirectCast(vDGItem.DataItem, DataRowView)
    vLabel.Text = HttpUtility.HtmlEncode(vDRView(mvAttr1).ToString) & "<br>" & HttpUtility.HtmlEncode(vDRView(mvAttr2).ToString)
  End Sub

End Class

#End Region

#Region "MemoTemplate"

Public Class MemoTemplate
  Implements ITemplate

  Private mvAttr As String

  Public Sub New(ByVal pAttr As String)
    mvAttr = pAttr
  End Sub

  'must implement following method
  Public Sub InstantiateIn(ByVal pContainer As Control) Implements ITemplate.InstantiateIn
    Dim vLabel As New Label
    AddHandler vLabel.DataBinding, AddressOf OnDataBinding
    pContainer.Controls.Add(vLabel)
  End Sub

  Public Sub OnDataBinding(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vLabel As Label = DirectCast(sender, Label)
    Dim vDGItem As DataGridItem = DirectCast(vLabel.NamingContainer, DataGridItem)
    Dim vDRView As DataRowView = DirectCast(vDGItem.DataItem, DataRowView)
    vLabel.Text = HttpUtility.HtmlEncode((mvAttr).ToString).Replace(vbCrLf, "<br>")
  End Sub
End Class

#End Region

#Region "DisplayTemplate"

Public Class DisplayTemplate
  Implements ITemplate

  Private mvAttr As String

  Public Sub New(ByVal pAttr As String)
    mvAttr = pAttr
  End Sub

  Public ReadOnly Property DataItem As String
    Get
      Return mvAttr
    End Get
  End Property

  Public Sub InstantiateIn(ByVal pContainer As Control) Implements ITemplate.InstantiateIn
    Dim vLabel As New Label
    AddHandler vLabel.DataBinding, AddressOf OnDataBinding
    pContainer.Controls.Add(vLabel)
  End Sub

  Public Sub OnDataBinding(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vLabel As Label = DirectCast(sender, Label)
    vLabel.Text = HttpUtility.HtmlEncode(DirectCast(DirectCast(vLabel.NamingContainer, DataGridItem).DataItem, DataRowView)(mvAttr).ToString)
  End Sub
End Class

#End Region

#Region "EditTemplate"

Public Class EditTemplate
  Implements ITemplate

  Private mvAttr As String
  Private mvPageNumber As Integer
  Private mvIsCommand As Boolean
  Private mvCommandName As String
  Private mvCommandArg As String
  Private mvCommandCondition As String = ""   'When to display the Edit hyper link
  Private mvDisplayAsButton As Boolean = False

  Public Sub New(ByVal pAttr As String, ByVal pPageNumber As Integer)
    mvAttr = pAttr
    mvPageNumber = pPageNumber
    mvIsCommand = False
    mvCommandArg = String.Empty
  End Sub
  Public Sub New(ByVal pAttr As String, ByVal pPageNumber As Integer, ByVal pIsCommand As Boolean, ByVal pCommandName As String, ByVal pCommandArg As String, ByVal pCommandCondition As String)
    Me.New(pAttr, pPageNumber, pIsCommand, pCommandName, pCommandArg, pCommandCondition, False)
  End Sub
  Public Sub New(ByVal pAttr As String, ByVal pPageNumber As Integer, ByVal pIsCommand As Boolean, ByVal pCommandName As String, ByVal pCommandArg As String, ByVal pCommandCondition As String, ByVal pDisplayAsButton As Boolean)
    mvAttr = pAttr
    mvPageNumber = pPageNumber
    mvIsCommand = pIsCommand
    mvCommandName = pCommandName
    mvCommandArg = pCommandArg
    mvCommandCondition = pCommandCondition
    mvDisplayAsButton = pDisplayAsButton
  End Sub

  'must implement following method
  Public Sub InstantiateIn(ByVal pContainer As Control) Implements ITemplate.InstantiateIn
    If Not mvIsCommand Then
      Dim vLabel As New Label
      AddHandler vLabel.DataBinding, AddressOf OnDataBinding
      pContainer.Controls.Add(vLabel)
    Else
      If mvDisplayAsButton Then
        Dim vButton As New Button
        AddHandler vButton.DataBinding, AddressOf OnDataBinding
        vButton.CommandName = mvCommandName
        vButton.CausesValidation = False

        vButton.CssClass = "Button"
        vButton.Text = mvCommandName
        pContainer.Controls.Add(vButton)
      Else
        Dim vLinkButton As New LinkButton
        AddHandler vLinkButton.DataBinding, AddressOf OnDataBinding
        vLinkButton.CommandName = mvCommandName
        vLinkButton.CausesValidation = False
        vLinkButton.Text = mvCommandName
        pContainer.Controls.Add(vLinkButton)
      End If
    End If
  End Sub

  Public Sub OnDataBinding(ByVal sender As Object, ByVal e As System.EventArgs)
    If Not mvIsCommand Then
      Dim vLabel As Label = DirectCast(sender, Label)
      Dim vDGItem As DataGridItem = DirectCast(vLabel.NamingContainer, DataGridItem)
      Dim vDRView As DataRowView = DirectCast(vDGItem.DataItem, DataRowView)
      Dim vResult As String = vDRView(mvAttr).ToString

      Dim vSB As StringBuilder = New StringBuilder("<a href='")

      vSB.Append("Default.aspx?pn=")
      vSB.Append(mvPageNumber.ToString)
      vSB.Append("&AC=")
      vSB.Append(vResult)
      vSB.Append("'>Edit</a>")
      vResult = vSB.ToString()
      vLabel.Text = vResult
    Else
      If mvCommandArg.Length > 0 OrElse mvCommandCondition.Length > 0 Then
        Dim vDGItem As DataGridItem
        Dim vDRView As DataRowView
        Dim vButton As Button = Nothing
        Dim vLinkButton As LinkButton = Nothing
        If mvDisplayAsButton Then
          vButton = DirectCast(sender, Button)
          vDGItem = DirectCast(vButton.NamingContainer, DataGridItem)
          vDRView = DirectCast(vDGItem.DataItem, DataRowView)
          If mvCommandArg.Length > 0 Then vButton.CommandArgument = vDRView(mvCommandArg).ToString
        Else
          vLinkButton = DirectCast(sender, LinkButton)
          vDGItem = DirectCast(vLinkButton.NamingContainer, DataGridItem)
          vDRView = DirectCast(vDGItem.DataItem, DataRowView)
          If mvCommandArg.Length > 0 Then vLinkButton.CommandArgument = vDRView(mvCommandArg).ToString
        End If

        If mvCommandCondition.Length > 0 Then
          Dim vHideLink As Boolean = True
          'Normally the data restriction filter is applied after adding this EditTemplale.
          'but to be on safe side, retain the existing data restriction.
          Dim vOldFilter As String = vDRView.DataView.RowFilter
          vDRView.DataView.RowFilter = If(vOldFilter.Length > 0, mvCommandCondition & " AND " & vOldFilter, mvCommandCondition)
          For Each vRowView As DataRowView In vDRView.DataView
            If vRowView Is vDRView Then
              vHideLink = False
              Exit For
            End If
          Next
          vDRView.DataView.RowFilter = vOldFilter
          If vHideLink Then
            If mvDisplayAsButton Then
              vButton.Visible = False
            Else
              vLinkButton.Text = ""
            End If
          End If
        End If
      End If
    End If
  End Sub
End Class

#End Region

#Region "CheckBoxTemplate"
Public Class CheckBoxTemplate
  Implements ITemplate

  Private mvAttr As String
  Public Sub New(ByVal pAttr As String)
    mvAttr = pAttr
  End Sub
  Public Sub InstantiateIn(ByVal pContainer As Control) Implements ITemplate.InstantiateIn
    Dim vCheckBox As CheckBox = New CheckBox()

    AddHandler vCheckBox.DataBinding, AddressOf OnDataBinding
    pContainer.Controls.Add(vCheckBox)
  End Sub
  Public Sub OnDataBinding(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vCheckBox As CheckBox = CType(sender, CheckBox)
    Dim vDGItem As DataGridItem = DirectCast(vCheckBox.NamingContainer, DataGridItem)
    Dim vDRView As DataRowView = DirectCast(vDGItem.DataItem, DataRowView)
    vCheckBox.Checked = BooleanValue(vDRView(mvAttr).ToString)
  End Sub
End Class
#End Region
#Region "AlternateDisplayFormatTemplate"

Public Class AlternateDisplayFormatTemplate
  Implements ITemplate

  Dim mvTemplateType As ListItemType
  Dim mvCols() As String
  Dim mvHeaders() As String
  Dim mvHasHeader As Boolean  'Used to remove the column if there are no headings
  Private Const HEADER As String = "<td style='vertical-align:top'><b>{0}</b></td><td>&nbsp;&nbsp;</td>"

  Sub New(ByVal type As ListItemType, Optional ByVal pColumns As String = "", Optional ByVal pHeaders As String = "")
    mvTemplateType = type
    mvCols = pColumns.Split(","c)
    mvHeaders = pHeaders.Split(","c)

    'Loop through the headers and decide if we need to display the header and spacer columns.
    'The empty columns should be removed as it occupies some space and messes up the alignment.
    For Each vHeader As String In mvHeaders
      If vHeader.Length > 0 Then
        mvHasHeader = True
        Exit For
      End If
    Next
  End Sub

  Sub InstantiateIn(ByVal container As Control) _
      Implements ITemplate.InstantiateIn
    Dim vLiteral As New Literal()
    vLiteral.ID = "litItemTemplate"
    Select Case mvTemplateType
      Case ListItemType.Item
        AddHandler vLiteral.DataBinding, AddressOf OnDataBinding
    End Select

    If mvCols.Length > 0 AndAlso mvCols(0) = "CheckColumn" Then
      Dim vCheckBox As CheckBox = New CheckBox()
      AddHandler vCheckBox.DataBinding, AddressOf OnCheckBoxDataBinding
      container.Controls.Add(vCheckBox)
    End If
    container.Controls.Add(vLiteral)
  End Sub

  Public Sub OnDataBinding(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vLiteral As Literal = CType(sender, Literal)
    Dim container As DataListItem = CType(vLiteral.NamingContainer, DataListItem)
    Dim vHtml As New StringBuilder
    Dim vImg As String = String.Empty

    vHtml.Append("<table>")
    'Create a row for each column
    Dim vImagePath As String = String.Empty
    Dim vHeaderHTML As String = String.Empty
    For vIndex As Integer = 0 To mvCols.Length - 1
      If mvHasHeader Then
        vHeaderHTML = String.Format(HEADER, mvHeaders(vIndex))
      Else
        vHeaderHTML = String.Empty
      End If
      Select Case mvCols(vIndex)
        Case "EventImage"
          vImagePath = String.Format("Images/Events/{0}", DataBinder.Eval(container.DataItem, "EventImage"))
          vImg = GetImage(vImagePath, DataBinder.Eval(container.DataItem, "EventImage").ToString, "Images/Events/default.png", "EventImage")
          vHtml.AppendFormat("<tr>{0}<td>{1}</td></tr>", vHeaderHTML, vImg)
        Case "ExamImage"
          vImagePath = String.Format("Images/Exams/{0}", DataBinder.Eval(container.DataItem, "ExamImage"))
          vImg = GetImage(vImagePath, DataBinder.Eval(container.DataItem, "ExamImage").ToString, "Images/Exams/default.png", "ExamImage")
          vHtml.AppendFormat("<tr>{0}<td>{1}</td></tr>", vHeaderHTML, vImg)
        Case "ProductImage"
          Dim vImageName As String = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.web_product_image_name)
          If vImageName.Length = 0 Then vImageName = "Product{0}.png"
          vImagePath = String.Format("Images/Products/" & vImageName, DataBinder.Eval(container.DataItem, "Product"))
          vImg = GetImage(vImagePath, DataBinder.Eval(container.DataItem, "Product").ToString, "Images/Products/default.png", "ProductImage")
          vHtml.AppendFormat("<tr>{0}<td>{1}</td></tr>", vHeaderHTML, vImg)
        Case "ImageName"
          vImagePath = String.Format("Images/Downloads/{0}", DataBinder.Eval(container.DataItem, "ImageName"))
          vImg = GetImage(vImagePath, DataBinder.Eval(container.DataItem, "ImageName").ToString, "Images/Downloads/default.png", "DownloadImage")
          vHtml.AppendFormat("<tr>{0}<td>{1}</td></tr>", vHeaderHTML, vImg)
        Case Else
          vHtml.AppendFormat("<tr>{0}<td>{1}</td></tr>", vHeaderHTML, DataBinder.Eval(container.DataItem, mvCols(vIndex)))
      End Select
    Next
    'Template table ends here
    vHtml.Append("</table>")
    vLiteral.Text = vHtml.ToString
  End Sub

  Public Sub OnCheckBoxDataBinding(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vCheckBox As CheckBox = CType(sender, CheckBox)
    Dim vDataListItem As DataListItem = DirectCast(vCheckBox.NamingContainer, DataListItem)
    Dim vDRView As DataRowView = DirectCast(vDataListItem.DataItem, DataRowView)
    vCheckBox.Checked = BooleanValue(vDRView("CheckColumn").ToString)
  End Sub

End Class
#End Region

