Imports System.ComponentModel

<TypeDescriptionProvider(GetType(AbstractControlDescriptionProvider(Of CareFDEControl, System.Windows.Forms.UserControl)))>
Friend MustInherit Class CareFDEControl

  Protected mvContactInfo As ContactInfo
  Private mvControlType As CareNetServices.FDEControlTypes

  Private mvInitialParameters As String = ""      'Initial parameters
  Private mvDefaultParameters As String = ""      'Default parameters
  Protected mvInitialSettings As String = ""      'Initial parameters with values
  Protected mvDefaultSettings As String = ""      'Default parameters with values
  Protected mvDefaultList As New ParameterList
  Private mvFDEPageNumber As Integer              'Unique ID of the FDEPage this control is sitting on
  Private mvFDEPageItemNumber As Integer          'Unique ID of the control
  Private mvSequenceNumber As Integer             'Control's sequence number on the FDEPage
  Private mvPageType As String
  Protected mvUserControlName As String
  Private mvEditing As Boolean = False
  Private mvHasCompleted As Boolean

  'Event handling
  Protected mvSupportsContactData As Boolean
  Protected mvSupportsAddressData As Boolean
  Protected mvSupportsSourceChanged As Boolean
  Protected mvSupportsTransactionDateChanged As Boolean
  Protected mvSupportsSelectionChanged As Boolean
  Protected mvSupportsClearBankDetails As Boolean
  Protected mvSupportsReferenceData As Boolean

  'Processing
  Protected WithEvents mvProductValidation As ProductValidation

  Protected mvCreditCardDetailsNumber As Integer
  Protected mvBankDetailsNumber As Integer
  Protected mvCreateContactAccount As String
  Protected mvNewBank As Boolean
  Protected mvRetainSource As Boolean

  'Sizing
  Private mvMouseDown As Boolean
  Private mvMouseLocation As Point
  Private mvSize As Size
  Private mvMousePos As Point
  Private mvSizing As Boolean
  Protected mvRequiredHeight As Integer
  Private mvSizeChanged As Boolean

  Private Const BUTTON_GAP As Integer = 8
  Private Const RIGHT_GAP As Integer = 12

  Private WithEvents mvFDEControlMenu As FDEControlMenu
  Friend Event ContactChanged(ByVal sender As Object, ByVal pContactNumber As Integer)
  Friend Event AddressChanged(ByVal sender As Object, ByVal pAddressNumber As Integer)
  Friend Event SourceChanged(ByVal sender As Object, ByVal pSourceCode As String, ByVal pDistributionCode As String, ByVal pIncentiveScheme As String)
  Friend Event TransactionDateChanged(ByVal sender As Object, ByVal pTransactionDate As String)
  Friend Event DeleteControl(ByVal sender As Object)
  Friend Event SelectedContactChanged(ByVal sender As Object, ByVal pContactNumber As Integer)
  Friend Event BankDetailsChange(ByVal sender As Object)
  Friend Event ReferenceMandatory(ByVal sender As Object, ByVal pMandatory As Boolean)
  Public Event EnableOtherModules(ByVal sender As Object, ByVal pEnable As Boolean)
  Public Event FormClosingAllowed(ByVal pValue As Boolean)
  Friend Event RegDonationAdded(ByVal sender As Object)

#Region " Initialisation "

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pEditing As Boolean)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    mvEditing = pEditing
    mvInitialSettings = pRow("InitialSettings").ToString    'pInitialSettings
    mvDefaultSettings = pRow("DefaultSettings").ToString    'pDefaultSettings
    mvDefaultList = GetParameterListFromSettings(mvDefaultSettings)
    InitialiseControls(pType, pRow, IntegerValue(pRow("SequenceNumber").ToString))
  End Sub

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pInitialSettings As String, ByVal pDefaultSettings As String, ByVal pFDEPageNumber As Integer, ByVal pSequenceNumber As Integer, ByVal pEditing As Boolean)

    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    mvEditing = pEditing
    mvInitialSettings = pInitialSettings
    mvDefaultSettings = pDefaultSettings
    mvDefaultList = GetParameterListFromSettings(mvDefaultSettings)
    mvFDEPageNumber = pFDEPageNumber
    InitialiseControls(pType, pRow, pSequenceNumber)
  End Sub

  Protected Sub InitialiseControls(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pSequenceNumber As Integer)
    mvPageType = pRow("FpPageType").ToString
    mvControlType = pType
    mvProductValidation = New ProductValidation(pType)
    mvSequenceNumber = pSequenceNumber
    mvUserControlName = pRow("FdeUserControl").ToString
    mvInitialParameters = pRow("InitialParameters").ToString
    mvDefaultParameters = pRow("DefaultParameters").ToString
    If pRow.Table.Columns.Contains("FdePageNumber") Then mvFDEPageNumber = IntegerValue(pRow("FdePageNumber").ToString)
    If pRow.Table.Columns.Contains("FdePageItemNumber") Then mvFDEPageItemNumber = IntegerValue(pRow("FdePageItemNumber").ToString)


    'Set-up the EditPanel
    If mvFDEPageItemNumber = 0 Then Save(mvSequenceNumber)
    Dim vList As ParameterList = GetControlParameters()
    Dim vDT As DataTable = DataHelper.GetFastDataEntryControlItems(vList)
    Dim vPanelItems As New PanelItems(mvUserControlName, vDT)
    vPanelItems.ResetDisplayGridHeight = pType <> CareNetServices.FDEControlTypes.AddRegularDonation  'Use the original display height for Add Regular Donation
    ResetPanelItems(vPanelItems)

    'set caption if control type is Product Sale
    If pType = CareNetServices.FDEControlTypes.ProductSale Then
      vPanelItems(0).ControlCaption = "Product Sale"
    End If

    epl.Init(New EditPanelInfo(mvUserControlName, vPanelItems))

    'Complete setting up of the control
    ChangeControlAnchors()

    If mvEditing Then
      mvFDEControlMenu = New FDEControlMenu(mvUserControlName, (mvInitialParameters.Length + mvDefaultParameters.Length) > 0)
      epl.ContextMenuStrip = mvFDEControlMenu
    Else
      Me.pnl.BorderStyle = System.Windows.Forms.BorderStyle.None     'Remove the border
    End If
    If pRow.Table.Columns.Contains("FdeItemHeight") Then
      Me.Height = IntegerValue(pRow("FdeItemHeight").ToString)
    Else
      Me.Height = epl.RequiredHeight
    End If
    If mvEditing = False Then SetDefaults()
    Me.epl.DataChanged = False
  End Sub

#End Region

#Region " Public Properties "

  Friend ReadOnly Property ControlSizeChanged() As Boolean
    Get
      Return mvSizeChanged
    End Get
  End Property

  Friend ReadOnly Property ControlType() As CareNetServices.FDEControlTypes
    Get
      Return mvControlType
    End Get
  End Property

  Friend ReadOnly Property SupportsClearBankDetails() As Boolean
    Get
      Return mvSupportsClearBankDetails
    End Get
  End Property

  Friend ReadOnly Property SupportsContactData() As Boolean
    Get
      Return mvSupportsContactData
    End Get
  End Property

  Friend ReadOnly Property SupportsSelectionChanged() As Boolean
    Get
      Return mvSupportsSelectionChanged
    End Get
  End Property

  Friend ReadOnly Property SupportsAddressData() As Boolean
    Get
      Return mvSupportsAddressData
    End Get
  End Property

  Friend ReadOnly Property SupportsReferenceData() As Boolean
    Get
      Return mvSupportsReferenceData
    End Get
  End Property

  Friend ReadOnly Property SupportsSourceChanged() As Boolean
    Get
      Return mvSupportsSourceChanged
    End Get
  End Property

  Friend ReadOnly Property SupportsTransactionDateChanged() As Boolean
    Get
      Return mvSupportsTransactionDateChanged
    End Get
  End Property

  Friend ReadOnly Property UserControlName() As String
    Get
      Return mvUserControlName
    End Get
  End Property

  Friend Overridable ReadOnly Property CanSubmit() As Boolean
    Get
      Return False
    End Get
  End Property

  Friend Overridable Property HasCompleted() As Boolean
    Get
      Return mvHasCompleted
    End Get
    Set(ByVal pValue As Boolean)
      mvHasCompleted = pValue
    End Set
  End Property

#End Region

#Region " Public Methods "

  Friend Overridable Sub ResizeControl(ByVal pWidth As Integer)
    Dim vEditButton As Control
    Dim vTop As Integer
    vEditButton = FindControl(epl, "Edit", False)
    Dim vButtonWidth As Integer
    If vEditButton IsNot Nothing Then
      vButtonWidth = vEditButton.Width + BUTTON_GAP
      vEditButton.Left = Me.Width - (vEditButton.Width + RIGHT_GAP)
    End If
    For Each vControl As Control In epl.Controls
      If TypeOf (vControl) Is TextLookupBox Then
        Dim vTextLookupBox As TextLookupBox = DirectCast(vControl, TextLookupBox)
        vTextLookupBox.SetTotalWidth(Me.Width - (vTextLookupBox.Left + vButtonWidth + RIGHT_GAP))
        If vTop = 0 Then vTop = vTextLookupBox.Top
      ElseIf TypeOf (vControl) Is DisplayGrid Then
        Dim vDisplayGrid As DisplayGrid = DirectCast(vControl, DisplayGrid)
        If mvControlType = CareNetServices.FDEControlTypes.AddRegularDonation Then
          vDisplayGrid.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        Else
          vDisplayGrid.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Bottom
        End If
        vDisplayGrid.SetBounds(vDisplayGrid.Left, vDisplayGrid.Top, Me.Width - (vDisplayGrid.Left + vButtonWidth + RIGHT_GAP), vDisplayGrid.Height)
        If vTop = 0 Then vTop = vDisplayGrid.Top
      ElseIf TypeOf (vControl) Is ComboBox Then
        If vTop = 0 Then vTop = vControl.Top
      End If
    Next
    If vEditButton IsNot Nothing AndAlso vTop > 0 Then vEditButton.Top = vTop

    If mvEditing = False Then
      Me.Height = epl.Height
      Me.pnl.Height = Me.Height
    End If
  End Sub

  Friend Overridable Sub RefreshContactData(ByVal pContactInfo As ContactInfo)
    mvContactInfo = pContactInfo
    mvProductValidation.ContactInfo = pContactInfo
    mvProductValidation.VATRateInfo = Nothing
  End Sub

  Friend Overridable Sub ChangeSelectedContact(ByVal pContactNumber As Integer)
    epl.FindPanelControl(Of TextLookupBox)("ContactNumber").Text = pContactNumber.ToString
  End Sub

  Friend Overridable Sub ResetBankDetails()
  End Sub
  Friend Overridable Sub RefreshDonationBalance()
  End Sub

  Friend Overridable Sub SetReferenceMandatory(ByVal pMandatory As Boolean)
  End Sub

  Friend Overridable Sub RefreshAddressData(ByVal pAddressNumber As Integer)
  End Sub

  Friend Overridable Sub RefreshSource(ByVal pSourceCode As String, ByVal pDistributionCode As String, ByVal pIncentiveScheme As String)
  End Sub

  Friend Overridable Sub SetAndRefreshSource(ByVal pSourceCode As String)
    SetValueRaiseChanged(epl, "Source", pSourceCode)
  End Sub

  Friend Overridable Sub RefreshTransactionDate(ByVal pTransactionDate As String)
  End Sub

  Friend Overridable Sub ResetIncentives()
  End Sub

  Friend Overridable Sub ButtonClicked(ByVal pParameterName As String)
  End Sub

  Friend Overridable Sub SetDefaults()
    epl.Populate(mvDefaultList)
    epl.DataChanged = False
  End Sub

  Friend Overridable Sub SetPartialDefaults()

  End Sub

  Friend Overridable Function BuildParameterList(ByRef pList As ParameterList) As Boolean
    pList.FillFromValueList(mvDefaultSettings)
    Dim vValid As Boolean = epl.AddValuesToList(pList, True, EditPanel.AddNullValueTypes.anvtCheckBoxesOnly)
    Return vValid
  End Function

  Friend Overridable Function CheckIncentives(ByRef pList As ParameterList) As Boolean
  End Function

  Friend Overridable Sub AddIncentives(ByVal pSequenceNumbers As String, ByVal pQuantity As String)
  End Sub

  Friend Sub ResetModule(ByVal pNextButton As Boolean)
    Dim vReset As Boolean = True
    Select Case mvControlType
      Case CareNetServices.FDEControlTypes.AddTransactionDetails
        vReset = Not pNextButton
        epl.EnableControlList("TransactionDate,Source,Mailing,TransactionAmount,TransactionOrigin,Notes,EligibleForGiftAid,Reference", vReset)
      Case CareNetServices.FDEControlTypes.ContactSelection
        vReset = Not pNextButton
        epl.EnableControl("ContactNumber", vReset)
      Case CareNetServices.FDEControlTypes.AddRegularDonation
        If Not pNextButton Then mvContactInfo = Nothing 'Clear mvContacInfo so that a default line is not added for this contact as it is not selected anymore
      Case CareNetServices.FDEControlTypes.Telemarketing, CareNetServices.FDEControlTypes.AddressDisplay
        vReset = Not pNextButton
    End Select
    Dim vDefaultList As New ParameterList
    Dim vClearAllFields As Boolean = True
    If mvControlType = CareNetServices.FDEControlTypes.AddDonationCC Or mvControlType = CareNetServices.FDEControlTypes.ProductSale Then
      If pNextButton Then
        vDefaultList.Add("PaymentMethod", epl.GetValue("PaymentMethod"))
        vDefaultList.Add("CardNumberNotRequired", epl.GetValue("CardNumberNotRequired"))
        vDefaultList.Add("SortCode", epl.GetValue("SortCode"))
        vDefaultList.Add("AccountNumber", epl.GetValue("AccountNumber"))
        vDefaultList.Add("CreditCardType", epl.GetValue("CreditCardType"))
        vDefaultList.Add("CreditCardNumber", epl.GetValue("CreditCardNumber"))
        vDefaultList.Add("IssueNumber", epl.GetValue("IssueNumber"))
        vDefaultList.Add("CardStartDate", epl.GetValue("CardStartDate"))
        vDefaultList.Add("CardExpiryDate", epl.GetValue("CardExpiryDate"))
        vDefaultList.Add("SecurityCode", epl.GetValue("SecurityCode"))
        epl.ClearPartialControlList(epl, "PaymentMethod,CardNumberNotRequired,SortCode,AccountNumber,CreditCardType,CreditCardNumber,IssueNumber,CardStartDate,CardExpiryDate,SecurityCode")
      End If
    End If
    mvRetainSource = pNextButton
    If vReset Then
      epl.Clear()
      SetDefaults()
      epl.SetOptionButtonTabStops(epl)
      mvBankDetailsNumber = 0
      If pNextButton Then
        If SupportsContactData AndAlso mvEditing = False Then RefreshContactData(mvContactInfo)
      Else
        epl.DataChanged = False
      End If
      mvHasCompleted = False
    End If
    If vDefaultList.Count > 0 Then
      epl.SetValue("PaymentMethod", vDefaultList("PaymentMethod"))
      epl.SetValue("CardNumberNotRequired", vDefaultList("CardNumberNotRequired"))
      epl.SetValue("SortCode", vDefaultList("SortCode"))
      epl.SetValue("AccountNumber", vDefaultList("AccountNumber"))
      epl.SetValue("CreditCardType", vDefaultList("CreditCardType"))
      epl.SetValue("CreditCardNumber", vDefaultList("CreditCardNumber"))
      epl.SetValue("IssueNumber", vDefaultList("IssueNumber"))
      epl.SetValue("CardStartDate", vDefaultList("CardStartDate"))
      epl.SetValue("CardExpiryDate", vDefaultList("CardExpiryDate"))
      epl.SetValue("SecurityCode", vDefaultList("SecurityCode"))
      epl.EnableControlList("PaymentMethod,CardNumberNotRequired,SortCode,AccountNumber,CreditCardType,CreditCardNumber,IssueNumber,CardStartDate,CardExpiryDate,SecurityCode", False)
    End If
  End Sub

  Friend Sub Delete()
    If mvFDEPageItemNumber > 0 Then
      Dim vList As New ParameterList(True, True)
      vList.IntegerValue("FdePageItemNumber") = mvFDEPageItemNumber
      DataHelper.DeleteFastDataEntryData(CareNetServices.XMLFastDataEntryTypes.fdePageControls, vList)
    End If
  End Sub

  Friend Sub Save(ByVal pSequenceNumber As Integer)
    mvSequenceNumber = pSequenceNumber
    Dim vList As ParameterList = GetControlParameters()
    Dim vReturnList As ParameterList = DataHelper.UpdateFastDataEntryData(CareNetServices.XMLFastDataEntryTypes.fdePageControls, vList)
    mvFDEPageItemNumber = vReturnList.IntegerValue("FdePageItemNumber")
    mvSizeChanged = False
  End Sub

  Friend Overridable Sub GetPaymentMethodParameters(ByVal pList As ParameterList)
  End Sub

  ''' <summary>Set the Direct Debit start date using the appropriate auto pay delay days.</summary>
  ''' <param name="pBaseDate">Base start date for the DD. Start date is calculated from this date.</param>
  ''' <param name="pParameterName">Name of DD start date control that will be updated with the calculated date.</param>
  Protected Sub SetDDStartDate(ByVal pBaseDate As Date, ByVal pParameterName As String)
    Dim vAutoPayDate As Date = pBaseDate
    Dim vBankAccount As String = epl.GetValue("BankAccount")

    If String.IsNullOrWhiteSpace(vBankAccount) Then
      'If we don't have a Bank Account then just use config
      'Once Bank Account has been set date will be re-calculated
      Dim vDays As Integer = IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_auto_pay_delay))
      If vDays > 0 Then
        vAutoPayDate = pBaseDate.AddDays(vDays) 'Just add days and don't worry about working days
      End If
    Else
      Dim vList As New ParameterList(True, True)
      vList("AutoPayDate") = pBaseDate.ToString(AppValues.DateFormat)
      vList("PaymentMethod") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_dd)
      vList("BankAccount") = vBankAccount
      If mvControlType.Equals(CareNetServices.FDEControlTypes.AddMemberDD) Then
        vList("MembershipType") = Me.GetOptionButtonValue("MembershipType")
        vList("Joined") = epl.GetValue("Joined")
      End If
      vAutoPayDate = DataHelper.GetPaymentPlanAutoPayDate(vList)
    End If

    epl.SetValue(pParameterName, vAutoPayDate.ToShortDateString)
  End Sub

  ''' <summary>Reset the controls DataChanged to be False.</summary>
  Friend Sub ResetDataChanged()
    epl.DataChanged = False
  End Sub

#End Region

#Region " Edit Panel Events "

  Private Sub epl_ButtonClicked(ByVal sender As Object, ByVal pParameterName As String) Handles epl.ButtonClicked
    If mvEditing = False Then
      If pParameterName = "Edit" AndAlso mvContactInfo IsNot Nothing Then
        Dim vType As CareServices.XMLContactDataSelectionTypes = CareServices.XMLContactDataSelectionTypes.xcdtNone
        Select Case mvControlType
          Case CareNetServices.FDEControlTypes.ActivityDisplay
            If mvDefaultList.ContainsKey("ActivityGroup") Then
              ShowDataSheet(Me.ParentForm, frmDataSheet.DataSheetTypes.dstActivities, mvContactInfo, "B", "", mvDefaultList("ActivityGroup"), "")
              RefreshContactData(mvContactInfo)
            End If
          Case CareNetServices.FDEControlTypes.AddressDisplay
            vType = CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses
          Case CareNetServices.FDEControlTypes.CommunicationsDisplay
            vType = CareServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers
          Case CareNetServices.FDEControlTypes.GiftAidDisplay
            vType = CareServices.XMLContactDataSelectionTypes.xcdtContactGiftAidDeclarations
          Case CareNetServices.FDEControlTypes.SuppressionDisplay
            If mvDefaultList.ContainsKey("SuppressionGroup") Then
              ShowDataSheet(Me.ParentForm, frmDataSheet.DataSheetTypes.dstSuppressions, mvContactInfo, "", "", mvDefaultList("SuppressionGroup"), "")
              RefreshContactData(mvContactInfo)
            End If
          Case CareNetServices.FDEControlTypes.AddRegularDonation
            Dim vAddRegDon As FDEAddRegularDonation = DirectCast(Me, FDEAddRegularDonation)
            Dim vRow As Integer
            Dim vGrid As DisplayGrid = DirectCast(FindControl(epl, "DetailLines"), DisplayGrid)
            If vGrid.DataRowCount = 0 Then
              vRow = -1
            Else
              vRow = vGrid.CurrentDataRow
            End If
            vAddRegDon.DetailLineDefaults("CheckForRowCount") = "Y" 'This will be used to use the same data source for both Detail Line grids (FDE and maintenance) 
            Dim vForm As New frmCardMaintenance(CType(Me.Parent.Parent, ThemedForm), mvContactInfo, CareNetServices.XMLMaintenanceControlTypes.xmctPaymentPlanProducts, vAddRegDon.PPDDataSet, vRow, vAddRegDon.DetailLineDefaults)
            vForm.SetModal = True
            If vForm.ShowDialog = DialogResult.Cancel Then
              vAddRegDon.RefreshDonationBalance()
            End If

        End Select
        If vType <> CareServices.XMLContactDataSelectionTypes.xcdtNone Then
          EditContactData(Nothing, mvContactInfo, vType, True)
          If mvControlType = CareNetServices.FDEControlTypes.AddressDisplay Then
            'Default Address may have changed
            Dim vAddressNumber As Integer = IntegerValue(epl.GetValue("AddressNumber"))
            If mvContactInfo.AddressNumber = vAddressNumber Then
              Dim vContactNumber As Integer = mvContactInfo.ContactNumber
              mvContactInfo = New ContactInfo(vContactNumber)
              mvProductValidation.ContactInfo = mvContactInfo
              If mvContactInfo.AddressNumber <> vAddressNumber Then epl.SetValue("AddressNumber", mvContactInfo.AddressNumber.ToString)
            End If
          End If
          RefreshContactData(mvContactInfo)
        End If
      Else
        ButtonClicked(pParameterName)
      End If
    ElseIf pParameterName = "Edit" Then
      If mvControlType = CareNetServices.FDEControlTypes.AddRegularDonation Then
        Dim vForm As New frmCardMaintenance(CType(Me.Parent.Parent, ThemedForm), Nothing, CareNetServices.XMLMaintenanceControlTypes.xmctPaymentPlanProducts, New DataSet, -1, Nothing)
        Dim vCustomiseMenu As New CustomiseMenu
        vCustomiseMenu.SetContext(vForm, CareServices.XMLMaintenanceControlTypes.xmctPaymentPlanProducts, "")
        vForm.SetCustomiseMenu(vCustomiseMenu, True)
        vForm.SetModal = True
        vForm.ShowDialog()
      End If
    End If
  End Sub

  Private Sub epl_ContactSelected(ByVal sender As Object, ByVal pContactNumber As Integer) Handles epl.ContactSelected
    If mvEditing = False Then FormHelper.ShowContactCardIndex(pContactNumber)
  End Sub

  Private Sub epl_ValidateItem(ByVal sender As Object, ByVal pParameterName As String, ByVal pValue As String, ByRef pValid As Boolean) Handles epl.ValidateItem
    If mvEditing = False Then
      Dim vEPL As EditPanel = DirectCast(sender, EditPanel)

      Select Case mvControlType
        Case CareNetServices.FDEControlTypes.AddDonationCC, CareNetServices.FDEControlTypes.ProductSale
          Select Case pParameterName
            Case "SortCode"
              SetBankDetails(epl, "SortCode", pValue)
            Case "AccountNumber"
              SetBankDetails(epl, "AccountNumber", pValue)
            Case "CreditCardNumber"
              Dim vCCType As EditPanel.CreditCardValidationTypes = EditPanel.CreditCardValidationTypes.ccvtStandard
              Select Case vEPL.ValidCreditCardNumber(vEPL.GetValue(pParameterName), vCCType)
                Case EditPanel.CreditCardValidationStatus.ccvsInvalidNumber
                  pValid = vEPL.SetErrorField(pParameterName, InformationMessages.ImInvalidCardNumber, True)
                Case EditPanel.CreditCardValidationStatus.ccvsNotNumeric
                  pValid = vEPL.SetErrorField(pParameterName, InformationMessages.ImCardNumberNotNumeric, True)
                Case EditPanel.CreditCardValidationStatus.ccvsValid
                  vEPL.SetErrorField(pParameterName, "")
              End Select
            Case "CardExpiryDate"
              If pValue.Length > 0 Then
                If IsNumeric(pValue) Then
                  If (New DateHelper(pValue, DateHelper.DateHelperCardDateType.dhcdtExpiryDate).DateValue) < DateTime.Now Then
                    pValid = vEPL.SetErrorField(pParameterName, GetInformationMessage(InformationMessages.ImExpiryDateMustBeInFuture), True)
                  End If
                Else
                  pValid = vEPL.SetErrorField(pParameterName, GetInformationMessage(InformationMessages.ImInvalidValue), True)
                End If
              End If
            Case "CardStartDate"
              If pValue.Length > 0 Then
                If IsNumeric(pValue) Then
                  If (New DateHelper(pValue, DateHelper.DateHelperCardDateType.dhcdtValidDate).DateValue) = Nothing Then
                    pValid = vEPL.SetErrorField(pParameterName, GetInformationMessage(InformationMessages.ImInvalidCardValidDate), True)
                  End If
                Else
                  pValid = vEPL.SetErrorField(pParameterName, GetInformationMessage(InformationMessages.ImInvalidValue), True)
                End If
              End If
            Case "DeceasedContactNumber"
              Dim vCheckBox As CheckBox = vEPL.FindPanelControl(Of CheckBox)("LineTypeG")
              If vCheckBox.Checked Then
                Dim vRow As DataRow = DataHelper.GetContactItem(CareServices.XMLContactDataSelectionTypes.xcdtContactInformation, IntegerValue(pValue))
                Dim vDeceasedStatus As String = AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.deceased_status)
                If vRow IsNot Nothing Then
                  If vRow("Status").ToString <> vDeceasedStatus Then
                    vEPL.SetErrorField("DeceasedContactNumber", InformationMessages.ImDeceasedContact, True)
                    pValid = False
                  End If
                End If
              End If
          End Select

        Case CareNetServices.FDEControlTypes.AddMemberDD
          Select Case pParameterName
            Case "ContactNumber"
              'This is the member and can not be a joint contact
              If pValue.Length > 0 Then
                Dim vContactInfo As New ContactInfo(IntegerValue(pValue))
                If vContactInfo.ContactType = ContactInfo.ContactTypes.ctJoint Then
                  pValid = vEPL.SetErrorField("ContactNumber", InformationMessages.ImJointContactCannotBeMember, True)
                End If
              End If
          End Select

        Case CareNetServices.FDEControlTypes.AddressDisplay
          If pParameterName = "AddressNumber" Then RaiseEvent AddressChanged(Me, IntegerValue(pValue))

        Case CareNetServices.FDEControlTypes.ContactSelection
          If pParameterName = "ContactNumber" Then RaiseEvent ContactChanged(Me, IntegerValue(pValue))

        Case CareNetServices.FDEControlTypes.AddRegularDonation
          If pParameterName = "OrderDate" Or pParameterName = "StartDate" Then 'pp date = order date

            Dim vDDDate As Date = CDate(vEPL.GetValue("StartDate"))
            Dim vPPDate As Date = CDate(vEPL.GetValue("OrderDate"))

            If (vDDDate < vPPDate) Then vEPL.SetErrorField(pParameterName, InformationMessages.ImInvalidDateRangeDonation, True)
          End If

          RaiseEvent RegDonationAdded(Me)

          If pParameterName = "Balance" Then
            If vEPL.GetValue("Amount").Length > 0 AndAlso DoubleValue(vEPL.GetValue("Amount")) <> DoubleValue(vEPL.GetValue("Balance")) Then
              pValid = False
            End If
          End If
      End Select
    End If
  End Sub

  Protected Sub RaiseValueChangedEvent(ByVal sender As Object, ByVal pParameterName As String, ByVal pValue As String)
    epl_ValueChanged(sender, pParameterName, pValue)
  End Sub

  Protected Sub RaiseSelectedContactChangedEvent(ByVal sender As Object, ByVal pContactNumber As Integer)
    RaiseEvent ContactChanged(sender, pContactNumber)
    RaiseEvent SelectedContactChanged(sender, pContactNumber)
  End Sub

  Protected Sub RaiseEnableOtherModulesEvent(ByVal sender As Object, ByVal pEnable As Boolean)
    RaiseEvent EnableOtherModules(sender, pEnable)
  End Sub

  Protected Sub RaiseFormClosingAllowedEvent(ByVal pValue As Boolean)
    RaiseEvent FormClosingAllowed(pValue)
  End Sub

  Private Sub epl_ValueChanged(ByVal sender As Object, ByVal pParameterName As String, ByVal pValue As String) Handles epl.ValueChanged
    If mvEditing = False Then
      Dim vEPL As EditPanel = DirectCast(sender, EditPanel)

      If pParameterName.Contains("_") Then
        'These are OptionButtons so get the ParameterName without the Value
        If pParameterName.StartsWith("MembershipType") Then pParameterName = "MembershipType"
        If pParameterName.StartsWith("PaymentFrequency") Then pParameterName = "PaymentFrequency"
        If pParameterName.StartsWith("DistributionCode") Then pParameterName = "DistributionCode"
      End If

      Select Case mvControlType
        Case CareNetServices.FDEControlTypes.AddDonationCC, CareNetServices.FDEControlTypes.ProductSale
          Select Case pParameterName
            Case "Amount"
              mvProductValidation.ValueChanged(vEPL, pParameterName, pValue)
            Case "CreditCardNumber"
              If pValue.Length > 0 AndAlso mvContactInfo IsNot Nothing Then
                Dim vList As New ParameterList(True)
                vList("CreditCardNumber") = pValue
                Dim vRow As DataRow = DataHelper.GetContactItem(CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCards, mvContactInfo.ContactNumber, vList, True)
                If vRow IsNot Nothing Then
                  vEPL.SetValue("CardExpiryDate", vRow("ExpiryDate").ToString)
                  vEPL.SetValue("CreditCardType", vRow("CreditCardType").ToString)
                  mvCreditCardDetailsNumber = IntegerValue(vRow("CreditCardDetailsNumber").ToString)
                Else
                  vEPL.SetValue("CardExpiryDate", "")
                End If
              End If
              epl.PanelInfo.PanelItems("CreditCardType").Mandatory = (pValue.Length > 0)
              epl.PanelInfo.PanelItems("CardExpiryDate").Mandatory = (pValue.Length > 0)

            Case "LineTypeG", "LineTypeH", "LineTypeS"
              If BooleanValue(pValue) Then

                Dim vType1 As String = "H"
                Dim vType2 As String = "S"
                Select Case pParameterName.Substring(pParameterName.Length - 1, 1).ToUpper
                  Case "H"
                    vType1 = "G"
                  Case "S"
                    vType2 = "G"
                End Select
                vEPL.SetValue("LineType" & vType1, "N")
                vEPL.SetValue("LineType" & vType2, "N")
                vEPL.EnableControl("DeceasedContactNumber", True)
                Dim vCheckBox As CheckBox = vEPL.FindPanelControl(Of CheckBox)("LineTypeG")
                Dim vString As String = vEPL.FindPanelControl(Of TextLookupBox)("DeceasedContactNumber").Text
                If vCheckBox.Checked AndAlso vString.Length > 0 Then
                  Dim vRow As DataRow = DataHelper.GetContactItem(CareServices.XMLContactDataSelectionTypes.xcdtContactInformation, IntegerValue(vString))
                  Dim vDeceasedStatus As String = AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.deceased_status)
                  If vRow IsNot Nothing Then
                    If vRow("Status").ToString <> vDeceasedStatus Then
                      vEPL.SetErrorField("DeceasedContactNumber", InformationMessages.ImDeceasedContact, True)
                    End If
                  End If
                Else
                  vEPL.SetErrorField("DeceasedContactNumber", "")
                End If

              End If
            Case "PaymentMethod"
              Me.IsCardTransaction = False
              Dim vPaymentMethod As ComboBox = vEPL.FindPanelControl(Of ComboBox)("PaymentMethod")
              If vPaymentMethod IsNot Nothing Then
                Select Case vPaymentMethod.SelectedValue.ToString
                  Case AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cc)
                    Me.IsCardTransaction = True
                    If WebBasedCardAuthoriser.IsAvailable Then
                      Dim fieldList As String() = {"CreditCardType", "CreditCardNumber", "CardExpiryDate", "IssueNumber", "CardStartDate"}
                      vEPL.ClearControlList(fieldList.AsCommaSeperated)
                      For Each fieldName In fieldList
                        epl.PanelInfo.PanelItems(fieldName).Mandatory = False
                        vEPL.SetErrorField(fieldName, "")
                      Next
                      vEPL.EnableControlList(fieldList.AsCommaSeperated, False)
                      vEPL.EnableControl("CardNumberNotRequired", True)
                    Else
                      If mvInitialSettings.Length > 0 Then
                        Dim vList As New ParameterList
                        vList.FillFromValueList(mvInitialSettings)
                        If vList.ContainsKey("OnlineCCAuthorisation") AndAlso BooleanValue(vList("OnlineCCAuthorisation")) Then
                          epl.PanelInfo.PanelItems("SecurityCode").Mandatory = False
                          vEPL.EnableControl("SecurityCode", True)
                        Else
                          vEPL.EnableControl("SecurityCode", False)
                        End If
                      End If
                      vEPL.EnableControlList("CreditCardType,CreditCardNumber,CardExpiryDate,IssueNumber,CardStartDate,CardNumberNotRequired", True)
                      If FindControl(vEPL, "CreditCardNumber", False) IsNot Nothing Then
                        vEPL.PanelInfo.PanelItems("CreditCardNumber").Mandatory = True
                      End If
                    End If
                    vEPL.ClearControlList("SortCode,AccountNumber")
                    vEPL.SetErrorField("SortCode", "")
                    vEPL.SetErrorField("AccountNumber", "")
                    vEPL.EnableControlList("AccountNumber,SortCode", False)
                    RaiseEvent ReferenceMandatory(Me, False)
                  Case AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_po), AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cheque), AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cash) '"PO", "CASH", "CHQ"
                    vEPL.SetErrorField("CreditCardNumber", "")
                    vEPL.SetErrorField("CardStartDate", "")
                    vEPL.SetErrorField("SortCode", "")
                    vEPL.SetErrorField("AccountNumber", "")
                    vEPL.EnableControlList("CreditCardType,CreditCardNumber,CardExpiryDate,IssueNumber,CardStartDate,SecurityCode,SortCode,AccountNumber,CardNumberNotRequired", False)
                    vEPL.ClearControlList("CardNumberNotRequired,SortCode,AccountNumber,CreditCardType,CreditCardNumber,CardExpiryDate,IssueNumber,CardStartDate,SecurityCode")
                    If vPaymentMethod.SelectedValue.ToString = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cheque) Then
                      vEPL.EnableControl("SortCode", True)
                      vEPL.EnableControl("AccountNumber", True)
                    End If
                    If vPaymentMethod.SelectedValue.ToString = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cheque) Then
                      RaiseEvent ReferenceMandatory(Me, True)
                    Else
                      RaiseEvent ReferenceMandatory(Me, False)
                    End If
                End Select
              End If
            Case "CardNumberNotRequired"
              Dim vCardNumberNotRequired As CheckBox = vEPL.FindPanelControl(Of CheckBox)("CardNumberNotRequired")
              If vCardNumberNotRequired.Enabled Then
                If vCardNumberNotRequired.Checked Then
                  If FindControl(vEPL, "CreditCardNumber", False) IsNot Nothing Then
                    vEPL.PanelInfo.PanelItems("CreditCardNumber").Mandatory = False
                    vEPL.SetErrorField("CreditCardNumber", "")
                    vEPL.EnableControlList("CreditCardType,CreditCardNumber,CardExpiryDate,IssueNumber,CardStartDate,SecurityCode", False)
                    vEPL.ClearControlList("CreditCardType,CreditCardNumber,CardExpiryDate,IssueNumber,CardStartDate,SecurityCode")
                    vEPL.SetErrorField("CreditCardType", "")
                    vEPL.SetErrorField("CreditCardNumber", "")
                    vEPL.SetErrorField("CardExpiryDate", "")
                    vEPL.SetErrorField("IssueNumber", "")
                    vEPL.SetErrorField("CardStartDate", "")
                    vEPL.SetErrorField("SecurityCode", "")
                  End If
                Else
                  If FindControl(vEPL, "CreditCardNumber", False) IsNot Nothing Then
                    vEPL.PanelInfo.PanelItems("CreditCardNumber").Mandatory = Not WebBasedCardAuthoriser.IsAvailable
                    vEPL.EnableControlList("CreditCardType,CreditCardNumber,CardExpiryDate,IssueNumber,CardStartDate", Not WebBasedCardAuthoriser.IsAvailable)
                    vEPL.SetErrorField("CreditCardNumber", "")
                    vEPL.SetErrorField("CardExpiryDate", "")
                    vEPL.SetErrorField("IssueNumber", "")
                    vEPL.SetErrorField("CardStartDate", "")
                    If mvInitialSettings.Length > 0 Then
                      Dim vList As New ParameterList
                      vList.FillFromValueList(mvInitialSettings)
                      If vList.ContainsKey("OnlineCCAuthorisation") AndAlso BooleanValue(vList("OnlineCCAuthorisation")) Then
                        epl.PanelInfo.PanelItems("SecurityCode").Mandatory = False
                        vEPL.EnableControl("SecurityCode", Not WebBasedCardAuthoriser.IsAvailable)
                      Else
                        vEPL.EnableControl("SecurityCode", False)
                      End If
                    End If
                End If
                End If
              End If
          End Select
        Case CareNetServices.FDEControlTypes.AddMemberDD, CareNetServices.FDEControlTypes.AddRegularDonation
          Select Case pParameterName
            'BR17159 
            Case "Balance"
              'if there is only 1 row in the grid AND the product rate is zero then update row with the value from the balance field
              Dim vGrid As DisplayGrid = DirectCast(FindControl(epl, "DetailLines"), DisplayGrid)

              If vGrid.RowCount = 1 Then
                If vGrid.GetValue(0, "Rate") = "0" Then
                  Dim vBalance As String = epl.PanelInfo.PanelItems("Balance").LastValue.ToString()
                  vGrid.SetValue(0, "Balance", vBalance)
                  vGrid.SetValue(0, "Amount", epl.PanelInfo.PanelItems("Amount").LastValue)
                End If

              ElseIf vGrid.RowCount = 0 Then
                'and default rate of default product is zero
                'add a new regular donation using defaults
                Dim vBalance As String = epl.PanelInfo.PanelItems("Balance").LastValue.ToString()
                Dim vAddRegDon As FDEAddRegularDonation = DirectCast(Me, FDEAddRegularDonation)
                vAddRegDon.SetDefaults()
                If vGrid.RowCount = 1 Then
                  vGrid.SetValue(0, "Balance", vBalance)
                  epl.SetValue("Balance", vBalance)
                End If
              ElseIf vGrid.RowCount > 1 Then
                Dim vAddRegDon As FDEAddRegularDonation = DirectCast(Me, FDEAddRegularDonation)
                vAddRegDon.RefreshDonationBalance()
              End If

            Case "AccountNumber"
              If (AppValues.DefaultCountryCode <> "CH" AndAlso AppValues.DefaultCountryCode <> "NL") Then
                If pValue.Length > 0 AndAlso pValue.Length < 8 Then
                  pValue = pValue.PadLeft(8, "0"c)
                  vEPL.SetValue("AccountNumber", pValue)
                End If
              End If
              If mvControlType = CareNetServices.FDEControlTypes.AddRegularDonation And pValue.Length > 0 AndAlso epl.GetValue("ReasonForDespatch") = AppValues.ControlValue(AppValues.ControlValues.o_reason) Then
                'Only change the ReasonForDespatch to DD Reason if the user has not changed it manually (the default is set to Other Reason)
                epl.SetValue("ReasonForDespatch", AppValues.ControlValue(AppValues.ControlValues.dd_reason))
              End If
              SetBankDetails(vEPL, pParameterName, pValue)
              vEPL.PanelInfo.PanelItems("SortCode").Mandatory = (pValue.Length > 0)
              vEPL.PanelInfo.PanelItems("AccountNumber").Mandatory = (pValue.Length > 0)
              vEPL.PanelInfo.PanelItems("AccountName").Mandatory = (pValue.Length > 0)
              vEPL.PanelInfo.PanelItems("BranchName").Mandatory = (pValue.Length > 0)
              vEPL.PanelInfo.PanelItems("StartDate").Mandatory = (pValue.Length > 0)
            Case "BankAccount"
              If AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.auto_pay_claim_date_method).ToUpper.Equals("D") Then
                Dim vCombo As ComboBox = vEPL.FindPanelControl(Of ComboBox)("ClaimDay")
                If vCombo IsNot Nothing Then
                  Dim vDT As DataTable = DirectCast(vCombo.DataSource, DataTable)
                  DefaultClaimDay(vDT, pValue, "DD")
                  vEPL.ValidateControl("ClaimDay")
                End If
              End If
              SetBankDetails(vEPL, pParameterName, pValue)
              Dim vBaseDate As Date = Today
              If mvControlType.Equals(CareNetServices.FDEControlTypes.AddMemberDD) Then
                Dim vJoined As String = epl.GetValue("Joined")
                If IsDate(vJoined) Then vBaseDate = DateValue(vJoined)
              End If
              SetDDStartDate(vBaseDate, "StartDate")

            Case "GiftMembership", "MemberContactNumber", "MembershipType", "Joined"
              ProcessMemberValuesChanged(vEPL, pParameterName, pValue)

            Case "SortCode"
              SetBankDetails(vEPL, pParameterName, pValue)
            Case "Amount"
              If mvControlType = CareNetServices.FDEControlTypes.AddRegularDonation Then vEPL.SetValue("Balance", pValue)
            Case "OrderDate"
              'payment plan start date, dd start date must be before start date
              vEPL.ValidateControl("OrderDate")
          End Select
          If pParameterName = "AccountNumber" OrElse pParameterName = "SortCode" Then
            vEPL.PanelInfo.PanelItems("SortCode").Mandatory = (pValue.Length > 0)
            vEPL.PanelInfo.PanelItems("AccountNumber").Mandatory = (pValue.Length > 0)
            vEPL.PanelInfo.PanelItems("BranchName").Mandatory = (pValue.Length > 0)
            vEPL.PanelInfo.PanelItems("AccountName").Mandatory = (pValue.Length > 0)
            If AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.auto_pay_claim_date_method) = "D" Then
              vEPL.PanelInfo.PanelItems("ClaimDay").Mandatory = (pValue.Length > 0)
            End If
          End If

        Case CareNetServices.FDEControlTypes.AddTransactionDetails
          Select Case pParameterName
            Case "Source"
              Dim vSource As TextLookupBox = vEPL.FindPanelControl(Of TextLookupBox)("Source")
              vEPL.SetValue("Mailing", vSource.GetDataRowItem("ThankYouLetter"))
              RaiseEvent SourceChanged(Me, pValue, vSource.GetDataRowItem("DistributionCode"), vSource.GetDataRowItem("IncentiveScheme"))

            Case "TransactionDate"
              If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.opt_fp_prevent_future_date) AndAlso _
                Date.Compare(CDate(pValue), Today) > 0 Then
                vEPL.SetErrorField("TransactionDate", InformationMessages.ImTransactionDateInFuture)
              Else
                RaiseEvent TransactionDateChanged(Me, pValue)
              End If
            Case "TransactionAmount"
              If FindControl(vEPL, "EligibleForGiftAid", False) IsNot Nothing Then
                If pValue >= AppValues.ControlValue(AppValues.ControlTables.covenant_controls, AppValues.ControlValues.gift_aid_minimum) Then
                  vEPL.SetValue("EligibleForGiftAid", "Y")
                Else
                  vEPL.SetValue("EligibleForGiftAid", "N")
                End If
              End If
          End Select
        Case CareNetServices.FDEControlTypes.ContactSelection
          Select Case pParameterName
            Case "ContactNumber"
              RaiseEvent BankDetailsChange(Me)
          End Select
        Case CareNetServices.FDEControlTypes.Telemarketing
          Select Case pParameterName
            Case "Campaign"
              Dim vRow As DataRow = epl.FindPanelControl(Of TextLookupBox)("Campaign").GetDataRow
              If vRow IsNot Nothing Then
                If pValue.Length > 0 Then 'not to clear the appeal and segment when defaulting the values
                  epl.SetValue("Appeal", "", , , , True)
                  epl.SetValue("Segment", "", , , , True)
                  epl.EnableControl("GetContact", False)
                Else
                  epl.EnableControl("Campaign", False)  'called from SetDefaults. The default value is set, so disable the control
                End If
                'Set the default outcome
                Dim vCampaignTopic As String = ""
                If vRow.Table.Columns.Contains("Topic") Then vCampaignTopic = vRow("Topic").ToString
                If vCampaignTopic.Length = 0 Then vCampaignTopic = AppValues.ConfigurationValue(AppValues.ConfigurationValues.phone_out_outcome_topic)
                epl.FindPanelControl(Of TextLookupBox)("Outcome").FillComboWithRestriction(vCampaignTopic)

                'Set the script
                If vRow.Table.Columns.Contains("HtmlText") Then
                  Dim vScriptBrowser As WebBrowser = DirectCast(FindControl(epl, "HtmlText"), WebBrowser)
                  vScriptBrowser.Document.Write(vRow("HtmlText").ToString)
                  vScriptBrowser.Refresh(WebBrowserRefreshOption.Normal) 'If the control is populated once then although calling OpenNew clears the data but sometime does not refresh it
                End If
              End If
            Case "Appeal"
              epl.SetValue("Segment", "", , , , True)
              epl.EnableControl("GetContact", False)
            Case "Segment"
              Dim vRow As DataRow = epl.FindPanelControl(Of TextLookupBox)("Segment").GetDataRow
              If vRow Is Nothing Then
                'Disable all controls
                epl.EnableControls(False, True)
                epl.EnableControl("Campaign", Not mvDefaultList.ContainsKey("Campaign"))
                epl.EnableControl("Appeal", Not mvDefaultList.ContainsKey("Appeal"))
                epl.EnableControlList("Segment,HtmlText", True)
              Else
                epl.EnableControls(False, True)
                epl.EnableControl("Campaign", Not mvDefaultList.ContainsKey("Campaign"))
                epl.EnableControl("Appeal", Not mvDefaultList.ContainsKey("Appeal"))
                epl.EnableControl("Segment", Not mvDefaultList.ContainsKey("Segment"))
                epl.EnableControlList("HtmlText,GetContact", True)
                If vRow("Source").ToString.Length > 0 Then RaiseEvent SourceChanged(Me, vRow("Source").ToString, "", "")
              End If
            Case "TopicGroup"
              Dim vDataSheet As TopicDataSheet = DirectCast(FindControl(Me, "TopicDataSheet", False), TopicDataSheet)
              vDataSheet.Init(New DocumentInfo(0, ""), pValue, Nothing, False)
            Case "Outcome"
              Dim vCallBackMinutes As String = epl.FindPanelControl(Of TextLookupBox)(pParameterName).GetDataRowItem("CallBackMinutes")
              epl.EnableControl("CallBackTime", vCallBackMinutes.Length > 0 AndAlso BooleanValue(mvDefaultList("OverwriteCallBackTime")))
              If vCallBackMinutes.Length > 0 Then

                'Get the current date as Required Date
                Dim vRequiredDate As Date = Now
                If (Weekday(vRequiredDate) = vbSaturday) Then vRequiredDate = vRequiredDate.AddDays(1)
                If (Weekday(vRequiredDate) = vbSunday) Then vRequiredDate = vRequiredDate.AddDays(1)

                'Get the current time as Required Time

                Dim vRequiredTime As Date = TimeValue(vRequiredDate.ToShortTimeString)
                Dim vStartOfDay As Date = TimeValue(AppValues.ControlValue(AppValues.ControlValues.start_of_day))
                Dim vEndOfDay As Date = TimeValue(AppValues.ControlValue(AppValues.ControlValues.end_of_day))
                If vRequiredTime < vStartOfDay Then vRequiredTime = vStartOfDay
                If vRequiredTime > vEndOfDay Then vRequiredTime = vEndOfDay

                'Check for Lunch Time
                Dim vStartOfLunch As Date
                Dim vEndOfLunch As Date
                If AppValues.ControlValue(AppValues.ControlValues.start_of_lunch).Length > 0 Then
                  vStartOfLunch = TimeValue(AppValues.ControlValue(AppValues.ControlValues.start_of_lunch))
                  vEndOfLunch = TimeValue(AppValues.ControlValue(AppValues.ControlValues.end_of_lunch))
                End If
                If vRequiredTime > vStartOfLunch And vRequiredTime < vEndOfLunch Then vRequiredTime = vEndOfLunch
                Dim vConsiderLunch As Boolean
                If vRequiredTime < vEndOfLunch Then vConsiderLunch = True

                'Combined Required Date and Required Time
                vRequiredDate = Date.Parse(vRequiredDate.ToShortDateString & " " & vRequiredTime.ToShortTimeString)

                'Now add time period which should be one of day,hour OR minutes
                Dim vTimeSpan As TimeSpan = TimeSpan.FromMinutes(DoubleValue(vCallBackMinutes))
                Dim vDurationDays As Integer = vTimeSpan.Days
                Dim vDurationHours As Integer = vTimeSpan.Hours
                Dim vDurationMinutes As Integer = vTimeSpan.Minutes

                If vDurationDays > 0 Then
                  vRequiredDate = vRequiredDate.AddDays(vDurationDays)
                  If (Weekday(vRequiredDate) = vbSaturday) Then vRequiredDate = vRequiredDate.AddDays(1)
                  If (Weekday(vRequiredDate) = vbSunday) Then vRequiredDate = vRequiredDate.AddDays(1)
                End If
                If vDurationHours > 0 Or vDurationMinutes > 0 Then
                  vRequiredTime = vRequiredTime.AddHours(vDurationHours)
                  vRequiredTime = vRequiredTime.AddMinutes(vDurationMinutes)
                  If vConsiderLunch = True AndAlso vRequiredTime < vEndOfDay Then
                    If vRequiredTime > vEndOfLunch Then
                      vRequiredTime = vRequiredTime.AddMinutes(vEndOfLunch.Subtract(vStartOfLunch).TotalMinutes)  'Only do this if the required time is for the same day
                    ElseIf vRequiredTime > vStartOfLunch And vRequiredTime < vEndOfLunch Then
                      vRequiredTime = vEndOfLunch 'check for lunch break
                    End If
                  End If
                  While vRequiredTime >= vEndOfDay
                    vRequiredTime = vStartOfDay.AddMinutes(vRequiredTime.Subtract(vEndOfDay).TotalMinutes)
                    If vRequiredTime > vStartOfLunch And vRequiredTime < vEndOfLunch Then vRequiredTime = vEndOfLunch 'check for lunch break
                    vRequiredDate = vRequiredDate.AddDays(1)
                    If (Weekday(vRequiredDate) = vbSaturday) Then vRequiredDate = vRequiredDate.AddDays(1)
                    If (Weekday(vRequiredDate) = vbSunday) Then vRequiredDate = vRequiredDate.AddDays(1)
                  End While
                End If
                'Now convert date and time back to a full date/time value
                epl.SetValue("CallBackTime", Date.Parse(vRequiredDate.ToShortDateString & " " & vRequiredTime.ToString("HH:mm")).ToString)
              End If
          End Select
      End Select

      '************************************************************************************************************************
      'Now process items which appear on many modules
      '************************************************************************************************************************
      mvProductValidation.ValueChanged(vEPL, pParameterName, pValue)

    End If
  End Sub

  Private Sub epl_ValidateAllItems(ByVal sender As Object, ByVal pList As CDBNETCL.ParameterList, ByRef pValid As Boolean) Handles epl.ValidateAllItems
    If mvEditing = False Then
      Dim vEPL As EditPanel = DirectCast(sender, EditPanel)

      Select Case mvControlType
        Case CareNetServices.FDEControlTypes.AddDonationCC, CareNetServices.FDEControlTypes.ProductSale
          'Amount
          If mvControlType = CareNetServices.FDEControlTypes.AddDonationCC Then 'product sale allows zero 
            If vEPL.GetValue("Amount").Length > 0 AndAlso DoubleValue(vEPL.GetValue("Amount")) = 0 Then
              pValid = vEPL.SetErrorField("Amount", InformationMessages.ImDonationAmountCannotBeZero)
            End If
          End If
          'ExpiryDate
          If vEPL.GetValue("CardExpiryDate").Length > 0 Then
            Dim vDateTime As DateTime = New DateHelper(vEPL.GetValue("CardExpiryDate"), DateHelper.DateHelperCardDateType.dhcdtExpiryDate).DateValue
            If Year(vDateTime) = 1 Then
              pValid = False
              vEPL.SetErrorField("CardExpiryDate", InformationMessages.ImExpiryDateMustBeInFuture)
            End If
          End If
          'SecurityCode
          Dim vTextBox As TextBox = vEPL.FindPanelControl(Of TextBox)("SecurityCode")
          If vTextBox.Enabled Then
            If vTextBox.Text.Length > 0 AndAlso vTextBox.Text.Length < 3 Then
              pValid = False
              vEPL.SetErrorField("SecurityCode", InformationMessages.ImInvalidCardSecurityCode)
            End If
          End If
          'ValidDate
          If vEPL.GetValue("CardStartDate").Length > 0 Then
            Dim vDateTime As DateTime = New DateHelper(vEPL.GetValue("CardStartDate"), DateHelper.DateHelperCardDateType.dhcdtValidDate).DateValue
            If Year(vDateTime) = 1 Then
              pValid = False
              vEPL.SetErrorField("CardStartDate", InformationMessages.ImInvalidCardValidDate, True)
            End If
          End If
          'PaymentMethod
          Dim vPaymentMethod As ComboBox = vEPL.FindPanelControl(Of ComboBox)("PaymentMethod")
          If vPaymentMethod IsNot Nothing AndAlso vPaymentMethod.SelectedValue.ToString = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cheque) Then
            RaiseEvent ReferenceMandatory(Me, True)
          End If
        Case CareNetServices.FDEControlTypes.AddMemberDD
          ValidateMembersPage(vEPL, pValid)

        Case CareNetServices.FDEControlTypes.AddRegularDonation
          Dim vAddRegDon As FDEAddRegularDonation = DirectCast(Me, FDEAddRegularDonation)
          If DataHelper.GetTableFromDataSet(vAddRegDon.PPDDataSet).Rows.Count > 0 Then
            Dim vBalance As Double
            For Each vRow As DataRow In DataHelper.GetTableFromDataSet(vAddRegDon.PPDDataSet).Rows
              vBalance += DoubleValue(vRow("Balance").ToString)
            Next
            If vBalance <> epl.GetDoubleValue("Balance") Then
              pValid = False
              vEPL.SetErrorField("Balance", GetInformationMessage(InformationMessages.ImPPBalanceNotMatched, vBalance.ToString("0.00")), True)
            End If
          End If
      End Select

      'Handle controls on multiple pages
      mvProductValidation.ValidateAllItems(vEPL, pList, pValid)
    End If
  End Sub

  Private Sub epl_GetCodeRestrictions(ByVal sender As Object, ByVal pParameterName As String, ByVal pList As CDBNETCL.ParameterList) Handles epl.GetCodeRestrictions
    GetCodeRestrictions(pParameterName, pList)
  End Sub

  Private Sub epl_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles epl.DoubleClick
    If mvEditing = True Then
      CustomiseControl()
    End If
  End Sub

  Friend Overridable Sub GetCodeRestrictions(ByVal pParameterName As String, ByVal pList As CDBNETCL.ParameterList)
  End Sub

#End Region

#Region " DragDrop and sizing Support "

  Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
    MyBase.OnMouseDown(e)
    If mvEditing AndAlso (e.Button = System.Windows.Forms.MouseButtons.Left And e.Y + 4 > Me.Height - 4) Then
      mvSize = Me.Size
      mvMousePos = e.Location
      mvSizing = True
    End If
  End Sub

  Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
    MyBase.OnMouseUp(e)
    mvSizing = False
    If Me.Cursor = Cursors.SizeNS Then Me.Cursor = Cursors.Default
  End Sub

  Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
    MyBase.OnMouseMove(e)
    If mvEditing Then
      If mvSizing Then
        Me.Height = mvSize.Height + e.Y - mvMousePos.Y
        mvRequiredHeight = Me.Height
        Me.Invalidate()
        mvSizeChanged = True
      ElseIf e.Y + 4 >= Me.Height - Me.Padding.Bottom Then
        Me.Cursor = Cursors.SizeNS
      Else
        If Me.Cursor = Cursors.SizeNS Then Me.Cursor = Cursors.Default
      End If
    End If
  End Sub

  Private Sub epl_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles epl.MouseMove
    If Me.Cursor = Cursors.SizeNS Then Me.Cursor = Cursors.Default
  End Sub

#End Region

#Region " Menu Handling "

  Private Sub mvFDEControlMenu_MenuSelected(ByVal pMenuItem As System.Windows.Forms.ToolStripMenuItem, ByVal pItem As FDEControlMenu.FDEControlmenuItems) Handles mvFDEControlMenu.MenuSelected
    Dim vBusyCursor As New BusyCursor()
    Try
      Select Case pItem
        Case FDEControlMenu.FDEControlmenuItems.Customise
          CustomiseControl()
        Case FDEControlMenu.FDEControlmenuItems.Revert
          If ShowQuestion(QuestionMessages.QmRevertModule, MessageBoxButtons.YesNo) = DialogResult.Yes Then
            RefreshControl(True)
          End If
        Case FDEControlMenu.FDEControlmenuItems.Parameters
          If ShowQuestion(QuestionMessages.QmChangeFDEParameters, MessageBoxButtons.YesNo) = DialogResult.Yes Then
            If DataHelper.DisplayControlParameters(Me.ParentForm, Me.UserControlName, mvInitialParameters, mvDefaultParameters, mvInitialSettings, mvDefaultSettings) Then
              'Parameters have been updated, need to remove existing FpControls and add new ones
              'First update the FdePageItem
              Dim vList As New ParameterList(True, True)
              vList.IntegerValue("FdePageItemNumber") = mvFDEPageItemNumber
              vList("InitialParameters") = mvInitialSettings
              vList("DefaultParameters") = mvDefaultSettings
              DataHelper.UpdateFastDataEntryData(CareNetServices.XMLFastDataEntryTypes.fdePageControls, vList)
              RefreshControl(True)
            End If
          End If
        Case FDEControlMenu.FDEControlmenuItems.Delete
          Dim vDelete As Boolean = True
          If My.Settings.ConfirmDelete AndAlso ShowQuestion(QuestionMessages.QmConfirmDeleteFDEModule, MessageBoxButtons.OKCancel) = DialogResult.Cancel Then vDelete = False
          If vDelete Then RaiseEvent DeleteControl(Me)
      End Select
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    Finally
      vBusyCursor.Dispose()
    End Try
  End Sub

#End Region

#Region " Private Methods "

  Protected Function GetParameterListFromSettings(ByVal pSettings As String) As ParameterList
    Dim vList As New ParameterList
    If pSettings.Length > 0 Then vList.FillFromValueList(pSettings)
    Return vList
  End Function

  Private Function GetControlParameters() As ParameterList
    Dim vList As New ParameterList(True, True)
    If mvFDEPageItemNumber > 0 Then vList.IntegerValue("FdePageItemNumber") = mvFDEPageItemNumber
    vList.IntegerValue("FdePageNumber") = mvFDEPageNumber
    vList.IntegerValue("SequenceNumber") = mvSequenceNumber
    vList("FdeUserControl") = mvUserControlName
    vList.IntegerValue("FdeItemHeight") = Me.Height
    vList.IntegerValue("FdeItemWidth") = Me.Width
    If mvInitialSettings.Length > 0 Then vList("InitialParameters") = mvInitialSettings
    If mvDefaultSettings.Length > 0 Then vList("DefaultParameters") = mvDefaultSettings
    Return vList
  End Function

  Private Sub CustomiseControl()
    Dim vForm As New frmModuleContent(mvFDEPageItemNumber, mvPageType, frmModuleContent.ModuleContentUsage.FastDataEntry)
    If vForm.ShowDialog(Me) = DialogResult.OK Then
      RefreshControl(False)
    End If
  End Sub

  Private Sub RefreshControl(ByVal pRevert As Boolean)
    Dim vList As New ParameterList(True, True)
    If pRevert Then
      'Revert FpControls
      vList.IntegerValue("FdePageItemNumber") = mvFDEPageItemNumber
      DataHelper.UpdateFastDataEntryData(CareNetServices.XMLFastDataEntryTypes.fdePageControlItems, vList)
    End If
    epl.Visible = False
    vList = GetControlParameters()
    Dim vDT As DataTable = DataHelper.GetFastDataEntryControlItems(vList)
    Dim vPanelItems As New PanelItems(mvUserControlName, vDT)
    vPanelItems.ResetDisplayGridHeight = mvControlType <> CareNetServices.FDEControlTypes.AddRegularDonation
    ResetPanelItems(vPanelItems)
    epl.Controls.Clear()
    epl.Init(New EditPanelInfo(mvUserControlName, vPanelItems))
    ChangeControlAnchors()
    Me.Height = epl.RequiredHeight
    ResizeControl(Me.Width)
    epl.Visible = True
    Save(mvSequenceNumber)
    If mvEditing = False Then SetDefaults()
  End Sub

  Protected Overridable Sub ProcessMemberValuesChanged(ByVal pEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String)
  End Sub

  Protected Overridable Sub ValidateMembersPage(ByVal pEPL As EditPanel, ByRef pValid As Boolean)
  End Sub

  Protected Sub SetValueRaiseChanged(ByVal pEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String, Optional ByVal pDisable As Boolean = False, Optional ByVal pUpdateLastValue As Boolean = False) Handles mvProductValidation.SetValueRaisedEvent
    pEPL.SetValue(pParameterName, pValue, pDisable, , , pUpdateLastValue)
    epl_ValueChanged(pEPL, pParameterName, pValue)
  End Sub

  Protected Sub SetBankDetails(ByVal pEPL As EditPanel, ByVal pParameterName As String, ByVal pValue As String)
    Dim vPayerContactNumber As Integer
    Dim vPayerAddressNumber As Integer
    If mvContactInfo IsNot Nothing Then
      vPayerContactNumber = mvContactInfo.ContactNumber
      vPayerAddressNumber = mvContactInfo.SelectedAddressNumber
    End If
    Dim vContactNumber As Integer = vPayerContactNumber
    AppHelper.SetBankDetails(pEPL, pParameterName, pValue, False, mvBankDetailsNumber, mvCreateContactAccount, mvNewBank, vPayerContactNumber, vPayerAddressNumber)
    Dim vContactInfo As ContactInfo = mvContactInfo
    If vContactInfo IsNot Nothing AndAlso (vPayerContactNumber > 0 AndAlso vContactInfo.ContactNumber <> vPayerContactNumber) Then
      vContactInfo = New ContactInfo(vPayerContactNumber)
    End If
    If vContactInfo IsNot Nothing Then  'FDE contact has not been selected
      If vPayerAddressNumber > 0 Then vContactInfo.SelectedAddressNumber = vPayerAddressNumber
      If pParameterName = "AccountNumber" AndAlso vPayerContactNumber > 0 AndAlso FindControl(pEPL, "AccountName", False) IsNot Nothing Then
        If pEPL.GetValue("AccountName").Length = 0 Then pEPL.SetValue("AccountName", vContactInfo.ContactName)
      End If
    End If
    If mvControlType = CareNetServices.FDEControlTypes.AddDonationCC Or mvControlType = CareNetServices.FDEControlTypes.ProductSale AndAlso vContactNumber <> vPayerContactNumber Then
      RaiseEvent ContactChanged(Me, vPayerContactNumber)
      RaiseEvent SelectedContactChanged(Me, vPayerContactNumber)
    End If
  End Sub

  Private Sub ResetPanelItems(ByVal pPanelItems As PanelItems)
    Select Case mvControlType
      Case CareNetServices.FDEControlTypes.AddDonationCC, CareNetServices.FDEControlTypes.ProductSale
        pPanelItems("DeceasedContactNumber").SetValidationData("contacts", "contact_number")
      Case CareNetServices.FDEControlTypes.AddMemberDD, CareNetServices.FDEControlTypes.AddRegularDonation
        If mvControlType = CareNetServices.FDEControlTypes.AddMemberDD Then
          pPanelItems("MemberContactNumber").SetValidationData("contacts", "contact_number")
          pPanelItems("MemberAddressNumber").SetValidationData("addresses", "address_number")
          Dim vMandatory As Boolean = False
          Dim vList As New ParameterList
          vList.FillFromValueList(mvDefaultSettings)
          If vList.ContainsKey("MembershipType") = False Then vMandatory = True
          pPanelItems("Joined").Mandatory = vMandatory
        Else
          pPanelItems("Balance").Mandatory = False
        End If
        pPanelItems("ClaimDay").Mandatory = True
        pPanelItems("SortCode").Mandatory = False
        pPanelItems("AccountNumber").Mandatory = False
        pPanelItems("AccountNumber").EntryLength = 8
        pPanelItems("BranchName").Mandatory = False
        pPanelItems("AccountName").Mandatory = False
      Case CareNetServices.FDEControlTypes.AddressDisplay
        pPanelItems("AddressNumber").SetValidationData("addresses", "address_number")
    End Select
  End Sub

  Private Sub ChangeControlAnchors()
    Select Case mvControlType
      Case CareNetServices.FDEControlTypes.ActivityDisplay, CareNetServices.FDEControlTypes.AddressDisplay, CareNetServices.FDEControlTypes.CommunicationsDisplay, _
           CareNetServices.FDEControlTypes.GiftAidDisplay, CareNetServices.FDEControlTypes.SuppressionDisplay
        Dim vControl As Control = FindControl(epl, "Edit", False)
        If vControl IsNot Nothing AndAlso TypeOf (vControl) Is Button Then
          vControl.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        End If
    End Select
    Select Case mvControlType
      Case CareNetServices.FDEControlTypes.AddressDisplay
        Dim vControl As Control = FindControl(epl, "AddressNumber", False)
        If vControl IsNot Nothing AndAlso TypeOf (vControl) Is TextLookupBox Then
          vControl.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        End If
      Case CareNetServices.FDEControlTypes.ContactSelection
        Dim vControl As Control = FindControl(epl, "ContactNumber", False)
        If vControl IsNot Nothing AndAlso TypeOf (vControl) Is TextLookupBox Then
          vControl.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        End If
    End Select
  End Sub

  Protected Function GetOptionButtonValue(ByVal pParameterName As String) As String
    Dim vValue As String = ""
    For Each vPanelItem As PanelItem In epl.PanelInfo.PanelItems
      If vPanelItem.ControlType = PanelItem.ControlTypes.ctOptionButton AndAlso vPanelItem.OptionParameterName = pParameterName Then
        vValue = epl.GetValue(pParameterName & "_" & vPanelItem.OptionButtonValue)
        If vValue.Length > 0 Then Exit For
      End If
    Next
    Return vValue
  End Function

  Protected Function GetMandatoryOptionButton(ByVal pList As ParameterList, ByRef pParameterName As String) As Boolean
    Dim vValid As Boolean = True
    Dim vValue As String = ""
    If pList.ContainsKey(pParameterName) Then vValue = pList(pParameterName)
    If vValue.Length = 0 Then vValue = GetOptionButtonValue(pParameterName)
    If vValue.Length > 0 Then
      pList(pParameterName) = vValue
    Else
      'See if item is mandatory and if so, return OptionButton name (for first item) plus vValid as False
      For Each vPanelItem As PanelItem In epl.PanelInfo.PanelItems
        If vPanelItem.ControlType = PanelItem.ControlTypes.ctOptionButton AndAlso vPanelItem.OptionParameterName = pParameterName Then
          If vPanelItem.Visible = True AndAlso vPanelItem.Mandatory = True Then
            pParameterName &= "_" & vPanelItem.OptionButtonValue
            vValid = False
            Exit For
          End If
        End If
      Next
    End If
    Return vValid
  End Function

  Protected Sub ClearOptionButtonError(ByVal pParameterName As String)
    For Each vPanelItem As PanelItem In epl.PanelInfo.PanelItems
      If vPanelItem.ControlType = PanelItem.ControlTypes.ctOptionButton AndAlso vPanelItem.OptionParameterName = pParameterName Then
        If vPanelItem.Visible = True AndAlso vPanelItem.Mandatory = True Then
          epl.SetErrorField(pParameterName & "_" & vPanelItem.OptionButtonValue, "")
          Exit For
        End If
      End If
    Next
  End Sub

  Public Overridable Overloads Property Enabled As Boolean
    Get
      Return MyBase.Enabled
    End Get
    Set(ByVal value As Boolean)
      MyBase.Enabled = value
    End Set
  End Property

  Private isCardTransactionField As Boolean = False
  Public Property IsCardTransaction As Boolean
    Get
      Return isCardTransactionField
    End Get
    Private Set(value As Boolean)
      isCardTransactionField = value
    End Set
  End Property
#End Region
End Class
