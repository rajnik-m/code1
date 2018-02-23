Namespace Access
  Public Class TraderControl

    Private Enum TraderControlFields
      'Attributes
      tcfControlType = 1
      tcfControlTop
      tcfControlLeft
      tcfControlWidth
      tcfControlHeight
      tcfControlCaption
      tcfCaptionWidth
      tcfHelpText
      tcfVisible
      tcfTableName
      tcfAttributeName
      tcfFieldType
      tcfEntryLength
      tcfNullsInvalid
      tcfPattern
      tcfValidationTable
      tcfValidationAttribute
      tcfCaseConversion
      tcfMinimumValue
      tcfMaximumValue
      tcfContactGroup
      tcfParameterName
      tcfDefaultValue
      'Other properties
      tcfValue
      tcfDescription
      tcfTabIndex
      tcfContactType
      tcfLastValue
      tcfLastEnabled
      tcfValRequired
      tcfNoDescription
      tcfIndex 'The Index in the Controls Collection
      tcfPageIndex 'The Index of the Control on the Page
      tcfPageType
      tcfPageCode
    End Enum

    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private Const SAGEPAYHOSTED As String = "SAGEPAYHOSTED"


    Private Sub InitClassFields()

      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .Add("control_type")
          .Add("control_top", CDBField.FieldTypes.cftLong)
          .Add("control_left", CDBField.FieldTypes.cftLong)
          .Add("control_width", CDBField.FieldTypes.cftLong)
          .Add("control_height", CDBField.FieldTypes.cftLong)
          .Add("control_caption")
          .Add("caption_width", CDBField.FieldTypes.cftInteger)
          .Add("help_text")
          .Add("visible")
          .Add("table_name")
          .Add("attribute_name")
          .Add("type")
          .Add("entry_length", CDBField.FieldTypes.cftInteger)
          .Add("nulls_invalid")
          .Add("pattern")
          .Add("validation_table")
          .Add("validation_attribute")
          .Add("case")
          .Add("minimum_value")
          .Add("maximum_value")
          .Add("contact_group")
          .Add("parameter_name")
          .Add("default_value")
          .Add("value")
          .Add("description")
          .Add("tab_index")
          .Add("contact_type", CDBField.FieldTypes.cftLong)
          .Add("last_value")
          .Add("last_enabled")
          .Add("val_required")
          .Add("no_description")
          .Add("index", CDBField.FieldTypes.cftLong)
          .Add("page_index", CDBField.FieldTypes.cftLong)
          .Add("page_type")
          .Add("page_code", CDBField.FieldTypes.cftInteger)

          .Item(TraderControlFields.tcfContactGroup).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataControlsContactGroup)
          .Item(TraderControlFields.tcfParameterName).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataControlParameterName)
        End With
      Else
        mvClassFields.ClearItems()
      End If

    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRS As CDBRecordSet, ByVal pTraderPageType As TraderPage.TraderPageType, ByVal pTabIndex As Integer, ByVal pNextTop As Integer, ByVal pControlIndex As Integer, ByVal pPageIndex As Integer, ByVal pTraderApplication As TraderApplication)
      'pControlIndex  = Index number of Control in Controls Collection
      'pPageIndex     = Index number of Control on Trader Page
      Dim vFields As CDBFields
      Dim vIndex As TraderControlFields

      mvEnv = pEnv
      InitClassFields()

      'Set ClassFields from RecordSet
      vFields = pRS.Fields
      With mvClassFields
        For vIndex = TraderControlFields.tcfControlType To TraderControlFields.tcfDefaultValue
          If vIndex = TraderControlFields.tcfFieldType Then
            .Item(vIndex).SetValue = CStr(vFields("type").ValueAsFieldType)
          ElseIf vIndex = TraderControlFields.tcfControlTop Then
            If pNextTop > 0 Then
              .Item(TraderControlFields.tcfControlTop).SetValue = CStr(pNextTop)
            Else
              .SetItem(vIndex, vFields)
            End If
          ElseIf (vIndex = TraderControlFields.tcfContactGroup Or vIndex = TraderControlFields.tcfParameterName) Then
            .SetOptionalItem(vIndex, vFields)
          Else
            .SetItem(vIndex, vFields)
          End If
        Next
      End With

      'Handle user specifically setting a field to be mandatory
      If vFields.ContainsKey("mandatory_item") = True AndAlso vFields("mandatory_item").Bool = True Then
        'Control customised to make it mandatory
        mvClassFields.Item(TraderControlFields.tcfNullsInvalid).Bool = True
      End If

      'Set other ClassFields
      mvClassFields.Item(TraderControlFields.tcfTabIndex).SetValue = CStr(pTabIndex)
      mvClassFields.Item(TraderControlFields.tcfIndex).SetValue = CStr(pControlIndex)
      mvClassFields.Item(TraderControlFields.tcfPageIndex).SetValue = CStr(pPageIndex)
      mvClassFields.Item(TraderControlFields.tcfPageCode).SetValue = vFields("fp_page_type").Value
      mvClassFields.Item(TraderControlFields.tcfPageType).SetValue = CStr(pTraderPageType)

      'Set Fields specific to the AttributeName
      With mvClassFields
        Select Case .Item(TraderControlFields.tcfAttributeName).Value
          Case "contact_number", "deceased_contact_number", "mailing_contact_number", "payee_contact_number", "booking_contact_number", "related_contact_number"
            .Item(TraderControlFields.tcfValidationTable).SetValue = "contacts"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "contact_number"
            If pTraderPageType = TraderPage.TraderPageType.tpTransactionDetails And .Item(TraderControlFields.tcfTableName).Value = "members" Then .Item(TraderControlFields.tcfAttributeName).SetValue = "member_contact_number"
            If pTraderPageType = TraderPage.TraderPageType.tpPurchaseOrderDetails And .Item(TraderControlFields.tcfAttributeName).Value = "payee_contact_number" And .Item(TraderControlFields.tcfVisible).Bool Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True '.NullsInvalid = "Y"
            If pTraderPageType = TraderPage.TraderPageType.tpProductDetails And .Item(TraderControlFields.tcfAttributeName).Value = "contact_number" And .Item(TraderControlFields.tcfTableName).Value = "batch_transaction_analysis" Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True ' .NullsInvalid = "Y"
            If pTraderPageType = TraderPage.TraderPageType.tpPaymentPlanProducts And .Item(TraderControlFields.tcfAttributeName).Value = "contact_number" And .Item(TraderControlFields.tcfTableName).Value = "order_details" Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True '.NullsInvalid = "Y"
            If pTraderPageType = TraderPage.TraderPageType.tpPaymentPlanDetailsMaintenance And .Item(TraderControlFields.tcfAttributeName).Value = "contact_number" And .Item(TraderControlFields.tcfTableName).Value = "order_details" Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True '.NullsInvalid = "Y"

          Case "service_contact_number"
            .Item(TraderControlFields.tcfValidationTable).SetValue = "contacts"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "service_contact_number" 'NB

          Case "address_number", "mailing_address_number", "payee_address_number", "booking_address_number", "work_address_number", "paybill_address_number"
            .Item(TraderControlFields.tcfValidationTable).SetValue = "addresses"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "address_number"
            If pTraderPageType = TraderPage.TraderPageType.tpTransactionDetails And .Item(TraderControlFields.tcfTableName).Value = "members" Then .Item(TraderControlFields.tcfAttributeName).SetValue = "member_address_number"
            If pTraderPageType = TraderPage.TraderPageType.tpPurchaseOrderDetails And .Item(TraderControlFields.tcfAttributeName).Value = "payee_address_number" And .Item(TraderControlFields.tcfVisible).Bool Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True '.NullsInvalid = "Y"

          Case "covenant_number"
            .Item(TraderControlFields.tcfNoDescription).Bool = True
            .Item(TraderControlFields.tcfValidationTable).SetValue = "covenants"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "covenant_number"
            If pTraderPageType <> TraderPage.TraderPageType.tpContactSelection And pTraderApplication.Payments = False Then .Item(TraderControlFields.tcfVisible).Bool = False

          Case "member_number"
            .Item(TraderControlFields.tcfNoDescription).Bool = True
            .Item(TraderControlFields.tcfValidationTable).SetValue = "members"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "member_number"
            If pTraderPageType <> TraderPage.TraderPageType.tpContactSelection And (pTraderApplication.Payments = False And pTraderApplication.ChangeMembership = False) Then .Item(TraderControlFields.tcfVisible).Bool = False '.Visible = False
            If PageCode = "MEM" Or PageCode = "CMT" Then
              If .Item(TraderControlFields.tcfTableName).Value = "members" Then
                .Item(TraderControlFields.tcfAttributeName).SetValue = "affiliated_member_number"
              Else
                'This is the MemberNumber field so always make it visible
                .Item(TraderControlFields.tcfVisible).Bool = True
              End If
            End If

          Case "order_number"
            .Item(TraderControlFields.tcfNoDescription).Bool = True
            .Item(TraderControlFields.tcfValidationTable).SetValue = "orders"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "order_number"
            If pTraderPageType <> TraderPage.TraderPageType.tpContactSelection And (pTraderApplication.Payments = False And pTraderApplication.ChangeMembership = False And pTraderApplication.CancelPaymentPlan = False) Then .Item(TraderControlFields.tcfVisible).Bool = False '.Visible = False

          Case "bankers_order_number"
            .Item(TraderControlFields.tcfNoDescription).Bool = True
            .Item(TraderControlFields.tcfValidationTable).SetValue = "bankers_orders"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "bankers_order_number"
            .Item(TraderControlFields.tcfNullsInvalid).Bool = False

          Case "direct_debit_number"
            .Item(TraderControlFields.tcfNoDescription).Bool = True
            .Item(TraderControlFields.tcfValidationTable).SetValue = "direct_debits"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "direct_debit_number"
            .Item(TraderControlFields.tcfNullsInvalid).Bool = False

          Case "credit_card_authority_number"
            .Item(TraderControlFields.tcfNoDescription).Bool = True
            .Item(TraderControlFields.tcfValidationTable).SetValue = "credit_card_authorities"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "credit_card_authority_number"
            .Item(TraderControlFields.tcfNullsInvalid).Bool = False

          Case "sales_ledger_account"
            .Item(TraderControlFields.tcfNoDescription).Bool = True
            .Item(TraderControlFields.tcfValidationTable).SetValue = "credit_customers"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "sales_ledger_account"
            If pTraderPageType = TraderPage.TraderPageType.tpCreditStatementGeneration Then .Item(TraderControlFields.tcfNullsInvalid).Bool = False '.NullsInvalid = "N"

          Case "reason_for_despatch"
            .Item(TraderControlFields.tcfValidationTable).SetValue = "reasons_for_despatch"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "reason_for_despatch"

          Case "event_number"
            .Item(TraderControlFields.tcfValidationTable).SetValue = "events"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "event_number"

          Case "option_number"
            .Item(TraderControlFields.tcfValidationTable).SetValue = "event_booking_options"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "option_number"

          Case "booking_number"
            If pTraderPageType = TraderPage.TraderPageType.tpEventBooking Or pTraderPageType = TraderPage.TraderPageType.tpProductDetails Then
              .Item(TraderControlFields.tcfNullsInvalid).Bool = False
              .Item(TraderControlFields.tcfValidationTable).SetValue = "event_bookings"
              .Item(TraderControlFields.tcfValidationAttribute).SetValue = "booking_number"
            End If

          Case "block_booking_number"
            .Item(TraderControlFields.tcfValidationTable).SetValue = "room_block_bookings"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "block_booking_number"

          Case "branch", "product", "rate", "expiry_date"
            If pTraderPageType <> TraderPage.TraderPageType.tpPostageAndPacking And pTraderPageType <> TraderPage.TraderPageType.tpPurchaseOrderProducts Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True '.NullsInvalid = "Y"
            If .Item(TraderControlFields.tcfAttributeName).Value = "expiry_date" AndAlso HostedAuthorisation(pTraderPageType) AndAlso _
              pTraderApplication.OnlineCCAuthorisation Then .Item(TraderControlFields.tcfVisible).Bool = False

          Case "distribution_code"
            .Item(TraderControlFields.tcfNullsInvalid).Bool = pTraderApplication.DistributionCodeMandatory

          Case "sales_contact_number"
            .Item(TraderControlFields.tcfNullsInvalid).Bool = pTraderApplication.SalesContactMandatory

          Case "card_number", "credit_card_number"
            If .Item(TraderControlFields.tcfAttributeName).Value = "card_number" And pTraderPageType <> TraderPage.TraderPageType.tpPostageAndPacking Then .Item(TraderControlFields.tcfNullsInvalid).SetValue = "Y" '.NullsInvalid = "Y"
            .Item(TraderControlFields.tcfCaseConversion).SetValue = "N"

            If HostedAuthorisation(pTraderPageType) AndAlso pTraderApplication.OnlineCCAuthorisation Then
              .Item(TraderControlFields.tcfVisible).Bool = False
            End If

          Case "quantity"
            If pTraderPageType <> TraderPage.TraderPageType.tpPostageAndPacking Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True
            If pTraderPageType <> TraderPage.TraderPageType.tpServiceBooking Then .Item(TraderControlFields.tcfFieldType).SetValue = CStr(CDBField.FieldTypes.cftInteger)

          Case "credit_card_type"
            If .Item(TraderControlFields.tcfTableName).Value = "card_sales" Then
              .Item(TraderControlFields.tcfNullsInvalid).Bool = True
              If pTraderPageType = TraderPage.TraderPageType.tpCardDetails Then
                .Item(TraderControlFields.tcfValidationTable).SetValue = ""
                .Item(TraderControlFields.tcfValidationAttribute).SetValue = ""
                .Item(TraderControlFields.tcfPattern).SetValue = "[CD]"
                .Item(TraderControlFields.tcfEntryLength).SetValue = CStr(1)
                If mvEnv.GetConfig("fp_card_sales_combined_claim") = "A" Or mvEnv.GetConfig("fp_card_sales_combined_claim") = "Y" Then .Item(TraderControlFields.tcfVisible).Bool = False '.Visible = False
              End If
            End If

            If HostedAuthorisation(pTraderPageType) AndAlso _
              pTraderApplication.OnlineCCAuthorisation Then
              .Item(TraderControlFields.tcfVisible).Bool = False
            End If

          Case "account_name"
            .Item(TraderControlFields.tcfFieldType).SetValue = CStr(CDBField.FieldTypes.cftCharacter)

          Case "account_number"
            If pTraderPageType = TraderPage.TraderPageType.tpDirectDebit Then
              'Limit to 8 - zero padded
              .Item(TraderControlFields.tcfEntryLength).SetValue = CStr(8)
            End If

          Case "source"
            If pTraderPageType = TraderPage.TraderPageType.tpMembership Or pTraderPageType = TraderPage.TraderPageType.tpChangeMembershipType Or pTraderPageType = TraderPage.TraderPageType.tpPaymentPlanDetails Then
              .Item(TraderControlFields.tcfNullsInvalid).Bool = True '.NullsInvalid = "Y"
            ElseIf pTraderPageType = TraderPage.TraderPageType.tpServiceBooking Then
              .Item(TraderControlFields.tcfNullsInvalid).Bool = .Item(TraderControlFields.tcfVisible).Bool
            End If

          Case "reference"
            Select Case pTraderPageType
              Case TraderPage.TraderPageType.tpCardDetails
                If HostedAuthorisation(pTraderPageType) AndAlso pTraderApplication.OnlineCCAuthorisation Then .Item(TraderControlFields.tcfVisible).Bool = False
              Case TraderPage.TraderPageType.tpCreditCustomer
                .Item(TraderControlFields.tcfNullsInvalid).Bool = Not (mvEnv.GetConfigOption("fp_cs_reference_optional"))
              Case TraderPage.TraderPageType.tpStandingOrder
                .Item(TraderControlFields.tcfNullsInvalid).Bool = False
            End Select

          Case "notes"
            If .Item(TraderControlFields.tcfTableName).Value = "batch_transaction_analysis" Then .Item(TraderControlFields.tcfVisible).Bool = pTraderApplication.AnalysisComments

          Case "text1", "text2", "text3", "text4", "text5"
            If pTraderPageType = TraderPage.TraderPageType.tpDirectDebit Then
              If (mvEnv.DefaultCountry = "CH" Or mvEnv.DefaultCountry = "NL") Then
                'Leave Visible property as it is
              Else
                'Always hide
                .Item(TraderControlFields.tcfVisible).Bool = False
              End If
            End If

          Case "percentage"
            .Item(TraderControlFields.tcfMaximumValue).SetValue = ""

          Case "legacy_number"
            .Item(TraderControlFields.tcfValidationTable).SetValue = "contact_legacies"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "legacy_number"

          Case "gaye_pledge_number"
            .Item(TraderControlFields.tcfNoDescription).Bool = True
            .Item(TraderControlFields.tcfValidationTable).SetValue = "gaye_pledges"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "gaye_pledge_number"

          Case "provisional"
            .Item(TraderControlFields.tcfVisible).Bool = False 'IncludeConfirmedTransactions And IncludeProvisionalTransactions

          Case "additional_reference_1", "additional_reference_2"
            If pTraderApplication.Voucher Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True

          Case "start_date"
            If pTraderPageType = TraderPage.TraderPageType.tpGiftAidDeclaration Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True
            If pTraderPageType = TraderPage.TraderPageType.tpPurchaseOrderDetails Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True
            If pTraderPageType = TraderPage.TraderPageType.tpCancelGiftAidDeclaration Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True

          Case "organisation_number", "agency_number", "payroll_organisation_number"
            .Item(TraderControlFields.tcfValidationTable).SetValue = "organisations"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "organisation_number"

          Case "output_group", "purchase_order_type"
            If pTraderPageType = TraderPage.TraderPageType.tpPurchaseOrderDetails And .Item(TraderControlFields.tcfVisible).Bool Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True

          Case "cancellation_reason"
            If pTraderPageType = TraderPage.TraderPageType.tpCancelPaymentPlan Or pTraderPageType = TraderPage.TraderPageType.tpCancelGiftAidDeclaration Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True

          Case "end_date"
            If pTraderPageType = TraderPage.TraderPageType.tpCancelGiftAidDeclaration Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True

          Case "declaration_number"
            .Item(TraderControlFields.tcfNoDescription).Bool = True
            If pTraderPageType = TraderPage.TraderPageType.tpCancelGiftAidDeclaration Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True

          Case "stock_sale"
            .Item(TraderControlFields.tcfVisible).Bool = False

          Case "claim_day"
            If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlAutoPayClaimDateMethod) = "D" Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True

          Case "product_number"
            If pTraderPageType = TraderPage.TraderPageType.tpConfirmProvisionalTransactions Then
              .Item(TraderControlFields.tcfNoDescription).Bool = True
              .Item(TraderControlFields.tcfValidationTable).SetValue = "product_numbers"
              .Item(TraderControlFields.tcfValidationAttribute).SetValue = "product_number"
            End If

          Case "company"
            If pTraderPageType = TraderPage.TraderPageType.tpCreditStatementGeneration Then .Item(TraderControlFields.tcfNullsInvalid).Bool = True

          Case "communication_number"
            .Item(TraderControlFields.tcfValidationTable).SetValue = "communications"
            .Item(TraderControlFields.tcfValidationAttribute).SetValue = "communication_number"

          Case "pledge_number"
            If pTraderPageType = TraderPage.TraderPageType.tpPostTaxPGPayment Then
              .Item(TraderControlFields.tcfNoDescription).Bool = True
              .Item(TraderControlFields.tcfValidationTable).SetValue = "post_tax_pg_pledges"
              .Item(TraderControlFields.tcfValidationAttribute).SetValue = "pledge_number"
            End If

          Case "security_code"
            .Item(TraderControlFields.tcfCaseConversion).SetValue = "N"
            If HostedAuthorisation(pTraderPageType) AndAlso _
              pTraderApplication.OnlineCCAuthorisation Then .Item(TraderControlFields.tcfVisible).Bool = False

          Case "collection_number"
            If .Item(TraderControlFields.tcfTableName).Value = "appeal_collections" Then
              .Item(TraderControlFields.tcfValidationTable).SetValue = "appeal_collections"
              .Item(TraderControlFields.tcfValidationAttribute).SetValue = "collection_number"
            End If

          Case "terms_number"
            If .Item(TraderControlFields.tcfTableName).Value = "credit_customers" Then
              .Item(TraderControlFields.tcfNullsInvalid).Bool = True
            End If

          Case "service_booking_number"
            If pTraderPageType = TraderPage.TraderPageType.tpProductDetails Then
              .Item(TraderControlFields.tcfValidationTable).SetValue = "service_bookings"
              .Item(TraderControlFields.tcfValidationAttribute).SetValue = "service_booking_number"
              .Item(TraderControlFields.tcfNullsInvalid).Bool = False '.NullsInvalid = "N"
            End If

          Case "issue_number", "valid_date", "authorisation_code"
            If HostedAuthorisation(pTraderPageType) AndAlso _
              pTraderApplication.OnlineCCAuthorisation Then
              .Item(TraderControlFields.tcfVisible).Bool = False
            End If

          Case "none"
            If (Not (HostedAuthorisation(pTraderPageType) AndAlso pTraderApplication.OnlineCCAuthorisation)) AndAlso String.Compare(.Item(TraderControlFields.tcfControlType).Value, "web", True) = 0 Then
              .Item(TraderControlFields.tcfVisible).Bool = False
            End If

          Case "exam_centre_code"
            If pTraderPageType = TraderPage.TraderPageType.tpExamBooking Then
              If pRS.Fields("mandatory_item").Value.Length = 0 Then
                'With no customisations in place, this should be non-mandatory
                .Item(TraderControlFields.tcfNullsInvalid).Bool = False
              Else
                'Once it has been customised, use the customisation
                .Item(TraderControlFields.tcfNullsInvalid).Bool = pRS.Fields("mandatory_item").Bool
              End If
            End If
        End Select
      End With

    End Sub

    Public ReadOnly Property PageCode() As String
      Get
        PageCode = mvClassFields.Item(TraderControlFields.tcfPageCode).Value
      End Get
    End Property

    Public ReadOnly Property PageType() As Integer
      Get
        PageType = mvClassFields.Item(TraderControlFields.tcfPageType).IntegerValue
      End Get
    End Property

    Public Property ControlType() As String
      Get
        ControlType = mvClassFields.Item(TraderControlFields.tcfControlType).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(TraderControlFields.tcfControlType).Value = Value
      End Set
    End Property

    Public ReadOnly Property ControlTypeCode() As String
      Get
        ControlTypeCode = Left(mvClassFields.Item(TraderControlFields.tcfControlType).Value, 3)
      End Get
    End Property

    Public ReadOnly Property ControlTypeModifier() As String
      Get
        ControlTypeModifier = Mid(mvClassFields.Item(TraderControlFields.tcfControlType).Value, 5)
      End Get
    End Property

    Public Property Top() As Integer
      Get
        Top = mvClassFields.Item(TraderControlFields.tcfControlTop).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(TraderControlFields.tcfControlTop).IntegerValue = Value
      End Set
    End Property

    'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public Property Left_Renamed() As Integer
      Get
        Left_Renamed = mvClassFields.Item(TraderControlFields.tcfControlLeft).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(TraderControlFields.tcfControlLeft).IntegerValue = Value
      End Set
    End Property

    Public Property Width() As Integer
      Get
        Width = mvClassFields.Item(TraderControlFields.tcfControlWidth).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(TraderControlFields.tcfControlWidth).IntegerValue = Value
      End Set
    End Property

    Public Property Height() As Integer
      Get
        Height = mvClassFields.Item(TraderControlFields.tcfControlHeight).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(TraderControlFields.tcfControlHeight).IntegerValue = Value
      End Set
    End Property

    Public Property Caption() As String
      Get
        Caption = mvClassFields.Item(TraderControlFields.tcfControlCaption).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(TraderControlFields.tcfControlCaption).Value = Value
      End Set
    End Property

    Public Property CaptionWidth() As Integer
      Get
        CaptionWidth = mvClassFields.Item(TraderControlFields.tcfCaptionWidth).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(TraderControlFields.tcfCaptionWidth).IntegerValue = Value
      End Set
    End Property

    Public Property HelpText() As String
      Get
        HelpText = mvClassFields.Item(TraderControlFields.tcfHelpText).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(TraderControlFields.tcfHelpText).Value = Value
      End Set
    End Property

    Public Property Visible() As Boolean
      Get
        Visible = mvClassFields.Item(TraderControlFields.tcfVisible).Bool
      End Get
      Set(ByVal Value As Boolean)
        mvClassFields.Item(TraderControlFields.tcfVisible).Bool = Value
      End Set
    End Property

    Public ReadOnly Property TableName() As String
      Get
        TableName = mvClassFields.Item(TraderControlFields.tcfTableName).Value
      End Get
    End Property

    Public ReadOnly Property AttributeName() As String
      Get
        AttributeName = mvClassFields.Item(TraderControlFields.tcfAttributeName).Value
      End Get
    End Property

    Public ReadOnly Property FieldType() As CDBField.FieldTypes
      Get
        FieldType = CType(mvClassFields.Item(TraderControlFields.tcfFieldType).IntegerValue, CDBField.FieldTypes)
      End Get
    End Property

    Public ReadOnly Property FieldTypeCode() As String
      Get
        Select Case mvClassFields.Item(TraderControlFields.tcfFieldType).Value
          Case CStr(CDBField.FieldTypes.cftCharacter)
            Return "C"
          Case CStr(CDBField.FieldTypes.cftLong)
            Return "L"
          Case CStr(CDBField.FieldTypes.cftMemo)
            Return "M"
          Case CStr(CDBField.FieldTypes.cftInteger)
            Return "I"
          Case CStr(CDBField.FieldTypes.cftNumeric)
            Return "N"
          Case CStr(CDBField.FieldTypes.cftDate)
            Return "D"
          Case CStr(CDBField.FieldTypes.cftTime)
            Return "T"
          Case Else
            Return "C"      'Add fix for compiler warnings
        End Select
      End Get
    End Property

    Public Property EntryLength() As Integer
      Get
        EntryLength = mvClassFields.Item(TraderControlFields.tcfEntryLength).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(TraderControlFields.tcfEntryLength).IntegerValue = Value
      End Set
    End Property

    Public Property NullsInvalid() As Boolean
      Get
        NullsInvalid = mvClassFields.Item(TraderControlFields.tcfNullsInvalid).Bool
      End Get
      Set(ByVal Value As Boolean)
        mvClassFields.Item(TraderControlFields.tcfNullsInvalid).Bool = Value
      End Set
    End Property

    Public ReadOnly Property Pattern() As String
      Get
        Pattern = mvClassFields.Item(TraderControlFields.tcfPattern).Value
      End Get
    End Property

    Public Property ValTable() As String
      Get
        ValTable = mvClassFields.Item(TraderControlFields.tcfValidationTable).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(TraderControlFields.tcfValidationTable).Value = Value
      End Set
    End Property

    Public Property ValAttribute() As String
      Get
        ValAttribute = mvClassFields.Item(TraderControlFields.tcfValidationAttribute).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(TraderControlFields.tcfValidationAttribute).Value = Value
      End Set
    End Property

    Public Property CaseConversion() As String
      Get
        CaseConversion = mvClassFields.Item(TraderControlFields.tcfCaseConversion).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(TraderControlFields.tcfCaseConversion).Value = Value
      End Set
    End Property

    Public ReadOnly Property MinimumValue() As String
      Get
        'Could be null
        MinimumValue = mvClassFields.Item(TraderControlFields.tcfMinimumValue).Value
      End Get
    End Property

    Public ReadOnly Property MaximumValue() As String
      Get
        'Could be null
        MaximumValue = mvClassFields.Item(TraderControlFields.tcfMaximumValue).Value
      End Get
    End Property

    Public Property ContactGroup() As String
      Get
        ContactGroup = mvClassFields.Item(TraderControlFields.tcfContactGroup).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(TraderControlFields.tcfContactGroup).Value = Value
      End Set
    End Property

    Public Property ParameterName() As String
      Get
        ParameterName = mvClassFields.Item(TraderControlFields.tcfParameterName).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(TraderControlFields.tcfParameterName).Value = Value
      End Set
    End Property

    Public Property Value() As String
      Get
        Value = mvClassFields.Item(TraderControlFields.tcfValue).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(TraderControlFields.tcfValue).Value = Value
      End Set
    End Property

    Public ReadOnly Property DoubleValue() As Double
      Get
        DoubleValue = mvClassFields.Item(TraderControlFields.tcfValue).DoubleValue
      End Get
    End Property

    Public ReadOnly Property LongValue() As Integer
      Get
        LongValue = mvClassFields.Item(TraderControlFields.tcfValue).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Bool() As Boolean
      Get
        Bool = mvClassFields.Item(TraderControlFields.tcfValue).Bool
      End Get
    End Property

    Public Property Desc() As String
      Get
        Desc = mvClassFields.Item(TraderControlFields.tcfDescription).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(TraderControlFields.tcfDescription).Value = Value
      End Set
    End Property

    Public Property TabIndex() As Integer
      Get
        TabIndex = mvClassFields.Item(TraderControlFields.tcfTabIndex).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(TraderControlFields.tcfTabIndex).IntegerValue = Value
      End Set
    End Property

    'UPGRADE_NOTE: CType was upgraded to CType_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public Property CType_Renamed() As Integer
      Get
        'ContactType
        CType_Renamed = mvClassFields.Item(TraderControlFields.tcfContactType).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(TraderControlFields.tcfContactType).IntegerValue = Value
      End Set
    End Property

    Public Property LastValue() As String
      Get
        LastValue = mvClassFields.Item(TraderControlFields.tcfLastValue).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(TraderControlFields.tcfLastValue).Value = Value
      End Set
    End Property

    Public ReadOnly Property LastLongValue() As Integer
      Get
        LastLongValue = mvClassFields.Item(TraderControlFields.tcfLastValue).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LastBool() As Boolean
      Get
        LastBool = mvClassFields.Item(TraderControlFields.tcfLastValue).Bool
      End Get
    End Property

    Public Property LastEnabled() As Boolean
      Get
        LastEnabled = mvClassFields.Item(TraderControlFields.tcfLastEnabled).Bool
      End Get
      Set(ByVal Value As Boolean)
        mvClassFields.Item(TraderControlFields.tcfLastEnabled).Bool = Value
      End Set
    End Property

    Public Property ValRequired() As Boolean
      Get
        ValRequired = mvClassFields.Item(TraderControlFields.tcfValRequired).Bool
      End Get
      Set(ByVal Value As Boolean)
        mvClassFields.Item(TraderControlFields.tcfValRequired).Bool = Value
      End Set
    End Property

    Public ReadOnly Property NoDescription() As Boolean
      Get
        NoDescription = mvClassFields.Item(TraderControlFields.tcfNoDescription).Bool
      End Get
    End Property

    Public ReadOnly Property Index() As Integer
      Get
        'The Index number of the Control in the Collection (starts at zero)
        Index = mvClassFields.Item(TraderControlFields.tcfIndex).IntegerValue
      End Get
    End Property

    Public ReadOnly Property PageIndex() As Integer
      Get
        'The Index number of the Control on the Trader Page (starts at zero)
        PageIndex = mvClassFields.Item(TraderControlFields.tcfPageIndex).IntegerValue
      End Get
    End Property

    Public ReadOnly Property DefaultValue() As String
      Get
        Return mvClassFields.Item(TraderControlFields.tcfDefaultValue).Value
      End Get
    End Property

    Private Function HostedAuthorisation(pTraderPageType As TraderPage.TraderPageType) As Boolean
      If pTraderPageType = TraderPage.TraderPageType.tpCardDetails AndAlso
        (String.Compare(mvEnv.GetConfig("fp_cc_authorisation_type"), "TNSHOSTED", True) = 0 OrElse
         (String.Compare(mvEnv.GetConfig("fp_cc_authorisation_type"), SAGEPAYHOSTED, True) = 0)) Then
        Return True
      Else
        Return False
      End If
    End Function
  End Class



End Namespace

