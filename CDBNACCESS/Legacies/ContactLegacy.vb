Namespace Access

  Public Class ContactLegacy
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum ContactLegacyFields
      AllFields = 0
      LegacyNumber
      ContactNumber
      LegacyId
      LegacyStatus
      Source
      SourceDate
      WillDate
      LastCodicilDate
      GrossEstateValue
      NetEstateValue
      TotalEstimatedValue
      AdminExpensesValue
      TaxValue
      OtherBequestsValue
      NetForProbate
      LiabilitiesValue
      DateOfDeath
      DeathNotificationSource
      DeathNotificationDate
      DateOfProbate
      NextBequestNumber
      MasterAction
      ReviewDate
      LegacyReviewReason
      AgencyNotificationDate
      AccountsReceived
      AccountsApproved
      AgeAtDeath
      LeadCharity
      InDispute
      LegacyDisputeReason
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("legacy_number", CDBField.FieldTypes.cftLong)
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("legacy_id")
        .Add("legacy_status")
        .Add("source")
        .Add("source_date", CDBField.FieldTypes.cftDate)
        .Add("will_date", CDBField.FieldTypes.cftDate)
        .Add("last_codicil_date", CDBField.FieldTypes.cftDate)
        .Add("gross_estate_value", CDBField.FieldTypes.cftNumeric)
        .Add("net_estate_value", CDBField.FieldTypes.cftNumeric)
        .Add("total_estimated_value", CDBField.FieldTypes.cftNumeric)
        .Add("admin_expenses_value", CDBField.FieldTypes.cftNumeric)
        .Add("tax_value", CDBField.FieldTypes.cftNumeric)
        .Add("other_bequests_value", CDBField.FieldTypes.cftNumeric)
        .Add("net_for_probate", CDBField.FieldTypes.cftNumeric)
        .Add("liabilities_value", CDBField.FieldTypes.cftNumeric)
        .Add("date_of_death", CDBField.FieldTypes.cftDate)
        .Add("death_notification_source", CDBField.FieldTypes.cftLong)
        .Add("death_notification_date", CDBField.FieldTypes.cftDate)
        .Add("date_of_probate", CDBField.FieldTypes.cftDate)
        .Add("next_bequest_number", CDBField.FieldTypes.cftInteger)
        .Add("master_action", CDBField.FieldTypes.cftLong)
        .Add("review_date", CDBField.FieldTypes.cftDate)
        .Add("legacy_review_reason")
        .Add("agency_notification_date", CDBField.FieldTypes.cftDate)
        .Add("accounts_received", CDBField.FieldTypes.cftDate)
        .Add("accounts_approved", CDBField.FieldTypes.cftDate)
        .Add("age_at_death", CDBField.FieldTypes.cftInteger)
        .Add("lead_charity")
        .Add("in_dispute")
        .Add("legacy_dispute_reason")

        .Item(ContactLegacyFields.LegacyNumber).PrimaryKey = True

        .SetControlNumberField(ContactLegacyFields.LegacyNumber, "LG")

        .SetUniqueField(ContactLegacyFields.ContactNumber)          'Only one legacy per contact
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "cl"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "contact_legacies"
      End Get
    End Property

    '--------------------------------------------------
    'Default constructor
    '--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public ReadOnly Property LegacyNumber() As Integer
      Get
        Return mvClassFields(ContactLegacyFields.LegacyNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields(ContactLegacyFields.ContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property LegacyId() As String
      Get
        Return mvClassFields(ContactLegacyFields.LegacyId).Value
      End Get
    End Property
    Public ReadOnly Property LegacyStatus() As String
      Get
        Return mvClassFields(ContactLegacyFields.LegacyStatus).Value
      End Get
    End Property
    Public ReadOnly Property Source() As String
      Get
        Return mvClassFields(ContactLegacyFields.Source).Value
      End Get
    End Property
    Public ReadOnly Property SourceDate() As String
      Get
        Return mvClassFields(ContactLegacyFields.SourceDate).Value
      End Get
    End Property
    Public ReadOnly Property WillDate() As String
      Get
        Return mvClassFields(ContactLegacyFields.WillDate).Value
      End Get
    End Property
    Public ReadOnly Property LastCodicilDate() As String
      Get
        Return mvClassFields(ContactLegacyFields.LastCodicilDate).Value
      End Get
    End Property
    Public ReadOnly Property GrossEstateValue() As Double
      Get
        Return mvClassFields(ContactLegacyFields.GrossEstateValue).DoubleValue
      End Get
    End Property
    Public ReadOnly Property NetEstateValue() As Double
      Get
        Return mvClassFields(ContactLegacyFields.NetEstateValue).DoubleValue
      End Get
    End Property
    Public ReadOnly Property TotalEstimatedValue() As Double
      Get
        Return mvClassFields(ContactLegacyFields.TotalEstimatedValue).DoubleValue
      End Get
    End Property
    Public ReadOnly Property AdminExpensesValue() As Double
      Get
        Return mvClassFields(ContactLegacyFields.AdminExpensesValue).DoubleValue
      End Get
    End Property
    Public ReadOnly Property TaxValue() As Double
      Get
        Return mvClassFields(ContactLegacyFields.TaxValue).DoubleValue
      End Get
    End Property
    Public ReadOnly Property OtherBequestsValue() As Double
      Get
        Return mvClassFields(ContactLegacyFields.OtherBequestsValue).DoubleValue
      End Get
    End Property
    Public ReadOnly Property NetForProbate() As Double
      Get
        Return mvClassFields(ContactLegacyFields.NetForProbate).DoubleValue
      End Get
    End Property
    Public ReadOnly Property LiabilitiesValue() As Double
      Get
        Return mvClassFields(ContactLegacyFields.LiabilitiesValue).DoubleValue
      End Get
    End Property
    Public ReadOnly Property DateOfDeath() As String
      Get
        Return mvClassFields(ContactLegacyFields.DateOfDeath).Value
      End Get
    End Property
    Public ReadOnly Property DeathNotificationSource() As Integer
      Get
        Return mvClassFields(ContactLegacyFields.DeathNotificationSource).IntegerValue
      End Get
    End Property
    Public ReadOnly Property DeathNotificationDate() As String
      Get
        Return mvClassFields(ContactLegacyFields.DeathNotificationDate).Value
      End Get
    End Property
    Public ReadOnly Property DateOfProbate() As String
      Get
        Return mvClassFields(ContactLegacyFields.DateOfProbate).Value
      End Get
    End Property
    Public ReadOnly Property NextBequestNumber() As Integer
      Get
        Return mvClassFields(ContactLegacyFields.NextBequestNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property MasterAction() As Integer
      Get
        Return mvClassFields(ContactLegacyFields.MasterAction).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ContactLegacyFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ContactLegacyFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property ReviewDate() As String
      Get
        Return mvClassFields(ContactLegacyFields.ReviewDate).Value
      End Get
    End Property
    Public ReadOnly Property LegacyReviewReason() As String
      Get
        Return mvClassFields(ContactLegacyFields.LegacyReviewReason).Value
      End Get
    End Property
    Public ReadOnly Property AgencyNotificationDate() As String
      Get
        Return mvClassFields(ContactLegacyFields.AgencyNotificationDate).Value
      End Get
    End Property
    Public ReadOnly Property AccountsReceived() As String
      Get
        Return mvClassFields(ContactLegacyFields.AccountsReceived).Value
      End Get
    End Property
    Public ReadOnly Property AccountsApproved() As String
      Get
        Return mvClassFields(ContactLegacyFields.AccountsApproved).Value
      End Get
    End Property
    Public ReadOnly Property AgeAtDeath() As Integer
      Get
        Return mvClassFields(ContactLegacyFields.AgeAtDeath).IntegerValue
      End Get
    End Property
    Public ReadOnly Property LeadCharity() As String
      Get
        Return mvClassFields(ContactLegacyFields.LeadCharity).Value
      End Get
    End Property
    Public ReadOnly Property InDispute() As String
      Get
        Return mvClassFields(ContactLegacyFields.InDispute).Value
      End Get
    End Property
    Public ReadOnly Property LegacyDisputeReason() As String
      Get
        Return mvClassFields(ContactLegacyFields.LegacyDisputeReason).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"

    Private mvSetGrossAmountFromAssets As Boolean

    Protected Overrides Sub SetDefaults()
      MyBase.SetDefaults()
      mvClassFields(ContactLegacyFields.InDispute).Value = "N"
      mvClassFields(ContactLegacyFields.LeadCharity).Value = "N"
      mvClassFields(ContactLegacyFields.NextBequestNumber).Value = "1"
    End Sub

    Protected Overrides Sub SetValid()
      MyBase.SetValid()
      If mvClassFields(ContactLegacyFields.SourceDate).Value = "" Then mvClassFields(ContactLegacyFields.SourceDate).Value = TodaysDate()
    End Sub

    Public Overrides Function GetAddRecordMandatoryParameters() As String
      Return "ContactNumber,LegacyStatus,Source"
    End Function

    Protected Overrides Sub PostValidateCreateParameters(ByVal pParameterList As CDBParameters)
      MyBase.PostValidateCreateParameters(pParameterList)
      ValidateContactExists()
      ValidateDates()
      mvSetGrossAmountFromAssets = False
    End Sub

    Protected Overrides Sub PostValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      MyBase.PostValidateUpdateParameters(pParameterList)
      ValidateDates()
      mvSetGrossAmountFromAssets = pParameterList.ParameterExists("UpdateGrossAmount").Bool
    End Sub

    Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Dim vCurrentStatus As New LegacyStatus(mvEnv)
      vCurrentStatus.Init(LegacyStatus)
      Dim vActivity As New ContactCategory(mvEnv)
      Dim vSuppression As New ContactSuppression(mvEnv)
      Dim vCurrentStatusCode As String = vCurrentStatus.Status
      Dim vOldStatusCode As String = ""
      mvEnv.Connection.StartTransaction()

      If mvSetGrossAmountFromAssets Then SetGrossAmountFromAssets()
      If mvClassFields(ContactLegacyFields.NetEstateValue).ValueChanged OrElse mvClassFields(ContactLegacyFields.OtherBequestsValue).ValueChanged Then
        SetResidualBequestAmounts()
      End If
      If Existing AndAlso mvClassFields(ContactLegacyFields.LegacyStatus).ValueChanged Then
        Dim vOldStatus As New LegacyStatus(mvEnv)
        vOldStatus.Init(mvClassFields(ContactLegacyFields.LegacyStatus).SetValue)
        vOldStatusCode = vOldStatus.Status
        If vOldStatus.Activity.Length > 0 AndAlso vOldStatus.ActivityValue.Length > 0 Then
          vActivity.SaveActivity(ContactCategory.ActivityEntryStyles.aesNormal, ContactNumber, vOldStatus.Activity, vOldStatus.ActivityValue, Source, "", TodaysDate)
        End If
        If vOldStatus.MailingSuppression.Length > 0 Then
          vSuppression.SaveSuppression(ContactSuppression.SuppressionEntryStyles.sesNormal Or ContactSuppression.SuppressionEntryStyles.sesNoInsertAllowed, ContactNumber, vOldStatus.MailingSuppression, SourceDate, TodaysDate, Source)
        End If
      ElseIf Existing Then
        vOldStatusCode = vCurrentStatusCode     'Existing and status not changed so don't update contact
      End If
      If Not Existing OrElse mvClassFields(ContactLegacyFields.LegacyStatus).ValueChanged Then
        If vCurrentStatus.Activity.Length > 0 AndAlso vCurrentStatus.ActivityValue.Length > 0 Then
          vActivity.SaveActivity(ContactCategory.ActivityEntryStyles.aesNormal, ContactNumber, vCurrentStatus.Activity, vCurrentStatus.ActivityValue, Source, SourceDate, Date.Parse(SourceDate).AddYears(99).ToString(CAREDateFormat))
        End If
        If vCurrentStatus.MailingSuppression.Length > 0 Then
          vSuppression.SaveSuppression(ContactSuppression.SuppressionEntryStyles.sesNormal, ContactNumber, vCurrentStatus.MailingSuppression, SourceDate, Date.Parse(SourceDate).AddYears(99).ToString(CAREDateFormat), Source)
        End If
      End If
      If vCurrentStatusCode <> vOldStatusCode Then
        Dim vContact As New Contact(mvEnv)
        vContact.Init(ContactNumber)
        If vContact.Existing Then
          vContact.SetStatus(ContactNumber,
                             vCurrentStatusCode,
                             "",
                             If(String.IsNullOrWhiteSpace(vContact.StatusReason),
                                ProjectText.String15027,
                                String.Empty))
        End If
      End If
      MyBase.Save(pAmendedBy, pAudit, pJournalNumber)
      mvEnv.Connection.CommitTransaction()
    End Sub

    Public Overrides Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add(mvClassFields(ContactLegacyFields.LegacyNumber).Name, LegacyNumber)
      If mvEnv.Connection.GetCount("legacy_bequests", vWhereFields) > 0 OrElse _
         mvEnv.Connection.GetCount("legacy_tax_certificates", vWhereFields) > 0 Then RaiseError(DataAccessErrors.daeCannotDeleteLegacy)
      Dim vCurrentStatus As New LegacyStatus(mvEnv)
      vCurrentStatus.Init(LegacyStatus)
      mvEnv.Connection.StartTransaction()
      If vCurrentStatus.Activity.Length > 0 AndAlso vCurrentStatus.ActivityValue.Length > 0 Then
        Dim vCategory As New ContactCategory(mvEnv)
        vCategory.DeleteActivity(ContactNumber, vCurrentStatus.Activity, vCurrentStatus.ActivityValue)
      End If
      If vCurrentStatus.MailingSuppression.Length > 0 Then
        Dim vSuppression As New ContactSuppression(mvEnv)
        vSuppression.DeleteSuppression(ContactNumber, vCurrentStatus.MailingSuppression)
      End If
      If vCurrentStatus.Status.Length > 0 Then
        Dim vContact As New Contact(mvEnv)
        vContact.SetStatus(ContactNumber, "", "", "")
      End If
      If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlLGAssetActivity).Length > 0 Then
        Dim vCategory As New ContactCategory(mvEnv)
        vCategory.DeleteActivity(ContactNumber, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlLGAssetActivity))
      End If
      MyBase.Delete(pAmendedBy, pAudit, pJournalNumber)
      mvEnv.Connection.CommitTransaction()
    End Sub

    Private Sub ValidateContactExists()
      Dim vContact As New Contact(mvEnv)
      vContact.Init(ContactNumber)
      If vContact.Existing = False Then Throw New CareException("Contact does not exist", 1, "")
    End Sub

    Private Sub ValidateDates()
      Dim vDateOfDeath As Date
      Dim vDateOfDeathValid As Boolean = Date.TryParse(DateOfDeath, vDateOfDeath)
      Dim vNotificationDate As Date
      Dim vNotificationDateValid As Boolean = Date.TryParse(DeathNotificationDate, vNotificationDate)
      Dim vLastCodicil As Date
      Dim vLastCodicilValid As Boolean = Date.TryParse(LastCodicilDate, vLastCodicil)
      Dim vWillDate As Date
      Dim vWillDateValid As Boolean = Date.TryParse(WillDate, vWillDate)
      Dim vProbateDate As Date
      Dim vProbateDateValid As Boolean = Date.TryParse(DateOfProbate, vProbateDate)
      Dim vAgencyNotification As Date
      Dim vAgencyNotificationValid As Boolean = Date.TryParse(AgencyNotificationDate, vAgencyNotification)
      Dim vAccountsReceived As Date
      Dim vAccountsReceivedValid As Boolean = Date.TryParse(AccountsReceived, vAccountsReceived)
      Dim vAccountsApproved As Date
      Dim vAccountsApprovedValid As Boolean = Date.TryParse(AccountsApproved, vAccountsApproved)

      'TODO Legacy Handle errors properly
      If vDateOfDeathValid AndAlso vDateOfDeath > Now Then
        RaiseError(DataAccessErrors.daeLegacyDeathDateFuture)                      'Date of Death cannot be in the future
      End If
      If vNotificationDateValid AndAlso (vDateOfDeath > vNotificationDate Or vLastCodicil > vNotificationDate Or vWillDate > vNotificationDate) Then
        RaiseError(DataAccessErrors.daeLegacyNotifcationDate)                      'Notification Date cannot be prior to Will Date, Last Codicil Date or Date of Death
      End If
      If vProbateDateValid AndAlso (vDateOfDeath > vProbateDate Or vLastCodicil > vProbateDate Or vWillDate > vProbateDate) Then
        RaiseError(DataAccessErrors.daeLegacyDateOfProbate)                        'Date of Probate cannot be prior to Will Date, Last Codicil Date or Date of Death
      End If
      If vAgencyNotificationValid AndAlso (vDateOfDeath > vAgencyNotification Or vLastCodicil > vAgencyNotification Or vWillDate > vAgencyNotification) Then
        RaiseError(DataAccessErrors.daeLegacyNotified)                             'Notified by Agency cannot be prior to Will Date, Last Codicil Date or Date of Death
      End If
      If vDateOfDeathValid AndAlso (vLastCodicil > vDateOfDeath Or vWillDate > vDateOfDeath) Then
        RaiseError(DataAccessErrors.daeLegacyDeathDate)                            'Date of Death cannot be prior to Will Date or Last Codicil Date
      End If
      If vLastCodicilValid And vWillDate > vLastCodicil Then
        RaiseError(DataAccessErrors.daeLegacyCodicilDate)                          'Last Codicil Date cannot be prior to Will Date
      End If
      If vAccountsReceivedValid And vAccountsReceived > Now Then
        RaiseError(DataAccessErrors.daeLegacyAccountsReceivedFuture)               'Accounts Received Date cannot be in the future
      End If
      If vAccountsApprovedValid And vAccountsApproved > Now Then
        RaiseError(DataAccessErrors.daeLegacyAccountsApprovedFuture)               'Accounts Approved Date cannot be in the future
      End If
      If vAccountsReceivedValid And vProbateDate > vAccountsReceived Then
        RaiseError(DataAccessErrors.daeLegacyAccountsReceived)                     'Accounts Received Date cannot be prior to Probate Date
      End If
      If vAccountsApprovedValid And vAccountsReceived > vAccountsApproved Then
        RaiseError(DataAccessErrors.daeLegacyAccountsApproved)                     'Accounts Approved Date cannot be prior to Accounts Received Date
      End If
    End Sub

    Public Function AllocateNextBequestNumber() As Integer
      Dim vWhereFields As New CDBFields
      Dim vCount As Integer
      Dim vRetries As Integer
      Dim vNumber As Integer
      Do
        vNumber = NextBequestNumber
        vWhereFields.Clear()
        vWhereFields.Add(mvClassFields.Item(ContactLegacyFields.LegacyNumber).Name, LegacyNumber)
        vWhereFields.Add(mvClassFields.Item(ContactLegacyFields.NextBequestNumber).Name, vNumber)
        mvClassFields.Item(ContactLegacyFields.NextBequestNumber).IntegerValue = vNumber + 1
        vCount = mvEnv.Connection.UpdateRecords(mvClassFields.DatabaseTableName, mvClassFields.UpdateFields, vWhereFields, False)
        vRetries += 1
        If vCount = 0 Then Init(LegacyNumber) 'Re-read the record
      Loop While vCount = 0 And vRetries < mvEnv.MaxRetries
      If vCount = 0 Then RaiseError(DataAccessErrors.daeAllocateBequestNumber, LegacyNumber.ToString)
      Return vNumber
    End Function

    Public Sub SetExpectedAmountFromBequests()
      'Set the Total Expected Amount on the Legacy according to the Bequest mix
      If Existing Then
        Dim vSQL As New SQLStatement(mvEnv.Connection, "SUM(expected_value) AS total_expected", "legacy_bequests", New CDBField("legacy_number", LegacyNumber), "")
        mvClassFields(ContactLegacyFields.TotalEstimatedValue).Value = vSQL.GetValue
      End If
    End Sub

    Public Sub SetGrossAmountFromAssets()
      'Set the Gross Amount on the Legacy according to the Assets
      If Existing Then
        Dim vWhereFields As New CDBFields()
        vWhereFields.Add("contact_number", ContactNumber)
        vWhereFields.Add("activity", mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlLGAssetActivity))
        Dim vSQL As New SQLStatement(mvEnv.Connection, "SUM(quantity) AS gross_amount", "contact_categories", vWhereFields, "")
        Dim vValue As String = vSQL.GetValue
        Dim vSumAssets As Double
        If vValue.Length > 0 Then vSumAssets = Double.Parse(vValue)
        Dim vAdjustment As Double = GrossEstateValue - vSumAssets
        If vAdjustment <> 0 Then
          mvClassFields(ContactLegacyFields.GrossEstateValue).DoubleValue = vSumAssets
          mvClassFields(ContactLegacyFields.NetEstateValue).DoubleValue = NetEstateValue - vAdjustment
          mvClassFields(ContactLegacyFields.NetForProbate).DoubleValue = NetForProbate - vAdjustment
          'Dont need to SetResidualBequestAmounts here as it will be done for us by the save method
        End If
      End If
    End Sub

    Public Sub SetResidualBequestAmounts()
      'Set Expected Value and amend Estimated Outstanding on all Residual Bequests for this
      'Legacy following an amendment to the Legacy Amounts. Need to look at each record
      'individually since the new Estimated Outstanding is based on the original value.
      'Note ref DOM, estimated_outstanding can go negative to indicate the charity owes some
      'monies back.
      If Existing Then
        Dim vWhereFields As New CDBFields()
        vWhereFields.Add("legacy_number", LegacyNumber)
        vWhereFields.Add("bequest_type", mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlLGResidualBequestType))
        Dim vBequest As New LegacyBequest(mvEnv)
        Dim vDataTable As System.Data.DataTable = vBequest.GetDataTable(vWhereFields)
        For Each vRow As System.Data.DataRow In vDataTable.Rows
          vBequest.InitFromDataRow(vRow)
          vBequest.SetExpectedValueAndOutstanding(NetEstateValue - OtherBequestsValue)
          vBequest.Save()
        Next
        If vDataTable.Rows.Count > 0 Then SetExpectedAmountFromBequests()
      End If
    End Sub

#End Region

  End Class
End Namespace
