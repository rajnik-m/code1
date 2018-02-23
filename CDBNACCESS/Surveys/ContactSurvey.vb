Namespace Access

  Public Class ContactSurvey
    Inherits CARERecord
    Implements IRecordCreate

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum ContactSurveyFields
      AllFields = 0
      CompletedOn
      ContactNumber
      ContactSurveyNumber
      CreatedBy
      CreatedOn
      Notes
      SentOn
      SurveyNumber
      SurveyVersionNumber
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("completed_on", CDBField.FieldTypes.cftDate)
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("contact_survey_number", CDBField.FieldTypes.cftLong)
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)
        .Add("notes", CDBField.FieldTypes.cftMemo)
        .Add("sent_on", CDBField.FieldTypes.cftDate)
        .Add("survey_number", CDBField.FieldTypes.cftLong)
        .Add("survey_version_number", CDBField.FieldTypes.cftLong)

        .Item(ContactSurveyFields.ContactSurveyNumber).PrimaryKey = True
        .SetControlNumberField(ContactSurveyFields.ContactSurveyNumber, "RS")
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "cs"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "contact_surveys"
      End Get
    End Property

    '--------------------------------------------------
    'Default constructor
    '--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    '--------------------------------------------------
    'AddDeleteCheckItems
    '--------------------------------------------------
    Public Overrides Sub AddDeleteCheckItems()
      AddCascadeDeleteItem("contact_survey_responses", "contact_survey_number")
    End Sub

    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ContactSurveyFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ContactSurveyFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property CompletedOn() As String
      Get
        Return mvClassFields(ContactSurveyFields.CompletedOn).Value
      End Get
    End Property
    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields(ContactSurveyFields.ContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ContactSurveyNumber() As Integer
      Get
        Return mvClassFields(ContactSurveyFields.ContactSurveyNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(ContactSurveyFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(ContactSurveyFields.CreatedOn).Value
      End Get
    End Property
    Public ReadOnly Property Notes() As String
      Get
        Return mvClassFields(ContactSurveyFields.Notes).Value
      End Get
    End Property
    Public ReadOnly Property SentOn() As String
      Get
        Return mvClassFields(ContactSurveyFields.SentOn).Value
      End Get
    End Property
    Public ReadOnly Property SurveyNumber() As Integer
      Get
        Return mvClassFields(ContactSurveyFields.SurveyNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property SurveyVersionNumber() As Integer
      Get
        Return mvClassFields(ContactSurveyFields.SurveyVersionNumber).IntegerValue
      End Get
    End Property
#End Region

    Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Save(pAmendedBy, pAudit, pJournalNumber, "")
    End Sub

    Public Overloads Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer, ByVal pSource As String)
      Dim vTrans As Boolean = False
      Try
        Dim vSurvey As New Survey(mvEnv)
        Dim vSurveyVersion As New SurveyVersion(mvEnv)
        vSurvey.Init(SurveyNumber)
        vSurveyVersion.Init(SurveyVersionNumber)

        Dim vSentCategory As ContactCategory = Nothing
        Dim vRespondedCategory As ContactCategory = Nothing

        If vSurvey.Existing AndAlso vSurveyVersion.Existing Then
          Dim vContact As New Contact(mvEnv)
          vContact.Init(ContactNumber)

          Dim vCategoryType As ContactCategory.ContactCategoryTypes = ContactCategory.ContactCategoryTypes.cctContact
          If vContact.Existing AndAlso vContact.ContactType = Contact.ContactTypes.ctcOrganisation Then vCategoryType = ContactCategory.ContactCategoryTypes.cctOrganisation

          If vSurvey.SentActivity.Length > 0 AndAlso vSurvey.SentActivityValue.Length > 0 Then
            If (IsDate(SentOn) = True AndAlso (Existing = False OrElse mvClassFields.Item(ContactSurveyFields.SentOn).ValueChanged = True)) _
            OrElse (IsDate(SentOn) = False AndAlso Existing = True AndAlso mvClassFields.Item(ContactSurveyFields.SentOn).ValueChanged = True) Then
              'Add the Category if this is a new record or the SentOn date has changed
              'OR delete the Category if this is an existing record and the SentOn date has been cleared
              If String.IsNullOrWhiteSpace(pSource) Then pSource = vSurveyVersion.Source
              If vCategoryType = ContactCategory.ContactCategoryTypes.cctOrganisation Then
                vSentCategory = New OrganisationCategory(mvEnv)
              Else
                vSentCategory = New ContactCategory(mvEnv)
              End If
              vSentCategory.Init()
              If Existing = True Then
                Dim vOldSentOn As String = mvClassFields.Item(ContactSurveyFields.SentOn).SetValue
                vSentCategory.Init(mvEnv, vCategoryType, ContactNumber, vSurvey.SentActivity, vSurvey.SentActivityValue, pSource, vOldSentOn, vOldSentOn)
              End If
            End If
          End If

          If vSurvey.RespondedActivity.Length > 0 AndAlso vSurvey.RespondedActivityValue.Length > 0 Then
            If IsDate(CompletedOn) Then
              'Once CompletedOn date has been set, update any ContactCategory to use the CompletedOn date
              'Any existing ContactCategory will have been created as Survey Answers were completed
              'Clearing the CompletedOn date will leave the Category unchanged
              vRespondedCategory = GetRespondedContactCategory(vSurvey.RespondedActivity, vSurvey.RespondedActivityValue, vSurveyVersion)
              If vRespondedCategory IsNot Nothing AndAlso vRespondedCategory.Existing = False Then vRespondedCategory = Nothing
            End If
          End If
        End If

        vTrans = mvEnv.Connection.StartTransaction()

        MyBase.Save(pAmendedBy, pAudit, pJournalNumber)

        If vSentCategory IsNot Nothing Then
          '(1) If SentOn date has changed, then change the Category
          If vSentCategory.Existing Then
            If IsDate(SentOn) Then
              '(1a) Update the existing Category
              vSentCategory.Update(SentOn, SentOn)
              If vSentCategory.IsValidForUpdate Then vSentCategory.Save(mvEnv.User.UserID, pAudit)
            Else
              '(1b) Delete the existing Category as the SentOn date has been cleared
              vSentCategory.Delete(mvEnv.User.UserID, pAudit)
            End If
          End If

          '(2) Create a new Category
          If IsDate(SentOn) = True AndAlso vSentCategory.Existing = False Then
            vSentCategory.SaveActivity(ContactCategory.ActivityEntryStyles.aesSmartClient, ContactNumber, vSurvey.SentActivity, vSurvey.SentActivityValue, pSource, SentOn, SentOn)
          End If
        End If

        If IsDate(CompletedOn) AndAlso vRespondedCategory IsNot Nothing Then
          If vRespondedCategory.Existing Then
            'Category has already been created by answering questions so just update it
            vRespondedCategory.Update(CompletedOn, CompletedOn)
            If vRespondedCategory.IsValidForUpdate Then vRespondedCategory.Save(mvEnv.User.UserID, pAudit)
          End If
        End If

        If vTrans Then mvEnv.Connection.CommitTransaction()

      Catch vEX As Exception
        If vTrans Then mvEnv.Connection.RollbackTransaction()
        Throw vEX
      End Try
    End Sub

    Public Function CreateInstance(ByVal pEnv As CDBEnvironment) As CARERecord Implements IRecordCreate.CreateInstance
      Return New ContactSurvey(pEnv)
    End Function

    Public Shared Function CreateInstance(ByVal pEnv As CDBEnvironment, ByVal pParameterList As CDBParameters) As ContactSurvey
      Dim vContactSurvey As New ContactSurvey(pEnv)
      vContactSurvey.Init(pParameterList)
      Return vContactSurvey
    End Function

   

    Protected Overrides Sub PreValidateCreateParameters(ByVal pParameterList As CDBParameters)
      'Mandatory
      ValidateContactParameter(pParameterList)
      ValidateCreateSurveyVersionParameter(pParameterList)
      ValidateSurveyParameter(pParameterList)
      'Optional
      ValidateDates(pParameterList)
    End Sub

    Protected Overrides Sub PreValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      'Mandatory
      ValidateUpdateSurveyVersionParameter(pParameterList)
      ValidateUpdateSurveyParameter(pParameterList)
      'Optional
      ValidateDates(pParameterList)
    End Sub

    Protected Overrides Sub PostValidateCreateParameters(ByVal pParameterList As CDBParameters)
      'Add code here to validate parameters passed to the create methods
      MyBase.PostValidateUpdateParameters(pParameterList)
    End Sub

    Protected Overrides Sub PostValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      'Add code here to validate parameters passed to the update methods
      MyBase.PostValidateUpdateParameters(pParameterList)
    End Sub
    Public Sub ValidateContactParameter(ByVal pParameterList As CDBParameters)
      Dim vContact As New Contact(mvEnv)
      vContact.Init(pParameterList("ContactNumber").IntegerValue)
      If Not vContact.Existing Then
        RaiseError(DataAccessErrors.daeContactNumberInvalid)
      End If
      Dim vSurveyContactGroupParameters As New CDBParameters
      vSurveyContactGroupParameters.Add("SurveyNumber", pParameterList("SurveyNumber").IntegerValue)
      vSurveyContactGroupParameters.Add("ContactGroup", vContact.ContactGroupCode)
      Dim vSurveyContactGroup As SurveyContactGroup = SurveyContactGroup.CreateInstance(mvEnv, vSurveyContactGroupParameters)
      If Not vSurveyContactGroup.Existing Then
        RaiseError(DataAccessErrors.daeSurveyInvalidContactGroup)
      End If
    End Sub

    Public Sub ValidateCreateSurveyVersionParameter(ByVal pParameterList As CDBParameters)
      Dim vSurveyVersion As SurveyVersion = SurveyVersion.CreateInstance(mvEnv, pParameterList)
      If Not vSurveyVersion.Existing Then
        RaiseError(DataAccessErrors.daeSurveyVersionNumberInvalid)
      End If
      If vSurveyVersion.SurveyNumber <> pParameterList("SurveyNumber").IntegerValue Then
        RaiseError(DataAccessErrors.daeSurveyVersionNumberInvalid)
      End If
    End Sub

    Public Sub ValidateSurveyParameter(ByVal pParameterList As CDBParameters)
      Dim vSurvey As Survey = Survey.CreateInstance(mvEnv, pParameterList)
      If Not vSurvey.Existing Then
        RaiseError(DataAccessErrors.daeSurveyNumberInvalid)
      End If
    End Sub

    Public Sub ValidateUpdateSurveyVersionParameter(pParameterList As CDBParameters)
      Dim vSurveyVersion As SurveyVersion = SurveyVersion.CreateInstance(mvEnv, pParameterList)
      If Not vSurveyVersion.Existing Then
        RaiseError(DataAccessErrors.daeSurveyVersionNumberInvalid)
      End If
      If vSurveyVersion.SurveyVersionNumber <> Me.SurveyVersionNumber Then
        RaiseError(DataAccessErrors.daeSurveyVersionNumberInvalid)
      End If
    End Sub
    Public Sub ValidateUpdateSurveyParameter(pParameterList As CDBParameters)
      Dim vSurvey As Survey = Survey.CreateInstance(mvEnv, pParameterList)
      If Not vSurvey.Existing Then
        RaiseError(DataAccessErrors.daeSurveyNumberInvalid)
      End If
      If vSurvey.SurveyNumber <> Me.SurveyNumber Then
        RaiseError(DataAccessErrors.daeSurveyNumberInvalid)
      End If
    End Sub

    ''' <summary>
    ''' Validate that SentOn is not after CompletedOn and that both are within Valid From and Valid to. 
    ''' </summary>
    ''' <param name="pParameterList"></param>
    ''' <remarks>Applies to create and update. The rules are the same.</remarks>
    Public Sub ValidateDates(ByVal pParameterList As CDBParameters)

      Dim vValidFrom As DateTime
      Dim vValidTo As DateTime
      Dim vSentOnDate As DateTime = Date.MinValue
      Dim vCompletedOnDate As DateTime = Date.MaxValue
      Dim vClosingDate As DateTime = Date.MaxValue
      Dim vSurveyVersion As New SurveyVersion(mvEnv)

      vSurveyVersion.Init(pParameterList("SurveyVersionNumber").IntegerValue)

      If String.IsNullOrEmpty(vSurveyVersion.ValidFrom) Then
        vValidFrom = Date.MinValue
      Else
        vValidFrom = CDate(vSurveyVersion.ValidFrom)
      End If
      If String.IsNullOrEmpty(vSurveyVersion.ValidTo) Then
        vValidTo = Date.MaxValue
      Else
        vValidTo = CDate(vSurveyVersion.ValidTo)
      End If

      If String.IsNullOrEmpty(vSurveyVersion.ClosingDate) Then
        vClosingDate = Date.MaxValue
      Else
        vClosingDate = CDate(vSurveyVersion.ClosingDate)
      End If

      If pParameterList.Exists("SentOn") Then
        If Not String.IsNullOrEmpty(pParameterList("SentOn").Value) Then
          vSentOnDate = CDate(pParameterList("SentOn").Value)
          If vSentOnDate < vValidFrom Then
            RaiseError(DataAccessErrors.daeSurveySentOnBeforeValidFrom)
          End If
          If vSentOnDate > vValidTo Then
            RaiseError(DataAccessErrors.daeSurveySentOnAfterValidTo)
          End If
        End If
      End If

      If pParameterList.Exists("CompletedOn") Then
        If Not String.IsNullOrEmpty(pParameterList("CompletedOn").Value) Then
          vCompletedOnDate = CDate(pParameterList("CompletedOn").Value)
          If vCompletedOnDate < vValidFrom Then
            RaiseError(DataAccessErrors.daeSurveyCompletedDateBeforeValidFrom)
          End If
          If vCompletedOnDate > vClosingDate Then
            RaiseError(DataAccessErrors.daeSurveyCompletedDateAfterClosingDate)
          End If
        End If
      End If

      If vClosingDate < vValidFrom Then
        RaiseError(DataAccessErrors.daeSurveyClosingDateBeforeValidFrom)
      End If

      If pParameterList.ContainsKey("SentOn") AndAlso pParameterList.ContainsKey("CompletedOn") Then
        If Not String.IsNullOrEmpty(pParameterList("SentOn").Value) And Not String.IsNullOrEmpty(pParameterList("CompletedOn").Value) Then
          If vSentOnDate > vCompletedOnDate Then
            RaiseError(DataAccessErrors.daeSurveySentOnAfterCompletedOn)
          End If
        End If
      End If

    End Sub

    ''' <summary>Gets the <see cref="ContactCategory">ContactCategory</see> for the Survey RespondedActivity and RespondedActivityValue.</summary>
    ''' <param name="pActivity">Survey Responded Activity</param>
    ''' <param name="pActivityValue">Survey Responded Activity Value</param>
    ''' <param name="pSurveyVersion">Survey Version being used</param>
    ''' <returns><see cref="ContactCategory">ContactCategory</see> initialised with first matching Category or a pre-initialised class if no category found.</returns>
    ''' <remarks>Used to find the first (newest) Contact Category</remarks>
    Friend Function GetRespondedContactCategory(ByVal pActivity As String, ByVal pActivityValue As String, ByVal pSurveyVersion As SurveyVersion) As ContactCategory
      Dim vCategory As ContactCategory = Nothing
      Dim vContact As New Contact(mvEnv)
      vContact.Init(ContactNumber)

      If Not (String.IsNullOrWhiteSpace(pActivity)) AndAlso Not (String.IsNullOrWhiteSpace(pActivityValue)) Then
        If vContact.Existing Then
          Dim vWhereFields As New CDBFields
          Dim vTableName As String = "contact_categories cc"
          If vContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
            vCategory = New OrganisationCategory(mvEnv)
            vWhereFields.Add("organisation_number", vContact.ContactNumber)
            vTableName = "organisation_categories cc"
          Else
            vCategory = New ContactCategory(mvEnv)
            vWhereFields.Add("contact_number", vContact.ContactNumber)
          End If

          'Find any existing Category so it can be updated
          vWhereFields.Add("activity", pActivity)
          vWhereFields.Add("activity_value", pActivityValue)
          vWhereFields.Add("source", pSurveyVersion.Source)

          Dim vCompletedOnDate As String = CompletedOn
          If IsDate(CompletedOn) AndAlso mvClassFields.Item(ContactSurveyFields.CompletedOn).ValueChanged Then
            'Look for Category with previous Completed On date
            vCompletedOnDate = mvClassFields.Item(ContactSurveyFields.CompletedOn).SetValue
          End If

          If IsDate(vCompletedOnDate) Then
            'If CompletedOn date is set Category will have that date
            vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, vCompletedOnDate)
            vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, vCompletedOnDate)
          Else
            If IsDate(pSurveyVersion.ValidFrom) Then vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, pSurveyVersion.ValidFrom, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
            If IsDate(pSurveyVersion.ValidTo) Then vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, pSurveyVersion.ValidTo, CDBField.FieldWhereOperators.fwoLessThanEqual)
          End If

          Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vCategory.GetRecordSetFields(), vTableName, vWhereFields, "valid_from DESC")
          Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
          If vRS.Fetch Then
            vCategory.InitFromRecordSet(vRS)
          End If
          vRS.CloseRecordSet()
        End If
      End If

      If vCategory Is Nothing Then
        If vContact.Existing AndAlso vContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
          vCategory = New OrganisationCategory(mvEnv)
        Else
          vCategory = New ContactCategory(mvEnv)
        End If
      End If

      Return vCategory

    End Function

  End Class
End Namespace