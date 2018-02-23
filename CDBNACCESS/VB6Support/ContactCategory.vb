Namespace Access

  Partial Public Class ContactCategory

    Public Enum ContactCategoryRecordSetTypes 'These are bit values
      ccatrtAll = &HFFS
      'ADD additional recordset types here
      ccatrtDetails = 1
      ccatrtDescriptions = &H100S
    End Enum

    Public Enum ContactCategoryTypes 'Keep same as contact class
      cctContact = 1
      cctOrganisation = 2
      cctPosition = 3
      cctExamCandidate = 4
    End Enum

    Private mvActivityDesc As String
    Private mvActivityValueDesc As String
    Private mvSourceDesc As String
    Private mvMergeActivity As Boolean

    Protected Overrides Sub ClearFields()
      mvActivityDesc = ""
      mvActivityValueDesc = ""
      mvSourceDesc = ""
      IsMerging = False
    End Sub

    Public Property IsMerging As Boolean
      Get
        Return mvMergeActivity
      End Get
      Private Set(value As Boolean)
        mvMergeActivity = value
      End Set
    End Property
    Protected Overrides Sub SetValid()
      MyBase.SetValid()
      If IsMerging = False Then CheckForQualifyingPositions(False)
    End Sub

    Protected Overrides Sub PostValidateCreateParameters(ByVal pParameterList As CDBParameters)
      If pParameterList.ParameterExists("MergeActivity").Bool Then IsMerging = True
      If pParameterList.ParameterExists("AmendedBy").Value.Length > 0 AndAlso IsDate(pParameterList.ParameterExists("AmendedOn").Value) Then
        mvOverrideAmended = True    'Use AmendedBy/On from pParameterList and do not re-set to current user & todays date
      End If
    End Sub

    Public Overloads Function GetRecordSetFields(ByVal pRSType As ContactCategoryRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If (pRSType And ContactCategoryRecordSetTypes.ccatrtAll) = ContactCategoryRecordSetTypes.ccatrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cc")
      Else
        If (pRSType And ContactCategoryRecordSetTypes.ccatrtDetails) = ContactCategoryRecordSetTypes.ccatrtDetails Then
          vFields = "cc." & mvClassFields(ContactCategoryFields.ContactNumber).Name & ",cc.activity,cc.activity_value,cc.quantity,cc.source,cc.valid_from,cc.valid_to,cc.amended_by,cc.amended_on"
          If mvClassFields(ContactCategoryFields.ActivityDate).InDatabase Then vFields = vFields & ",activity_date"
        End If
      End If
      If (pRSType And ContactCategoryRecordSetTypes.ccatrtDescriptions) = ContactCategoryRecordSetTypes.ccatrtDescriptions Then
        vFields = vFields & ",activity_desc,activity_value_desc,source_desc"
      End If
      Return vFields
    End Function

    Public Overloads Sub Init(ByVal pEnv As CDBEnvironment, ByVal pType As ContactCategoryTypes, ByVal pContactNumber As Integer, ByVal pActivity As String, ByVal pActivityValue As String, ByVal pSource As String, ByVal pValidFrom As String, ByVal pValidTo As String)
      InitFromType(pType)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add(mvClassFields(ContactCategoryFields.ContactNumber).Name, pContactNumber)
      vWhereFields.Add(mvClassFields(ContactCategoryFields.Activity).Name, pActivity)
      vWhereFields.Add(mvClassFields(ContactCategoryFields.ActivityValue).Name, pActivityValue)
      vWhereFields.Add(mvClassFields(ContactCategoryFields.Source).Name, pSource)
      vWhereFields.Add(mvClassFields(ContactCategoryFields.ValidFrom).Name, CDBField.FieldTypes.cftDate, pValidFrom)
      vWhereFields.Add(mvClassFields(ContactCategoryFields.ValidTo).Name, CDBField.FieldTypes.cftDate, pValidTo)
      InitWithPrimaryKey(vWhereFields)
    End Sub

    Private Sub InitFromType(ByVal pType As ContactCategoryTypes)
      CheckClassFields()
      Select Case pType
        Case ContactCategoryTypes.cctOrganisation
          mvClassFields(ContactCategoryFields.ContactCategoryNumber).SetName("organisation_category_number")
          mvClassFields(ContactCategoryFields.ContactNumber).SetName("organisation_number")
          mvClassFields.DatabaseTableName = "organisation_categories"
          mvClassFields.SetControlNumberField(ContactCategoryFields.ContactCategoryNumber, "CCG")
        Case ContactCategoryTypes.cctPosition
          mvClassFields(ContactCategoryFields.ContactCategoryNumber).SetName("contact_position_activity_id")
          mvClassFields(ContactCategoryFields.ContactNumber).SetName("contact_position_number")
          mvClassFields.DatabaseTableName = "contact_position_activities"
          mvClassFields.SetControlNumberField(ContactCategoryFields.ContactCategoryNumber, "PCG")
        Case ContactCategoryTypes.cctExamCandidate
          mvClassFields(ContactCategoryFields.ContactCategoryNumber).SetName("exam_candidate_activity_id")
          mvClassFields(ContactCategoryFields.ContactNumber).SetName("exam_booking_unit_id")
          mvClassFields.DatabaseTableName = "exam_candidate_activities"
          mvClassFields.SetControlNumberField(ContactCategoryFields.ContactCategoryNumber, "XCG")
      End Select
    End Sub

    Public Overloads Sub Init(ByVal pContactNumber As Integer, ByVal pActivity As String, ByVal pActivityValue As String)
      CheckClassFields()
      Dim vWhereFields As New CDBFields
      vWhereFields.Add(mvClassFields(ContactCategoryFields.ContactNumber).Name, pContactNumber)
      vWhereFields.Add(mvClassFields(ContactCategoryFields.Activity).Name, pActivity)
      vWhereFields.Add(mvClassFields(ContactCategoryFields.ActivityValue).Name, pActivityValue)
      InitWithPrimaryKey(vWhereFields)
    End Sub

    Public Overloads Sub Init(ByVal pContactNumber As Integer, ByVal pActivity As String, ByVal pActivityValue As String, ByVal pSource As String)
      CheckClassFields()
      Dim vWhereFields As New CDBFields
      vWhereFields.Add(mvClassFields(ContactCategoryFields.ContactNumber).Name, pContactNumber)
      vWhereFields.Add(mvClassFields(ContactCategoryFields.Activity).Name, pActivity)
      vWhereFields.Add(mvClassFields(ContactCategoryFields.ActivityValue).Name, pActivityValue)
      vWhereFields.Add(mvClassFields(ContactCategoryFields.Source).Name, pSource)
      InitWithPrimaryKey(vWhereFields)
    End Sub

    Public Overloads Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pType As ContactCategoryTypes, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ContactCategoryRecordSetTypes)
      InitFromType(pType)
      InitFromRecordSet(pRecordSet, pRSType)
    End Sub

    Public Overloads Sub InitFromRecordSet(ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ContactCategoryRecordSetTypes)
      Dim vFields As CDBFields

      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If ((pRSType And ContactCategoryRecordSetTypes.ccatrtAll) = ContactCategoryRecordSetTypes.ccatrtAll) Or ((pRSType And ContactCategoryRecordSetTypes.ccatrtDetails) = ContactCategoryRecordSetTypes.ccatrtDetails) Then
          .SetItem(ContactCategoryFields.ContactNumber, vFields)
          .SetItem(ContactCategoryFields.Activity, vFields)
          .SetItem(ContactCategoryFields.ActivityValue, vFields)
          .SetItem(ContactCategoryFields.Quantity, vFields)
          .SetItem(ContactCategoryFields.Source, vFields)
          .SetItem(ContactCategoryFields.ValidFrom, vFields)
          .SetItem(ContactCategoryFields.ValidTo, vFields)
          .SetOptionalItem(ContactCategoryFields.ActivityDate, vFields)
          .SetItem(ContactCategoryFields.AmendedBy, vFields)
          .SetItem(ContactCategoryFields.AmendedOn, vFields)
        End If
        If (pRSType And ContactCategoryRecordSetTypes.ccatrtAll) = ContactCategoryRecordSetTypes.ccatrtAll Then
          .SetItem(ContactCategoryFields.Notes, vFields)
        End If
        If (pRSType And ContactCategoryRecordSetTypes.ccatrtDescriptions) = ContactCategoryRecordSetTypes.ccatrtDescriptions Then
          mvActivityDesc = vFields("activity_desc").Value
          mvActivityValueDesc = vFields("activity_value_desc").Value
          mvSourceDesc = vFields("source_desc").Value
        End If
      End With
    End Sub

    Public Sub ContactTypeSaveActivity(ByVal pContactType As Contact.ContactTypes, ByVal pNumber As Integer, ByVal pActivity As String, ByVal pActivityValue As String, ByVal pSource As String, ByVal pValidFrom As String, ByVal pValidTo As String, ByVal pQty As String, ByVal pStyle As ActivityEntryStyles, Optional ByVal pNotes As String = "", Optional ByVal pAmendedOn As String = "", Optional ByVal pAmendedBy As String = "", Optional ByVal pActivityDate As String = "", Optional ByVal pResponseChannel As String = "")
      If pContactType = Contact.ContactTypes.ctcOrganisation Then
        Dim vCC As New OrganisationCategory(mvEnv)
        vCC.SaveActivity(pStyle, pNumber, pActivity, pActivityValue, pSource, pValidFrom, pValidTo, pQty, pNotes, pAmendedOn, pAmendedBy, pActivityDate, pResponseChannel)
      Else
        SaveActivity(pStyle, pNumber, pActivity, pActivityValue, pSource, pValidFrom, pValidTo, pQty, pNotes, pAmendedOn, pAmendedBy, pActivityDate, pResponseChannel)
      End If
    End Sub

    Protected Overridable Sub CheckForQualifyingPositions(ByVal pDelete As Boolean)
      If mvEnv.GetConfigOption("cd_use_qualifying_positions") = True AndAlso _
        Activity = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDaysRemainingActivity) AndAlso _
        ActivityValue = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDaysRemainingActivityVal) Then
        If pDelete = False Then
          'User is recording/updating a Current Registration: Ensure they have recorded a maximum registration activity & qty does not exceed it.
          Dim vMaxRegistrationCategory As New ContactCategory(mvEnv)
          vMaxRegistrationCategory.Init(ContactNumber, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlMaxPermittedDaysActivity), mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlMaxPermittedDaysActivityVal))
          If vMaxRegistrationCategory.Existing = False Then
            RaiseError(DataAccessErrors.daeMissingMaximumRegistrationActivity)
          Else
            If mvClassFields.Item(ContactCategoryFields.Quantity).ValueChanged = True Then
              If DoubleValue(Quantity) > Val(vMaxRegistrationCategory.Quantity) Then
                RaiseError(DataAccessErrors.daeExceedsMaximumRegistrationActivity)
              End If
            End If
          End If
        End If
      End If
    End Sub

    Public Overloads Sub Update(ByVal pActivity As String, ByVal pActivityValue As String, ByVal pSource As String, ByVal pValidFrom As String, ByVal pValidTo As String, ByVal pQuantity As String, ByVal pNotes As String, Optional ByVal pActivityDate As String = "", Optional ByVal pResponseChannel As String = "")
      With mvClassFields
        .Item(ContactCategoryFields.Activity).Value = pActivity
        .Item(ContactCategoryFields.ActivityValue).Value = pActivityValue
        .Item(ContactCategoryFields.Quantity).Value = pQuantity
        .Item(ContactCategoryFields.Source).Value = pSource
        .Item(ContactCategoryFields.ValidFrom).Value = pValidFrom
        .Item(ContactCategoryFields.ValidTo).Value = pValidTo
        .Item(ContactCategoryFields.ActivityDate).Value = pActivityDate
        .Item(ContactCategoryFields.ResponseChannel).Value = pResponseChannel
        .Item(ContactCategoryFields.Notes).Value = pNotes
      End With
    End Sub

    Public Overloads Sub Update(ByVal pValidFrom As String, ByVal pValidTo As String)
      With mvClassFields
        If pValidFrom.Length > 0 Then .Item(ContactCategoryFields.ValidFrom).Value = pValidFrom
        If pValidTo.Length > 0 Then .Item(ContactCategoryFields.ValidTo).Value = pValidTo
      End With
    End Sub

    Public Function IsValidForUpdate() As Boolean
      Dim vValid As Boolean = True
      With mvClassFields
        If .Item(ContactCategoryFields.Activity).ValueChanged OrElse .Item(ContactCategoryFields.ActivityValue).ValueChanged OrElse _
           .Item(ContactCategoryFields.Source).ValueChanged OrElse .Item(ContactCategoryFields.ValidFrom).ValueChanged OrElse _
           .Item(ContactCategoryFields.ValidTo).ValueChanged Then
          vValid = (Exists(Activity, ActivityValue, Source, ValidFrom, ValidTo) = False)
        End If
      End With
      Return vValid
    End Function

    Public Function Exists(ByVal pActivity As String, ByVal pActivityValue As String, Optional ByVal pSource As String = "", Optional ByVal pValidFrom As String = "", Optional ByVal pValidTo As String = "") As Boolean
      Dim vFields As New CDBFields
      With vFields
        .Add(mvClassFields.Item(ContactCategoryFields.ContactNumber).Name, CDBField.FieldTypes.cftLong, ContactNumber)
        .Add(mvClassFields.Item(ContactCategoryFields.Activity).Name, pActivity)
        .Add(mvClassFields.Item(ContactCategoryFields.ActivityValue).Name, pActivityValue)
        If pSource.Length > 0 AndAlso Not mvEnv.GetConfigOption("activity_exclude_source_check", False) Then .Add(mvClassFields.Item(ContactCategoryFields.Source).Name, pSource)
        If pValidTo.Length > 0 Then
          .Add(mvClassFields.Item(ContactCategoryFields.ValidFrom).Name, CDBField.FieldTypes.cftDate, pValidTo, CDBField.FieldWhereOperators.fwoLessThanEqual)
        End If
        If pValidFrom.Length > 0 Then
          .Add(mvClassFields.Item(ContactCategoryFields.ValidTo).Name, CDBField.FieldTypes.cftDate, pValidFrom, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        End If
        Dim vExistingValidFrom As String = ""
        Dim vExistingValidTo As String = ""
        Dim vExistingSource As String = ""
        With mvClassFields
          If Not (.Item(ContactCategoryFields.Activity).ValueChanged OrElse .Item(ContactCategoryFields.ActivityValue).ValueChanged OrElse .Item(ContactCategoryFields.Source).ValueChanged) _
            OrElse mvEnv.GetConfigOption("activity_exclude_source_check", False) AndAlso .Item(ContactCategoryFields.Source).SetValue <> pSource Then
            'Make sure that the GetCount does not find the record we are trying to amend
            'BR17032: Additionally where the 'activity_exclude_source_check' config is set and the source is changed ensure that we exclude the record we are trying to amend based on validfrom/validto
            vExistingValidFrom = .Item(ContactCategoryFields.ValidFrom).SetValue
            vExistingValidTo = .Item(ContactCategoryFields.ValidTo).SetValue
            If mvEnv.GetConfigOption("activity_exclude_source_check", False) AndAlso .Item(ContactCategoryFields.Source).SetValue <> pSource Then vExistingSource = .Item(ContactCategoryFields.Source).SetValue
          End If
        End With
        If vExistingValidFrom.Length > 0 OrElse vExistingValidTo.Length > 0 Then
          Dim vWhereOperator As CDBField.FieldWhereOperators = CDBField.FieldWhereOperators.fwoNotEqual
          If (vExistingValidFrom.Length > 0 AndAlso vExistingValidTo.Length > 0) OrElse (vExistingValidFrom.Length > 0 OrElse vExistingValidTo.Length > 0 AndAlso vExistingSource.Length > 0) Then vWhereOperator = CDBField.FieldWhereOperators.fwoNOT Or CDBField.FieldWhereOperators.fwoOpenBracket
          If vExistingValidFrom.Length > 0 Then .Add(mvClassFields.Item(ContactCategoryFields.ValidFrom).Name & "#2", CDBField.FieldTypes.cftDate, vExistingValidFrom, vWhereOperator)
          If vExistingValidTo.Length > 0 Then
            If vExistingValidFrom.Length > 0 Then
              If vExistingSource.Length > 0 Then .Add(mvClassFields.Item(ContactCategoryFields.Source).Name, CDBField.FieldTypes.cftCharacter, vExistingSource)
              vWhereOperator = CDBField.FieldWhereOperators.fwoCloseBracket
            ElseIf vExistingSource.Length > 0 Then
              .Add(mvClassFields.Item(ContactCategoryFields.Source).Name, CDBField.FieldTypes.cftCharacter, vExistingSource, vWhereOperator)
              vWhereOperator = CDBField.FieldWhereOperators.fwoCloseBracket
            End If
            .Add(mvClassFields.Item(ContactCategoryFields.ValidTo).Name & "#2", CDBField.FieldTypes.cftDate, vExistingValidTo, vWhereOperator)
          ElseIf vExistingSource.Length > 0 AndAlso vExistingValidFrom.Length > 0 Then
            .Add(mvClassFields.Item(ContactCategoryFields.Source).Name, CDBField.FieldTypes.cftCharacter, vExistingSource, CDBField.FieldWhereOperators.fwoCloseBracket)
          End If
        End If
      End With
      Return mvEnv.Connection.GetCount(mvClassFields.DatabaseTableName, vFields) > 0
    End Function
  End Class

End Namespace