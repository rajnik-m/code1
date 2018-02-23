Namespace Access

  Partial Public Class ContactSuppression

    Public Enum ContactSuppressionTypes 'Keep same as contact class
      cstContact = 1
      cstOrganisation = 2
    End Enum

    Public Overrides Sub Init(ByVal pType As Integer)
      CheckClassFields()
      If pType = ContactSuppressionTypes.cstOrganisation Then
        mvClassFields(ContactSuppressionFields.ContactNumber).SetName("organisation_number")
        mvClassFields.DatabaseTableName = "organisation_suppressions"
      End If
    End Sub

    Public Overloads Sub Init(ByVal pType As ContactSuppressionTypes, ByVal pContactNumber As Integer, ByVal pSuppression As String, ByVal pValidFrom As String, ByVal pvalidTo As String)
      CheckClassFields()
      If pType = ContactSuppressionTypes.cstOrganisation Then
        mvClassFields(ContactSuppressionFields.ContactNumber).SetName("organisation_number")
        mvClassFields.DatabaseTableName = "organisation_suppressions"
      End If
      Dim vWhereFields As New CDBFields
      vWhereFields.Add(mvClassFields(ContactSuppressionFields.ContactNumber).Name, pContactNumber)
      vWhereFields.Add(mvClassFields(ContactSuppressionFields.MailingSuppression).Name, pSuppression)
      vWhereFields.Add(mvClassFields(ContactSuppressionFields.ValidFrom).Name, CDBField.FieldTypes.cftDate, pValidFrom)
      vWhereFields.Add(mvClassFields(ContactSuppressionFields.ValidTo).Name, CDBField.FieldTypes.cftDate, pvalidTo)
      MyBase.InitWithPrimaryKey(vWhereFields)
    End Sub
    Public Shared Sub ContactTypeSaveSuppression(ByVal pEnv As CDBEnvironment, ByVal pStyle As SuppressionEntryStyles, ByVal pContactType As Contact.ContactTypes, ByVal pNumber As Integer, ByVal pSuppression As String, Optional ByVal pValidFrom As String = "", Optional ByVal pValidTo As String = "", Optional ByRef pAmendedOn As String = "", Optional ByRef pAmendedBy As String = "", Optional pSource As String = "")
      If pContactType = Contact.ContactTypes.ctcOrganisation Then
        Dim vSuppression As New OrganisationSuppression(pEnv)
        vSuppression.SaveSuppression(pStyle, pNumber, pSuppression, pValidFrom, pValidTo, pAmendedOn, pAmendedBy, pSource)
      Else
        Dim vSuppression As New ContactSuppression(pEnv)
        vSuppression.SaveSuppression(pStyle, pNumber, pSuppression, pValidFrom, pValidTo, pAmendedOn, pAmendedBy, pSource)
      End If
    End Sub

    Public Overloads Sub Create(ByVal pContactNumber As Integer, ByVal pSuppression As String, ByVal pValidFrom As String, ByVal pValidTo As String, Optional ByVal pNotes As String = "", Optional ByVal pSource As String = "", Optional ByVal pResponseChannel As String = "")
      With mvClassFields
        .Item(ContactSuppressionFields.ContactNumber).IntegerValue = pContactNumber
        .Item(ContactSuppressionFields.MailingSuppression).Value = pSuppression
        .Item(ContactSuppressionFields.ValidFrom).Value = pValidFrom
        .Item(ContactSuppressionFields.ValidTo).Value = pValidTo
        .Item(ContactSuppressionFields.Notes).Value = pNotes
        .Item(ContactSuppressionFields.Source).Value = pSource
        .Item(ContactSuppressionFields.ResponseChannel).Value = pResponseChannel
      End With
    End Sub

    Public Overloads Sub Update(ByVal pValidFrom As String, ByVal pValidTo As String, Optional ByVal pNotes As String = "", Optional ByVal pSource As String = "", Optional ByVal pResponseChannel As String = "")
      'This function is only used by WEB Services at present
      With mvClassFields
        .Item(ContactSuppressionFields.ValidFrom).Value = pValidFrom
        .Item(ContactSuppressionFields.ValidTo).Value = pValidTo
        .Item(ContactSuppressionFields.Notes).Value = pNotes
        .Item(ContactSuppressionFields.Source).Value = pSource
        .Item(ContactSuppressionFields.ResponseChannel).Value = pResponseChannel
      End With
    End Sub

    Public Function SuppressionExists(pContact As Contact, pSuppressionCode As String, pValidFrom As String, pValidTo As String) As Boolean
      Dim vWhereFields As New CDBFields
      Dim vTableName As String
      vWhereFields.Add("mailing_suppression", pSuppressionCode)
      'If any existing record ends on or after the new start date (or is null)
      vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, pValidFrom, CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      vWhereFields.Add("valid_to#2", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      'And starts on or before the new end date (or is null)
      vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, pValidTo, CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoLessThanEqual)
      vWhereFields.Add("valid_from#2", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      If pContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        vWhereFields.Add("organisation_number", pContact.ContactNumber)
        vTableName = "organisation_suppressions cs"
      Else
        vWhereFields.Add("contact_number", pContact.ContactNumber)
        vTableName = "contact_suppressions cs"
      End If
      Dim vSQL As New SQLStatement(mvEnv.Connection, GetRecordSetFields, vTableName, vWhereFields)
      Dim vRS As CDBRecordSet = vSQL.GetRecordSet
      Dim vSuppressionExists As Boolean
      While vRS.Fetch
        If mvExisting = False Then
          vSuppressionExists = True
          Exit While
        Else
          If vRS.Fields("valid_from").Value <> ValidFrom OrElse vRS.Fields("valid_to").Value <> ValidTo Then
            vSuppressionExists = True
            Exit While
          End If
        End If
      End While
      Return vSuppressionExists
    End Function


  End Class

End Namespace
