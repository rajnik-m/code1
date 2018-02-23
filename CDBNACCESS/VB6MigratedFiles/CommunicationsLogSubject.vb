

Namespace Access
  Public Class CommunicationsLogSubject

    Public Enum CommunicationsLogSubjectRecordSetTypes 'These are bit values
      clsrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CommunicationsLogSubjectFields
      clsfAll = 0
      clsfCommunicationsLogNumber
      clsfTopic
      clsfSubTopic
      clsfPrimary
      clsfAmendedOn
      clsfAmendedBy
      clsfQuantity
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "communications_log_subjects"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("communications_log_number", CDBField.FieldTypes.cftLong)
          .Add("topic")
          .Add("sub_topic")
          .Add("primary")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("amended_by")
          .Add("quantity", CDBField.FieldTypes.cftNumeric)

          .Item(CommunicationsLogSubjectFields.clsfCommunicationsLogNumber).SetPrimaryKeyOnly()
          .Item(CommunicationsLogSubjectFields.clsfTopic).SetPrimaryKeyOnly()
          .Item(CommunicationsLogSubjectFields.clsfSubTopic).SetPrimaryKeyOnly()

          .Item(CommunicationsLogSubjectFields.clsfPrimary).SpecialColumn = True
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CommunicationsLogSubjectFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(CommunicationsLogSubjectFields.clsfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CommunicationsLogSubjectFields.clsfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CommunicationsLogSubjectRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CommunicationsLogSubjectRecordSetTypes.clsrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cls")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCommunicationsLogNumber As Integer = 0, Optional ByRef pTopic As String = "", Optional ByRef pSubTopic As String = "")
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      mvEnv = pEnv
      If pCommunicationsLogNumber > 0 Then
        vWhereFields.Add("communications_log_number", CDBField.FieldTypes.cftLong, pCommunicationsLogNumber)
        If Len(pTopic) > 0 And Len(pSubTopic) > 0 Then
          vWhereFields.Add("topic", CDBField.FieldTypes.cftCharacter, pTopic)
          vWhereFields.Add("sub_topic", CDBField.FieldTypes.cftCharacter, pSubTopic)
        Else
          vWhereFields.Add("primary", CDBField.FieldTypes.cftCharacter, "Y").SpecialColumn = True
          vWhereFields.TableAlias = "cls"
        End If
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CommunicationsLogSubjectRecordSetTypes.clsrtAll) & " FROM communications_log_subjects cls WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CommunicationsLogSubjectRecordSetTypes.clsrtAll)
        Else
          InitClassFields()
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CommunicationsLogSubjectRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CommunicationsLogSubjectFields.clsfCommunicationsLogNumber, vFields)
        .SetItem(CommunicationsLogSubjectFields.clsfTopic, vFields)
        .SetItem(CommunicationsLogSubjectFields.clsfSubTopic, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CommunicationsLogSubjectRecordSetTypes.clsrtAll) = CommunicationsLogSubjectRecordSetTypes.clsrtAll Then
          .SetItem(CommunicationsLogSubjectFields.clsfPrimary, vFields)
          .SetItem(CommunicationsLogSubjectFields.clsfAmendedOn, vFields)
          .SetItem(CommunicationsLogSubjectFields.clsfAmendedBy, vFields)
          .SetItem(CommunicationsLogSubjectFields.clsfQuantity, vFields)
        End If
      End With
    End Sub

    Public Sub Update(ByRef pTopic As String, ByRef pSubTopic As String, ByRef pQuantity As String)
      With mvClassFields
        .Item(CommunicationsLogSubjectFields.clsfTopic).Value = pTopic
        .Item(CommunicationsLogSubjectFields.clsfSubTopic).Value = pSubTopic
        If Len(pQuantity) > 0 Then .Item(CommunicationsLogSubjectFields.clsfQuantity).Value = pQuantity
      End With
    End Sub

    Public Sub MakePrimary()
      mvClassFields.Item(CommunicationsLogSubjectFields.clsfPrimary).Value = "Y"
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields

      SetValid(CommunicationsLogSubjectFields.clsfAll)
      If mvExisting And mvClassFields.Item(CommunicationsLogSubjectFields.clsfPrimary).ValueChanged And Primary = True Then
        vWhereFields.Add("communications_log_number", CDBField.FieldTypes.cftLong, CommunicationsLogNumber)
        vUpdateFields.Add("primary", CDBField.FieldTypes.cftCharacter, "N").SpecialColumn = True
        mvEnv.Connection.UpdateRecords((mvClassFields.DatabaseTableName), vUpdateFields, vWhereFields)
      End If
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByRef pCommsLogNumber As Integer, ByRef pTopic As String, ByRef pSubTopic As String, ByRef pPrimary As Boolean, ByRef pQuantity As String)
      With mvClassFields
        .Item(CommunicationsLogSubjectFields.clsfCommunicationsLogNumber).IntegerValue = pCommsLogNumber
        .Item(CommunicationsLogSubjectFields.clsfTopic).Value = pTopic
        .Item(CommunicationsLogSubjectFields.clsfSubTopic).Value = pSubTopic
        .Item(CommunicationsLogSubjectFields.clsfPrimary).Bool = pPrimary
        If Len(pQuantity) > 0 Then .Item(CommunicationsLogSubjectFields.clsfQuantity).DoubleValue = Val(pQuantity)
      End With
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(CommunicationsLogSubjectFields.clsfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CommunicationsLogSubjectFields.clsfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CommunicationsLogNumber() As Integer
      Get
        CommunicationsLogNumber = mvClassFields.Item(CommunicationsLogSubjectFields.clsfCommunicationsLogNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Primary() As Boolean
      Get
        Primary = mvClassFields.Item(CommunicationsLogSubjectFields.clsfPrimary).Bool
      End Get
    End Property

    Public ReadOnly Property Quantity() As Double
      Get
        Quantity = mvClassFields.Item(CommunicationsLogSubjectFields.clsfQuantity).DoubleValue
      End Get
    End Property

    Public ReadOnly Property QuantitySet() As Boolean
      Get
        QuantitySet = mvClassFields.Item(CommunicationsLogSubjectFields.clsfQuantity).Value <> ""
      End Get
    End Property

    Public ReadOnly Property SubTopic() As String
      Get
        SubTopic = mvClassFields.Item(CommunicationsLogSubjectFields.clsfSubTopic).Value
      End Get
    End Property

    Public ReadOnly Property Topic() As String
      Get
        Topic = mvClassFields.Item(CommunicationsLogSubjectFields.clsfTopic).Value
      End Get
    End Property
  End Class
End Namespace
