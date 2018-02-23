Imports System.Linq


Namespace Access
  Public Class ContactExternalLink
    Implements IDbLoadable, IDbSelectable

    Public Enum ContactExternalLinkRecordSetTypes 'These are bit values
      celrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Protected Friend Enum ContactExternalLinkFields
      celfAll = 0
      celfContactNumber
      celfDataSource
      celfExternalReference
      celfAmendedBy
      celfAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Public Sub New()
      'this class doesn't derive from CARERecord so we have to keep a parameterless constructor
    End Sub

    Public Sub New(pEnv As CDBEnvironment)
      Me.Environment = pEnv
    End Sub

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "contact_external_links"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("data_source")
          .Add("external_reference")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(ContactExternalLinkFields.celfDataSource).SetPrimaryKeyOnly()
        mvClassFields.Item(ContactExternalLinkFields.celfExternalReference).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As ContactExternalLinkFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(ContactExternalLinkFields.celfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(ContactExternalLinkFields.celfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As ContactExternalLinkRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = ContactExternalLinkRecordSetTypes.celrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cel")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pDataSource As String = "", Optional ByRef pExternalReference As String = "", Optional ByRef pContactNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      mvEnv = pEnv
      If Len(pDataSource) > 0 And ((Len(pExternalReference) > 0) Or pContactNumber > 0) Then
        vWhereFields.Add("data_source", CDBField.FieldTypes.cftCharacter, pDataSource)
        If pExternalReference.Length > 0 Then vWhereFields.Add("external_reference", CDBField.FieldTypes.cftCharacter, pExternalReference)
        If pContactNumber > 0 Then vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ContactExternalLinkRecordSetTypes.celrtAll) & " FROM contact_external_links WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, ContactExternalLinkRecordSetTypes.celrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ContactExternalLinkRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(ContactExternalLinkFields.celfDataSource, vFields)
        .SetItem(ContactExternalLinkFields.celfExternalReference, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And ContactExternalLinkRecordSetTypes.celrtAll) = ContactExternalLinkRecordSetTypes.celrtAll Then
          .SetItem(ContactExternalLinkFields.celfContactNumber, vFields)
          .SetItem(ContactExternalLinkFields.celfAmendedBy, vFields)
          .SetItem(ContactExternalLinkFields.celfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Create(ByRef pContactNumber As Integer, ByRef pDataSource As String, ByRef pExternalReference As String)
      With mvClassFields
        .Item(ContactExternalLinkFields.celfContactNumber).IntegerValue = pContactNumber
        .Item(ContactExternalLinkFields.celfDataSource).Value = pDataSource
        .Item(ContactExternalLinkFields.celfExternalReference).Value = pExternalReference
      End With
    End Sub

    Public Sub Update(Optional ByRef pDataSource As String = "", Optional ByRef pExternalReference As String = "")
      With mvClassFields
        If pDataSource.Length > 0 Then .Item(ContactExternalLinkFields.celfDataSource).Value = pDataSource
        If pExternalReference.Length > 0 Then .Item(ContactExternalLinkFields.celfExternalReference).Value = pExternalReference
      End With
    End Sub

    Public Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(ContactExternalLinkFields.celfAll)
      mvClassFields.VerifyUnique(mvEnv.Connection)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
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
        AmendedBy = mvClassFields.Item(ContactExternalLinkFields.celfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(ContactExternalLinkFields.celfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(ContactExternalLinkFields.celfContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property DataSource() As String
      Get
        DataSource = mvClassFields.Item(ContactExternalLinkFields.celfDataSource).Value
      End Get
    End Property

    Public ReadOnly Property ExternalReference() As String
      Get
        ExternalReference = mvClassFields.Item(ContactExternalLinkFields.celfExternalReference).Value
      End Get
    End Property
    Protected Friend Property ClassFields As ClassFields
      Get
        If mvClassFields Is Nothing Then InitClassFields()
        Return mvClassFields
      End Get
      Private Set(value As ClassFields)
        mvClassFields = value
      End Set
    End Property

    Public Property Environment As CDBEnvironment
      Get
        Return mvEnv
      End Get
      Set(value As CDBEnvironment)
        mvEnv = value
      End Set
    End Property

    Private Sub InitFromDataRow(ByVal pDataRow As DataRow, ByVal pUseProperName As Boolean) 'Copied from CARERecord
      InitClassFields()
      mvExisting = True
      Dim vName As String
      For Each vClassField As ClassField In mvClassFields
        If pUseProperName Then vName = vClassField.ProperName Else vName = vClassField.Name
        If pDataRow.Table.Columns.Contains(vName) Then
          vClassField.SetValue = pDataRow.Item(vName).ToString
        End If
      Next
    End Sub

    Public Sub LoadFromRow(pRow As DataRow) Implements IDbLoadable.LoadFromRow
      InitFromDataRow(pRow, False)
    End Sub
    Public ReadOnly Property FieldNames As String Implements IDbSelectable.DbFieldNames
      Get
        Return Me.ClassFields.FieldNames(mvEnv, Me.ClassFields.TableAlias)
      End Get
    End Property

    Public ReadOnly Property AliasedTableName As String Implements IDbSelectable.DbAliasedTableName
      Get
        Return Me.ClassFields.TableNameAndAlias
      End Get
    End Property
    Protected Friend Function CreateWhere(pFieldIndexes As IEnumerable(Of Integer)) As CDBFields
      Dim vWhere As New CDBFields
      pFieldIndexes.ToList().ForEach(Sub(vIndex) vWhere.Add(ClassFields(vIndex).Name, ClassFields(vIndex).FieldType, ClassFields(vIndex).Value))
      Return vWhere
    End Function
  End Class
End Namespace
