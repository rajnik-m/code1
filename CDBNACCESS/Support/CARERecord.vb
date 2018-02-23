Imports System.Data
Imports System.Linq
Imports System.Collections.Generic
Imports Advanced.Data.Merge

Namespace Access

  Public Interface IRecordCreate
    Function CreateInstance(ByVal pEnv As CDBEnvironment) As CARERecord
  End Interface

  Public MustInherit Class CARERecord
    Implements ICARERecord, IDbLoadable, IDbSelectable

    'Standard Class Setup
    Protected mvEnv As CDBEnvironment
    Protected mvClassFields As ClassFields
    Protected mvExisting As Boolean
    Protected mvMaintenanceAttributes As CollectionList(Of MaintenanceAttribute)
    Protected mvOverrideAmended As Boolean
    Protected mvOverrideCreated As Boolean

    Public Enum MaintenanceTypes
      SelectData
      Insert
      Update
      Delete
    End Enum

    Protected Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = DatabaseTableName 'Set the table name
          .TableAlias = TableAlias 'Set the alias
          AddFields() 'This should add an entry for each field in the table
          If SupportsAmendedOnAndBy Then
            .Add("amended_by").PrefixRequired = True
            .Add("amended_on", CDBField.FieldTypes.cftDate).PrefixRequired = True
          End If
          AddAdditionalFields()
        End With
        AddDeleteCheckItems()
      Else
        mvClassFields.ClearItems()
      End If
      ClearFields()
      mvOverrideAmended = False
      mvOverrideCreated = False
      mvExisting = False
    End Sub

    Protected Sub CheckClassFields()
      If mvClassFields Is Nothing Then InitClassFields()
    End Sub

    Public Sub InitFromMaintenanceData(ByVal pTableName As String) Implements ICARERecord.InitFromMaintenanceData
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        Dim vClassField As ClassField
        With mvClassFields
          .DatabaseTableName = pTableName 'Set the table name
          .TableAlias = pTableName        'Set the alias
          For Each vMaintenanceAttribute As MaintenanceAttribute In MaintenanceAttributes
            vClassField = .Add(vMaintenanceAttribute.AttributeName, vMaintenanceAttribute.FieldType)
            vClassField.PrimaryKey = vMaintenanceAttribute.PrimaryKey
            vClassField.SpecialColumn = mvEnv.Connection.IsSpecialColumn(vMaintenanceAttribute.AttributeName)
          Next
          AddAdditionalFields()
        End With
        mvExisting = False
      End If
    End Sub

    Public ReadOnly Property MaintenanceAttributes() As CollectionList(Of MaintenanceAttribute) Implements ICARERecord.MaintenanceAttributes
      Get
        If mvMaintenanceAttributes Is Nothing Then InitMaintenanceAttributes(DatabaseTableName)
        Return mvMaintenanceAttributes
      End Get
    End Property

    Private Sub InitMaintenanceAttributes(ByVal pTableName As String)
      Dim vMA As New MaintenanceAttribute(mvEnv)
      Dim vRecordSet As CDBRecordSet
      vRecordSet = New SQLStatement(mvEnv.Connection, vMA.GetRecordSetFields, "maintenance_attributes ma", New CDBField("table_name", pTableName), "sequence_number").GetRecordSet
      mvMaintenanceAttributes = New CollectionList(Of MaintenanceAttribute)
      While vRecordSet.Fetch
        vMA = New MaintenanceAttribute(mvEnv)
        vMA.InitFromRecordSet(vRecordSet)
        If Not mvMaintenanceAttributes.ContainsKey(vMA.AttributeName) Then mvMaintenanceAttributes.Add(vMA.AttributeName, vMA)
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Protected MustOverride Sub AddFields()
    Protected MustOverride ReadOnly Property DatabaseTableName() As String
    Protected MustOverride ReadOnly Property TableAlias() As String
    Protected MustOverride ReadOnly Property SupportsAmendedOnAndBy() As Boolean

    Public Sub New(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
    End Sub

    Public Sub SetEnvironment(ByVal pEnv As CDBEnvironment) Implements ICARERecord.SetEnvironment
      Environment = pEnv
    End Sub

    Public Overridable ReadOnly Property NeedsMaintenanceInfo() As Boolean Implements ICARERecord.NeedsMaintenanceInfo
      Get
        Return False
      End Get
    End Property

    Public Overridable Function KeyValueRequired(ByVal pField As String) As Boolean Implements ICARERecord.KeyValueRequired
      'Override this function if its ok for an attribute that is part of the primary key to have a null value
      Return mvClassFields(pField).PrimaryKey
    End Function


    Protected Overridable Sub AddAdditionalFields()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Protected Overridable Sub ClearFields()
      'Add code here to clear any additional values held by the class
    End Sub

    Protected Overridable Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Public Overridable Function GetAddRecordMandatoryParameters() As String Implements ICARERecord.GetAddRecordMandatoryParameters
      Return ""
    End Function

    Public Overridable ReadOnly Property IsMaintenanceTable() As Boolean Implements ICARERecord.IsMaintenanceTable
      Get
        Return False
      End Get
    End Property

    Public Overridable ReadOnly Property NoUniqueKey() As Boolean Implements ICARERecord.NoUniqueKey
      Get
        CheckClassFields()
        Dim vFields As New CDBFields
        For Each vClassField As ClassField In mvClassFields
          If vClassField.PrimaryKey Then
            Return False
          End If
        Next
        Return True
      End Get
    End Property

    Public Overridable Sub PreValidateParameterList(ByVal pType As MaintenanceTypes, ByVal pParameterList As CDBParameters)
      'Add code here to validate parameters passed to the create methods
    End Sub
    Protected Overridable Sub PreValidateCreateParameters(ByVal pParameterList As CDBParameters)
      'Add code here to validate parameters passed to the create methods
    End Sub

    Protected Overridable Sub PreValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      'Add code here to validate parameters passed to the update methods
    End Sub

    Protected Overridable Sub PostValidateCreateParameters(ByVal pParameterList As CDBParameters)
      'Add code here to validate parameters passed to the create methods
    End Sub

    Protected Overridable Sub PostValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      'Add code here to validate parameters passed to the update methods
    End Sub

    Protected Overridable Sub ValidateRecordExists(ByVal pPrimaryKey As Integer, ByVal pRecordName As String, ByVal pRecord As CARERecord)
      pRecord.Init(pPrimaryKey)
      If pRecord.Existing = False Then RaiseError(DataAccessErrors.daeRecordDoesNotExists, pRecordName)
    End Sub

    Protected Overridable Sub SetValid()
      'Add code here to ensure all values are valid before saving
      If SupportsAmendedOnAndBy And mvOverrideAmended = False Then
        mvClassFields.Item("amended_on").Value = TodaysDate()
        mvClassFields.Item("amended_by").Value = If(mvEnv.InitialisingDatabase AndAlso String.IsNullOrWhiteSpace(mvEnv.User.UserID), "dbinit", mvEnv.User.UserID)
      End If
      mvClassFields.SetControlNumber(mvEnv)
      If mvExisting = False AndAlso mvClassFields.ContainsKey("created_by") AndAlso Not mvOverrideCreated Then
        mvClassFields.Item("created_by").Value = If(mvEnv.InitialisingDatabase AndAlso String.IsNullOrWhiteSpace(mvEnv.User.UserID), "dbinit", mvEnv.User.UserID)
        If mvClassFields.ContainsKey("created_on") Then mvClassFields.Item("created_on").Value = TodaysDateAndTime()
      End If
    End Sub

    Public Sub SetControlNumber() Implements ICARERecord.SetControlNumber
      mvClassFields.SetControlNumber(mvEnv)
    End Sub

    Public Overridable Function GetRecordSetFields() As String Implements ICARERecord.GetRecordSetFields
      CheckClassFields()
      Return mvClassFields.FieldNames(mvEnv, TableAlias)
    End Function

    Public Overridable Function GetRecordSetFieldsExclude(ByVal pExcludeFields As List(Of ClassField)) As String Implements ICARERecord.GetRecordSetFieldsExclude
      CheckClassFields()
      Return mvClassFields.FieldNames(mvEnv, TableAlias, pExcludeFields)
    End Function

    Public Sub Init() Implements ICARERecord.Init
      InitClassFields()
      SetDefaults()
    End Sub

    Public Overridable Sub Init(ByVal pPrimaryKeyValue As Integer) Implements ICARERecord.Init
      Init()
      Dim vWhereFields As New CDBFields
      Dim vPrimaryKey As ClassField = mvClassFields.GetUniquePrimaryKey
      If vPrimaryKey.FieldType <> CDBField.FieldTypes.cftInteger AndAlso vPrimaryKey.FieldType <> CDBField.FieldTypes.cftLong Then RaiseError(DataAccessErrors.daePrimaryKeyWrongDataType, mvClassFields.DatabaseTableName)
      vWhereFields.Add(vPrimaryKey.Name, CDBField.FieldTypes.cftLong, pPrimaryKeyValue.ToString)
      InitWithPrimaryKey(vWhereFields)
    End Sub

    Public Sub Init(ByVal pPrimaryKeyValue As String) Implements ICARERecord.Init
      Init()
      Dim vPrimaryKey As ClassField = mvClassFields.GetUniquePrimaryKey
      If vPrimaryKey.FieldType <> CDBField.FieldTypes.cftCharacter Then RaiseError(DataAccessErrors.daePrimaryKeyWrongDataType, mvClassFields.DatabaseTableName)
      Dim vWhereFields As New CDBFields(New CDBField(vPrimaryKey.Name, pPrimaryKeyValue))
      InitWithPrimaryKey(vWhereFields)
    End Sub

    Public Overridable Sub InitForUpdate(ByVal pParams As CDBParameters) Implements ICARERecord.InitForUpdate
      Init()
      InitWithPrimaryKey(GetUniqueKeyFields(pParams))
    End Sub

    Public Sub Init(ByVal pParams As CDBParameters) Implements ICARERecord.Init
      Init()
      InitWithPrimaryKey(GetUniqueKeyFields(pParams))
    End Sub

    Public Sub InitWithPrimaryKey(ByVal pWhereFields As CDBFields) Implements ICARERecord.InitWithPrimaryKey
      Dim vSQL As New SQLStatement(mvEnv.Connection, GetRecordSetFields(), mvClassFields.TableNameAndAlias, pWhereFields)
      Dim vRecordSet As CDBRecordSet = vSQL.GetRecordSet
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(vRecordSet)
      Else
        Init()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub Clone(ByVal pRecord As CARERecord) Implements ICARERecord.Clone   'Use when the primary key is a control number
      Clone(pRecord, 0)
    End Sub

    Public Sub Clone(ByVal pRecord As CARERecord, ByVal pPrimaryKeyValue As Integer) Implements ICARERecord.Clone
      CopyValues(pRecord)
      Dim vClassField As ClassField = mvClassFields.GetUniquePrimaryKey
      vClassField.IntegerValue = pPrimaryKeyValue
    End Sub

    Public Sub Clone(ByVal pRecord As CARERecord, ByVal pPrimaryKeyValue As String) Implements ICARERecord.Clone
      CopyValues(pRecord)
      Dim vClassField As ClassField = mvClassFields.GetUniquePrimaryKey
      vClassField.Value = pPrimaryKeyValue
    End Sub

    Public Sub Clone(ByVal pRecord As CARERecord, ByVal pParams As CDBParameters) Implements ICARERecord.Clone
      Clone(pRecord, 0)
      Update(pParams, False)
    End Sub

    Public Sub CopyValues(ByVal pRecord As CARERecord) Implements ICARERecord.CopyValues
      Init()
      For vIndex As Integer = 1 To mvClassFields.Count
        mvClassFields(vIndex).Value = pRecord.mvClassFields(vIndex).Value
      Next
    End Sub

    Public Sub Create(ByVal pParameterList As CDBParameters, ByVal pPrimaryKeyValue As Integer) Implements ICARERecord.Create
      Init()
      PreValidateCreateParameters(pParameterList)
      Dim vClassField As ClassField = mvClassFields.GetUniquePrimaryKey
      vClassField.IntegerValue = pPrimaryKeyValue
      Update(pParameterList, False)
      PostValidateCreateParameters(pParameterList)
    End Sub

    Public Sub Create(ByVal pParameterList As CDBParameters) Implements ICARERecord.Create
      Init()
      PreValidateCreateParameters(pParameterList)
      For Each vClassField As ClassField In mvClassFields
        If vClassField.PrimaryKey = True Then
          If pParameterList.ContainsKey(vClassField.ParameterName) Then vClassField.Value = pParameterList(vClassField.ParameterName).Value
        End If
      Next
      Update(pParameterList, False)
      PostValidateCreateParameters(pParameterList)
    End Sub

    Public Overridable Sub Update(ByVal pParameterList As CDBParameters) Implements ICARERecord.Update
      Update(pParameterList, True)
    End Sub

    Private Sub Update(ByVal pParameterList As CDBParameters, ByVal pValidate As Boolean)
      If pValidate Then PreValidateUpdateParameters(pParameterList)
      For Each vClassField As ClassField In mvClassFields
        If vClassField.PrimaryKey = False AndAlso (vClassField.NonUpdatable = False OrElse pValidate = False) Then
          If pParameterList.ContainsKey(vClassField.ParameterName) Then vClassField.Value = pParameterList(vClassField.ParameterName).Value
        End If
      Next
      If pValidate Then PostValidateUpdateParameters(pParameterList)
    End Sub

    Public Overridable Sub InitFromRecordSet(ByVal pRecordSet As CDBRecordSet) Implements ICARERecord.InitFromRecordSet
      InitClassFields()
      Dim vFields As CDBFields = pRecordSet.Fields
      mvExisting = True
      For Each vClassField As ClassField In mvClassFields
        If vClassField.InDatabase AndAlso vClassField.FieldType <> CDBField.FieldTypes.cftBulk Then
          If vClassField.FieldType = CDBField.FieldTypes.cftBinary Then
            vClassField.SetValue = vFields(vClassField.Name).Value
            vClassField.ByteValue = vFields(vClassField.Name).ByteValue
          Else
            vClassField.SetValue = vFields(vClassField.Name).Value
          End If
        Else
          If vFields.Exists(vClassField.Name) Then
            vClassField.SetValue = vFields(vClassField.Name).Value
          End If
        End If
      Next
    End Sub

    Protected Sub InitFromRecordSetFields(ByVal pRecordSet As CDBRecordSet, ByVal pFields As String)
      InitFromRecordSetFields(pRecordSet, pFields, "")
    End Sub

    Protected Sub InitFromRecordSetFields(ByVal pRecordSet As CDBRecordSet, ByVal pFields As String, ByVal pPrefixName As String)
      InitClassFields()
      Dim vFields As CDBFields = pRecordSet.Fields
      mvExisting = True
      Dim vFieldNames As String() = pFields.Split(","c)
      Dim vCFieldName As String
      For Each vFieldName As String In vFieldNames
        Dim vPos As Integer = vFieldName.IndexOf(".")
        If vPos >= 0 Then vFieldName = vFieldName.Substring(vPos + 1)
        If pPrefixName.Length > 0 Then
          vCFieldName = vFieldName.Replace(pPrefixName & "_", "")
        Else
          vCFieldName = vFieldName
        End If
        mvClassFields(vCFieldName).SetValue = vFields(vFieldName).Value
      Next
    End Sub

    Public Sub InitFromXMLNode(ByVal pNode As Xml.XmlNode) Implements ICARERecord.InitFromXMLNode
      CheckClassFields()
      Dim vAttr As Xml.XmlAttribute
      For vIndex As Integer = 0 To pNode.Attributes.Count - 1
        vAttr = pNode.Attributes(vIndex)
        mvClassFields(AttributeName(vAttr.Name)).Value = vAttr.InnerText
      Next
    End Sub

#Region "Delete Handling"

    Private mvDeleteCheckItems As List(Of DeleteCheckItem)
    Private mvCascadeDeleteItems As List(Of DeleteCheckItem)
    Private mvRelations As IEnumerable(Of MaintenanceRelation)

    <MergeDeleteMethod()>
    Public Sub Delete() Implements ICARERecord.Delete
      Delete("", False, 0)
    End Sub
    Public Sub Delete(ByVal pAmendedBy As String) Implements ICARERecord.Delete
      Delete(pAmendedBy, False, 0)
    End Sub
    Public Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean) Implements ICARERecord.Delete
      Delete(pAmendedBy, pAudit, 0)
    End Sub
    Public Overridable Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer) Implements ICARERecord.Delete
      Dim vCheckFailItem As DeleteCheckItem = GetFirstDeleteCheckItem()
      If vCheckFailItem IsNot Nothing Then
        RaiseError(DataAccessErrors.daeReferencedInOtherTable, vCheckFailItem.Description)
      End If
      Dim vTransactionStarted As Boolean
      If mvCascadeDeleteItems IsNot Nothing Then
        If mvEnv.Connection.InTransaction = False Then
          mvEnv.Connection.StartTransaction()
          vTransactionStarted = True
        End If
        CascadeDeleteDependentItems()
      End If
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit, pJournalNumber)
      If vTransactionStarted Then mvEnv.Connection.CommitTransaction()
    End Sub

    Public Sub DeleteByForeignKey(ByVal pField As CDBField) Implements ICARERecord.DeleteByForeignKey
      Dim vWhereFields As New CDBFields(pField)
      mvEnv.Connection.DeleteRecords(DatabaseTableName, vWhereFields, False)
    End Sub
    Public Sub DeleteByForeignKeys(ByVal pWhereFields As CDBFields) Implements ICARERecord.DeleteByForeignKeys
      mvEnv.Connection.DeleteRecords(DatabaseTableName, pWhereFields, False)
    End Sub

    Public Overridable Sub AddDeleteCheckItems() Implements ICARERecord.AddDeleteCheckItems

    End Sub

    Protected Sub AddDeleteCheckItem(ByVal pTableName As String, ByVal pAttributeName As String, ByVal pDescription As String)
      If mvDeleteCheckItems Is Nothing Then mvDeleteCheckItems = New List(Of DeleteCheckItem)
      mvDeleteCheckItems.Add(New DeleteCheckItem(pTableName, pAttributeName, pDescription))
    End Sub

    Protected Sub AddCascadeDeleteItem(ByVal pTableName As String, ByVal pAttributeName As String)
      If mvCascadeDeleteItems Is Nothing Then mvCascadeDeleteItems = New List(Of DeleteCheckItem)
      mvCascadeDeleteItems.Add(New DeleteCheckItem(pTableName, pAttributeName, ""))
    End Sub

    ''' <summary>
    ''' Runs through the list of DeleteCheckItems and returns the first one that fails.  The DeleteCheckItems are set up at the derived class by calling AddDeleteCheckItem.  If none fail then it returns Nothing.
    ''' </summary>
    ''' <returns>DeleteCheckItem class representing the table and attribute for the relationship that fails.</returns>
    ''' <remarks>This is an overridable method that is called by the Delete method.  In most cases it will be sufficient to check for dependent records by using AddDeleteCheckItem.
    ''' If however the check has to be more complex then this method can be overridden.
    ''' </remarks>
    Protected Overridable Function GetFirstDeleteCheckItem() As DeleteCheckItem
      Dim vRtn As DeleteCheckItem = Nothing
      If mvDeleteCheckItems IsNot Nothing Then
        For Each vCheckItem As DeleteCheckItem In mvDeleteCheckItems
          Dim vWhereFields As New CDBFields
          Dim vPrimaryKey As ClassField = mvClassFields.GetUniquePrimaryKey
          If vPrimaryKey.FieldType = CDBField.FieldTypes.cftInteger AndAlso vPrimaryKey.Value.Contains(",") Then
            vWhereFields.Add(vCheckItem.AttributeName, vPrimaryKey.FieldType, vPrimaryKey.Value, CDBField.FieldWhereOperators.fwoIn)
          Else
            vWhereFields.Add(vCheckItem.AttributeName, vPrimaryKey.FieldType, vPrimaryKey.Value)
          End If
          If mvEnv.Connection.GetCount(vCheckItem.TableName, vWhereFields) > 0 Then
            vRtn = vCheckItem
            Exit For
          End If
        Next
      End If
      Return vRtn
    End Function
    Protected Overridable Sub CascadeDeleteDependentItems()
      For Each vCheckItem As DeleteCheckItem In mvCascadeDeleteItems
        Dim vWhereFields As New CDBFields
        Dim vPrimaryKey As ClassField = mvClassFields.GetUniquePrimaryKey
        vWhereFields.Add(vCheckItem.AttributeName, vPrimaryKey.FieldType, vPrimaryKey.Value)
        mvEnv.Connection.DeleteRecords(vCheckItem.TableName, vWhereFields, False)
      Next
    End Sub
    Protected Class DeleteCheckItem
      Private mvTableName As String
      Private mvAttributeName As String
      Private mvDescription As String

      Public Sub New(ByVal pTableName As String, ByVal pAttributeName As String, ByVal pDescription As String)
        mvTableName = pTableName
        mvAttributeName = pAttributeName
        mvDescription = pDescription
      End Sub

      Public ReadOnly Property TableName As String
        Get
          Return mvTableName
        End Get
      End Property

      Public ReadOnly Property AttributeName As String
        Get
          Return mvAttributeName
        End Get
      End Property

      Public ReadOnly Property Description As String
        Get
          Return mvDescription
        End Get
      End Property

    End Class
#End Region

    Public Function RecordExists(ByVal pField As CDBField) As Boolean Implements ICARERecord.RecordExists
      Dim vWhereFields As New CDBFields
      vWhereFields.Add(pField)
      Return RecordExists(vWhereFields)
    End Function
    Public Function RecordExists(ByVal pWhereFields As CDBFields) As Boolean Implements ICARERecord.RecordExists
      If mvClassFields Is Nothing Then Init()
      Return mvEnv.Connection.GetCount(mvClassFields.DatabaseTableName, pWhereFields) > 0
    End Function

    <MergeSaveMethod()>
    Public Sub Save() Implements ICARERecord.Save
      Save("", False, 0)
    End Sub
    Public Sub Save(ByVal pAmendedBy As String) Implements ICARERecord.Save
      Save(pAmendedBy, False, 0)
    End Sub
    Public Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean) Implements ICARERecord.Save
      Save(pAmendedBy, pAudit, 0)
    End Sub
    Public Overridable Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer) Implements ICARERecord.Save
      SetValid()
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit, pJournalNumber)
    End Sub
    ''' <summary>
    ''' Save with amendment history
    ''' </summary>
    ''' <param name="pAmendedBy"></param>
    ''' <param name="pAudit"></param>
    ''' <param name="pJournalNumber"></param>
    ''' <param name="pForceAmendmentHistory"></param>
    ''' <remarks></remarks>
    Public Overridable Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer, pForceAmendmentHistory As Boolean) Implements ICARERecord.Save
      SetValid()
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit, pJournalNumber, pForceAmendmentHistory)
    End Sub

    Public ReadOnly Property Existing() As Boolean Implements ICARERecord.Existing
      Get
        Return mvExisting
      End Get
    End Property

    Public Overridable ReadOnly Property DataTable() As CDBDataTable Implements ICARERecord.DataTable
      Get
        CheckClassFields()
        Return mvClassFields.DataTable
      End Get
    End Property

    Public Overridable ReadOnly Property DataTableColumnNames() As String Implements ICARERecord.DataTableColumnNames
      Get
        Return mvClassFields.DataTableColumnNames
      End Get
    End Property

    Public Sub AddDataColumns(ByVal pDataTable As CDBDataTable) Implements ICARERecord.AddDataColumns
      AddDataColumns(pDataTable, False)
    End Sub

    Public Sub AddDataColumns(ByVal pDataTable As CDBDataTable, ByVal pUseProperNames As Boolean) Implements ICARERecord.AddDataColumns
      If mvClassFields Is Nothing Then Init()
      For Each vClassField As ClassField In mvClassFields
        If pUseProperNames Then
          pDataTable.AddColumn(vClassField.ProperName, vClassField.FieldType)
        Else
          pDataTable.AddColumn(vClassField.Name, vClassField.FieldType)
        End If
      Next
    End Sub

    Public Sub AddDataRow(ByVal pDataTable As CDBDataTable) Implements ICARERecord.AddDataRow
      AddDataRow(pDataTable, False)
    End Sub

    Public Sub AddDataRow(ByVal pDataTable As CDBDataTable, ByVal pUseProperNames As Boolean) Implements ICARERecord.AddDataRow
      Dim vRow As CDBDataRow = pDataTable.AddRow
      For Each vClassField As ClassField In mvClassFields
        If pUseProperNames Then
          vRow.Item(vClassField.ProperName) = vClassField.Value
        Else
          vRow.Item(vClassField.Name) = vClassField.Value
        End If
      Next
    End Sub

    Public Function FieldValueString(ByVal pFieldName As String) As String Implements ICARERecord.FieldValueString
      Return mvClassFields(pFieldName).Value
    End Function

    Public Function FieldValueInteger(ByVal pFieldName As String) As Integer Implements ICARERecord.FieldValueInteger
      Return mvClassFields(pFieldName).IntegerValue
    End Function

    Public Function FieldValueChanged(ByVal pFieldName As String) As Boolean Implements ICARERecord.FieldValueChanged
      Return mvClassFields(pFieldName).ValueChanged
    End Function

    Public Overridable Function Validate() As Boolean Implements ICARERecord.Validate
      Dim vValid As Boolean = True
      If mvMaintenanceAttributes Is Nothing Then InitMaintenanceAttributes(DatabaseTableName)
      For Each vMA As MaintenanceAttribute In mvMaintenanceAttributes
        If vMA.NullsInvalid = True AndAlso (vMA.AttributeName <> "amended_by" AndAlso vMA.AttributeName <> "amended_on") Then
          If mvClassFields.ContainsKey(vMA.AttributeName) Then
            Dim vClassField As ClassField = mvClassFields(vMA.AttributeName)
            If vClassField.Value.Length = 0 Then
              vValid = False
            End If
          Else
            vValid = False
          End If
          If vValid = False Then Exit For
        End If
      Next
      Return vValid
    End Function

    Public Function GetList(Of ItemType As CARERecord)(ByVal pItem As IRecordCreate, ByVal pWhereFields As CDBFields) As List(Of ItemType) Implements ICARERecord.GetList
      Dim vSQL As New SQLStatement(mvEnv.Connection, GetRecordSetFields, String.Format("{0} {1}", DatabaseTableName, TableAlias), pWhereFields)
      Dim vRS As CDBRecordSet = vSQL.GetRecordSet
      Dim vList As New List(Of ItemType)
      While vRS.Fetch
        Dim vRecord As ItemType = DirectCast(pItem.CreateInstance(mvEnv), ItemType)
        vRecord.InitFromRecordSet(vRS)
        vList.Add(vRecord)
      End While
      vRS.CloseRecordSet()
      Return vList
    End Function

    Public Function GetDataTableFromList(ByVal pList As List(Of CARERecord)) As CDBDataTable Implements ICARERecord.GetDataTableFromList
      If pList.Count > 0 Then
        Dim vDataTable As CDBDataTable = pList(0).DataTable
        For vIndex As Integer = 1 To pList.Count - 1
          pList(vIndex).mvClassFields.AddToDataTable(vDataTable)
        Next
        Return vDataTable
      Else
        Return DataTable
      End If
    End Function

    Public Overridable Function GetUniqueKeyFields(ByVal pParams As CDBParameters) As CDBFields Implements ICARERecord.GetUniqueKeyFields
      CheckClassFields()
      Dim vFields As New CDBFields
      For Each vClassField As ClassField In mvClassFields
        If vClassField.PrimaryKey Then
          vFields.Add(New CDBField(vClassField.Name, vClassField.FieldType, pParams(vClassField.ProperName).Value))
        End If
      Next
      Return vFields
    End Function

    Public Function GetUniqueKeyFields() As CDBFields Implements ICARERecord.GetUniqueKeyFields
      CheckClassFields()
      Dim vFields As New CDBFields
      For Each vClassField As ClassField In mvClassFields
        If vClassField.PrimaryKey Then
          vFields.Add(New CDBField(vClassField.Name, vClassField.FieldType, vClassField.Value))
        End If
      Next
      Return vFields
    End Function

    Public Function GetValuesAsFields() As CDBFields Implements ICARERecord.GetValuesAsFields
      CheckClassFields()
      Dim vFields As New CDBFields
      For Each vClassField As ClassField In mvClassFields
        If vClassField.Value.Length > 0 Then
          vFields.Add(New CDBField(vClassField.Name, vClassField.FieldType, vClassField.Value))
        End If
      Next
      Return vFields
    End Function

    Public Overridable Function GetUpdateKeyFieldNames() As String Implements ICARERecord.GetUpdateKeyFieldNames
      Return GetUniqueKeyFieldNames()
    End Function

    Public Overridable Function GetUniqueKeyFieldNames() As String Implements ICARERecord.GetUniqueKeyFieldNames
      CheckClassFields()
      Dim vNames As New StringBuilder
      Dim vAdded As Boolean
      For Each vClassField As ClassField In mvClassFields
        If vClassField.PrimaryKey Then
          If vAdded Then vNames.Append(",")
          vNames.Append(vClassField.ProperName)
          vAdded = True
        End If
      Next
      Return vNames.ToString
    End Function

    Public Overridable Function GetUniqueKeyParameters() As CDBParameters Implements ICARERecord.GetUniqueKeyParameters
      CheckClassFields()
      Dim vParams As New CDBParameters
      For Each vClassField As ClassField In mvClassFields
        If vClassField.PrimaryKey Then
          vParams.Add(vClassField.ProperName, vClassField.FieldType, vClassField.Value)
        End If
      Next
      Return vParams
    End Function

    Public Overridable Function GetAddParameters() As CDBParameters Implements ICARERecord.GetAddParameters
      CheckClassFields()
      Dim vMandatoryNames As String = GetAddRecordMandatoryParameters()
      Dim vList As New ArrayListEx(vMandatoryNames)
      Dim vParams As New CDBParameters
      For Each vClassField As ClassField In mvClassFields
        If Not vClassField.PrimaryKey AndAlso vClassField.Name <> "amended_by" AndAlso vClassField.Name <> "amended_on" Then
          Dim vParam As CDBParameter = vParams.Add(vClassField.ProperName, vClassField.FieldType, vClassField.Value)
          If vList.Contains(vClassField.ProperName) Then vParam.Mandatory = True
        ElseIf vList.Contains(vClassField.ProperName) Then
          Dim vParam As CDBParameter = vParams.Add(vClassField.ProperName, vClassField.FieldType, vClassField.Value)
          vParam.Mandatory = True
        End If
      Next
      Return vParams
    End Function

    Public Overridable Function GetUpdateParameters() As CDBParameters Implements ICARERecord.GetUpdateParameters
      CheckClassFields()
      Dim vMandatoryNames As String = GetUpdateKeyFieldNames()
      Dim vList As New ArrayListEx(vMandatoryNames)
      Dim vParams As New CDBParameters
      For Each vClassField As ClassField In mvClassFields
        If Not vClassField.PrimaryKey AndAlso vClassField.NonUpdatable = False AndAlso vClassField.Name <> "amended_by" AndAlso vClassField.Name <> "amended_on" Then
          Dim vParam As CDBParameter = vParams.Add(vClassField.ProperName, vClassField.FieldType, vClassField.Value)
          If vList.Contains(vClassField.ProperName) Then vParam.Mandatory = True
        ElseIf vList.Contains(vClassField.ProperName) Then
          Dim vParam As CDBParameter = vParams.Add(vClassField.ProperName, vClassField.FieldType, vClassField.Value)
          vParam.Mandatory = True
        ElseIf vList.Contains("Old" & vClassField.ProperName) Then
          Dim vParam As CDBParameter = vParams.Add("Old" & vClassField.ProperName, vClassField.FieldType, vClassField.Value)
          vParam.Mandatory = True
        End If
      Next
      Return vParams
    End Function

    Public Sub SaveAmendedOnChanges() Implements ICARERecord.SaveAmendedOnChanges
      mvClassFields.SaveAmendedOnChanges = True
    End Sub

    Protected Sub ReadBulkAttribute(ByVal pIndex As Integer)
      If mvClassFields(pIndex).FieldType <> CDBField.FieldTypes.cftBulk Then
        Throw New IndexOutOfRangeException
      Else
        Dim vRS As CDBRecordSet = New SQLStatement(mvEnv.Connection, mvClassFields(pIndex).Name, DatabaseTableName, GetUniqueKeyFields).GetRecordSet(CDBConnection.RecordSetOptions.NoDataTable)
        If vRS.Fetch = True Then
          mvClassFields(pIndex).Value = vRS.Fields(1).Value
        End If
        vRS.CloseRecordSet()
      End If
    End Sub

    ''' <summary>
    ''' The <see cref="CDBEnvironment"/> object associated with this entity
    ''' </summary>
    ''' <value>The <see cref="CDBEnvironment"/> object associated with this entity</value>
    ''' <returns>The <see cref="CDBEnvironment"/> object associated with this entity</returns>
    Public Property Environment As CDBEnvironment Implements ICARERecord.Environment
      Get
        Return mvEnv
      End Get
      Private Set(pEnv As CDBEnvironment)
        mvEnv = pEnv
      End Set
    End Property

#Region "Methods from IDEM"

    Public Overridable Sub SetDataRow(ByVal pDataRow As DataRow)
      For Each vClassField As ClassField In mvClassFields
        If pDataRow.Table.Columns.Contains(vClassField.ProperName) Then
          If vClassField.Value.Length > 0 Then
            pDataRow.Item(vClassField.ProperName) = vClassField.Value
          Else
            'Set to null?
          End If
        End If
      Next
    End Sub

    Public Function GetBlankDataSet() As DataSet
      Return GetDataSet(New CDBFields(New CDBField("1", 0)), True, "", 0)
    End Function

    Public Function GetDataSet(ByVal pAddColumnTable As Boolean) As DataSet
      InitClassFields()
      Return GetDataSet(Nothing, pAddColumnTable, DefaultOrderByClause, 0)
    End Function

    Protected Overridable Function DefaultOrderByClause() As String
      Return ""
    End Function

    Public Function GetDataSet(ByVal pWhereFields As CDBFields, ByVal pAddColumnTable As Boolean, ByVal pOrderBy As String, ByVal pMaxRecords As Integer) As DataSet
      Return GetDataSet(pWhereFields, pAddColumnTable, pOrderBy, pMaxRecords, False)
    End Function

    Public Function GetDataSet(ByVal pWhereFields As CDBFields, ByVal pAddColumnTable As Boolean, ByVal pOrderBy As String, ByVal pMaxRecords As Integer, ByVal pKeepAttributeNames As Boolean) As DataSet
      CheckClassFields()
      Dim vSQL As New SQLStatement(mvEnv.Connection, GetRecordSetFields, DatabaseTableName & " " & TableAlias, pWhereFields)
      vSQL.OrderBy = pOrderBy
      vSQL.MaxRows = pMaxRecords
      Dim vDataSet As DataSet = mvEnv.Connection.GetDataSet(vSQL)
      If vDataSet.Tables.Contains("Table") Then
        Dim vTable As DataTable = vDataSet.Tables("Table")
        vTable.TableName = "DataRow"                          'If it has the default table name change it to match our expectations
        If Not pKeepAttributeNames Then
          For Each vCol As DataColumn In vTable.Columns
            vCol.ColumnName = ProperName(vCol.ColumnName)
          Next
        End If
        SetCaptionNames()
        If pAddColumnTable Then
          Dim vColumnTable As DataTable = NewColumnTable()
          For Each vClassField As ClassField In mvClassFields
            Dim vDataType As String = "Char"
            Select Case vClassField.FieldType
              Case CDBField.FieldTypes.cftInteger, CDBField.FieldTypes.cftIdentity
                vDataType = "Integer"
              Case CDBField.FieldTypes.cftLong
                vDataType = "Long"
              Case CDBField.FieldTypes.cftNumeric
                vDataType = "Numeric"
              Case CDBField.FieldTypes.cftDate
                vDataType = "Date"
              Case CDBField.FieldTypes.cftTime
                vDataType = "DateTime"
              Case Else
                vDataType = "Char"
            End Select
            AddDataColumn(vColumnTable, vClassField.ProperName, vClassField.Caption, vDataType)
          Next
          vDataSet.Tables.Add(vColumnTable)
        End If
      End If
      Return vDataSet
    End Function

    Public Function GetAttributeNames() As DataTable
      Return mvEnv.Connection.GetAttributeNames(DatabaseTableName)
    End Function

    Public Function GetDataTable() As DataTable
      Return GetDataTable(Nothing, 0)
    End Function

    Public Function GetDataTable(ByVal pWhereFields As CDBFields) As DataTable
      Return GetDataTable(pWhereFields, 0)
    End Function

    Public Function GetDataTable(ByVal pWhereFields As CDBFields, ByVal pMaxRows As Integer) As DataTable
      Dim vDataSet As DataSet = GetDataSet(pWhereFields, False, DefaultOrderByClause, pMaxRows)
      If vDataSet.Tables.Contains("DataRow") Then
        Return vDataSet.Tables("DataRow")
      Else
        Return Nothing
      End If
    End Function

    Public Shared Function NewColumnTable() As DataTable
      Dim vTable As New DataTable("Column")
      vTable.Columns.AddRange(New DataColumn() _
      {
        New DataColumn("Name"),
        New DataColumn("Heading"),
        New DataColumn("Visible"),
        New DataColumn("DataType")
      })
      Return vTable
    End Function

    Public Shared Sub AddDataColumn(ByVal pTable As DataTable, ByVal pName As String, ByVal pHeading As String, Optional ByVal pDataType As String = "Char", Optional ByVal pVisible As String = "Y")
      Dim vRow As DataRow = pTable.NewRow
      With vRow
        .Item("Name") = pName
        .Item("Heading") = pHeading
        .Item("DataType") = pDataType
        .Item("Visible") = pVisible
      End With
      pTable.Rows.Add(vRow)
    End Sub

    Public Sub InitFromDataRow(ByVal pDataRow As DataRow)
      InitFromDataRow(pDataRow, True)
    End Sub

    Public Overridable Sub InitFromDataRow(ByVal pDataRow As DataRow, ByVal pUseProperName As Boolean)
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

    Protected Overridable Sub SetCaptionNames()

    End Sub


#End Region

    Public Shared Function GetBulkCopyDataTable(pRecords As IEnumerable(Of CARERecord)) As DataTable
      Dim vTable As New DataTable

      If pRecords IsNot Nothing AndAlso pRecords.Count > 0 Then
        Dim vFirstRecord As CARERecord = pRecords(0)
        vTable = pRecords(0).GetDataTable(Nothing, 1)


        Dim vFieldName As String = ""
        For Each vField As ClassField In vFirstRecord.mvClassFields
          vFieldName = vField.ProperName
          vTable.Columns(vFieldName).ColumnName = vField.Name
        Next

        vTable.Rows.Clear()

        'Populate the DataTable from the list of records
        pRecords.ToList().ForEach(Sub(vItem)
                                    vItem.SetValid()
                                    Dim vRow As DataRow = vTable.NewRow()
                                    Dim vField As String = ""
                                    For vIndex As Integer = 1 To vItem.mvClassFields.Count
                                      vField = vItem.mvClassFields.Item(vIndex).Name
                                      If vItem.mvClassFields.Item(vIndex).Value.Length = 0 Then
                                        vRow(vField) = DBNull.Value
                                      Else
                                        vRow(vField) = vItem.mvClassFields.Item(vIndex).Value
                                      End If
                                    Next
                                    vTable.Rows.Add(vRow)
                                  End Sub)

      End If
      Return vTable

    End Function

    ''' <summary>
    ''' Return a CDBFields class containing the field name and value of every field index that is passed
    ''' </summary>
    ''' <param name="pFieldIndexes">A list(Of Integer) of all field indexes that are to be included in the CDBFields clause</param>
    ''' <returns>This method is useful when a record needs to be queried that matches the values of a composite key stored in the current object</returns>
    Protected Function CreateWhere(pFieldIndexes As IEnumerable(Of Integer)) As CDBFields
      Dim vWhere As New CDBFields
      pFieldIndexes.ToList().ForEach(Sub(vIndex) vWhere.Add(ClassFields(vIndex).Name, ClassFields(vIndex).FieldType, ClassFields(vIndex).Value))
      Return vWhere
    End Function
    ''' <summary>
    ''' Returns a CDBParameters instance with the values of your class for the field indexes that you pass.
    ''' </summary>
    ''' <param name="pFieldIndexes">An IEnumerable of the Fields indexes used to access the ClassFields property</param>
    ''' <example>
    ''' Dim vGAClaimLine as DeclarationLinesUnclaimed = Me.GetRelatedInstance(Of DeclarationLinesUnclaimed)({BatchTransactionAnalysisFields.BatchNumber, BatchTransactionAnalysisFields.TransationNumber, BatchTransactionAnalysisFields.LineNumber}) 'Try to get the instance before you create it
    ''' If vGAClaimLine Is Nothing Then vGAClaimline = New DeclarationLinesUnclaimed(Me.Environment)
    ''' vGAClaimLine.Create(Me.CreateParams({BatchTransactionAnalysisFields.BatchNumber, BatchTransactionAnalysisFields.TransationNumber, BatchTransactionAnalysisFields.LineNumber}) 'Initialise the new class with the values from my class
    ''' </example>
    ''' <returns>an instance of the CDBParameters class with the values from your class</returns>
    Protected Function CreateParams(pFieldIndexes As IEnumerable(Of Integer)) As CDBParameters
      Dim vResult As New CDBParameters
      pFieldIndexes.ToList().ForEach(Sub(pIndex) vResult.Add(ClassFields(pIndex).ParameterName, ClassFields(pIndex).Value))
      Return vResult
    End Function
    ''' <summary>
    ''' Returns an instance of T with the values within the fields specified by the pRelatedFieldIndexes parameter
    ''' </summary>
    ''' <typeparam name="T">The CARERecord type that you want </typeparam>
    ''' <param name="pRelatedFieldIndexes">An IEnumerable of the Fields indexes used to access the ClassFields property</param>
    ''' <remarks>
    ''' This method will create a CDBFields instance with the fields and values specified in the pRelatedFieldIndexes parameter
    ''' It will then try to instantiate your class by calling InitWithPrimaryKey.  If the CARERecord.Existing returns True then
    ''' the return will be the instance.  If it is False then the return will be Nothing
    ''' </remarks>
    ''' <returns>The CARERecord that you have specified if it exists</returns>
    Public Function GetRelatedInstance(Of T As CARERecord)(pRelatedFieldIndexes As IEnumerable(Of Integer)) As T
      Dim vWhere As CDBFields = CreateWhere(pRelatedFieldIndexes)
      Return CARERecordFactory.Instance.GetInstance(Of T)(Me.Environment, vWhere)
    End Function
    ''' <summary>
    ''' Returns an IList of the CARERecord type you have specified for the values that match the CDBFields you have specified
    ''' </summary>
    ''' <typeparam name="T">The CARERecord Type that you want as a collection</typeparam>
    ''' <param name="pWhere">The CDBFields that you want the CARERecord instances to match</param>
    ''' <returns>an IList of the CARERecord Type that you want</returns>
    Public Function GetList(Of T As CARERecord)(pWhere As CDBFields) As IList(Of T)
      Dim vResult As IEnumerable(Of T) = CARERecordFactory.Instance.GetEnumerable(Of T)(Me.Environment, pWhere)
      Return If(vResult IsNot Nothing, vResult.ToList(), New List(Of T))
    End Function
    ''' <summary>
    ''' Returns a List of CARERecords with the values within the fields specified by the pRelatedFieldIndexes parameter
    ''' </summary>
    ''' <typeparam name="T">The CARERecord type that you want as a collection</typeparam>
    ''' <param name="pRelatedFieldIndexes">An IEnumerable of the Fields indexes used to access the ClassFields property</param>
    ''' <remarks>
    ''' This method will create a Where clause as a CDBFields instance with the fields and values specified in the pRelatedFieldIndexes parameter.
    ''' It will then try to instantiate your list by calling GetDataTable() and then running an InitFromDataRow for each row in the DataTable.
    ''' The List will only contain the records that match the values you have passed.
    ''' </remarks>
    ''' <returns>A List of the CARERecords that match the values in your class.  If no record matches the values then the list will be empty.</returns>
    Public Function GetRelatedList(Of T As CARERecord)(pRelatedFieldIndexes As IEnumerable(Of Integer)) As IList(Of T)
      Dim vWhere As CDBFields = CreateWhere(pRelatedFieldIndexes)
      Dim vResult As IEnumerable(Of T) = CARERecordFactory.Instance.GetEnumerable(Of T)(Me.Environment, vWhere)
      Return If(vResult IsNot Nothing, vResult.ToList(), New List(Of T))
    End Function

    Public Sub LoadFromRow(pRow As DataRow) Implements IDbLoadable.LoadFromRow
      Dim vUseProperNames As Boolean = False
      vUseProperNames = pRow IsNot Nothing AndAlso Not pRow.Table.Columns.Cast(Of DataColumn).Any(Function(pColumn) pColumn.ColumnName.Contains("_"))
      Me.InitFromDataRow(pRow, vUseProperNames)
    End Sub

    ''' <summary>
    ''' The Property accessor for the mvClassFields member variable.  Use this instead of accessing the member variable directly.
    ''' </summary>
    ''' <returns></returns>
    Protected Friend ReadOnly Property ClassFields As ClassFields
      Get
        CheckClassFields()
        Return mvClassFields
      End Get
    End Property

    Public ReadOnly Property FieldNames As String Implements IDbSelectable.DbFieldNames
      Get
        Return Me.ClassFields.FieldNames(Me.Environment, Me.TableAlias)
      End Get
    End Property

    Public ReadOnly Property AliasedTableName As String Implements IDbSelectable.DbAliasedTableName
      Get
        Return Me.ClassFields.TableNameAndAlias
      End Get
    End Property

    Public ReadOnly Property UniquePrimaryKey As ClassField
      Get
        Return Me.ClassFields.GetUniquePrimaryKey(True)
      End Get
    End Property

    Public ReadOnly Property IsDirty As Boolean
      Get
        Return Me.ClassFields.FieldsChanged()
      End Get
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns>Collection of MaintenanceRelation that have this CARERecod as their Primary Table in table Maintenance_Relations</returns>
    Public Property Relations As IEnumerable(Of MaintenanceRelation)
      Get
        If mvRelations Is Nothing Then
          Dim vQBERecord As New MaintenanceRelation(Me.Environment)
          vQBERecord.InitClassFields()
          vQBERecord.PrimaryTableName = Me.DatabaseTableName
          mvRelations = vQBERecord.GetRelatedList(Of MaintenanceRelation)({MaintenanceRelation.MaintenanceRelationFields.PrimaryTableName})
        End If
        Return mvRelations
      End Get
      Private Set(value As IEnumerable(Of MaintenanceRelation))

      End Set
    End Property
    ''' <summary>
    ''' Will check to see if this record has related records on other tables. Uses maintenence_relations table. 
    ''' If a record is found an exception is raised. Intended to stop the creation of orphans when deleting a CARERecord.
    ''' </summary>
    ''' <param name="pClassFields"> The ClassFields for this CARERecord usually mvClassFields</param>
    Public Sub CheckUsedElsewhere(pClassFields As ClassFields)
      Dim vAction As New Action(Of MaintenanceRelation)(Sub(pInstance) RaiseError(DataAccessErrors.daeRecordCannotBeDeleted, pInstance.RelatedTableDesc, pInstance.RelatedAttributeDesc))
      MaintenanceRelation.ValidateRelations(pClassFields, Me.Relations, vAction)
    End Sub

  End Class
End Namespace