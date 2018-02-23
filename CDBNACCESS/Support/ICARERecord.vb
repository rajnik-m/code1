Public Interface ICARERecord

  Sub InitFromMaintenanceData(ByVal pTableName As String)
  ReadOnly Property MaintenanceAttributes() As CollectionList(Of MaintenanceAttribute)
  Sub SetEnvironment(ByVal pEnv As CDBEnvironment)
  ReadOnly Property NeedsMaintenanceInfo() As Boolean
  Function KeyValueRequired(ByVal pField As String) As Boolean
  Function GetAddRecordMandatoryParameters() As String
  ReadOnly Property IsMaintenanceTable() As Boolean
  ReadOnly Property NoUniqueKey() As Boolean
  Sub SetControlNumber()
  Function GetRecordSetFields() As String
  Function GetRecordSetFieldsExclude(ByVal pExcludeFields As List(Of ClassField)) As String
  Sub Init()
  Sub Init(ByVal pPrimaryKeyValue As Integer)
  Sub Init(ByVal pPrimaryKeyValue As String)
  Sub InitForUpdate(ByVal pParams As CDBParameters)
  Sub Init(ByVal pParams As CDBParameters)
  Sub InitWithPrimaryKey(ByVal pWhereFields As CDBFields)
  Sub Clone(ByVal pRecord As CARERecord)   'Use when the primary key is a control number
  Sub Clone(ByVal pRecord As CARERecord, ByVal pPrimaryKeyValue As Integer)
  Sub Clone(ByVal pRecord As CARERecord, ByVal pPrimaryKeyValue As String)
  Sub Clone(ByVal pRecord As CARERecord, ByVal pParams As CDBParameters)
  Sub CopyValues(ByVal pRecord As CARERecord)
  Sub Create(ByVal pParameterList As CDBParameters, ByVal pPrimaryKeyValue As Integer)
  Sub Create(ByVal pParameterList As CDBParameters)
  Sub Update(ByVal pParameterList As CDBParameters)
  Sub InitFromRecordSet(ByVal pRecordSet As CDBRecordSet)
  Sub InitFromXMLNode(ByVal pNode As Xml.XmlNode)
  Sub Delete()
  Sub Delete(ByVal pAmendedBy As String)
  Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean)
  Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
  Sub DeleteByForeignKey(ByVal pField As CDBField)
  Sub DeleteByForeignKeys(ByVal pWhereFields As CDBFields)
  Sub AddDeleteCheckItems()
  Function RecordExists(ByVal pField As CDBField) As Boolean
  Function RecordExists(ByVal pWhereFields As CDBFields) As Boolean
  Sub Save()
  Sub Save(ByVal pAmendedBy As String)
  Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean)
  Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
  Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer, pForceAmendmentHistory As Boolean)
  ReadOnly Property Existing() As Boolean
  ReadOnly Property DataTable() As CDBDataTable
  ReadOnly Property DataTableColumnNames() As String
  Sub AddDataColumns(ByVal pDataTable As CDBDataTable)
  Sub AddDataColumns(ByVal pDataTable As CDBDataTable, ByVal pUseProperNames As Boolean)
  Sub AddDataRow(ByVal pDataTable As CDBDataTable)
  Sub AddDataRow(ByVal pDataTable As CDBDataTable, ByVal pUseProperNames As Boolean)
  Function FieldValueString(ByVal pFieldName As String) As String
  Function FieldValueInteger(ByVal pFieldName As String) As Integer
  Function FieldValueChanged(ByVal pFieldName As String) As Boolean
  Function Validate() As Boolean
  Function GetList(Of ItemType As CARERecord)(ByVal pItem As IRecordCreate, ByVal pWhereFields As CDBFields) As List(Of ItemType)
  Function GetDataTableFromList(ByVal pList As List(Of CARERecord)) As CDBDataTable
  Function GetUniqueKeyFields() As CDBFields
  Function GetUniqueKeyFields(ByVal pParams As CDBParameters) As CDBFields
  Function GetValuesAsFields() As CDBFields
  Function GetUpdateKeyFieldNames() As String
  Function GetUniqueKeyFieldNames() As String
  Function GetUniqueKeyParameters() As CDBParameters
  Function GetAddParameters() As CDBParameters
  Function GetUpdateParameters() As CDBParameters
  Sub SaveAmendedOnChanges()
  Property Environment As CDBEnvironment

End Interface