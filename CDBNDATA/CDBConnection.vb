Imports System.Data.Common
Imports System.IO

Namespace Data

  Partial Public MustInherit Class CDBConnection
    Implements IDisposable

#Region "Enums"

    Public Enum RDBMSTypes
      rdbmsUnknown
      rdbmsSqlServer
      rdbmsOracle
    End Enum

    Public Enum RowRestrictionTypes
      UseTopN
      UseRownum
    End Enum

    Protected Enum DataAccessModes
      damNormal = 0       'Normal operation
      damTest             'Disable all Inserts, Updates and Deletes
      damGenerateSQL      'Generate SQL for Inserts, Updates and Deletes
    End Enum

    <Flags()>
    Public Enum SQLLoggingModes
      None
      Insert = 1
      Update = 2
      Delete = 4
      [Select] = 8
      Configs = 16
      AllSql = 1 + 2 + 4 + 8
      All = 1 + 2 + 4 + 8 + 16
      Timed = 32
      Mail = 64
    End Enum

    Public Enum cdbExecuteConstants
      sqlShowError = 0                  'Show errors
      sqlIgnoreError                    'Ignore all errors 
      sqlIgnoreDuplicate                'Ignore error if duplicate value inserted
    End Enum

    Public Enum DatabaseHintTypes
      dhtFullTableScanOracle8
      dhtOptionForceOrder
      dhtUseHashOracle9
      dhtNoMergeOracle9
    End Enum

    Public Enum RecordSetOptions As Integer
      None                  'Default behaviour - can use a datatable does not need multiple result sets
      NoDataTable           'Used for retrieving bulk data - does not use a datatable - does not need multiple result sets
      MultipleResultSets    'Used by reporting etc. Switches to MARS if appropriate
    End Enum

#End Region

#Region "Class Variables"

    Protected Const MAX_LEN_SQL_IDENTIFIER As Integer = 30
    Protected Const SQL_LOG_FILENAME As String = "CDBSQL.SQL"
    Protected Const SQL_MAIL_FILENAME As String = "CDBMAIL.SQL"

    Protected mvConnection As DbConnection
    Protected mvRDBMSType As RDBMSTypes
    Protected mvConnectionOpen As Boolean
    Protected mvDataAccessMode As DataAccessModes
    Protected mvGeneratedSQL As String
    Protected mvSQLLoggingMode As SQLLoggingModes
    Protected mvInTransaction As Boolean
    Protected mvTransaction As DbTransaction
    Protected mvRecordSets As New List(Of CDBRecordSet)
    Protected mvSQLLogQueueName As String
    Protected mvLastInsertErrorIsDuplicate As Boolean
    Protected mvDebugSQLMessages As Boolean = True

    Private mvUserID As String
    Private mvLogMessage As String
    Private mvSW As Stopwatch
    Private mvTransactionSW As Stopwatch
    Private mvCurrentSQL As String
    Private mvUnicodeFields As SortedList

    Public Enum DatabaseOptions
      ForXML
      ForXML_XSINIL
    End Enum

#End Region

#Region "Virtual Methods"
    Public Event RowsCopied(ByVal pRowsCopied As Integer)


    Public MustOverride Sub OpenConnection(ByVal pConnect As String, ByVal pLogname As String, ByVal pPassWord As String, ByVal pNeedMultipleResultSets As Boolean, ByVal pOverrideConnectString As Boolean)
    Public MustOverride Function GetTableNames(Optional ByVal pSchemaName As String = "") As DataTable
    Public MustOverride Function TableExists(ByVal pTableName As String) As Boolean
    Public MustOverride Function AttributeExists(ByVal pTableName As String, ByVal pAttributeName As String) As Boolean
    Public MustOverride Function GetAttributeNames(ByVal pTableName As String) As DataTable
    Public MustOverride Function GetIdentityColumn(ByVal pTableName As String) As String
    Protected MustOverride Function GetRecordSet(ByVal pSQL As String, ByVal pTimeout As Integer, ByVal pOptions As RecordSetOptions) As CDBRecordSet
    Public MustOverride Function GetRecordSetAnsiJoins(ByVal pSQL As String, ByVal pTimeout As Integer, ByVal pOptions As RecordSetOptions) As CDBRecordSet
    Public MustOverride Function GetDataSet(ByVal pSQL As SQLStatement) As DataSet
    Public MustOverride Function ProcessAnsiJoins(ByVal pSQL As String) As String
    Public MustOverride Sub CreateView(ByVal pUserName As String, ByVal pViewName As String, ByVal pSQL As String)
    Public MustOverride Sub DropView(ByVal pViewName As String)
    Public MustOverride Function DBLTrim(ByVal pExpression As String) As String
    Public MustOverride Function DBRTrim(ByVal pExpression As String) As String
    Public MustOverride Function DBLeft(ByVal pExpression As String, ByVal pLength As String) As String
    Public MustOverride Function DBIndexOf(ByVal pSearchString As String, ByVal pExpression As String) As String
    Public MustOverride Function DBSubString(ByVal pExpression As String, ByVal pStart As String, ByVal pLength As String) As String
    Public MustOverride Function DBCollateString() As String
    Public MustOverride Function DBConcatString() As String
    Public MustOverride Function DBDate() As String
    Public MustOverride Function DBAge() As String
    Public MustOverride Function DBIsNull(ByVal pExpression As String, ByVal pReplacement As String) As String
    Public MustOverride Function DBLength(ByVal pExpression As String) As String
    Public MustOverride Function DBLPad(ByVal pExpression As String, ByVal pLength As Integer) As String
    Public MustOverride Function DBAddYears(ByVal pDateAttribute As String, ByVal pYearsAttribute As String) As String
    Public MustOverride Function DBAddMonths(ByVal pDateAttribute As String, ByVal pMonthsAttribute As String) As String
    Public MustOverride Function DBAddWeeks(ByVal pDateAttribute As String, ByVal pWeeksAttribute As String) As String
    Public MustOverride Function DBHint(ByVal pHintType As DatabaseHintTypes, ByVal pTableName As String, ByVal UseHints As Boolean) As String
    Public MustOverride Function DBMonthDiff(ByVal pEarlierDate As String, ByVal pLaterDate As String) As String
    Public MustOverride Function DBToNumber(ByVal pExpression As String) As String
    Public MustOverride Function DBToString(ByVal pExpression As String) As String
    Public MustOverride Function DBMaxToString(ByVal pExpression As String) As String
    ''' <summary>
    ''' Converts a datetime table attribute to a date-only expression.
    ''' </summary>
    ''' <remarks>
    ''' Only use this function if your database attribute is a datetime data type.  To convert a string expression to a date, use <see cref="DBToToDate(String)"/>
    ''' Most DBMSes will use the same SQL to represent DBDateTimeAttribToDate and DbToDate.  Not all though, I'll let you guess which one.
    ''' Clue: It's not SQL Server.
    ''' </remarks>
    ''' <param name="pAttributeName">The name of the datetime database column that you want to convert to a date-only expression</param>
    ''' <returns></returns>
    Public MustOverride Function DBDateTimeAttribToDate(ByVal pAttributeName As String) As String
    Public MustOverride Function DBToDate(ByVal pExpression As String) As String
    Protected MustOverride Function TableOwnerPrefix() As String
    Public MustOverride Function NativeDataType(ByVal pField As CDBField) As String
    Public MustOverride ReadOnly Property RowRestrictionType() As RowRestrictionTypes
    Public MustOverride Sub AppendDateTime(ByVal pSQL As StringBuilder, ByVal pValue As String)
    Public MustOverride Function BulkCopyTable(ByVal pSourceConnection As CDBConnection, ByVal pTableName As String) As Integer
    Public MustOverride Function BulkCopyData(ByVal pSourceConnection As CDBConnection, ByVal pDestinationTableName As String, ByVal pSQL As String) As Integer
    Public MustOverride Function BulkCopyData(ByVal pSourceConnection As CDBConnection, ByVal pDestinationTableName As String, pTable As DataTable) As Integer
    Public MustOverride Function SupportsOption(ByVal pOption As DatabaseOptions) As Boolean
    Friend MustOverride Function GetDBParameter(ByVal pFieldName As String, ByVal pFieldValue As String, ByRef pParamName As String) As DbParameter
    Friend MustOverride Function GetDBBulkParameter(ByVal pFieldName As String, ByVal pFieldValue As String, ByRef pParamName As String) As DbParameter
    Friend MustOverride Function GetDBParameterFromFile(ByVal pFieldName As String, ByVal pFilename As String, ByRef pParamName As String) As DbParameter
    Public MustOverride Function GetDBParameterFromByteArray(ByVal pFieldName As String, ByVal pValue As Byte(), ByRef pParamName As String) As DbParameter
    Public MustOverride Function IsCaseSensitive() As Boolean
    Public MustOverride Function SupportsNoLock() As Boolean
    Public MustOverride Function DBForceOrder() As String
    Public MustOverride Sub DropTable(ByVal pTableName As String)
    Public MustOverride Function DeleteAllRecords(ByVal pTable As String) As Integer
    Public MustOverride Function InsertRecord(ByVal pTableName As String, ByVal pFields As CDBFields, ByVal pIgnoreDuplicates As Boolean) As Boolean
    Public MustOverride Function UpdateRecords(ByVal pTableName As String, ByVal pUpdateFields As CDBFields, ByVal pWhereFields As CDBFields, Optional ByVal pErrorIfNoRecords As Boolean = True) As Integer
    Public MustOverride Sub ComputeTableStatistics(ByVal pTableName As String)
    Public MustOverride Function GetIndexNames(ByVal pTableName As String) As DataTable
    Public MustOverride Function GetIndexColumns(ByVal pTableName As String, ByVal pIndexName As String) As DataTable
    Friend MustOverride Function PreProcessMaxRows(ByVal pSQLStatement As SQLStatement) As SQLStatement
    Public MustOverride Function UseTableSpaces() As Boolean
    Public MustOverride Function UnicodePerformed(ByVal pTableName As String, ByVal pAttributeName As String) As Boolean
    Friend MustOverride Sub AddUnicodeValue(ByVal pSQL As StringBuilder, ByVal pValue As String)
    Public MustOverride Function IsUserDBA(ByVal pLogname As String, ByVal pTableName As String) As Boolean
    Public MustOverride Function DBYear(ByVal pDateString As String) As String
    Public MustOverride Function DBMonth(ByVal pDateString As String) As String
    Public MustOverride Function DBDay(ByVal pDateString As String) As String
    Public MustOverride Function DBReplaceLineFeedWithSpace(ByVal pFieldName As String) As String
    Public MustOverride Function GetBinaryDBParameter(ByVal pFieldName As String, ByVal pValue As Byte(), ByRef pParamName As String) As DbParameter
    Public MustOverride Function IsBaseTable(schemaName As String, tableName As String) As Boolean
    Public MustOverride Function GetPrimaryKeyColumns(ByVal pTableName As String) As DataTable

    ''' <summary>Bulk Update and Insert Data into a Table.</summary>
    ''' <param name="pSQLStatement">The SQLStatement used to populate the DataTable.  Used so that the same SQL is used for the original data and for the changed data.</param>
    ''' <param name="pDataTable">The data to be inserted or updated. This must have the RowState set correctly on each row.</param>
    ''' <remarks>This has been briefly tested in SQLServer only and will need thorough testing</remarks>
    Public MustOverride Sub BulkUpdate(ByVal pSQLStatement As SQLStatement, ByVal pDataTable As DataTable)

    ''' <summary>Add a column to the table that can contain GUID values.
    ''' </summary>
    ''' <param name="pTableName">Table requiring the new column</param>
    ''' <param name="pColumnName">Column to be added.  This must not already exist.</param>
    Public Sub AddGUIDColumnToTable(ByVal pTableName As String, ByVal pColumnName As String)
      Dim vSQL As String = String.Format("ALTER TABLE {0} ADD {1} {2}", pTableName, pColumnName, NativeDataType(CDBField.FieldTypes.cftGUID))
      Me.ExecuteSQL(vSQL)
    End Sub

    ''' <summary>Update the column with a GUID in every row.</summary>
    ''' <param name="pTableName">Table to be updated.</param>
    ''' <param name="pColumnName">Column to be populated with a GUID</param>
    Public MustOverride Sub PopulateGUIDColumn(ByVal pTableName As String, ByVal pColumnName As String)

#End Region

#Region "Constructors and Factory"

    Public Sub New(ByVal pSQLLoggingMode As SQLLoggingModes, ByVal pSQLLogQueueName As String)
      mvSQLLoggingMode = pSQLLoggingMode
      mvSQLLogQueueName = pSQLLogQueueName
    End Sub

    Public Shared Function GetCDBConnection(ByVal pType As CDBConnection.RDBMSTypes, ByVal pSQLLoggingMode As SQLLoggingModes, ByVal pSQLLogQueueName As String) As CDBConnection
      If pType = CDBConnection.RDBMSTypes.rdbmsSqlServer Then
        Return New CDBSQLServerConnection(pSQLLoggingMode, pSQLLogQueueName)
      Else
        Return New CDBOracleConnection(pSQLLoggingMode, pSQLLogQueueName)
      End If
    End Function

#End Region

#Region "Connection methods"

    Protected Function GetFullConnectString(ByVal pConnect As String, ByVal pDefaultLogname As String, ByVal pDefaultPassWord As String, ByVal pOverrideConnectString As Boolean) As String
      Dim vItems() As String = pConnect.Split(";"c)
      If pOverrideConnectString Then
        For vIndex As Integer = 0 To vItems.Length - 1
          If vItems(vIndex).ToLower.StartsWith("user id=") Then
            vItems(vIndex) = "user id=" & pDefaultLogname
          ElseIf vItems(vIndex).ToLower.StartsWith("password=") Then
            vItems(vIndex) = "password=" & pDefaultPassWord
          ElseIf vItems(vIndex).ToLower.StartsWith("integrated security=") Then
            vItems(vIndex) = String.Empty
          End If
        Next
        pConnect = String.Join(";", vItems)
      End If
      If pConnect.ToLower.Contains("user id=") Then
        For vIndex As Integer = 0 To vItems.Length - 1
          If vItems(vIndex).Trim.ToLower.StartsWith("user id") Then
            mvUserID = vItems(vIndex).Trim.Substring(7).TrimStart("= ".ToCharArray)
          End If
        Next
      Else
        pConnect = String.Concat(pConnect, ";user id=", pDefaultLogname)
        mvUserID = pDefaultLogname
      End If
      If Not pConnect.ToLower.Contains("password=") Then pConnect = String.Concat(pConnect, ";password=", pDefaultPassWord)
      Return pConnect
    End Function

    Public Sub CloseConnection()
      If mvConnectionOpen Then
        If (mvSQLLoggingMode And SQLLoggingModes.All) = SQLLoggingModes.All Then LogSQL("DISCONNECT")
        mvConnection.Close()
        mvConnection.Dispose()
        mvConnection = Nothing
        mvConnectionOpen = False
        Debug.Print("DISCONNECT")
      End If
    End Sub

    Public ReadOnly Property ConnectionOpen() As Boolean
      Get
        Return mvConnectionOpen
      End Get
    End Property

    Public ReadOnly Property UserID() As String
      Get
        Return mvUserID
      End Get
    End Property

#End Region

#Region "Count methods"

    Public Function GetCountFromStatement(ByVal pSQLStatement As SQLStatement) As Integer
      Dim vCount As Integer
      Dim vRecordSet As CDBRecordSet = GetRecordSet(pSQLStatement.CountSQL, pSQLStatement.Timeout, RecordSetOptions.None)
      If vRecordSet.Fetch Then
        vCount = IntegerValue(vRecordSet.Fields(1).Value)
      End If
      vRecordSet.CloseRecordSet()
      Return vCount
    End Function

    Public Function GetCount(ByVal pTable As String, ByVal pWhereFields As CDBFields) As Integer
      Dim vNoRecords As Boolean
      Dim vCount As Integer

      Dim vSQLStatement As New SQLStatement(Me, "count(*) AS record_count", pTable, pWhereFields)
      Dim vRecordSet As CDBRecordSet = vSQLStatement.GetRecordSet
      If vRecordSet.Fetch() Then
        vCount = vRecordSet.Fields(1).IntegerValue
      Else
        vNoRecords = True
      End If
      vRecordSet.CloseRecordSet()
      If vNoRecords Then RaiseError(DataAccessErrors.daeCountNoRecords, vSQLStatement.SQL)
      Return vCount
    End Function

#End Region

    Friend Sub RowsCopiedHandler(ByVal sender As Object, ByVal e As SqlClient.SqlRowsCopiedEventArgs)
      RaiseEvent RowsCopied(CInt(e.RowsCopied))
    End Sub

    Public Function GetSelectSQLCSC() As String
      Dim vAddCSC As Boolean
      If mvRDBMSType = RDBMSTypes.rdbmsSqlServer Then
        vAddCSC = True
        For Each vRecordSet As CDBRecordSet In mvRecordSets
          If vRecordSet.Active = True Then
            vAddCSC = False
            Exit For
          End If
        Next
      End If
      If vAddCSC Then
        Return "SELECT /* CDB.NET */ "
      Else
        Return "SELECT "
      End If
    End Function

    Public Function NullsSortAtEnd() As Boolean
      If mvRDBMSType = RDBMSTypes.rdbmsSqlServer Then
        Return False
      ElseIf mvRDBMSType = RDBMSTypes.rdbmsOracle Then
        Return True
      End If
    End Function

#Region "RecordSets"

    Public Function GetDataTable(ByVal pSQL As SQLStatement) As DataTable
      Dim vDataSet As DataSet = GetDataSet(pSQL)
      If vDataSet.Tables.Count > 0 Then
        Return vDataSet.Tables(0)
      Else
        Return Nothing
      End If
    End Function

    Public Function GetRecordSet(ByVal pSQL As String) As CDBRecordSet
      'SHOULD ONLY USED BY VB6 MIGRATEDCODE
      Return GetRecordSet(pSQL, 0, RecordSetOptions.None)
    End Function

    Public Function GetRecordSet(ByVal pSQL As String, ByVal pOptions As RecordSetOptions) As CDBRecordSet
      'SHOULD ONLY USED BY VB6 MIGRATEDCODE
      Return GetRecordSet(pSQL, 0, pOptions)
    End Function

    Public Function GetRecordSet(ByVal pSQLStatement As SQLStatement, ByVal pTimeout As Integer) As CDBRecordSet
      If pSQLStatement.MaxRows > 0 Then pSQLStatement = PreProcessMaxRows(pSQLStatement)
      Return GetRecordSet(pSQLStatement.SQL, pTimeout, RecordSetOptions.None)
    End Function

    Public Function GetRecordSet(ByVal pSQLStatement As SQLStatement, ByVal pTimeout As Integer, ByVal pOptions As RecordSetOptions) As CDBRecordSet
      If pSQLStatement.MaxRows > 0 Then pSQLStatement = PreProcessMaxRows(pSQLStatement)
      Return GetRecordSet(pSQLStatement.SQL, pTimeout, pOptions)
    End Function

    Friend ReadOnly Property RecordSets() As List(Of CDBRecordSet)
      Get
        Return mvRecordSets
      End Get
    End Property

    Friend Overridable Sub NotifyRecordSetClosed(ByVal pRecordSet As CDBRecordSet)
      LogRecordSetClosed()
    End Sub

#End Region

#Region "Insert Update Delete and Execute"
    Public ReadOnly Property LastGeneratedSQL() As String
      Get
        Return mvGeneratedSQL
      End Get
    End Property
    Public ReadOnly Property IsLastErrorDuplicate() As Boolean
      Get
        Return mvLastInsertErrorIsDuplicate
      End Get
    End Property

    Public Function InsertRecord(ByVal pTableName As String, ByVal pFields As CDBFields) As Boolean
      Dim vSQL As New StringBuilder
      Dim vGotAttr As Boolean
      Dim vDocName As String
      Dim vNeedData As Boolean
      Dim vIsFile As Boolean
      Dim vRows As Integer
      Dim vDBParameters As New CollectionList(Of DbParameter)

      mvLastInsertErrorIsDuplicate = False
      vSQL.Append("INSERT INTO ")
      vSQL.Append(pTableName)
      vSQL.Append(" (")
      For Each vField As CDBField In pFields
        If vGotAttr Then vSQL.Append(", ")
        If vField.SpecialColumn Then
          vSQL.Append(DBSpecialCol("", vField.Name))
        Else
          vSQL.Append(vField.Name)
        End If
        vGotAttr = True
      Next
      vGotAttr = False
      vSQL.Append(") VALUES (")
      For Each vField As CDBField In pFields
        If vGotAttr Then vSQL.Append(", ")
        vGotAttr = True
        If vField.Value.Length = 0 Then
          vSQL.Append("null")
        Else
          Select Case vField.FieldType
            Case CDBField.FieldTypes.cftDate
              AppendDate(vSQL, vField.Value)
            Case CDBField.FieldTypes.cftTime
              AppendDateTime(vSQL, vField.Value)
            Case CDBField.FieldTypes.cftCharacter
              vSQL.Append("'")
              vSQL.Append(vField.Value.Replace("'", "''"))
              vSQL.Append("'")
            Case CDBField.FieldTypes.cftMemo
              Dim vParamName As String = ""
              vDBParameters.Add(vField.Name, GetDBParameter(vField.Name, vField.Value, vParamName))
              If mvDataAccessMode = DataAccessModes.damGenerateSQL Then
                vSQL.Append("'")
                vSQL.Append(vField.Value.Replace("'", "''"))
                vSQL.Append("'")
              Else
                vSQL.Append(vParamName)
              End If
            Case CDBField.FieldTypes.cftNumeric
              vSQL.Append(vField.FixedValue)
            Case CDBField.FieldTypes.cftInteger, CDBField.FieldTypes.cftLong, CDBField.FieldTypes.cftBit
              vSQL.Append(vField.Value)
            Case CDBField.FieldTypes.cftBulk
              vDocName = vField.Value
              vSQL.Append("?")
              vNeedData = True
            Case CDBField.FieldTypes.cftFile
              vDocName = vField.Value
              If FileLen(vDocName) = 0 Then
                vSQL.Append("null")
              Else
                vSQL.Append("?")
                vIsFile = True
                vNeedData = True
              End If
            Case CDBField.FieldTypes.cftUnicode
              AddUnicodeValue(vSQL, vField.Value)
            Case CDBField.FieldTypes.cftBinary
              Dim vParamName As String = vField.DBParam.ParameterName
              vDBParameters.Add(vParamName, vField.DBParam)
              vSQL.Append(vParamName)
          End Select
        End If
      Next
      vSQL.Append(")")

      Select Case mvDataAccessMode
        Case DataAccessModes.damTest
          Debug.Print("TEST " & vSQL.ToString)
          vRows = 1 'Assume that it affected some rows although we cannot tell
        Case DataAccessModes.damGenerateSQL
          mvGeneratedSQL = vSQL.ToString
          vRows = 1 'Assume that it affected some rows although we cannot tell
        Case DataAccessModes.damNormal
          If mvDebugSQLMessages Then Debug.Print(vSQL.ToString)
          If (mvSQLLoggingMode And SQLLoggingModes.Insert) > 0 Then LogSQL(vSQL.ToString)
          If vNeedData = True Then
            'If vIsFile Then
            '  vReturn = CDBODBCExecuteWithDoc(mvODBCEnv, mvDBHandle, vSQL, vDocName, vRows)
            'Else
            '  If vCharData Then
            '    vReturn = CDBODBCExecuteWithChar(mvODBCEnv, mvDBHandle, vSQL, vDocName, vRows)
            '  Else
            '    vReturn = CDBODBCExecuteWithString(mvODBCEnv, mvDBHandle, vSQL, vDocName, vRows)
            '  End If
            'End If
            'If vReturn > 0 Then CheckError(mvDBHandle, vReturn, vSQL)
          Else
            vRows = ExecuteNonQuery(vSQL.ToString, 0, vDBParameters)
            If vRows < 1 Then RaiseError(DataAccessErrors.daeInsertFailed, pTableName, vSQL.ToString)
          End If
      End Select

    End Function

    Protected Function DoUpdateRecords(ByVal pTableName As String, ByVal pUpdateFields As CDBFields, ByVal pWhereFields As CDBFields, Optional ByVal pErrorIfNoRecords As Boolean = True) As Integer
      Dim vRecordChanged As Boolean
      Dim vSQL As New StringBuilder
      Dim vGotAttr As Boolean
      Dim vRows As Integer
      Dim vDbParameters As New CollectionList(Of DbParameter)

      For Each vField As CDBField In pUpdateFields
        If vField.Name <> "amended_on" Or vField.Name <> "amended_by" Then
          vRecordChanged = True
          Exit For
        End If
      Next
      If vRecordChanged Then
        vSQL.Append("UPDATE ")
        vSQL.Append(pTableName)
        vSQL.Append(" SET ")
        For Each vField As CDBField In pUpdateFields
          If vGotAttr Then
            vSQL.Append(", ")
          Else
            vGotAttr = True
          End If
          If vField.SpecialColumn Then
            vSQL.Append(DBSpecialCol("", vField.Name))
            vSQL.Append(" = ")
          Else
            vSQL.Append(vField.Name)
            vSQL.Append(" = ")
          End If
          If vField.Value.Length = 0 Then
            vSQL.Append("null")
          Else
            Select Case vField.FieldType
              Case CDBField.FieldTypes.cftDate
                AppendDate(vSQL, vField.Value)
              Case CDBField.FieldTypes.cftTime
                AppendDateTime(vSQL, vField.Value)
              Case CDBField.FieldTypes.cftCharacter
                vSQL.Append("'")
                vSQL.Append(vField.Value.Replace("'", "''"))
                vSQL.Append("'")
              Case CDBField.FieldTypes.cftMemo
                If mvDataAccessMode = DataAccessModes.damGenerateSQL Then
                  vSQL.Append("'")
                  vSQL.Append(vField.Value.Replace("'", "''"))
                  vSQL.Append("'")
                Else
                  Dim vParamName As String = ""
                  vDbParameters.Add(vField.Name, GetDBParameter(vField.Name, vField.Value, vParamName))
                  vSQL.Append(vParamName)
                End If
              Case CDBField.FieldTypes.cftNumeric
                vSQL.Append(vField.FixedValue)
              Case CDBField.FieldTypes.cftInteger, CDBField.FieldTypes.cftLong, CDBField.FieldTypes.cftBit
                vSQL.Append(vField.Value)
              Case CDBField.FieldTypes.cftBulk
                Dim vParamName As String = ""
                vDbParameters.Add(vField.Name, GetDBBulkParameter(vField.Name, vField.Value, vParamName))
                vSQL.Append(vParamName)
              Case CDBField.FieldTypes.cftFile
                Dim vParamName As String = ""
                vDbParameters.Add(vField.Name, GetDBParameterFromFile(vField.Name, vField.Value, vParamName))
                vSQL.Append(vParamName)
              Case CDBField.FieldTypes.cftUnicode
                AddUnicodeValue(vSQL, vField.Value)
              Case CDBField.FieldTypes.cftBinary
                Dim vParamName As String = vField.DBParam.ParameterName
                vDbParameters.Add(vParamName, vField.DBParam)
                vSQL.Append(vParamName)
            End Select
          End If
        Next
        Dim vWhere As String = WhereClause(pWhereFields)
        If vWhere.Length = 0 Then
          RaiseError(DataAccessErrors.daeCannotUpdate, pTableName)
        Else
          vSQL.Append(" WHERE ")
          vSQL.Append(vWhere)

          Select Case mvDataAccessMode
            Case DataAccessModes.damTest
              Debug.Print("TEST " & vSQL.ToString)
              vRows = 1 'Assume that it affected some rows although we cannot tell
            Case DataAccessModes.damGenerateSQL
              mvGeneratedSQL = vSQL.ToString
              vRows = 1 'Assume that it affected some rows although we cannot tell
            Case DataAccessModes.damNormal
              If mvDebugSQLMessages Then Debug.Print(vSQL.ToString)
              If (mvSQLLoggingMode And SQLLoggingModes.Update) = SQLLoggingModes.Update Then LogSQL(vSQL.ToString)
              vRows = ExecuteNonQuery(vSQL.ToString, 0, vDbParameters)
              If vRows < 1 And pErrorIfNoRecords Then RaiseError(DataAccessErrors.daeUpdateFailed, pTableName, vSQL.ToString)
          End Select
        End If
      End If
      Return vRows
    End Function

    Public Function DeleteRecords(ByVal pTableName As String, ByVal pWhereFields As CDBFields) As Integer
      Return DeleteRecords(pTableName, pWhereFields, True)
    End Function

    Public Function DeleteRecords(ByVal pTableName As String, ByVal pWhereFields As CDBFields, ByVal pErrorIfNoRecordsDeleted As Boolean) As Integer
      Dim vSQL As New StringBuilder
      Dim vRows As Integer

      vSQL.Append("DELETE FROM ")
      vSQL.Append(pTableName)
      vSQL.Append(" WHERE ")
      Dim vWhere As String = WhereClause(pWhereFields)
      If vWhere.Length = 0 Then
        RaiseError(DataAccessErrors.daeCannotDelete, pTableName)
      Else
        vSQL.Append(vWhere)
        Select Case mvDataAccessMode
          Case DataAccessModes.damTest
            Debug.Print("TEST " & vSQL.ToString)
            vRows = 1 'Assume that it affected some rows although we cannot tell
          Case DataAccessModes.damGenerateSQL
            mvGeneratedSQL = vSQL.ToString
            vRows = 1 'Assume that it affected some rows although we cannot tell
          Case DataAccessModes.damNormal
            If mvDebugSQLMessages Then Debug.Print(vSQL.ToString)
            If (mvSQLLoggingMode And SQLLoggingModes.Delete) = SQLLoggingModes.Delete Then LogSQL(vSQL.ToString)
            vRows = ExecuteNonQuery(vSQL.ToString, 0)
            If vRows < 1 And pErrorIfNoRecordsDeleted Then RaiseError(DataAccessErrors.daeUpdateFailed, pTableName, vSQL.ToString)
        End Select
      End If
      Return vRows
    End Function

    Public Function ExecuteSQL(ByVal pSQL As String, Optional ByVal pFlags As cdbExecuteConstants = cdbExecuteConstants.sqlShowError) As Integer
      Try
        Select Case mvDataAccessMode
          Case DataAccessModes.damTest
            Debug.Print("TEST " & pSQL)
            Return 1 'Assume that it affected some rows although we cannot tell
          Case DataAccessModes.damNormal
            If mvDebugSQLMessages Then Debug.Print(pSQL)
            If (mvSQLLoggingMode And SQLLoggingModes.AllSql) = SQLLoggingModes.AllSql Then LogSQL(pSQL)
            Return ExecuteNonQuery(pSQL.ToString, 0)
        End Select
      Catch vEx As Exception
        If pFlags = cdbExecuteConstants.sqlIgnoreError Then
          Return 0
        Else
          Throw vEx
        End If
      End Try
    End Function

    Private Function ExecuteNonQuery(ByVal pSQL As String, ByVal pTimeout As Integer) As Integer
      Return ExecuteNonQuery(pSQL, pTimeout, Nothing)
    End Function

    Private Function ExecuteNonQuery(ByVal pSQL As String, ByVal pTimeout As Integer, ByVal pDBParameters As CollectionList(Of DbParameter)) As Integer
      Using vCommand As DbCommand = mvConnection.CreateCommand
        vCommand.CommandText = pSQL
        vCommand.CommandType = CommandType.Text
        vCommand.CommandTimeout = pTimeout
        If pDBParameters IsNot Nothing AndAlso pDBParameters.Count > 0 Then
          For Each vDBParam As DbParameter In pDBParameters
            vCommand.Parameters.Add(vDBParam)
          Next
        End If
        vCommand.Connection = mvConnection
        If mvInTransaction Then vCommand.Transaction = mvTransaction
        Try
          Return vCommand.ExecuteNonQuery()
        Catch vEx As SqlClient.SqlException
          If vEx.Number >= 50000 Then
            Throw New CareException(vEx.Errors(0).Message, DataAccessErrors.daeODBCUserDefinedError, vEx.Source)
          Else
            Throw
          End If
        End Try
      End Using
    End Function

    Public Function ExecuteReader(ByVal pSQL As String) As IDataReader
      Return ExecuteReader(pSQL, 0, Nothing)
    End Function

    Private Function ExecuteReader(ByVal pSQL As String, ByVal pTimeout As Integer) As IDataReader
      Return ExecuteReader(pSQL, pTimeout, Nothing)
    End Function

    Private Function ExecuteReader(ByVal pSQL As String, ByVal pTimeout As Integer, ByVal pDBParameters As CollectionList(Of DbParameter)) As IDataReader
      Using vCommand As DbCommand = mvConnection.CreateCommand
        vCommand.CommandText = pSQL
        vCommand.CommandType = CommandType.Text
        vCommand.CommandTimeout = pTimeout
        If pDBParameters IsNot Nothing AndAlso pDBParameters.Count > 0 Then
          For Each vDBParam As DbParameter In pDBParameters
            vCommand.Parameters.Add(vDBParam)
          Next
        End If
        vCommand.Connection = mvConnection
        If mvInTransaction Then vCommand.Transaction = mvTransaction
        Try
          Return vCommand.ExecuteReader()
        Catch vEx As SqlClient.SqlException
          If vEx.Number >= 50000 Then
            Throw New CareException(vEx.Errors(0).Message, DataAccessErrors.daeODBCUserDefinedError, vEx.Source)
          Else
            Throw
          End If
        End Try
      End Using
    End Function

    Public Function CreateCommand() As DbCommand
      Dim vCommand As DbCommand = mvConnection.CreateCommand
      If mvInTransaction Then
        vCommand.Transaction = Me.Transaction
      End If
      Return vCommand
    End Function
#End Region

#Region "Table and Index Schema methods"

    Friend Function GetDataReader(ByVal pSQL As String) As DbDataReader
      Using vCommand As DbCommand = mvConnection.CreateCommand
        vCommand.CommandText = pSQL
        vCommand.CommandType = CommandType.Text
        vCommand.Connection = mvConnection
        Return vCommand.ExecuteReader()
      End Using
    End Function

    Public Sub AddColumnFromField(ByVal pTableName As String, ByVal vField As CDBField)
      Dim vSQL As New StringBuilder
      vSQL.Append("ALTER TABLE ")
      vSQL.Append(TableOwnerPrefix)
      vSQL.Append(pTableName)
      vSQL.Append(" ADD ")
      vSQL.Append(vField.Name)
      vSQL.Append(" ")
      vSQL.Append(NativeDataType(vField))
      If vField.Mandatory Then vSQL.Append(" NOT NULL")
      ExecuteSQL(vSQL.ToString)
    End Sub

    Public Sub DropColumn(ByVal pTableName As String, ByVal vColumnName As String)
      Dim vSQL As New StringBuilder
      vSQL.Append("ALTER TABLE ")
      vSQL.Append(TableOwnerPrefix)
      vSQL.Append(pTableName)
      vSQL.Append(" DROP COLUMN ")
      vSQL.Append(vColumnName)
      ExecuteSQL(vSQL.ToString)
    End Sub

    Public Function GetTriggers() As DataTable
      Dim vSQL As String = "SELECT /* CDB.NET */ trigger_name = name, trigger_owner = USER_NAME(uid), table_name = OBJECT_NAME(parent_obj), " &
                           "isupdate = OBJECTPROPERTY( id, 'ExecIsUpdateTrigger'), isdelete = OBJECTPROPERTY( id, 'ExecIsDeleteTrigger'), " &
                           "isinsert = OBJECTPROPERTY( id, 'ExecIsInsertTrigger'), isafter = OBJECTPROPERTY( id, 'ExecIsAfterTrigger'), " &
                           "isinsteadof = OBJECTPROPERTY( id, 'ExecIsInsteadOfTrigger'), " &
                           "status = CASE OBJECTPROPERTY(id, 'ExecIsTriggerDisabled') WHEN 1 THEN 'Disabled' ELSE 'Enabled' END " &
                           "FROM sysobjects WHERE type = 'TR' ORDER BY table_name"
      Dim vSQLStatement As New SQLStatement(Me, vSQL)
      Dim vDataSet As DataSet = Me.GetDataSet(vSQLStatement)
      If vDataSet.Tables.Contains("Table") Then
        Return vDataSet.Tables("Table")
      Else
        Return Nothing
      End If
    End Function

    Public Sub CreateTrigger(ByVal pTableName As String, ByVal pTriggerName As String, ByVal pType As TriggerTypes, ByVal pSQL As String)
      Dim vSQL As New StringBuilder

      With vSQL
        .Append("CREATE TRIGGER ")
        .Append(pTriggerName)
        .Append(" ON ")
        .Append(pTableName)
        .Append(" FOR ")
        Select Case pType
          Case TriggerTypes.Insert
            .Append("INSERT")
          Case TriggerTypes.Update
            .Append("UPDATE")
          Case TriggerTypes.Delete
            .Append("DELETE")
        End Select
        .Append(" AS ")
        .Append(pSQL)
        ExecuteSQL(.ToString)
      End With
    End Sub

    Public Sub DropTrigger(ByVal pTriggerName As String)
      ExecuteSQL(String.Format("DROP TRIGGER {0}", pTriggerName))
    End Sub

    Public Overridable Function IndexExists(ByVal pTableName As String, pAttributes As IList(Of String)) As Boolean
      Dim vFound As Boolean = False
      Dim vHashName As String = GetIndexName(pTableName, pAttributes)
      Dim vTable As DataTable = GetIndexNames(pTableName)
      For Each vRow As DataRow In vTable.Rows
        If vRow("INDEX_NAME").ToString.Equals(vHashName, StringComparison.InvariantCultureIgnoreCase) Then
          vFound = True
        End If
      Next
      Return vFound
    End Function

    Public MustOverride Function IndexIsUnique(ByVal pTableName As String, pAttributes As IList(Of String)) As Boolean

    Public Sub AddAttribute(ByVal pTableName As String, ByVal pField As CDBField)
      Dim vSQL As New StringBuilder
      vSQL.Append("ALTER TABLE ")
      vSQL.Append(TableOwnerPrefix)
      vSQL.Append(pTableName)
      vSQL.Append(" ADD ")
      vSQL.Append(pField.Name)
      vSQL.Append(" ")
      vSQL.Append(NativeDataType(pField))
      If pField.Mandatory Then vSQL.Append(" NOT NULL")
      ExecuteSQL(vSQL.ToString)
    End Sub

    Public Sub CreateTableFromFields(ByVal pTableName As String, ByVal pFields As CDBFields)
      'Creates a table from a fields collection.
      'Attributes will be created as non-mandatory with the datatype parameter(s) being held in vField.Value, i.e. the length of a character attribute
      'For a decimal attribute vField.Value will hold the total number of digits including those after the decimal place; the number of digits after the decimal will always be two
      If TableExists(pTableName) Then ExecuteSQL("DROP TABLE " & pTableName)
      Dim vSQL As New StringBuilder
      vSQL.Append("CREATE TABLE ")
      vSQL.Append(TableOwnerPrefix)
      vSQL.Append(pTableName)
      vSQL.Append("(")
      Dim vNeedSeparator As Boolean = False
      For Each vField As CDBField In pFields
        If vNeedSeparator Then vSQL.Append(", ")
        vSQL.Append(vField.Name)
        vSQL.Append(" ")
        vSQL.Append(NativeDataType(vField))
        If vField.Mandatory Then vSQL.Append(" NOT NULL")
        vNeedSeparator = True
      Next
      vSQL.Append(")")
      ExecuteSQL(vSQL.ToString)
    End Sub

    Public Enum TriggerTypes As Integer
      Insert
      Update
      Delete
    End Enum

    Public Function GetTriggerName(ByVal pTableName As String, ByVal pPrefix As String, ByVal pType As TriggerTypes) As String
      Dim vHashName As New StringBuilder
      With vHashName
        .Append(pPrefix)
        .Append("_")
        Select Case pType
          Case TriggerTypes.Insert
            .Append("insert")
          Case TriggerTypes.Update
            .Append("update")
          Case TriggerTypes.Delete
            .Append("delete")
        End Select
        .Append("_")
        .Append(GetHashName(pTableName))
        Return vHashName.ToString
      End With
    End Function

    Public Sub CreateIndex(ByVal pUnique As Boolean, ByVal pTable As String, pAttributes As IList(Of String))
      For Each vStatement As String In CreateIndexSql(pUnique, pTable, pAttributes)
        ExecuteSQL(vStatement)
      Next vStatement
    End Sub

    Public MustOverride Function CreateIndexSql(ByVal pUnique As Boolean, ByVal pTable As String, pAttributes As IList(Of String)) As IList(Of String)

    Public Sub DropIndex(ByVal pTable As String, pAttributes As IList(Of String))
      Dim vHashName As String = GetIndexName(pTable, pAttributes)
      DropIndexByName(pTable, vHashName)
    End Sub

    Public MustOverride Function DropIndexSql(ByVal pTable As String, ByVal pName As String) As IList(Of String)

    Public Sub DropIndexByName(ByVal pTable As String, ByVal pName As String)
      For Each vStatement As String In DropIndexSql(pTable, pName)
        ExecuteSQL(vStatement, cdbExecuteConstants.sqlIgnoreError)
      Next vStatement
    End Sub

    Public Function GetIndexName(ByVal pTable As String, pAttributes As IList(Of String)) As String
      If String.IsNullOrWhiteSpace(pTable) Then
        Throw New ArgumentException("Table name cannot be blank", "pTable")
      End If
      If pAttributes.Count < 1 Then
        Throw New ArgumentException("At leaszt one attribute must be specified", "pAttributes")
      End If
      Dim vHashName As New StringBuilder(GetHashName(pTable))
      For Each vAttribute As String In pAttributes
        If String.IsNullOrWhiteSpace(vAttribute) Then
          Throw New ArgumentException("All attributes passed must have names", "pAttributes")
        End If
        vHashName.Append(GetHashName(vAttribute))
      Next vAttribute
      Return If(vHashName.Length > 30, vHashName.ToString.Substring(0, 30), vHashName.ToString)
    End Function

    Private Function GetHashName(ByVal pString As String) As String
      Dim vHashName As New StringBuilder
      Dim vHashCode As Integer
      Dim vWordIndex As Integer
      Dim vChar As String
      For vIndex As Integer = 1 To pString.Length
        vChar = pString(vIndex - 1)
        If vWordIndex = 0 Then vHashName.Append(vChar)
        If vChar = "_" Then
          vWordIndex = 0
        Else
          vWordIndex += 1
          If vWordIndex > 1 Then
            If vWordIndex < 5 Then vHashCode = vHashCode * vWordIndex
            vHashCode = vHashCode + Asc(vChar)
          End If
        End If
      Next
      Dim vHashDigits As String = Left$(Format$(vHashCode), 2) & Right$(Format$(vHashCode), 2)
      Dim vPos As Integer = InStr(pString, "smcam_smapp_")
      If vPos > 0 Then
        vHashDigits = Format$(Val(Mid$(pString, vPos + 12)))
        If Len(vHashDigits) < 4 Then vHashDigits = Right$("0000" & vHashDigits, 4)
      End If
      Return vHashName.Append(vHashDigits).ToString
    End Function

#End Region

#Region "Where clause and Database Specific SQL"

    Public Overridable Function IsSpecialColumn(ByVal pName As String) As Boolean
      'Adding a new attributes to the following should also be added in CDBNETCL.DBSpecialCol
      Select Case pName
        Case "case", "current", "distributed", "expression", "external", "function",
             "module", "number", "permanent", "prefix", "primary", "priority", "reference", "when"
          Return True
        Case Else
          Return False
      End Select
    End Function

    Private Sub AppendDate(ByVal pSQL As StringBuilder, ByVal pValue As String)
      Dim vDate As Date = CDate(pValue)
      With (pSQL)
        .Append("'")
        .Append(vDate.Day)
        .Append(" ")
        .Append(GetMonthAbbreviation(vDate))
        .Append(" ")
        .Append(vDate.Year.ToString("D4")) 'returns the date as a 4 digit number with padded zeroes.  Useful when the date is DateTime.MinValue (i.e. 1 Jan 0001)
        .Append("'")
      End With
    End Sub

    Protected Function GetMonthAbbreviation(ByVal pDate As Date) As String
      'Should be english
      Dim vMonths() As String = {"jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"}
      Return vMonths(pDate.Month - 1)
    End Function

    Public Function SQLLiteral(ByVal pOperator As String, ByVal pValue As Date) As String
      Return SQLLiteral(pOperator, CDBField.FieldTypes.cftDate, pValue.ToString(CAREDateFormat))
    End Function

    Public Function SQLLiteral(ByVal pOperator As String, ByVal pFieldType As CDBField.FieldTypes, ByVal pValue As String) As String
      'Given a value return the string required to use as a literal in a SQL statement
      'Add a leading space but no trailing space
      Dim vPos As Integer
      Dim vPos1 As Integer

      If pValue.Length = 0 Then
        Select Case pOperator
          Case "=", "<="
            Return " IS NULL"
          Case "<>", ">="
            Return " IS NOT NULL"
          Case Else
            Return ""
            'Debug.Assert False        'Happens if just getting a literal to substitute (Reports)
        End Select
      Else
        Dim vSQL As New StringBuilder
        If pOperator.Length > 0 Then
          vSQL.Append(" ")
          vSQL.Append(pOperator)
        End If
        Select Case pFieldType
          Case CDBField.FieldTypes.cftCharacter, CDBField.FieldTypes.cftMemo, CDBField.FieldTypes.cftUnicode
            vPos = InStr(pValue, "'")
            If vPos > 0 Then
              vPos1 = InStr(pValue, "','")
              If vPos1 > 0 Then vPos = 0
            End If
            If vPos > 0 Then
              vSQL.Append(" '")
              vSQL.Append(pValue.Replace("'", "''"))
              vSQL.Append("'")
            Else
              If pFieldType = CDBField.FieldTypes.cftUnicode AndAlso RDBMSType = RDBMSTypes.rdbmsSqlServer Then
                vSQL.Append(" N'")
              Else
                vSQL.Append(" '")
              End If
              vSQL.Append(pValue)
              vSQL.Append("'")
            End If
          Case CDBField.FieldTypes.cftDate
            If pValue = "today" Then pValue = TodaysDate()
            AppendDate(vSQL, pValue)
          Case CDBField.FieldTypes.cftTime
            AppendDateTime(vSQL, pValue)
          Case Else
            vSQL.Append(" ")
            vSQL.Append(pValue)
        End Select
        Return vSQL.ToString
      End If
    End Function

    Public Function DBAttrName(ByVal pAttrName As String) As String
      If pAttrName.Length > MAX_LEN_SQL_IDENTIFIER Then
        Return pAttrName.Substring(0, MAX_LEN_SQL_IDENTIFIER)
      Else
        Return pAttrName
      End If
    End Function

    Public Function DBLike(ByVal pValue As String, Optional ByVal pFieldType As CDBField.FieldTypes = CDBField.FieldTypes.cftCharacter) As String
      'This function returns a string consisting of the SQL
      'keyword LIKE followed by the appropriate string
      'with the wildcards replaced as expected by the
      'specific back-end we are dealing with
      'The string passed to this function is expected to
      'contain the JET type syntax for like ie. * and ? for the wildcard characters
      'NOTE DBLike should be a case insensitive comparison
      Dim vUnicodePrefix As String = ""
      If pFieldType = CDBField.FieldTypes.cftUnicode AndAlso RDBMSType = RDBMSTypes.rdbmsSqlServer Then vUnicodePrefix = "N"
      Dim vLikeValue As String = pValue.Replace("'", "''").Replace("*", "%").Replace("?", "_")
      Return " LIKE " & vUnicodePrefix & "'" & vLikeValue & "'"
    End Function

    Public Function DBLikeOrEqual(ByVal pValue As String, Optional ByVal pFieldType As CDBField.FieldTypes = CDBField.FieldTypes.cftCharacter) As String
      Dim vUnicodePrefix As String = ""
      If pFieldType = CDBField.FieldTypes.cftUnicode AndAlso RDBMSType = RDBMSTypes.rdbmsSqlServer Then vUnicodePrefix = "N"
      If pValue.Contains("*") OrElse pValue.Contains("?") OrElse pValue.Contains("%") Then
        Return DBLike(pValue, pFieldType)
      Else
        Return " = " & vUnicodePrefix & "'" & pValue.Replace("'", "''") & "'"
      End If
    End Function

    Public Function DBSortByNullsFirst() As String
      If NullsSortAtEnd() Then
        Return " DESC "
      Else
        Return ""
      End If
    End Function

    Public Function DBOrderByNullsFirstDesc(ByRef pAttributeName As String) As String
      Return DBOrderByDate(pAttributeName, True, True)
    End Function

    Public Function DBOrderByNullsFirstAsc(ByRef pAttributeName As String) As String
      Return DBOrderByDate(pAttributeName, True, False)
    End Function

    Public Function DBOrderByNullsLastDesc(ByRef pAttributeName As String) As String
      Return DBOrderByDate(pAttributeName, False, True)
    End Function

    Public Function DBOrderByNullsLastAsc(ByRef pAttributeName As String) As String
      Return DBOrderByDate(pAttributeName, False, False)
    End Function

    Private Function DBOrderByDate(ByRef pAttributeName As String, ByVal pNullsFirst As Boolean, ByVal pDescending As Boolean) As String
      Dim vBuilder As New StringBuilder
      AppendDate(vBuilder, New Date(If(pNullsFirst Xor pDescending, 1900, 2999), 1, 1).ToString(CAREDateFormat))
      Return DBIsNull(pAttributeName, vBuilder.ToString) & If(pDescending, " DESC", "")
    End Function

    Public Function DBSpecialCol(ByVal pAttribute As String) As String
      Return DBSpecialCol(Nothing, pAttribute)
    End Function

    Public Function DBSpecialCol(ByVal pTable As String, ByVal pAttribute As String) As String
      'Given a column and table name return the string required to reference it in a database
      Dim vIgnore As Boolean
      Dim vName As New StringBuilder

      If pTable Is Nothing AndAlso pAttribute.Contains(".") Then
        pTable = pAttribute.Split("."c)(0)
        pAttribute = pAttribute.Split("."c)(1)
      End If
      If Not IsSpecialColumn(pAttribute) Then
        If pAttribute.IndexOf(" ") >= 0 Then
          'If the attribute has a space in it then it is being used in a non core table so special it
        Else
          If pAttribute.Length <= MAX_LEN_SQL_IDENTIFIER AndAlso mvDebugSQLMessages = True Then Debug.Print("Unknown special column " & pAttribute)
          vIgnore = True
        End If
      End If
      If pTable IsNot Nothing AndAlso pTable.Length > 0 Then
        vName.Append(pTable)
        vName.Append(".")
      End If
      If vIgnore Then
        vName.Append(DBAttrName(pAttribute))
      Else
        vName.Append("""")
        vName.Append(pAttribute)
        vName.Append("""")
      End If
      Return vName.ToString
    End Function

    Public Function WhereClause(ByVal pFields As CDBFields) As String
      Dim vWhere As New StringBuilder
      Dim vAnd As Boolean
      ' TA 15/5 Requirement for CDBField.FieldWhereOperators.fwoNullOrValue removed but left code in case needed in future
      ' But it will need testing as code was not used.

      For Each vField As CDBField In pFields
        Dim vName As String = If(vField.FieldType = CDBField.FieldTypes.cftMemo AndAlso Me.RDBMSType = RDBMSTypes.rdbmsOracle,
                                 String.Format("to_char({0})", vField.Name),
                                 vField.Name)
        Dim vPos As Integer = (vName.IndexOf("#", 0) + 1)
        If vPos > 0 Then vName = vName.Substring(0, vPos - 1)
        Dim vWhereOperator As CDBField.FieldWhereOperators = (vField.WhereOperator And CDBField.FieldWhereOperators.fwoOperatorOnly) 'Get just the operator part
        If vWhereOperator = CDBField.FieldWhereOperators.fwoBetweenTo Then vAnd = False
        If vAnd Then
          If (vField.WhereOperator And CDBField.FieldWhereOperators.fwoOR) > 0 Then
            vWhere.Append(" OR ")
          Else
            vWhere.Append(" AND ")
          End If
        Else
          vAnd = True
        End If
        If (vField.WhereOperator And CDBField.FieldWhereOperators.fwoNOT) = CDBField.FieldWhereOperators.fwoNOT Then vWhere.Append(" NOT ")
        If (vField.WhereOperator And CDBField.FieldWhereOperators.fwoOpenBracketTwice) = CDBField.FieldWhereOperators.fwoOpenBracketTwice Then
          vWhere.Append(" (( ")
        ElseIf (vField.WhereOperator And CDBField.FieldWhereOperators.fwoOpenBracket) = CDBField.FieldWhereOperators.fwoOpenBracket Then
          vWhere.Append(" ( ")
        End If
        Dim vOperator As String
        Select Case vWhereOperator
          Case CDBField.FieldWhereOperators.fwoEqual
            vOperator = "="
          Case CDBField.FieldWhereOperators.fwoGreaterThan
            vOperator = ">"
          Case CDBField.FieldWhereOperators.fwoNullOrGreaterThan
            vOperator = ">"
          Case CDBField.FieldWhereOperators.fwoNullOrEqual
            vOperator = "="
          Case CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual
            vOperator = ">="
          Case CDBField.FieldWhereOperators.fwoGreaterThanEqual
            vOperator = ">="
          Case CDBField.FieldWhereOperators.fwoLessThan
            vOperator = "<"
          Case CDBField.FieldWhereOperators.fwoNullOrLessThan
            vOperator = "<"
          Case CDBField.FieldWhereOperators.fwoLessThanEqual
            vOperator = "<="
          Case CDBField.FieldWhereOperators.fwoNullOrLessThanEqual
            vOperator = "<="
          Case CDBField.FieldWhereOperators.fwoNotEqual, CDBField.FieldWhereOperators.fwoNullOrNotEqual
            vOperator = "<>"
          Case CDBField.FieldWhereOperators.fwoNotLike
            vOperator = "LIKE"
            '    Case CDBField.FieldWhereOperators.fwoNullOrValue
            '      vOperator = "="
          Case CDBField.FieldWhereOperators.fwoIn
            vOperator = ""
          Case CDBField.FieldWhereOperators.fwoNotIn
            vOperator = ""
          Case CDBField.FieldWhereOperators.fwoInOrEqual
            If (vField.Value.IndexOf(",", 0) + 1) > 0 Then
              vOperator = ""
            Else
              vOperator = "="
            End If
          Case CDBField.FieldWhereOperators.fwoBetweenFrom
            vOperator = "BETWEEN"
          Case CDBField.FieldWhereOperators.fwoBetweenTo
            vOperator = "AND"
          Case CDBField.FieldWhereOperators.fwoLike
            vOperator = "LIKE"
          Case CDBField.FieldWhereOperators.fwoLikeOrEqual
            vOperator = "LIKE"
          Case CDBField.FieldWhereOperators.fwoExist
            vOperator = "EXISTS "
          Case Else
            vOperator = ""
        End Select
        If vOperator.Length > 0 Then
          Select Case vWhereOperator
            Case CDBField.FieldWhereOperators.fwoExist
              vWhere.Append(vOperator)
              vWhere.Append("( ")
              vWhere.Append(vField.Value)
              vWhere.Append(") ")
            Case CDBField.FieldWhereOperators.fwoBetweenTo
              vWhere.Append(SQLLiteral(vOperator, vField.FieldType, vField.Value))
            Case CDBField.FieldWhereOperators.fwoLikeOrEqual
              vWhere.Append(vName)
              If IsUnicodeField(vName) Then
                vWhere.Append(DBLikeOrEqual(vField.Value, CDBField.FieldTypes.cftUnicode))
              Else
                vWhere.Append(DBLikeOrEqual(vField.Value))
              End If
            Case CDBField.FieldWhereOperators.fwoLike
              If vField.SpecialColumn Then
                vWhere.Append(DBSpecialCol(pFields.TableAlias, vName))
              Else
                vWhere.Append(vName)
              End If
              If IsUnicodeField(vName) Then
                vWhere.Append(DBLike(vField.Value, CDBField.FieldTypes.cftUnicode))
              Else
                vWhere.Append(DBLike(vField.Value))
              End If
              '        vWhere = vWhere & vName & DBLike(vField.Value)
            Case CDBField.FieldWhereOperators.fwoNotLike
              vWhere.Append(vName)
              vWhere.Append(" NOT")
              If IsUnicodeField(vName) Then
                vWhere.Append(DBLike(vField.Value, CDBField.FieldTypes.cftUnicode))
              Else
                vWhere.Append(DBLike(vField.Value))
              End If
              '      Case CDBField.FieldWhereOperators.fwoNullOrValue
              '        vWhere = vWhere & "(" & vName & SQLLiteral(vOperator, vField.FieldType, gvNull)
              '        vWhere = vWhere & " OR " & vName & SQLLiteral(vOperator, vField.FieldType, vField.Value) & ")"
            Case CDBField.FieldWhereOperators.fwoNullOrEqual, CDBField.FieldWhereOperators.fwoNullOrGreaterThan, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual,
            CDBField.FieldWhereOperators.fwoNullOrLessThan, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual, CDBField.FieldWhereOperators.fwoNullOrNotEqual
              vWhere.Append("(")
              If vField.SpecialColumn Then vWhere.Append(DBSpecialCol(pFields.TableAlias, vName)) Else vWhere.Append(vName)
              vWhere.Append(SQLLiteral("=", vField.FieldType, ""))
              vWhere.Append(" OR ")
              If vField.SpecialColumn Then vWhere.Append(DBSpecialCol(pFields.TableAlias, vName)) Else vWhere.Append(vName)
              vWhere.Append(SQLLiteral(vOperator, vField.FieldType, vField.Value))
              vWhere.Append(")")
            Case Else
              If vField.SpecialColumn Then
                vWhere.Append(DBSpecialCol(pFields.TableAlias, vName))
              Else
                vWhere.Append(vName)
              End If
              vWhere.Append(SQLLiteral(vOperator, vField.FieldType, vField.Value))
          End Select
        Else
          Select Case vWhereOperator
            Case CDBField.FieldWhereOperators.fwoOperatorOnly
              '
            Case CDBField.FieldWhereOperators.fwoNotIn
              vWhere.Append(vName)
              vWhere.Append(" NOT IN ( ")
              vWhere.Append(FormatInList(vField))
              vWhere.Append(" ) ")
            Case Else
              vWhere.Append(vName)
              vWhere.Append(" IN ( ")
              vWhere.Append(FormatInList(vField))
              vWhere.Append(" ) ")
          End Select
        End If
        If (vField.WhereOperator And CDBField.FieldWhereOperators.fwoCloseBracket) = CDBField.FieldWhereOperators.fwoCloseBracket Then
          vWhere.Append(" ) ")
        End If
        If (vField.WhereOperator And CDBField.FieldWhereOperators.fwoCloseBracketTwice) = CDBField.FieldWhereOperators.fwoCloseBracketTwice Then
          vWhere.Append(" )) ")
        End If
      Next
      Return vWhere.ToString
    End Function

    Private Function FormatInList(ByVal pField As CDBField) As String
      Select Case pField.FieldType
        Case CDBField.FieldTypes.cftDate, CDBField.FieldTypes.cftTime
          Dim vValues() As String = pField.Value.Split(","c)
          Dim vResult As New StringBuilder
          Dim vAddComma As Boolean
          For Each vItem As String In vValues
            If vItem.StartsWith("'") Then vItem = vItem.Substring(1, vItem.Length - 2)
            If vAddComma Then vResult.Append(",")
            If pField.FieldType = CDBField.FieldTypes.cftDate Then
              AppendDate(vResult, vItem)
            Else
              AppendDateTime(vResult, vItem)
            End If
            vAddComma = True
          Next
          Return vResult.ToString
        Case Else
          Return pField.Value
      End Select
    End Function
    Public Sub PopulateUnicode()
      If mvUnicodeFields Is Nothing Then
        mvUnicodeFields = New SortedList
        mvUnicodeFields.Add("forenames", "forenames")
        mvUnicodeFields.Add("initials", "initials")
        mvUnicodeFields.Add("surname", "surname")
        mvUnicodeFields.Add("salutation", "salutation")
        mvUnicodeFields.Add("label_name", "label_name")
        mvUnicodeFields.Add("preferred_forename", "preferred_forename")
        mvUnicodeFields.Add("surname_prefix", "surname_prefix")
        mvUnicodeFields.Add("informal_salutation", "informal_salutation")
        mvUnicodeFields.Add("name", "name")
        mvUnicodeFields.Add("sort_name", "sort_name")
        mvUnicodeFields.Add("abbreviation", "abbreviation")
        mvUnicodeFields.Add("search_name", "search_name")
      End If
    End Sub
    Public Function IsUnicodeField(ByVal pFieldName As String) As Boolean
      'This function will check if given name is unicode or not
      'Below mentioned fields are unicode in database.
      If RDBMSType = RDBMSTypes.rdbmsSqlServer AndAlso mvUnicodeFields.Contains(pFieldName) Then
        Return True
      Else
        Return False
      End If
    End Function

    Public Function NativeDataType(ByVal pFieldType As CDBField.FieldTypes, Optional ByVal pFieldLength As Integer = 0) As String
      Dim vField As New CDBField("", pFieldType)
      vField.Value = pFieldLength.ToString
      Return NativeDataType(vField)
    End Function

#End Region

#Region "Transaction Handling"

    Public ReadOnly Property InTransaction() As Boolean
      Get
        Return mvInTransaction
      End Get
    End Property

    Public Overridable ReadOnly Property MergeTerminator As String
      Get
        Return ";" 'According to the ANSI default, a merge statement should be terminated by a semi-colon.  Some DBMS (ORACLE) don't handle it.
      End Get
    End Property

    Public Function StartTransaction() As Boolean
      If Not mvInTransaction Then
        Select Case mvDataAccessMode
          Case DataAccessModes.damTest
            Debug.Print("TEST Start Transaction")
          Case DataAccessModes.damNormal
            If mvDebugSQLMessages Then Debug.Print("Start Transaction")
            If (mvSQLLoggingMode And SQLLoggingModes.AllSql) = SQLLoggingModes.AllSql Then
              LogSQL("Transaction Started")
              mvTransactionSW = New Stopwatch
              mvTransactionSW.Start()
            End If
            mvTransaction = mvConnection.BeginTransaction
        End Select
        mvInTransaction = True
        Return True                   'We started a transaction
      End If
    End Function

    Public Sub CommitTransaction()
      If mvInTransaction Then
        Select Case mvDataAccessMode
          Case DataAccessModes.damTest
            Debug.Print("TEST Commit Transaction")
          Case DataAccessModes.damNormal
            If mvDebugSQLMessages Then Debug.Print("Commit Transaction")
            mvTransaction.Commit()
            If (mvSQLLoggingMode And SQLLoggingModes.AllSql) = SQLLoggingModes.AllSql Then
              mvTransactionSW.Stop()
              LogSQL(String.Format("Transaction Committed - Duration {0}", mvTransactionSW.ElapsedMilliseconds))
            End If
            mvTransaction = Nothing
        End Select
      End If
      mvInTransaction = False
    End Sub

    Public Sub RollbackTransaction(Optional ByVal pReportError As Boolean = True)
      If mvInTransaction Then
        Select Case mvDataAccessMode
          Case DataAccessModes.damTest
            Debug.Print("TEST Rollback Transaction")
          Case DataAccessModes.damNormal
            If mvDebugSQLMessages Then Debug.Print("Rollback Transaction")
            mvTransaction.Rollback()
            If (mvSQLLoggingMode And SQLLoggingModes.AllSql) = SQLLoggingModes.AllSql Then
              mvTransactionSW.Stop()
              LogSQL(String.Format("Transaction Rolled Back - Duration {0}", mvTransactionSW.ElapsedMilliseconds))
            End If
            mvTransaction = Nothing
        End Select
      End If
      mvInTransaction = False
    End Sub

#End Region

#Region "Logging"

    Public Sub LogMailSQL(ByVal pSQL As String)
      If (mvSQLLoggingMode And SQLLoggingModes.AllSql) = SQLLoggingModes.AllSql OrElse
        (mvSQLLoggingMode And SQLLoggingModes.Mail) = SQLLoggingModes.Mail Then
        Dim vMessage As New SQLMessage(SQLMessage.SQLMessageTypes.MailSQL, String.Format("{0} {1} {2}", Now.ToString(CAREDateFormat), Now.ToLongTimeString, pSQL))
      End If
    End Sub

    Public Sub LogSQL(ByVal pSQL As String)
      If mvSQLLogQueueName.Length > 0 Then
        Try
          Dim mvMessageQueue As System.Messaging.MessageQueue = New System.Messaging.MessageQueue(mvSQLLogQueueName)
          If mvMessageQueue IsNot Nothing Then
            Dim vMessage As New System.Messaging.Message
            Dim vDataTime As String = String.Format("{0} {1}", Now.ToString(CAREDateFormat), Now.ToLongTimeString)
            vMessage.Label = String.Format("NG SQL Log TID {0} on {1}", System.Threading.Thread.CurrentThread.ManagedThreadId, vDataTime)
            vMessage.Body = String.Format("{0} TID {1} {2}", vDataTime, System.Threading.Thread.CurrentThread.ManagedThreadId, pSQL)
            mvMessageQueue.Send(vMessage)
          End If
        Catch vEx As Exception
          Debug.Print(vEx.Message)
        End Try
      Else
        Dim vMessage As New SQLMessage(SQLMessage.SQLMessageTypes.OtherSQL, String.Format("{0} {1} {2}", Now.ToString(CAREDateFormat), Now.ToLongTimeString, pSQL))
      End If
    End Sub

    Friend Sub LogRecordSetOpened(ByVal pSQL As String)
      If mvDebugSQLMessages Then Debug.Print(pSQL)
      If mvSQLLoggingMode <> SQLLoggingModes.None Then
        mvCurrentSQL = pSQL
        If (mvSQLLoggingMode And SQLLoggingModes.[Select]) = SQLLoggingModes.[Select] Then LogSQL(mvCurrentSQL)
        If (mvSQLLoggingMode And SQLLoggingModes.Timed) = SQLLoggingModes.Timed Then
          mvSW = New Stopwatch
          mvSW.Start()
        End If
      End If
    End Sub

    Friend Sub LogRecordSetClosed()
      If (mvSQLLoggingMode And SQLLoggingModes.Timed) = SQLLoggingModes.Timed Then
        mvSW.Stop()
        Dim vMessage As String = String.Format("Record Set Duration {0} - SQL: {1}", mvSW.ElapsedMilliseconds, mvCurrentSQL)
        Debug.Print(vMessage)
        LogSQL(vMessage)
      End If
    End Sub

    Private Class SQLMessage
      Public Enum SQLMessageTypes As Integer
        OtherSQL
        MailSQL
      End Enum

      Private mvMessage As String
      Private mvMessageType As SQLMessageTypes

      Public Sub New(ByVal pMessageType As SQLMessageTypes, ByVal pMessage As String)
        mvMessageType = pMessageType
        mvMessage = pMessage
        Dim vThreadStart As System.Threading.ThreadStart = New Threading.ThreadStart(AddressOf WriteLogMessage)
        Dim vThread As New Threading.Thread(vThreadStart)
        vThread.Start()
      End Sub

      Public Sub WriteLogMessage()
        Dim vBaseName As String
        Select Case mvMessageType
          Case SQLMessageTypes.MailSQL
            vBaseName = SQL_MAIL_FILENAME
          Case Else
            vBaseName = SQL_LOG_FILENAME
        End Select
        Dim vPath As String = System.IO.Path.GetTempPath
        Dim vFileName As String = String.Format("{0}{1}", vPath, vBaseName)
        Dim vMutex As New Threading.Mutex(False, vBaseName)
        vMutex.WaitOne(-1, False)
        Dim vWriter As IO.StreamWriter = Nothing
        Try
          Debug.Print("Writing Message " & mvMessage)
          vWriter = My.Computer.FileSystem.OpenTextFileWriter(vFileName, True)
          vWriter.WriteLine(mvMessage)
        Finally
          If vWriter IsNot Nothing Then vWriter.Close()
        End Try
        vMutex.ReleaseMutex()
      End Sub

    End Class

    Public ReadOnly Property Transaction As DbTransaction
      Get
        Return mvTransaction
      End Get
    End Property
#End Region

    Public MustOverride Function GetTableConverter(tableName As String,
                                                   changedColumns As List(Of TableConverter.ColumnDescriptor),
                                                   logfile As StreamWriter) As TableConverter

    Public MustOverride ReadOnly Property ConcatonateOperator As String

    Public MustOverride Function CopyColumnsToNewTable(pSelectionSQL As SQLStatement, pDestinationTableName As String) As Integer

    Public MustOverride Function CopyColumnsToNewTable(pSelectionSQL As String, pDestinationTableName As String) As Integer

#Region "IDisposable Support"
    Private disposedValue As Boolean

    Protected Overridable Sub Dispose(disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          Try
            Me.CloseConnection()
          Catch ex As Exception
            'Swallow any exceptions, we can't do much else
          End Try
        End If
      End If
      Me.disposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub
#End Region

  End Class
End Namespace