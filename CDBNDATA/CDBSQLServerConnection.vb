Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.IO

Namespace Data

  Friend Class CDBSQLServerConnection
    Inherits CDBConnection

    Private mvSQLServerMajorVersion As Integer
    Private mvConnectionString As String
    Private mvUseMultipleResultSets As Boolean
    Private mvAdditionalConnections As New List(Of DbConnection)

    Public Sub New(ByVal pSQLLoggingMode As SQLLoggingModes, ByVal pSQLLogQueueName As String)
      MyBase.New(pSQLLoggingMode, pSQLLogQueueName)
      mvRDBMSType = RDBMSTypes.rdbmsSqlServer
    End Sub

    Public Overrides Sub OpenConnection(ByVal pConnect As String, ByVal pDefaultLogname As String, ByVal pDefaultPassWord As String, ByVal pNeedMultipleResultSets As Boolean, ByVal pOverrideConnectString As Boolean)
      Dim vConnectString As String = GetFullConnectString(pConnect, pDefaultLogname, pDefaultPassWord, pOverrideConnectString)
      If pNeedMultipleResultSets Then vConnectString &= ";MultipleActiveResultSets=true"
      mvConnection = New SqlConnection(vConnectString)
      mvConnectionString = vConnectString
      If (mvSQLLoggingMode And SQLLoggingModes.All) = SQLLoggingModes.All Then LogSQL("OPEN CONNECTION " & pConnect)
      mvConnection.Open()
      Debug.Print("OPEN CONNECTION " & mvConnectionString)
      Dim vVersion As String = mvConnection.ServerVersion
      Dim vItems() As String = vVersion.Split("."c)
      mvSQLServerMajorVersion = IntegerValue(vItems(0))
      If pNeedMultipleResultSets AndAlso mvSQLServerMajorVersion > 8 Then
        mvUseMultipleResultSets = True
      Else
        mvUseMultipleResultSets = False
      End If
      mvConnectionOpen = True
    End Sub

#Region "Table and Index Schema methods"

    Public Overrides Function TableExists(ByVal pTableName As String) As Boolean
      Dim vRestriction(3) As String
      vRestriction(2) = pTableName
      Dim vTable As DataTable = mvConnection.GetSchema("Tables", vRestriction)
      If vTable.Rows.Count > 0 Then Return True
    End Function

    Public Overrides Function UnicodePerformed(ByVal pTableName As String, ByVal pAttributeName As String) As Boolean
      Dim vRestriction(3) As String
      vRestriction(2) = pTableName
      vRestriction(3) = pAttributeName
      Dim vTable As DataTable = mvConnection.GetSchema("Columns", vRestriction)
      If vTable.Rows.Count > 0 Then
        If vTable.Rows(0).Item("Data_Type").ToString = "nvarchar" Then
          Return True
        End If
      End If
    End Function
    Public Overrides Function AttributeExists(ByVal pTableName As String, ByVal pAttributeName As String) As Boolean
      Dim vRestriction(3) As String
      vRestriction(2) = pTableName
      vRestriction(3) = pAttributeName
      Dim vTable As DataTable = mvConnection.GetSchema("Columns", vRestriction)
      If vTable.Rows.Count > 0 Then Return True
    End Function

    Public Overrides Function BulkCopyData(ByVal pSourceConnection As CDBConnection, ByVal pDestinationTableName As String, ByVal pSQL As String) As Integer

      Dim vReader As DbDataReader = pSourceConnection.GetDataReader(pSQL)

      Return BulkCopyData(pSourceConnection, pDestinationTableName, vReader)

    End Function
    Public Overrides Function BulkCopyData(ByVal pSourceConnection As CDBConnection, ByVal pDestinationTableName As String, ByVal pTable As DataTable) As Integer

      Dim vReader As DbDataReader = pTable.CreateDataReader()

      Return BulkCopyData(pSourceConnection, pDestinationTableName, vReader)

    End Function

    Public Overloads Function BulkCopyData(ByVal pSourceConnection As CDBConnection, ByVal pDestinationTableName As String, ByVal pReader As DbDataReader) As Integer
      Dim vCount As Integer
      Using vCommandRowCount As New SqlCommand("SELECT COUNT(*) FROM " & pDestinationTableName, DirectCast(mvConnection, SqlConnection))
        vCommandRowCount.CommandTimeout = 0
        vCount = CInt(vCommandRowCount.ExecuteScalar())
        Using vBulkCopy As New SqlClient.SqlBulkCopy(DirectCast(mvConnection, SqlConnection), SqlBulkCopyOptions.TableLock, Nothing)
          vBulkCopy.DestinationTableName = pDestinationTableName
          vBulkCopy.BatchSize = 1000
          vBulkCopy.NotifyAfter = 5000
          vBulkCopy.BulkCopyTimeout = 3000     '5 minutes
          AddHandler vBulkCopy.SqlRowsCopied, AddressOf RowsCopiedHandler
          Dim vReader As IDataReader = DirectCast(pReader, IDataReader)
          Try
            For vIndex As Integer = 0 To vReader.FieldCount - 1
              vBulkCopy.ColumnMappings.Add(New SqlBulkCopyColumnMapping(vReader.GetName(vIndex), vReader.GetName(vIndex)))
            Next
            vBulkCopy.WriteToServer(vReader)
          Finally
            vReader.Close()
          End Try
        End Using
        Return CInt(vCommandRowCount.ExecuteScalar()) - vCount
      End Using
    End Function

    Public Overrides Function BulkCopyTable(ByVal pSourceConnection As CDBConnection, ByVal pTableName As String) As Integer
      Dim vCount As Integer
      Using vBulkCopy As New SqlClient.SqlBulkCopy(DirectCast(mvConnection, SqlConnection), SqlBulkCopyOptions.TableLock Or SqlBulkCopyOptions.KeepIdentity, Nothing)
        vBulkCopy.DestinationTableName = pTableName
        vBulkCopy.BatchSize = 1000
        vBulkCopy.NotifyAfter = 5000
        vBulkCopy.BulkCopyTimeout = 3000     '50 minutes
        AddHandler vBulkCopy.SqlRowsCopied, AddressOf RowsCopiedHandler
        ExecuteSQL("TRUNCATE TABLE " & pTableName)
        Dim vReader As SqlDataReader = DirectCast(pSourceConnection.GetDataReader("SELECT * FROM " & pTableName), SqlDataReader)
        Try
          For vIndex As Integer = 0 To vReader.FieldCount - 1
            vBulkCopy.ColumnMappings.Add(New SqlBulkCopyColumnMapping(vReader.GetName(vIndex), vReader.GetName(vIndex)))
          Next
          vBulkCopy.WriteToServer(vReader)
        Finally
          vReader.Close()
        End Try
        Using vCommandRowCount As New SqlCommand("SELECT COUNT(*) FROM " & pTableName, DirectCast(mvConnection, SqlConnection))
          vCommandRowCount.CommandTimeout = 0
          vCount = CInt(vCommandRowCount.ExecuteScalar())
        End Using
      End Using
      Return vCount
    End Function

    Public Overrides Function GetTableNames(Optional ByVal pSchemaName As String = "") As DataTable
      Dim vRestriction(3) As String
      Dim vTable As DataTable = mvConnection.GetSchema("Tables", vRestriction)
      Dim vRows As DataRow() = vTable.Select("TABLE_TYPE <> 'BASE TABLE' AND TABLE_TYPE <> 'VIEW'")
      For Each vRow As DataRow In vRows
        vTable.Rows.Remove(vRow)
      Next
      For vIndex As Integer = vTable.Columns.Count - 1 To 0 Step -1
        If vTable.Columns(vIndex).ColumnName <> "TABLE_NAME" AndAlso vTable.Columns(vIndex).ColumnName <> "TABLE_SCHEMA" Then
          vTable.Columns.Remove(vTable.Columns(vIndex))
        End If
      Next
      If vTable.Columns.Contains("TABLE_SCHEMA") Then vTable.Columns("TABLE_SCHEMA").ColumnName = "OWNER" 'Make consistent with Oracle
      Return vTable
    End Function

    Public Overrides Function GetAttributeNames(ByVal pTableName As String) As DataTable
      Dim vRestriction(3) As String
      vRestriction(2) = pTableName
      Dim vTable As DataTable = mvConnection.GetSchema("Columns", vRestriction)
      For vIndex As Integer = vTable.Columns.Count - 1 To 0 Step -1
        Select Case vTable.Columns(vIndex).ColumnName
          Case "COLUMN_NAME", "DATA_TYPE"
            'leave
          Case "TABLE_SCHEMA"
            vTable.Columns(vIndex).ColumnName = "OWNER"
          Case "IS_NULLABLE"
            vTable.Columns(vIndex).ColumnName = "NULLABLE"
          Case "CHARACTER_MAXIMUM_LENGTH"
            vTable.Columns(vIndex).ColumnName = "LENGTH"
          Case "NUMERIC_PRECISION"
            vTable.Columns(vIndex).ColumnName = "PRECISION"
          Case "NUMERIC_SCALE"
            vTable.Columns(vIndex).ColumnName = "SCALE"
          Case "ORDINAL_POSITION"
            vTable.Columns(vIndex).ColumnName = "POSITION"
          Case Else
            vTable.Columns.Remove(vTable.Columns(vIndex))
        End Select
      Next
      For Each vRow As DataRow In vTable.Rows
        vRow.Item("NULLABLE") = vRow.Item("NULLABLE").ToString.Substring(0, 1)
        Select Case vRow.Item("DATA_TYPE").ToString
          Case "varchar"
            If vRow.Field(Of Integer)("LENGTH") < 1 Then
              vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftMemo)
            Else
              vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftCharacter)
            End If
          Case "smallint"
            vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftInteger)
          Case "bit"
            vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftBit)
          Case "int"
            vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftLong)
          Case "datetime"
            vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftTime)
          Case "text"
            vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftMemo)
          Case "decimal", "numeric"
            vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftNumeric)
          Case "image", "varbinary"
            vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftBulk)
          Case "nvarchar"
            vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftUnicode)
          Case "binary"
            vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftBinary)
          Case Else
            Debug.Print("Unknown Data Type")
        End Select
      Next
      Return vTable
    End Function

    Public Overrides Function GetIndexNames(ByVal pTableName As String) As DataTable
      'Dim vRTable As DataTable = mvConnection.GetSchema("Restrictions")

      Dim vRestriction(3) As String
      vRestriction(2) = pTableName
      Dim vTable As DataTable = mvConnection.GetSchema("Indexes", vRestriction)
      For vIndex As Integer = vTable.Columns.Count - 1 To 0 Step -1
        If vTable.Columns(vIndex).ColumnName.ToUpper <> "INDEX_NAME" Then
          vTable.Columns.Remove(vTable.Columns(vIndex))
        Else
          vTable.Columns(vIndex).ColumnName = "INDEX_NAME"
        End If
      Next
      Dim vDataColumn As New DataColumn("UNIQUE")
      vTable.Columns.Add(vDataColumn)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("INDEXPROPERTY(OBJECT_ID(tab.name),ind.name,'IsUnique')", CDBField.FieldTypes.cftInteger, "1")
      vWhereFields.Add("INDEXPROPERTY(OBJECT_ID(tab.name),ind.name,'IsStatistics')", CDBField.FieldTypes.cftInteger, "0")
      vWhereFields.Add("tab.name", pTableName)
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("sysindexes ind", "ind.id", "tab.id")
      Dim vSQL As New SQLStatement(Me, "ind.name", "sysobjects tab", vWhereFields, "", vAnsiJoins)
      Dim vUniqueDataTable As DataTable = vSQL.GetDataTable
      For Each vUniqueRow As DataRow In vUniqueDataTable.Rows
        For Each vIndexRow As DataRow In vTable.Rows
          If vUniqueRow("name").ToString = vIndexRow("INDEX_NAME").ToString Then
            vIndexRow("UNIQUE") = "Y"
            Exit For
          End If
        Next
      Next
      Return vTable
    End Function

    Public Overrides Function GetIndexColumns(ByVal pTableName As String, ByVal pIndexName As String) As DataTable
      Dim vRTable As DataTable = mvConnection.GetSchema("Restrictions")

      Dim vRestriction(4) As String
      vRestriction(2) = pTableName
      vRestriction(3) = pIndexName
      Dim vTable As DataTable = mvConnection.GetSchema("IndexColumns", vRestriction)
      Dim vRows() As DataRow = vTable.Select("", "ordinal_position ASC")
      For Each vRow As DataRow In vRows
        Dim vValues() As Object = vRow.ItemArray
        vTable.Rows.Remove(vRow)
        vTable.Rows.Add(vValues)
      Next
      For vIndex As Integer = vTable.Columns.Count - 1 To 0 Step -1
        If vTable.Columns(vIndex).ColumnName.ToUpper = "COLUMN_NAME" Then
          vTable.Columns(vIndex).ColumnName = "COLUMN_NAME"
        Else
          vTable.Columns.Remove(vTable.Columns(vIndex))
        End If
      Next
      Return vTable
    End Function

    Public Overrides Function GetIdentityColumn(ByVal pTableName As String) As String
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("COLUMNPROPERTY(object_id(TABLE_NAME), COLUMN_NAME, 'IsIdentity')", 1)
      vWhereFields.Add("table_name", pTableName)
      Dim vSQL As New SQLStatement(Me, "column_name", "INFORMATION_SCHEMA.COLUMNS", vWhereFields)
      Return vSQL.GetValue
    End Function

    Public Overrides Sub ComputeTableStatistics(ByVal pTableName As String)

    End Sub

    Public Overrides Function DropIndexSql(ByVal pTable As String, ByVal pName As String) As IList(Of String)
      Dim vResult As New List(Of String)
      If mvSQLServerMajorVersion > 8 Then
        vResult.Add(String.Format("DROP INDEX {0} ON {1}", pName, pTable))
      Else
        vResult.Add(String.Format("DROP INDEX {0}.{1}", pTable, pName))
      End If
      Return vResult
    End Function

    Public Overrides Sub CreateView(ByVal pUserName As String, ByVal pViewName As String, ByVal pSQL As String)
      ExecuteSQL("CREATE VIEW dbo." & pViewName & " AS " & pSQL)
    End Sub

    Public Overrides Sub DropView(ByVal pViewName As String)
      ExecuteSQL("DROP VIEW " & pViewName, cdbExecuteConstants.sqlIgnoreError)
    End Sub

    Public Overrides Function IsUserDBA(ByVal pLogname As String, ByVal pTableName As String) As Boolean
      Return True
    End Function

    Public Overrides Function CreateIndexSql(ByVal pUnique As Boolean, ByVal pTable As String, pAttributes As IList(Of String)) As IList(Of String)
      Dim vResult As New List(Of String)
      Dim vHashName As String = GetIndexName(pTable, pAttributes)
      Dim vSQL As String = String.Format("CREATE {0}INDEX {1} ON {2} ({3})",
                                         If(pUnique, "UNIQUE ", String.Empty),
                                         vHashName,
                                         pTable,
                                         Function(attributes As IList(Of String)) As String
                                           Dim result As String = String.Empty
                                           For Each attribute As String In attributes
                                             result &= String.Format("{0}{1}", If(String.IsNullOrWhiteSpace(result), String.Empty, ", "), attribute)
                                           Next attribute
                                           Return result
                                         End Function(pAttributes))
      vResult.Add(vSQL)
      Return vResult.AsReadOnly
    End Function

#End Region

    Public Overrides Function GetDataSet(ByVal pSQL As SQLStatement) As DataSet
      Debug.Print(pSQL.SQL)
      If mvSQLLoggingMode <> SQLLoggingModes.None Then
        If (mvSQLLoggingMode And SQLLoggingModes.[Select]) = SQLLoggingModes.[Select] Then LogSQL(pSQL.SQL)
      End If
      Using vDA As New SqlDataAdapter(pSQL.SQL, DirectCast(mvConnection, SqlConnection))
        If mvInTransaction Then vDA.SelectCommand.Transaction = DirectCast(mvTransaction, SqlTransaction)
        Dim vDS As New DataSet
        vDA.Fill(vDS)
        Return vDS
      End Using
    End Function

    Public Overrides Function IndexIsUnique(pTableName As String, pAttributes As IList(Of String)) As Boolean
      Dim vIndexName As String = GetIndexName(pTableName, pAttributes)
      Dim vSql As New StringBuilder
      vSql.AppendLine("SELECT is_unique ")
      vSql.AppendLine("FROM   sys.indexes ")
      vSql.AppendLine("WHERE  name = @IndexName")
      Using vCommand As New SqlCommand(vSql.ToString, CType(mvConnection, SqlConnection))
        vCommand.Parameters.Add("@IndexName", SqlDbType.VarChar, 128).Value = vIndexName
        Using vData As New DataTable
          vData.Load(vCommand.ExecuteReader())
          If vData.Rows.Count = 0 Then
            Throw New InvalidOperationException("Index not found")
          End If
          Return CInt(vData.Rows(0)(0)) <> 0
        End Using
      End Using
    End Function

    Friend Overrides Sub NotifyRecordSetClosed(ByVal pRecordSet As CDBRecordSet)
      MyBase.NotifyRecordSetClosed(pRecordSet)
      If mvAdditionalConnections.Contains(pRecordSet.DbConnection) Then
        pRecordSet.DbConnection.Close()
        mvAdditionalConnections.Remove(pRecordSet.DbConnection)
        Debug.Print("Additional Connection Count " & mvAdditionalConnections.Count)
      End If
    End Sub

    Protected Overrides Function GetRecordSet(ByVal pSQL As String, ByVal pTimeOut As Integer, ByVal pOptions As RecordSetOptions) As CDBRecordSet
      Try
        'If we are trying to open a cursor and we have not specified MARS and the database is capable then switch the connection to MARS
        If pOptions = RecordSetOptions.MultipleResultSets AndAlso mvUseMultipleResultSets = False AndAlso mvSQLServerMajorVersion > 8 Then
          CloseConnection()
          OpenConnection(mvConnectionString, "", "", True, False)
        ElseIf pOptions = RecordSetOptions.MultipleResultSets AndAlso mvSQLServerMajorVersion <= 8 Then
          pOptions = RecordSetOptions.None
        End If

        Dim vRecordSet As CDBRecordSet = Nothing
        Using vCommand As DbCommand = mvConnection.CreateCommand
          vCommand.CommandText = pSQL
          vCommand.CommandType = CommandType.Text
          vCommand.CommandTimeout = pTimeOut
          Dim vRecordSetIsActive As Boolean
          If mvUseMultipleResultSets = False Then
            For Each vOpenRecordSet As CDBRecordSet In mvRecordSets
              If vOpenRecordSet.Active Then
                vRecordSetIsActive = True
                Exit For
              End If
            Next
          End If
          If vRecordSetIsActive Then
            'A record set is already open for this connection
            'Open another database connection
            Debug.Assert(False, "Opening Additional Database Connection for: " & pSQL)
            Dim vConnection As SqlConnection = New SqlConnection(mvConnectionString)
            If (mvSQLLoggingMode And SQLLoggingModes.All) = SQLLoggingModes.All Then LogSQL("OPEN ADDITIONAL SQL CONNECTION")
            vConnection.Open()
            mvAdditionalConnections.Add(vConnection)
            vCommand.Connection = vConnection
          Else
            vCommand.Connection = mvConnection
            If mvInTransaction Then vCommand.Transaction = DirectCast(mvTransaction, SqlTransaction)
          End If
          LogRecordSetOpened(pSQL)
          If pOptions = RecordSetOptions.None Then
            Using vDA As New SqlDataAdapter(CType(vCommand, SqlCommand))
              Dim vDataTable As New DataTable
              vDA.Fill(vDataTable)
              vRecordSet = New CDBSQLServerRecordSet(vDataTable, Me, vCommand.Connection)
            End Using
          Else
            Dim vReader As SqlDataReader = DirectCast(vCommand.ExecuteReader(), SqlDataReader)
            vRecordSet = New CDBSQLServerRecordSet(vReader, Me, vCommand.Connection)
          End If
          mvRecordSets.Add(vRecordSet)
          Return vRecordSet
        End Using
      Catch vSQLEx As SqlException
        If vSQLEx.Number = -2 Then
          RaiseError(DataAccessErrors.daeProcessTimeout)
        Else
          RaiseError(DataAccessErrors.daeDataAccessError, vSQLEx.Message, pSQL)
        End If
      Catch vEx As Exception
        RaiseError(DataAccessErrors.daeDataAccessError, vEx.Message, pSQL)
      End Try
      Return Nothing
    End Function

    Public Overrides Function GetRecordSetAnsiJoins(ByVal pSQL As String, ByVal pTimeout As Integer, ByVal pOptions As RecordSetOptions) As CDBRecordSet
      Return GetRecordSet(pSQL, pTimeout, pOptions)
    End Function

    Public Overrides Function IsCaseSensitive() As Boolean
      Return False
    End Function

    Public Overrides Function SupportsNoLock() As Boolean
      Return True
    End Function

    Public Overrides Function IsSpecialColumn(ByVal pName As String) As Boolean
      Select Case pName
        Case "case", "current", "distributed", "expression", "external", "function",
             "module", "number", "permanent", "prefix", "primary", "reference", "when"
          Return True
        Case Else
          Return False
      End Select
    End Function

    Public Overrides ReadOnly Property RowRestrictionType() As CDBConnection.RowRestrictionTypes
      Get
        Return RowRestrictionTypes.UseTopN
      End Get
    End Property
    Friend Overrides Sub AddUnicodeValue(ByVal pSQL As StringBuilder, ByVal pValue As String)
      With (pSQL)
        .Append("N'")
        .Append(pValue.Replace("'", "''"))
        .Append("'")
      End With
    End Sub
    Public Overrides Sub AppendDateTime(ByVal pSQL As StringBuilder, ByVal pValue As String)
      With (pSQL)
        If pValue.Length >= 8 Then
          Dim vDate As Date = CDate(pValue)
          .Append("'")
          .Append(vDate.Day)
          .Append(" ")
          .Append(GetMonthAbbreviation(vDate))
          .Append(" ")
          .Append(vDate.Year.ToString("D4"))
          .Append(" ")
          .Append(vDate.Hour)
          .Append(":")
          .Append(vDate.Minute)
          .Append(":")
          .Append(vDate.Second)
          .Append("'")
        Else
          .Append("'")
          .Append(pValue)
          .Append("'")
        End If
      End With
    End Sub

    Public Overrides Function ProcessAnsiJoins(ByVal pSQL As String) As String
      Return pSQL
    End Function

    Public Overrides Function DBLTrim(ByVal pExpression As String) As String
      Return String.Format(" LTRIM({0}) ", pExpression)
    End Function
    Public Overrides Function DBRTrim(ByVal pExpression As String) As String
      Return String.Format(" RTRIM({0}) ", pExpression)
    End Function
    Public Overrides Function DBLeft(ByVal pExpression As String, ByVal pLength As String) As String
      Return String.Format(" LEFT({0},{1}) ", pExpression, pLength)
    End Function
    Public Overrides Function DBIndexOf(ByVal pSearchString As String, ByVal pExpression As String) As String
      Return String.Format(" CHARINDEX({0},{1}) ", pSearchString, pExpression)
    End Function
    Public Overrides Function DBSubString(ByVal pExpression As String, ByVal pStart As String, ByVal pLength As String) As String
      Return String.Format("SUBSTRING({0},{1},{2})", pExpression, pStart, pLength)
    End Function
    Public Overrides Function DBCollateString() As String
      If mvSQLServerMajorVersion >= 8 Then
        Return " COLLATE Latin1_General_Bin"
      Else
        Return ""
      End If
    End Function
    Public Overrides Function DBConcatString() As String
      Return " + "
    End Function

    Public Overrides Function DBAddMonths(ByVal pDateAttribute As String, ByVal pMonthsAttribute As String) As String
      Return " DATEADD(mm," & pMonthsAttribute & "," & pDateAttribute & ") "
    End Function

    Public Overrides Function DBAddYears(ByVal pDateAttribute As String, ByVal pYearsAttribute As String) As String
      Return " DATEADD(yy," & pYearsAttribute & "," & pDateAttribute & ") "
    End Function

    Public Overrides Function DBAddWeeks(pDateAttribute As String, pWeeksAttribute As String) As String
      Return " DATEADD(wk," & pWeeksAttribute & "," & pDateAttribute & ") "
    End Function

    Public Overrides Function DBAge() As String
      Return " DATEPART(yy, GETDATE() - date_of_birth) - 1900"
    End Function

    Public Overrides Function DBDate() As String
      Return " GETDATE() "
    End Function

    Public Overrides Function DBHint(ByVal pHintType As DatabaseHintTypes, ByVal pTableName As String, ByVal pUseHint As Boolean) As String
      Select Case pHintType
        Case DatabaseHintTypes.dhtOptionForceOrder
          Return " OPTION (FORCE ORDER) "
      End Select
      Return ""
    End Function

    Public Overrides Function DBMonthDiff(ByVal pEarlierDate As String, ByVal pLaterDate As String) As String
      Return " DATEDIFF(m," & pEarlierDate & "," & pLaterDate & ") "
    End Function

    Public Overrides Function DBToNumber(ByVal pExpression As String) As String
      Return String.Format("CONVERT(FLOAT, {0})", pExpression)
    End Function

    Public Overrides Function DBToString(ByVal pExpression As String) As String
      Return String.Format(" CONVERT(VARCHAR, {0}) ", pExpression)
    End Function

    Public Overrides Function DBMaxToString(ByVal pExpression As String) As String
      Return String.Format(" CONVERT(VARCHAR(MAX), {0}) ", pExpression)
    End Function

    Public Overrides Function DBDateTimeAttribToDate(pAttributeName As String) As String
      Return String.Format(" CONVERT(DATE, {0})", pAttributeName)
    End Function
    Public Overrides Function DBToDate(pExpression As String) As String
      Return String.Format(" CONVERT(DATE, {0})", pExpression)
    End Function

    Public Overrides Function DBIsNull(ByVal pExpression As String, ByVal pReplacement As String) As String
      Return String.Format(" ISNULL({0},{1}) ", pExpression, pReplacement)
    End Function

    Public Overrides Function DBLength(ByVal pExpression As String) As String
      Return String.Format(" len({0}) ", pExpression)
    End Function

    Public Overrides Function DBLPad(ByVal pExpression As String, ByVal pLength As Integer) As String
      Return String.Format(" replicate('0',{0}-len({1}))+{1} ", pLength, pExpression)
    End Function

    Public Overrides Function DBForceOrder() As String
      Return " OPTION (FORCE ORDER) "
    End Function

    Public Overrides Function DBYear(ByVal pDateString As String) As String
      'Expects a date (could be a date attribute name) and returns YEAR ('1998-03-07')
      Dim vDate As Date = Nothing
      If Date.TryParse(pDateString, vDate) Then
        'We have a date so put in correct format
        pDateString = "'" & vDate.Year & "-" & vDate.Month & "-" & vDate.Day & "'"
      End If
      Return String.Format("YEAR ({0})", pDateString)
    End Function

    Public Overrides Function DBMonth(ByVal pDateString As String) As String
      'Expects a date (could be a date attribute name) and returns MONTH ('1998-03-07')
      Dim vDate As Date = Nothing
      If Date.TryParse(pDateString, vDate) Then
        'We have a date so put in correct format
        pDateString = "'" & vDate.Year & "-" & vDate.Month & "-" & vDate.Day & "'"
      End If
      Return String.Format("MONTH ({0})", pDateString)
    End Function

    Public Overrides Function DBDay(ByVal pDateString As String) As String
      'Expects a date (could be a date attribute name) and returns DAY ('1998-03-07')
      Dim vDate As Date = Nothing
      If Date.TryParse(pDateString, vDate) Then
        'We have a date so put in correct format
        pDateString = "'" & vDate.Year & "-" & vDate.Month & "-" & vDate.Day & "'"
      End If
      Return String.Format("DAY ({0})", pDateString)
    End Function

    Public Overrides Function DBReplaceLineFeedWithSpace(ByVal pFieldName As String) As String
      Return String.Format("replace(cast({0} as varchar(max)),char(10),' ')", pFieldName)
    End Function

    Public Overrides Function SupportsOption(ByVal pOption As DatabaseOptions) As Boolean
      Select Case pOption
        Case DatabaseOptions.ForXML
          Return True
        Case DatabaseOptions.ForXML_XSINIL
          If Not mvSQLServerMajorVersion = 8 Then Return True
        Case Else
          Return False
      End Select
    End Function

    Protected Overrides Function TableOwnerPrefix() As String
      Return "dbo."
    End Function

    Public Overrides Function NativeDataType(ByVal pField As CDBField) As String
      Select Case pField.FieldType
        Case CDBField.FieldTypes.cftCharacter
          Return "varchar(" & pField.Value & ")"
        Case CDBField.FieldTypes.cftInteger
          Return "smallint"
        Case CDBField.FieldTypes.cftLong
          Return "int"
        Case CDBField.FieldTypes.cftNumeric
          Return "decimal(" & pField.Value & ",2)"
        Case CDBField.FieldTypes.cftMemo
          Return "varchar(max)"
        Case CDBField.FieldTypes.cftDate, CDBField.FieldTypes.cftTime
          Return "datetime"
        Case CDBField.FieldTypes.cftBulk
          Return "varbinary(max)"
        Case CDBField.FieldTypes.cftIdentity
          Return "int identity(1,1)"
        Case CDBField.FieldTypes.cftBit
          Return "bit"
        Case CDBField.FieldTypes.cftUnicode
          Return "nvarchar(" & pField.Value & ")"
        Case CDBField.FieldTypes.cftBinary
          Return "binary(" & pField.Value & ")"
        Case CDBField.FieldTypes.cftGUID
          Return "uniqueidentifier"
        Case Else
          Return "UNKNOWN DATA TYPE"
      End Select
    End Function

    Friend Overrides Function GetDBParameter(ByVal pFieldName As String, ByVal pFieldValue As String, ByRef pParamName As String) As System.Data.Common.DbParameter
      pParamName = "@" & pFieldName.Replace("_", "")
      Dim vDBParameter As New SqlParameter(pParamName, pFieldValue.Replace(vbCr, ""))
      Return vDBParameter
    End Function

    Friend Overrides Function GetDBBulkParameter(ByVal pFieldName As String, ByVal pFieldValue As String, ByRef pParamName As String) As System.Data.Common.DbParameter
      pParamName = "@" & pFieldName.Replace("_", "")
      Dim vValue As String = pFieldValue.Replace(vbCr, "")
      Dim vEncoding As New System.Text.ASCIIEncoding
      Dim vBytes() As Byte = vEncoding.GetBytes(vValue)
      Dim vDBParameter As New SqlParameter(pParamName, SqlDbType.Image, vBytes.Length, ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, vBytes)
      Return vDBParameter
    End Function

    Friend Overrides Function GetDBParameterFromFile(ByVal pFieldName As String, ByVal pFilename As String, ByRef pParamName As String) As System.Data.Common.DbParameter
      pParamName = "@" & pFieldName.Replace("_", "")
      Dim vFS As New IO.FileStream(pFilename, IO.FileMode.Open)
      Dim vBuffer(CInt(vFS.Length - 1)) As Byte
      vFS.Read(vBuffer, 0, CInt(vFS.Length))
      vFS.Close()
      Dim vDBParameter As New SqlParameter(pParamName, SqlDbType.Image, vBuffer.Length, ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, vBuffer)
      Return vDBParameter
    End Function

    Public Overrides Sub DropTable(ByVal pTableName As String)
      ExecuteSQL("DROP TABLE " & pTableName, cdbExecuteConstants.sqlIgnoreError)
    End Sub

    Public Overrides Function DeleteAllRecords(ByVal pTable As String) As Integer
      Dim vSQL As String = "TRUNCATE TABLE " & pTable
      Select Case mvDataAccessMode
        Case DataAccessModes.damTest
          Debug.Print("TEST " & vSQL)
          Return 1 'Assume that it affected some rows although we cannot tell
        Case Else
          If mvDebugSQLMessages Then Debug.Print(vSQL)
          If (mvSQLLoggingMode And SQLLoggingModes.Delete) > 0 Then LogSQL(vSQL)
          Return ExecuteSQL(vSQL)
      End Select
    End Function

    Public Overrides Function InsertRecord(ByVal pTableName As String, ByVal pFields As CDBFields, ByVal pIgnoreDuplicates As Boolean) As Boolean
      Try
        Return InsertRecord(pTableName, pFields)
      Catch ex As SqlClient.SqlException
        If ex.Number = 2601 AndAlso pIgnoreDuplicates Then
          mvLastInsertErrorIsDuplicate = True 'Ignore duplicate
        Else
          Throw ex
        End If
      End Try
    End Function

    Public Overrides Function UpdateRecords(ByVal pTableName As String, ByVal pUpdateFields As CDBFields, ByVal pWhereFields As CDBFields, Optional ByVal pErrorIfNoRecords As Boolean = True) As Integer
      Try
        Return DoUpdateRecords(pTableName, pUpdateFields, pWhereFields, pErrorIfNoRecords)
      Catch ex As SqlClient.SqlException
        If ex.Number = 2601 AndAlso pErrorIfNoRecords = False Then
          mvLastInsertErrorIsDuplicate = True 'Ignore duplicate
        Else
          Throw ex
        End If
      End Try
    End Function
    Public Overrides Function UseTableSpaces() As Boolean
      Return False
    End Function
    Friend Overrides Function PreProcessMaxRows(ByVal pSQLStatement As SQLStatement) As SQLStatement
      Return pSQLStatement
    End Function

    Public Overrides Function GetDBParameterFromByteArray(pFieldName As String, pValue() As Byte, ByRef pParamName As String) As DbParameter
      pParamName = "@" & pFieldName.Replace("_", "")
      Dim vDBParameter As New SqlParameter(pParamName, SqlDbType.Image, pValue.Length, ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, pValue)
      Return vDBParameter
    End Function

    Public Overrides Function GetBinaryDBParameter(ByVal pFieldName As String, ByVal pValue As Byte(), ByRef pParamName As String) As DbParameter
      'For password encryption BR19442 in Oracle a new data type (Binary) was added to code and it is stored as a blob in the database
      'In order for curent GetDBParameterFromByteArray code to work for Oracle database, a new function was required
      Return GetDBParameterFromByteArray(pFieldName, pValue, pParamName)
    End Function

    Public Overrides Function GetTableConverter(tableName As String,
                                                changedColumns As List(Of TableConverter.ColumnDescriptor),
                                                logfile As StreamWriter) As TableConverter
      Return New SqlServerTableConverter(tableName, changedColumns, Me, logfile)
    End Function

    Private ReadOnly Property Connection As SqlClient.SqlConnection
      Get
        Return DirectCast(mvConnection, SqlClient.SqlConnection)
      End Get
    End Property

    Public Overrides ReadOnly Property ConcatonateOperator As String
      Get
        Return "+"
      End Get
    End Property

    Public Overrides Function IsBaseTable(schemaName As String, tableName As String) As Boolean
      Dim table As DataTable = mvConnection.GetSchema("Tables", {Nothing, schemaName, tableName})
      If table.Rows.Count < 1 Then
        Throw New InvalidOperationException("Attempt to find determine the type of a non-existent table")
      End If
      If table.Rows.Count > 1 Then
        Throw New InvalidOperationException("Attempt to find determine the type of a non-unique table")
      End If
      Return DirectCast(table.Rows(0)("TABLE_TYPE"), String).Equals("BASE TABLE", StringComparison.InvariantCultureIgnoreCase)
    End Function
    Public Overrides Function GetPrimaryKeyColumns(pTableName As String) As DataTable
      Dim vFields As String = "columns.name as COLUMN_NAME"
      Dim vJoins As New AnsiJoins
      vJoins.Add(New AnsiJoin("sys.index_columns", "index_columns.[object_id]", "indexes.[object_id]", "index_columns.index_id", "indexes.index_id"))
      vJoins.Add(New AnsiJoin("sys.columns", "columns.[object_id]", "index_columns.[object_id]", "columns.column_id", "index_columns.column_id"))
      vJoins.Add(New AnsiJoin("sys.tables", "columns.[object_id]", "tables.[object_id]"))
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("indexes.is_primary_key", 1)
      vWhereFields.Add("tables.name", pTableName)
      Dim vSql As New SQLStatement(Me, vFields, "sys.indexes", vWhereFields, "", vJoins)
      Dim vPKColumns As DataTable = vSql.GetDataTable
      vPKColumns.TableName = "PRIMARY_KEY"
      Return vPKColumns
    End Function
    Public Overrides Function CopyColumnsToNewTable(pSelectionSQL As SQLStatement, pDestinationTableName As String) As Integer
      Dim vSQL As String = pSelectionSQL.SQL
      Return CopyColumnsToNewTable(vSQL, pDestinationTableName)
    End Function
    Public Overrides Function CopyColumnsToNewTable(pSelectionSQL As String, pDestinationTableName As String) As Integer
      Dim vCopySQL As New StringBuilder
      vCopySQL.Append(" INTO ")
      vCopySQL.Append(pDestinationTableName)
      vCopySQL.Append(" FROM")
      Dim vSQL As String = pSelectionSQL.Replace("FROM", vCopySQL.ToString())
      Dim vInt As Integer = ExecuteSQL(vSQL)
      Return vInt
    End Function

    ''' <summary>Bulk Update and Insert Data into a Table.</summary>
    ''' <param name="pSQLStatement">The SQLStatement used to populate the DataTable.  Used so that the same SQL is used for the original data and for the changed data.</param>
    ''' <param name="pDataTable">The data to be inserted or updated. This must have the RowState set correctly on each row.</param>
    Public Overrides Sub BulkUpdate(ByVal pSQLStatement As SQLStatement, ByVal pDataTable As DataTable)
      Using vSQLAdapter As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(pSQLStatement.SQL, Me.Connection)
        If Me.InTransaction Then
          'If we are currently in a transaction we need to set the transaction against the SelectCommand object
          Dim vCmd As SqlCommand = vSQLAdapter.SelectCommand
          If vCmd IsNot Nothing Then
            vCmd.Transaction = CType(Me.Transaction, SqlTransaction)
          End If
        End If
        Using New SqlCommandBuilder(vSQLAdapter)
          'Insert / Update the DataTable into the DataBase
          'The update requires primary key data on the database table and in the DataTable
          vSQLAdapter.Update(pDataTable)
        End Using
      End Using
    End Sub

    Public Overrides Sub PopulateGUIDColumn(pTableName As String, pColumnName As String)
      Me.ExecuteSQL(String.Format("UPDATE {0} SET {1} = newid()", pTableName, pColumnName))
    End Sub
    Public Overrides ReadOnly Property MergeTerminator As String
      Get
        Return ";" ' merge terminator required in SQL Server but not in ORACLE
      End Get
    End Property
  End Class
End Namespace