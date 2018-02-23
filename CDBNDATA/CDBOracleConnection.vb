Imports System.Data.OracleClient
Imports System.Data.Common
Imports System.Linq
Imports System.IO

Namespace Data

  Friend Class CDBOracleConnection
    Inherits CDBConnection

    Private mvOracleVersion As Integer
    Private mvTSChecked As Boolean
    Private mvUseTableSpaces As Boolean

    Public Sub New(ByVal pSQLLoggingMode As SQLLoggingModes, ByVal pSQLLogQueueName As String)
      MyBase.New(pSQLLoggingMode, pSQLLogQueueName)
      mvRDBMSType = RDBMSTypes.rdbmsOracle
    End Sub

    Public Overrides Sub OpenConnection(ByVal pConnect As String, ByVal pDefaultLogname As String, ByVal pDefaultPassWord As String, ByVal pNeedMultipleResultSets As Boolean, ByVal pOverrideConnectString As Boolean)
      mvConnection = New OracleConnection(GetFullConnectString(pConnect, pDefaultLogname, pDefaultPassWord, pOverrideConnectString))
      If (mvSQLLoggingMode And SQLLoggingModes.All) = SQLLoggingModes.All Then LogSQL("OPEN CONNECTION " & pConnect)
      mvConnection.Open()
      Using vCmd As DbCommand = mvConnection.CreateCommand()
        vCmd.CommandText = "ALTER SESSION SET NLS_SORT=BINARY_CI"
        vCmd.ExecuteNonQuery()
        vCmd.CommandText = "ALTER SESSION SET NLS_COMP=LINGUISTIC"
        vCmd.ExecuteNonQuery()
      End Using
      mvConnectionOpen = True
    End Sub

#Region "Table and Index Schema methods"

    Public Overrides Function TableExists(ByVal pTableName As String) As Boolean
      Dim vRestriction(1) As String
      vRestriction(1) = pTableName.ToUpper
      Dim vTable As DataTable = mvConnection.GetSchema("Tables", vRestriction)
      If vTable.Rows.Count > 0 Then Return True
    End Function

    Public Overrides Function AttributeExists(ByVal pTableName As String, ByVal pAttributeName As String) As Boolean
      'Dim vRTable As DataTable = mvConnection.GetSchema("Restrictions")

      Dim vRestriction(2) As String
      vRestriction(1) = pTableName.ToUpper
      vRestriction(2) = pAttributeName.ToUpper
      Dim vTable As DataTable = mvConnection.GetSchema("Columns", vRestriction)
      If vTable.Rows.Count > 0 Then Return True
    End Function
    Public Overrides Function UnicodePerformed(ByVal pTableName As String, ByVal pAttributeName As String) As Boolean
      Dim vRestriction(2) As String
      vRestriction(1) = pTableName
      vRestriction(2) = pAttributeName
      Dim vTable As DataTable = mvConnection.GetSchema("Columns", vRestriction)
      If vTable.Rows.Count > 0 Then Return True
    End Function
    Friend Overrides Sub AddUnicodeValue(ByVal pSQL As StringBuilder, ByVal pValue As String)
      With (pSQL)
        .Append("N'")
        .Append(pValue.Replace("'", "''"))
        .Append("'")
      End With
    End Sub
    Public Overrides Function BulkCopyData(ByVal pSourceConnection As CDBConnection, ByVal pDestinationTableName As String, ByVal pSQL As String) As Integer
      Throw New InvalidOperationException
    End Function

    Public Overrides Function BulkCopyData(pSourceConnection As CDBConnection, pDestinationTableName As String, pReader As DataTable) As Integer
      Throw New InvalidOperationException
    End Function

    Public Overrides Function BulkCopyTable(ByVal pSourceConnection As CDBConnection, ByVal pTableName As String) As Integer
      Throw New InvalidOperationException
    End Function

    Public Overrides Function GetTableNames(Optional ByVal pSchemaName As String = "") As DataTable
      Dim vRestriction(1) As String
      Dim vTable As DataTable = mvConnection.GetSchema("Tables", vRestriction)
      Dim vRows As DataRow() = vTable.Select("TYPE <> 'User'")

      For Each vRow As DataRow In vRows
        vTable.Rows.Remove(vRow)
      Next
      If pSchemaName.Length > 0 Then
        vRows = vTable.Select("OWNER <> '" & pSchemaName & "'")
        For Each vRow As DataRow In vRows
          vTable.Rows.Remove(vRow)
        Next
      End If
      For vIndex As Integer = vTable.Columns.Count - 1 To 0 Step -1
        If vTable.Columns(vIndex).ColumnName <> "TABLE_NAME" AndAlso vTable.Columns(vIndex).ColumnName <> "OWNER" Then
          vTable.Columns.Remove(vTable.Columns(vIndex))
        End If
      Next
      If Not String.IsNullOrWhiteSpace(pSchemaName) Then
        Using vViewTable As DataTable = mvConnection.GetSchema("Views", vRestriction)
          For Each vRow As DataRow In vViewTable.Select("OWNER = '" & pSchemaName & "'")
            Dim vNewRow As DataRow = vTable.NewRow()
            vNewRow("TABLE_NAME") = vRow("VIEW_NAME")
            vNewRow("OWNER") = vRow("OWNER")
            vTable.Rows.Add(vNewRow)
          Next vRow
        End Using
      End If
      Return vTable
    End Function

    Public Overrides Function GetAttributeNames(ByVal pTableName As String) As DataTable
      Dim vRestriction(2) As String
      vRestriction(1) = pTableName.ToUpper
      Dim vTable As DataTable = mvConnection.GetSchema("Columns", vRestriction)
      For vIndex As Integer = vTable.Columns.Count - 1 To 0 Step -1
        Select Case vTable.Columns(vIndex).ColumnName
          Case "COLUMN_NAME", "DATA_TYPE", "PRECISION", "SCALE", "LENGTH", "NULLABLE", "OWNER"
            'leave
          Case "DATATYPE"
            vTable.Columns(vIndex).ColumnName = "DATA_TYPE"
          Case "ID"
            vTable.Columns(vIndex).ColumnName = "POSITION"
          Case Else
            vTable.Columns.Remove(vTable.Columns(vIndex))
        End Select
      Next
      For Each vRow As DataRow In vTable.Rows
        vRow.Item("COLUMN_NAME") = vRow.Item("COLUMN_NAME").ToString.ToLower
        Select Case vRow.Item("DATA_TYPE").ToString
          Case "VARCHAR2"
            vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftCharacter)
          Case "NUMBER"
            If vRow.Item("SCALE").ToString = "0" Then
              vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftLong)
            Else
              vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftNumeric)
            End If
          Case "DATE"
            vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftTime)
          Case "LONG", "CLOB"
            vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftMemo)
          Case "LONG RAW"
            vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftBulk)
          Case "BINARY", "BLOB"
            If Not vRow.Field(Of String)("COLUMN_NAME").Equals("password", StringComparison.InvariantCultureIgnoreCase) Then
              vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftBulk)
            Else
              vRow.Item("DATA_TYPE") = CDBField.GetFieldTypeCode(CDBField.FieldTypes.cftBinary)
            End If
          Case Else
            Debug.Print("Unknown Data Type")
        End Select
      Next
      Return vTable
    End Function

    Public Overrides Function GetIndexNames(ByVal pTableName As String) As DataTable
      Dim vSql As New StringBuilder
      vSql.AppendLine("SELECT ai.index_name, ")
      vSql.AppendLine("       CASE ai.index_type ")
      vSql.AppendLine("         WHEN 'NORMAL' THEN ai.uniqueness ")
      vSql.AppendLine("         WHEN 'FUNCTION-BASED NORMAL' THEN ")
      vSql.AppendLine("           CASE ")
      vSql.AppendLine("             WHEN EXISTS (SELECT index_name ")
      vSql.AppendLine("                          FROM   all_indexes ")
      vSql.AppendLine("                          WHERE  index_name = concat(substr(ai.index_name,1,28), '_U' )) THEN 'UNIQUE' ")
      vSql.AppendLine("             ELSE ")
      vSql.AppendLine("               'NONUNIQUE' ")
      vSql.AppendLine("           END ")
      vSql.AppendLine("         ELSE ")
      vSql.AppendLine("           'Error' ")
      vSql.AppendLine("       END UNIQUENESS ")
      vSql.AppendLine("FROM   all_indexes ai ")
      vSql.AppendLine("WHERE  ai.table_name = :TableName ")
      vSql.AppendLine("       AND ai.index_name NOT LIKE '%_U'")
      Dim vTable As New DataTable
      Using vCommand As DbCommand = Me.CreateCommand
        vCommand.CommandType = CommandType.Text
        vCommand.CommandText = vSql.ToString
        DirectCast(vCommand, OracleCommand).Parameters.Add(":TableName", OracleType.VarChar, 128).Value = pTableName
        Using vReader As IDataReader = vCommand.ExecuteReader
          vTable.Load(vReader)
        End Using
      End Using
      vTable.Columns("UNIQUENESS").ColumnName = "UNIQUE"
      For Each vRow As DataRow In vTable.Rows
        If vRow("UNIQUE").ToString = "UNIQUE" Then
          vRow("UNIQUE") = "Y"
        Else
          vRow("UNIQUE") = "N"
        End If
      Next
      For vIndex As Integer = vTable.Columns.Count - 1 To 0 Step -1
        If vTable.Columns(vIndex).ColumnName.ToUpper <> "INDEX_NAME" Then
          If vTable.Columns(vIndex).ColumnName <> "UNIQUE" Then vTable.Columns.Remove(vTable.Columns(vIndex))
        Else
          vTable.Columns(vIndex).ColumnName = "INDEX_NAME"
        End If
      Next
      Return vTable
    End Function

    Public Overrides Function GetIndexColumns(ByVal pTableName As String, ByVal pIndexName As String) As DataTable
      Dim vSql As New StringBuilder
      vSql.AppendLine("SELECT ai.index_name, ")
      vSql.AppendLine("       CASE ai.index_type ")
      vSql.AppendLine("         WHEN 'NORMAL' THEN ai.uniqueness ")
      vSql.AppendLine("         WHEN 'FUNCTION-BASED NORMAL' THEN ")
      vSql.AppendLine("           CASE ")
      vSql.AppendLine("             WHEN EXISTS (SELECT index_name ")
      vSql.AppendLine("                          FROM   all_indexes ")
      vSql.AppendLine("                          WHERE  index_name = concat(substr(ai.index_name,1,28), '_U' )) THEN 'UNIQUE' ")
      vSql.AppendLine("             ELSE ")
      vSql.AppendLine("               'NONUNIQUE' ")
      vSql.AppendLine("           END ")
      vSql.AppendLine("         ELSE ")
      vSql.AppendLine("           'Error' ")
      vSql.AppendLine("       END UNIQUENESS, ")
      vSql.AppendLine("       aic.column_position, ")
      vSql.AppendLine("       aic.column_name, ")
      vSql.AppendLine("       aie.column_expression, ")
      vSql.AppendLine("       aic.descend ")
      vSql.AppendLine("FROM   all_ind_columns aic ")
      vSql.AppendLine("       INNER JOIN all_indexes ai ")
      vSql.AppendLine("         ON ai.index_name = aic.index_name ")
      vSql.AppendLine("       LEFT OUTER JOIN all_ind_expressions aie ")
      vSql.AppendLine("         ON aie.index_name = aic.index_name ")
      vSql.AppendLine("            AND aie.column_position = aic.column_position ")
      vSql.AppendLine("WHERE  aic.table_name = :TableName ")
      vSql.AppendLine("       AND aic.index_name = :IndexName ")
      vSql.AppendLine("ORDER BY aic.column_position")
      Dim vTable As New DataTable
      Using vCommand As DbCommand = Me.CreateCommand
        vCommand.CommandType = CommandType.Text
        vCommand.CommandText = vSql.ToString
        DirectCast(vCommand, OracleCommand).Parameters.Add(":TableName", OracleType.VarChar, 128).Value = pTableName
        DirectCast(vCommand, OracleCommand).Parameters.Add(":IndexName", OracleType.VarChar, 128).Value = pIndexName
        Using vReader As IDataReader = vCommand.ExecuteReader
          vTable.Load(vReader)
        End Using
      End Using
      For Each vRow As DataRow In vTable.Rows
        If Not vRow.IsNull("COLUMN_EXPRESSION") Then
          Dim vExpresssion As String = CStr(vRow("COLUMN_EXPRESSION")).Trim
          If vExpresssion.StartsWith("NLSSORT(""", StringComparison.InvariantCultureIgnoreCase) AndAlso
            vExpresssion.EndsWith(""",'nls_sort=''BINARY_CI''')", StringComparison.InvariantCultureIgnoreCase) Then
            vRow("COLUMN_NAME") = vExpresssion.Substring(9, vExpresssion.LastIndexOf("""") - 9)
          End If
        End If
      Next vRow
      For vIndex As Integer = vTable.Columns.Count - 1 To 0 Step -1
        If vTable.Columns(vIndex).ColumnName.Equals("COLUMN_NAME", StringComparison.InvariantCultureIgnoreCase) Then
          vTable.Columns(vIndex).ColumnName = "COLUMN_NAME"
        Else
          vTable.Columns.Remove(vTable.Columns(vIndex))
        End If
      Next
      Return vTable
    End Function

    Public Overrides Function GetIdentityColumn(ByVal pTableName As String) As String
      Return ""
    End Function

    Public Overrides Sub ComputeTableStatistics(ByVal pTableName As String)
      If (OracleVersion() >= 8 And OracleVersion() < 10) Then
        ExecuteSQL("begin dbms_stats.gather_table_stats( ownname=>'CARE_ADMIN', tabname=>'" & pTableName & "', cascade=> TRUE); end;")
      End If
    End Sub

    Private Function OracleVersion() As Integer
      If mvOracleVersion = 0 Then
        Dim vVersion As String = GetValue("SELECT version FROM v$instance")
        Dim vPos As Integer = vVersion.IndexOf(".")
        If vPos > 0 Then vVersion = vVersion.Substring(0, vPos)
        mvOracleVersion = CInt(vVersion)
      End If
      Return mvOracleVersion
    End Function

    Public Overrides Function DropIndexSql(ByVal pTable As String, ByVal pName As String) As IList(Of String)
      Dim vResult As New List(Of String)
      Dim vSQL As New StringBuilder
      vSQL.AppendLine(String.Format("DROP INDEX {0}", GetUniqueVarientIndexName(pName)))
      vResult.Add(vSQL.ToString)
      vSQL.Length = 0
      vSQL.AppendLine(String.Format("DROP INDEX {0}", pName))
      vResult.Add(vSQL.ToString)
      Return vResult
    End Function

    Private Function GetUniqueVarientIndexName(pIndexName As String) As String
      Return If(pIndexName.Length > 28, pIndexName.Substring(0, 28) & "_U", pIndexName & "_U")
    End Function

    Public Overrides Sub CreateView(ByVal pUserName As String, ByVal pViewName As String, ByVal pSQL As String)
      ExecuteSQL("CREATE VIEW " & pViewName & " AS " & ProcessAnsiJoins(pSQL) & " WITH READ ONLY")
      ExecuteSQL("CREATE PUBLIC SYNONYM " & pViewName & " FOR " & pUserName & "." & pViewName)
      ExecuteSQL("GRANT SELECT ON " & pViewName & " TO care_user")
    End Sub

    Public Overrides Sub DropView(ByVal pViewName As String)
      ExecuteSQL("DROP VIEW " & pViewName, cdbExecuteConstants.sqlIgnoreError)
      ExecuteSQL("DROP PUBLIC SYNONYM " & pViewName, cdbExecuteConstants.sqlIgnoreError)
    End Sub

    Public Overrides Function IsUserDBA(ByVal pLogname As String, ByVal pTableName As String) As Boolean
      Dim vDataTable As DataTable = GetTableNames(pLogname.ToUpper)
      Dim vRows As DataRow() = vDataTable.Select(String.Format("TABLE_NAME = '{0}'", pTableName))
      If vRows.Length > 0 Then
        Return True
      Else
        Return False
      End If
    End Function
    Public Overrides Function CreateIndexSql(pUnique As Boolean, pTable As String, pAttributes As IList(Of String)) As IList(Of String)
      Dim vResult As New List(Of String)
      Dim vUniqueRequired As Boolean = pUnique
      Dim vSQL As New StringBuilder
      Dim vHashName As String = GetIndexName(pTable, pAttributes)
      vSQL.AppendLine("SELECT column_name, ")
      vSQL.AppendLine("       data_type ")
      vSQL.AppendLine("FROM   all_tab_columns ")
      vSQL.AppendLine(String.Format("WHERE  table_name = '{0}' ", pTable))
      vSQL.AppendLine(String.Format("       AND column_name IN ({0})",
                                    Function(attributes As IList(Of String)) As String
                                      Dim result As String = String.Empty
                                      For Each attribute As String In attributes
                                        result &= String.Format("{0}'{1}'", If(String.IsNullOrWhiteSpace(result), String.Empty, ", "), attribute)
                                      Next attribute
                                      Return result
                                    End Function(pAttributes)))
      Using vCommand As DbCommand = Me.CreateCommand
        vCommand.CommandText = vSQL.ToString
        Dim vColumnMap As New Dictionary(Of String, KeyValuePair(Of String, String))(StringComparer.InvariantCultureIgnoreCase)
        Using vColumnData As New DataTable()
          Using vReader As IDataReader = vCommand.ExecuteReader
            vColumnData.Load(vReader)
          End Using
          For Each vColumnDefinition As KeyValuePair(Of String, String) In From vRow As DataRow In vColumnData.AsEnumerable
                                                                           Select New KeyValuePair(Of String, String)(vRow.Field(Of String)("column_name"),
                                                                                                                      vRow.Field(Of String)("data_type"))
            vColumnMap.Add(vColumnDefinition.Key, vColumnDefinition)
          Next vColumnDefinition
        End Using
        If vUniqueRequired AndAlso
           (From vRow As KeyValuePair(Of String, String) In vColumnMap.Values
            Where vRow.Value.IndexOf("char", StringComparison.InvariantCultureIgnoreCase) >= 0
            Select vRow.Key).Count > 0 Then
          vSQL.Length = 0
          vSQL.Append(String.Format("CREATE {0}INDEX {1} ",
                                    If(vUniqueRequired, "UNIQUE ", String.Empty),
                                    GetUniqueVarientIndexName(vHashName)))
          vSQL.Append(String.Format(" ON {0} ({1})",
                                    pTable,
                                    Function(attributes As IList(Of String)) As String
                                      Dim result As String = String.Empty
                                      For Each attribute As String In attributes
                                        result &= String.Format("{0}{1}", If(String.IsNullOrWhiteSpace(result), String.Empty, ", "), attribute)
                                      Next attribute
                                      Return result
                                    End Function(pAttributes)))
          vResult.Add(vSQL.ToString)
          vUniqueRequired = False
        End If
        vSQL.Length = 0
        vSQL.Append(String.Format("CREATE {0}INDEX {1} ",
                                  If(vUniqueRequired, "UNIQUE ", String.Empty),
                                  vHashName))
        vSQL.Append(String.Format(" ON {0}({1})",
                                  pTable,
                                  Function(attributes As IList(Of String)) As String
                                    Dim result As String = String.Empty
                                    For Each attribute As String In attributes
                                      result &= String.Format("{0}{1}", If(String.IsNullOrWhiteSpace(result), String.Empty, ", "), CaseInsensitveForm(vColumnMap(attribute).Key, vColumnMap(attribute).Value))
                                    Next attribute
                                    Return result
                                  End Function(pAttributes)))
        vResult.Add(vSQL.ToString)
        Return vResult.AsReadOnly
      End Using
    End Function

    Private Function CaseInsensitveForm(pAttribute As String, pDataType As String) As String
      Return If(pDataType.IndexOf("char", StringComparison.InvariantCultureIgnoreCase) >= 0,
                String.Format("NLSSORT(""{0}"", 'nls_sort=''BINARY_CI''')", pAttribute),
                String.Format("""{0}""", pAttribute))
    End Function

    Public Overrides Function IndexExists(pTableName As String, pAttributes As IList(Of String)) As Boolean
      Return Not String.IsNullOrWhiteSpace(GetExistingIndexName(pTableName, pAttributes))
    End Function

    Public Overrides Function IndexIsUnique(pTableName As String, pAttributes As IList(Of String)) As Boolean
      Dim vIndexName As String = GetExistingIndexName(pTableName, pAttributes)
      If String.IsNullOrWhiteSpace(vIndexName) Then
        Throw New InvalidOperationException("Index not found")
      End If
      Dim vSql As New StringBuilder
      vSql.AppendLine("SELECT uniqueness ")
      vSql.AppendLine("FROM   all_indexes ")
      vSql.AppendLine("WHERE  (index_name = :IndexName")
      vSql.AppendLine("        OR index_name = :UniqueIndexName)")
      vSql.AppendLine("       AND uniqueness = 'UNIQUE'")
      Using vCommand As DbCommand = Me.CreateCommand
        vCommand.CommandType = CommandType.Text
        vCommand.CommandText = vSql.ToString
        DirectCast(vCommand, OracleCommand).Parameters.Add(":IndexName", OracleType.VarChar, 128).Value = vIndexName
        DirectCast(vCommand, OracleCommand).Parameters.Add(":UniqueIndexName", OracleType.VarChar, 128).Value = GetUniqueVarientIndexName(vIndexName)
        Using vData As New DataTable
          Using vReader As IDataReader = vCommand.ExecuteReader()
            vData.Load(vReader)
          End Using
          Return Not (vData.Rows.Count = 0)
        End Using
      End Using
    End Function

    Private Function GetExistingIndexName(pTableName As String, pAttributes As IList(Of String)) As String
      Dim vResult As String = String.Empty
      For Each vIndexName As String In (From indexRow As DataRow In GetIndexNames(pTableName).AsEnumerable
                                        Select indexRow.Field(Of String)("INDEX_NAME"))
        Dim vIndexColumns As New List(Of String)(From indexRow As DataRow In GetIndexColumns(pTableName, vIndexName).AsEnumerable
                                                 Select indexRow.Field(Of String)("COLUMN_NAME"))
        Dim vIsRequiredIndex As Boolean = True
        For Each vAttribute As String In pAttributes
          If Not vIndexColumns.Contains(vAttribute, StringComparer.InvariantCultureIgnoreCase) Then
            vIsRequiredIndex = False
          End If
        Next vAttribute
        For Each vAttribute As String In vIndexColumns
          If Not pAttributes.Contains(vAttribute, StringComparer.InvariantCultureIgnoreCase) Then
            vIsRequiredIndex = False
          End If
        Next vAttribute
        If vIsRequiredIndex Then
          vResult = vIndexName
        End If
      Next vIndexName
      Return vResult
    End Function

#End Region

    Public Overrides Function GetRecordSetAnsiJoins(ByVal pSQL As String, ByVal pTimeout As Integer, ByVal pOptions As RecordSetOptions) As CDBRecordSet
      Return GetRecordSet(pSQL, pTimeout, pOptions)
    End Function

    Public Overrides Function GetDataSet(ByVal pSQL As SQLStatement) As DataSet
      Debug.Print(pSQL.SQL)
      If mvSQLLoggingMode <> SQLLoggingModes.None Then
        If (mvSQLLoggingMode And SQLLoggingModes.[Select]) = SQLLoggingModes.[Select] Then LogSQL(pSQL.SQL)
      End If
      Dim vDS As New DataSet
      Using vDA As New OracleDataAdapter(pSQL.SQL, DirectCast(mvConnection, OracleConnection))
        If mvInTransaction Then vDA.SelectCommand.Transaction = DirectCast(mvTransaction, OracleTransaction)
        vDA.Fill(vDS)
      End Using
      Return vDS
    End Function

    Protected Overrides Function GetRecordSet(ByVal pSQL As String, ByVal pTimeout As Integer, ByVal pOptions As RecordSetOptions) As CDBRecordSet
      Try
        If pSQL.EndsWith(";") Then pSQL = pSQL.TrimEnd(";"c)
        Using vCommand As DbCommand = Me.CreateCommand()
          vCommand.CommandText = pSQL
          vCommand.CommandType = CommandType.Text
          vCommand.CommandTimeout = pTimeout
          LogRecordSetOpened(pSQL)
          Dim vReader As IDataReader = vCommand.ExecuteReader
          Dim vRecordSet As CDBRecordSet = New CDBOracleRecordSet(CType(vReader, OracleDataReader), Me, vCommand.Connection)
          mvRecordSets.Add(vRecordSet)
          Return vRecordSet
        End Using
      Catch vSQLEx As OracleException
        If vSQLEx.Message.Contains("timeout") Then
          RaiseError(DataAccessErrors.daeProcessTimeout)
        Else
          RaiseError(DataAccessErrors.daeDataAccessError, vSQLEx.Message, pSQL)
        End If
      Catch vEx As Exception
        RaiseError(DataAccessErrors.daeDataAccessError, vEx.Message, pSQL)
      End Try
      Return Nothing
    End Function

    Public Overrides Function IsCaseSensitive() As Boolean
      Return True
    End Function

    Public Overrides Function SupportsNoLock() As Boolean
      Return False
    End Function

    Public Overrides Function IsSpecialColumn(ByVal pName As String) As Boolean
      Select Case pName
        Case "current", "number"
          Return True
        Case Else
          Return False
      End Select
    End Function

    Public Overrides ReadOnly Property RowRestrictionType() As CDBConnection.RowRestrictionTypes
      Get
        Return RowRestrictionTypes.UseRownum
      End Get
    End Property

    Public Overrides Sub AppendDateTime(ByVal pSQL As StringBuilder, ByVal pValue As String)
      With (pSQL)
        If pValue.Length >= 8 Then
          Dim vDate As Date = CDate(pValue)
          'to_date( '04 mar 2008 16:10:00', 'DD MON YYYY HH24:MI:SS' )
          .Append("to_date( '")
          .Append(vDate.Day.ToString("00"))
          .Append(" ")
          .Append(GetMonthAbbreviation(vDate))
          .Append(" ")
          .Append(vDate.Year)
          .Append(" ")
          .Append(vDate.Hour.ToString("00"))
          .Append(":")
          .Append(vDate.Minute.ToString("00"))
          .Append(":")
          .Append(vDate.Second.ToString("00"))
          .Append("' , 'DD MON YYYY HH24:MI:SS' )")
        Else
          .Append("'")
          .Append(pValue)
          .Append("'")
        End If
      End With
    End Sub

    Private Enum SQLItemTypes
      sitNone
      sitOpenParen
      sitCloseParen
      sitSubSelect
      sitLeftOuterJoin
      sitRightOuterJoin
      sitInnerJoin
      sitOn
      sitWhere
      sitAnd
      sitOrderBy
      sitGroupBy
      sitUnion
    End Enum

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
      Return String.Format(" SUBSTR({0},1,{1}) ", pExpression, pLength)
    End Function
    Public Overrides Function DBIndexOf(ByVal pSearchString As String, ByVal pExpression As String) As String
      Return String.Format(" INSTR({0},{1}) ", pExpression, pSearchString)
    End Function
    Public Overrides Function DBSubString(ByVal pExpression As String, ByVal pStart As String, ByVal pLength As String) As String
      Return String.Format("SUBSTR({0},{1},{2})", pExpression, pStart, pLength)
    End Function
    Public Overrides Function DBCollateString() As String
      Return ""
    End Function
    Public Overrides Function DBConcatString() As String
      Return " || "
    End Function

    Public Overrides Function DBAddMonths(ByVal pDateAttribute As String, ByVal pMonthsAttribute As String) As String
      Return String.Format(" ADD_MONTHS({0},{1}) ", pDateAttribute, pMonthsAttribute)
    End Function

    Public Overrides Function DBAddYears(ByVal pDateAttribute As String, ByVal pYearsAttribute As String) As String
      Return String.Format(" ADD_MONTHS({0},{1} * 12) ", pDateAttribute, pYearsAttribute)
    End Function

    Public Overrides Function DBAddWeeks(ByVal pDateAttribute As String, ByVal pWeeksAttribute As String) As String
      Return String.Format(" {0} + ({1} * 7) ", pDateAttribute, pWeeksAttribute)
    End Function

    Public Overrides Function DBDate() As String
      Return " SYSDATE "
    End Function

    Public Overrides Function DBAge() As String
      Return " FLOOR(MONTHS_BETWEEN(SYSDATE, date_of_birth)/12) "
    End Function

    Public Overrides Function DBHint(ByVal pHintType As DatabaseHintTypes, ByVal pTableName As String, ByVal pUseHint As Boolean) As String
      Select Case pHintType
        Case DatabaseHintTypes.dhtFullTableScanOracle8
          Debug.Assert(pTableName.Length > 0)
          If pTableName.Length > 0 And OracleVersion() <= 8 Then      'Oracle 8 or less
            Return " /*+ FULL(" & pTableName & ") */ "
          End If
        Case DatabaseHintTypes.dhtUseHashOracle9
          Debug.Assert(pTableName.Length > 0)
          If pTableName.Length > 0 And OracleVersion() >= 9 And pUseHint Then       'Oracle 9 or greater
            Return " /*+ USE_HASH(" & pTableName & ") */ "
          End If
        Case DatabaseHintTypes.dhtNoMergeOracle9
          Debug.Assert(pTableName.Length > 0)
          If pTableName.Length > 0 And OracleVersion() >= 9 And pUseHint Then       'Oracle 9 or greater
            Return " /*+ NO_MERGE(" & pTableName & ") */ "
          End If
      End Select
      Return ""
    End Function

    Public Overrides Function DBReplaceLineFeedWithSpace(ByVal pFieldName As String) As String
      Return String.Format("replace({0},chr(10),' ')", pFieldName)
    End Function

    Public Overrides Function DBMonthDiff(ByVal pEarlierDate As String, ByVal pLaterDate As String) As String
      Return String.Format(" months_between({0},{1}) ", pLaterDate, pEarlierDate)
    End Function

    Public Overrides Function DBToNumber(ByVal pExpression As String) As String
      Return String.Format("TO_NUMBER({0})", pExpression)
    End Function

    Public Overrides Function DBToString(ByVal pExpression As String) As String
      Return String.Format(" TO_CHAR({0}) ", pExpression)
    End Function

    Public Overrides Function DBToDate(pExpression As String) As String
      Return String.Format(" TO_DATE({0}, '{1}') ", pExpression, CAREDateFormat)
    End Function
    Public Overrides Function DBDateTimeAttribToDate(pDateTimeAttributeName As String) As String
      Return String.Format(" Trunc({0})", pDateTimeAttributeName)
    End Function
    Public Overrides Function DBMaxToString(ByVal pExpression As String) As String
      Return String.Format(" '' ", pExpression)
    End Function

    Public Overrides Function DBIsNull(ByVal pExpression As String, ByVal pReplacement As String) As String
      Return String.Format(" NVL({0},{1}) ", pExpression, pReplacement)
    End Function

    Public Overrides Function DBLength(ByVal pExpression As String) As String
      Return String.Format(" LENGTH({0}) ", pExpression)
    End Function

    Public Overrides Function DBLPad(ByVal pExpression As String, ByVal pLength As Integer) As String
      Return String.Format(" LPAD({0},{1},'0') ", pExpression, pLength)
    End Function

    Public Overrides Function DBForceOrder() As String
      Return ""
    End Function

    Public Overrides Function DBYear(ByVal pDateString As String) As String
      'Expects a date (could be a date attribute name) and returns EXTRACT(YEAR FROM '1998-03-07')
      Dim vDate As Date = Nothing
      If Date.TryParse(pDateString, vDate) Then
        'We have a date so put in correct format
        pDateString = "'" & vDate.Year & "-" & vDate.Month & "-" & vDate.Day & "'"
      End If
      Return String.Format("EXTRACT (YEAR FROM {0})", pDateString)
    End Function

    Public Overrides Function DBMonth(ByVal pDateString As String) As String
      'Expects a date (could be a date attribute name) and returns EXTRACT(MONTH FROM '1998-03-07')
      Dim vDate As Date = Nothing
      If Date.TryParse(pDateString, vDate) Then
        'We have a date so put in correct format
        pDateString = "'" & vDate.Year & "-" & vDate.Month & "-" & vDate.Day & "'"
      End If
      Return String.Format("EXTRACT (MONTH FROM {0})", pDateString)
    End Function

    Public Overrides Function DBDay(ByVal pDateString As String) As String
      'Expects a date (could be a date attribute name) and returns EXTRACT(DAY FROM '1998-03-07')
      Dim vDate As Date = Nothing
      If Date.TryParse(pDateString, vDate) Then
        'We have a date so put in correct format
        pDateString = "'" & vDate.Year & "-" & vDate.Month & "-" & vDate.Day & "'"
      End If
      Return String.Format("EXTRACT (DAY FROM {0})", pDateString)
    End Function

    Public Overrides Function SupportsOption(ByVal pOption As DatabaseOptions) As Boolean
      Select Case pOption
        Case DatabaseOptions.ForXML, DatabaseOptions.ForXML_XSINIL
          Return False
        Case Else
          Return False
      End Select
    End Function

    Protected Overrides Function TableOwnerPrefix() As String
      Return ""
    End Function

    Public Overrides Function NativeDataType(ByVal pField As CDBField) As String
      Select Case pField.FieldType
        Case CDBField.FieldTypes.cftCharacter
          Return "varchar2(" & pField.Value & ")"
        Case CDBField.FieldTypes.cftInteger, CDBField.FieldTypes.cftLong
          Return "integer"
        Case CDBField.FieldTypes.cftNumeric
          Return "number(" & pField.Value & ",2)"
        Case CDBField.FieldTypes.cftMemo
          Return "clob"
        Case CDBField.FieldTypes.cftDate, CDBField.FieldTypes.cftTime
          Return "date"
        Case CDBField.FieldTypes.cftBulk
          Return "blob"
        Case CDBField.FieldTypes.cftUnicode
          Return "varchar2(" & pField.Value & ")"
        Case CDBField.FieldTypes.cftBinary
          Return "blob"
        Case CDBField.FieldTypes.cftGUID
          Return "raw(16)"
        Case Else
          Return "UNKNOWN DATA TYPE"
      End Select
    End Function

    Friend Overrides Function GetDBParameter(ByVal pFieldName As String, ByVal pFieldValue As String, ByRef pParamName As String) As System.Data.Common.DbParameter
      pParamName = ":" & pFieldName.Replace("_", "")
      Dim vDBParameter As New OracleParameter(pParamName, pFieldValue.Replace(vbCr, ""))
      Return vDBParameter
    End Function

    Friend Overrides Function GetDBBulkParameter(ByVal pFieldName As String, ByVal pFieldValue As String, ByRef pParamName As String) As System.Data.Common.DbParameter
      pParamName = ":" & pFieldName.Replace("_", "")
      Dim vValue As String = pFieldValue.Replace(vbCr, "")
      Dim vEncoding As New System.Text.ASCIIEncoding
      Dim vBytes() As Byte = vEncoding.GetBytes(vValue)
      Dim vDBParameter As New OracleParameter(pParamName, OracleType.LongRaw, vBytes.Length, ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, vBytes)
      Return vDBParameter
    End Function

    Friend Overrides Function GetDBParameterFromFile(ByVal pFieldName As String, ByVal pFilename As String, ByRef pParamName As String) As System.Data.Common.DbParameter
      pParamName = ":" & pFieldName.Replace("_", "")
      Dim vFS As New IO.FileStream(pFilename, IO.FileMode.Open)
      Dim vBuffer(CInt(vFS.Length - 1)) As Byte
      vFS.Read(vBuffer, 0, CInt(vFS.Length))
      vFS.Close()
      Dim vDBParameter As New OracleParameter(pParamName, OracleType.LongRaw, vBuffer.Length, ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, vBuffer)
      Return vDBParameter
    End Function

    Public Overrides Sub DropTable(ByVal pTableName As String)
      ExecuteSQL("DROP TABLE " & pTableName, cdbExecuteConstants.sqlIgnoreError)
    End Sub

    Public Overrides Function DeleteAllRecords(ByVal pTable As String) As Integer
      Dim vSQL As String = "TRUNCATE TABLE " & GetIndexSchemaPrefix(pTable, "") & pTable
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

    Private Function GetIndexSchemaPrefix(ByVal pTable As String, ByVal pAttr As String) As String
      Dim vSchema As String = ""
      Dim vRestriction(2) As String
      vRestriction(1) = pTable.ToUpper
      If pAttr.Length > 0 Then vRestriction(2) = pAttr.ToUpper
      Dim vTable As DataTable = mvConnection.GetSchema("Columns", vRestriction)
      If vTable.Rows.Count > 0 Then
        vSchema = vTable.Rows(0).Item("OWNER").ToString()
      End If
      If vSchema.Length > 0 Then vSchema = vSchema & "."
      Return vSchema
    End Function

    Public Overrides Function InsertRecord(ByVal pTableName As String, ByVal pFields As CDBFields, ByVal pIgnoreDuplicates As Boolean) As Boolean
      Try
        Return InsertRecord(pTableName, pFields)
      Catch ex As OracleClient.OracleException
        If ex.Message.Contains("unique constraint") AndAlso pIgnoreDuplicates Then
          mvLastInsertErrorIsDuplicate = True  'Ignore duplicate
        Else
          Throw ex
        End If
      End Try
    End Function

    Public Overrides Function UpdateRecords(ByVal pTableName As String, ByVal pUpdateFields As CDBFields, ByVal pWhereFields As CDBFields, Optional ByVal pErrorIfNoRecords As Boolean = True) As Integer
      Try
        Return DoUpdateRecords(pTableName, pUpdateFields, pWhereFields, pErrorIfNoRecords)
      Catch ex As OracleClient.OracleException
        If ex.Message.Contains("unique constraint") AndAlso pErrorIfNoRecords = False Then
          mvLastInsertErrorIsDuplicate = True  'Ignore duplicate
        Else
          Throw ex
        End If
      End Try
    End Function

    Friend Overrides Function PreProcessMaxRows(ByVal pSQLStatement As SQLStatement) As SQLStatement
      'Oracle assigns a rownum before any sorting or aggregation so we need to:
      '1) Remove the MaxRows and build the SQL statement
      '2) Add the MaxRows and build a new SQL statement with the original as a nested select
      'The end result is SQL that look like this:
      'SELECT * FROM (....oiginal SQL statement....) WHERE rownum < x
      Dim vMaxRows As Integer = pSQLStatement.MaxRows
      If vMaxRows > 0 Then
        pSQLStatement.MaxRows = 0
        Dim vSQLStatement As New SQLStatement(Me, "SELECT * FROM ( " & pSQLStatement.SQL & " )")
        vSQLStatement.MaxRows = vMaxRows
        Return vSQLStatement
      Else
        Return pSQLStatement
      End If
    End Function
    Public Overrides Function UseTableSpaces() As Boolean
      Dim vErrorMsg As String
      Try
        If mvTSChecked = False Then
          ExecuteSQL("CREATE TABLE TABLESPACETEST ( TEST_FIELD VARCHAR2(1) ) TABLESPACE LARGE_TABLES", cdbExecuteConstants.sqlShowError)
          ExecuteSQL("DROP TABLE TABLESPACETEST", cdbExecuteConstants.sqlIgnoreError)
          mvTSChecked = True
        End If
      Catch ex As Exception
        vErrorMsg = ex.Message
        If InStr(vErrorMsg, "does not exist") > 0 Then
          mvUseTableSpaces = False
        Else
          mvUseTableSpaces = True
        End If
        mvTSChecked = True
      End Try
      Return mvUseTableSpaces
    End Function

    Public Overrides Function GetDBParameterFromByteArray(pFieldName As String, pValue As Byte(), ByRef pParamName As String) As DbParameter
      Dim vDBParameter As New OracleParameter
      vDBParameter = CType(CreateDBParameterFromByteArray(pFieldName, pValue, pParamName, OracleType.LongRaw), OracleParameter)
      Return vDBParameter
    End Function

    Private Function CreateDBParameterFromByteArray(pFieldName As String, pValue As Byte(), ByRef pParamName As String, ByVal pDataType As OracleType) As DbParameter
      'For password encryption BR19442 a new data type (Binary) was added to code and it is stored as a blob in the database
      'In order for curent GetDBParameterFromByteArray code to work for Oracle database, 2 new functions were required
      pParamName = ":" & pFieldName.Replace("_", "")
      Dim vDBParameter As New OracleParameter(pParamName, pDataType, pValue.Length, ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, pValue)
      Return vDBParameter
    End Function

    Public Overrides Function GetBinaryDBParameter(ByVal pFieldName As String, ByVal pValue As Byte(), ByRef pParamName As String) As DbParameter
      'For password encryption BR19442 a new data type (Binary) was added to code and it is stored as a blob in the database
      'In order for curent GetDBParameterFromByteArray code to work, 2 new functions were required
      Return CreateDBParameterFromByteArray(pFieldName, pValue, pParamName, OracleType.Blob)
    End Function

    Public Overrides Function GetTableConverter(tableName As String,
                                                changedColumns As List(Of TableConverter.ColumnDescriptor),
                                                logfile As StreamWriter) As TableConverter
      Return New OracleTableConverter(tableName, changedColumns, Me, logfile)
    End Function

    Private ReadOnly Property Connection As OracleClient.OracleConnection
      Get
        Return DirectCast(mvConnection, OracleClient.OracleConnection)
      End Get
    End Property

    Public Overrides ReadOnly Property ConcatonateOperator As String
      Get
        Return "||"
      End Get
    End Property

    Public Overrides Function IsBaseTable(schemaName As String, tableName As String) As Boolean
      Dim table As DataTable = mvConnection.GetSchema("Tables", {schemaName, tableName})
      If table.Rows.Count > 1 Then
        Throw New InvalidOperationException("Attempt to find determine the type of a non-unique table")
      End If
      Return table.Rows.Count = 1
    End Function
    Public Overrides Function GetPrimaryKeyColumns(pTableName As String) As DataTable
      Dim vRestriction(2) As String
      vRestriction(1) = pTableName
      Dim vPrimaryKeyIndex As DataTable = mvConnection.GetSchema("PrimaryKeys", vRestriction)
      Dim vPKColumns As New DataTable
      vPKColumns.TableName = "PRIMARY_KEY"
      vPKColumns.Columns.Add("COLUMN_NAME")
      If vPrimaryKeyIndex.Rows.Count > 0 Then
        Dim vColumns As DataTable = GetIndexColumns(pTableName, vPrimaryKeyIndex.Rows(0)("INDEX_NAME").ToString)
        For Each vRow As DataRow In vColumns.Rows
          vPKColumns.Rows.Add(vRow("COLUMN_NAME"))
        Next
      End If
      Return vPKColumns
    End Function
    Public Overrides Function CopyColumnsToNewTable(pSelectionSQL As SQLStatement, pDestinationTableName As String) As Integer
      Dim vSQL As String = pSelectionSQL.SQL
      Return CopyColumnsToNewTable(vSQL, pDestinationTableName)
    End Function
    ''' <summary>
    ''' Use Oracle CREATE TABLE AS to copy data from a SELECT Query into a Table
    ''' </summary>
    ''' <param name="pSelectionSQL"></param>
    ''' <param name="pDestinationTableName"></param>
    ''' <returns></returns>
    ''' <remarks>If you only want to create a table make sure the SELECT statement has a WHERE clause than </remarks>
    Public Overrides Function CopyColumnsToNewTable(pSelectionSQL As String, pDestinationTableName As String) As Integer
      Dim vCopySQL As New StringBuilder
      vCopySQL.Append("CREATE TABLE ")
      vCopySQL.Append(pDestinationTableName)
      vCopySQL.Append(" AS (")
      vCopySQL.Append(pSelectionSQL)
      vCopySQL.Append(")")
      Dim vSQL As String = vCopySQL.ToString()
      Dim vInt As Integer = ExecuteSQL(vSQL)
      Return vInt
    End Function


    ''' <summary>Bulk Update and Insert Data into a Table.</summary>
    ''' <param name="pSQLStatement">The SQLStatement used to populate the DataTable.  Used so that the same SQL is used for the original data and for the changed data.</param>
    ''' <param name="pDataTable">The data to be inserted or updated. This must have the RowState set correctly on each row.</param>
    Public Overrides Sub BulkUpdate(ByVal pSQLStatement As SQLStatement, ByVal pDataTable As DataTable)
      Using vSQLAdapter As OracleDataAdapter = New OracleDataAdapter(pSQLStatement.SQL, Me.Connection)
        If Me.InTransaction Then
          'If we are currently in a transaction we need to set the transaction against the SelectCommand object
          Dim vCmd As OracleCommand = vSQLAdapter.SelectCommand
          If vCmd IsNot Nothing Then
            vCmd.Transaction = CType(Me.Transaction, OracleTransaction)
          End If
        End If
        Using New OracleCommandBuilder(vSQLAdapter)
          'Insert / Update the DataTable into the DataBase
          'The update requires primary key data on the database table and in the DataTable
          vSQLAdapter.Update(pDataTable)
        End Using
      End Using
    End Sub

    Public Overrides Sub PopulateGUIDColumn(pTableName As String, pColumnName As String)
      Dim vSQL As String = String.Format("UPDATE {0} SET {1} = sys_guid()", pTableName, pColumnName)
      Me.ExecuteSQL(vSQL)
    End Sub

    Public Overrides ReadOnly Property MergeTerminator As String
      Get
        Return String.Empty 'The ORACLE command doesn't run Merge statements if they're terminated with a semi-colon as specified by the ANSI standard
      End Get
    End Property
  End Class
End Namespace