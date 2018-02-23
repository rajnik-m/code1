Imports System.Data.Common
Imports CARE.Data
Imports System.IO

Public Class OracleTableConverter
  Inherits TableConverter

  Private Const TEMPORARY_TABLE_NAME As String = "x_temporary_conversion_table_x"
  Private Const TEMPORARY_REPLACED_TABLE_NAME As String = "x_temporary_replaced_table_x"

  Private Shared SQL As New StringBuilder

  Private indexesField As New Dictionary(Of String, List(Of String))
  Private originalSchemaField As List(Of ColumnDescriptor) = Nothing

  Shared Sub New()
    SQL.AppendLine("select cols.COLUMN_NAME, ")
    SQL.AppendLine("       cols.DATA_TYPE, ")
    SQL.AppendLine("       cols.DATA_LENGTH, ")
    SQL.AppendLine("       cols.DATA_PRECISION, ")
    SQL.AppendLine("       cols.DATA_SCALE, ")
    SQL.AppendLine("       cols.NULLABLE ")
    SQL.AppendLine("from   SYS.ALL_TAB_COLUMNS cols ")
    SQL.AppendLine("where  cols.TABLE_NAME = :tableName ")
    SQL.AppendLine("order by COLUMN_ID")
  End Sub

  Protected Friend Sub New(tableName As String, changedColumns As List(Of ColumnDescriptor), connection As CDBConnection, logfile As StreamWriter)
    MyBase.New(tableName, changedColumns, connection, logfile)
  End Sub

  Protected Overrides Function GetExistingTableSchema() As IList(Of TableConverter.ColumnDescriptor)
    If originalSchemaField Is Nothing Then
      originalSchemaField = New List(Of ColumnDescriptor)
      Using statement As DbCommand = Me.Connection.CreateCommand
        statement.CommandText = SQL.ToString
        Dim parameter As DbParameter = statement.CreateParameter()
        parameter.ParameterName = "TableName"
        parameter.Value = Me.TableName
        statement.Parameters.Add(parameter)
        Using schemaData As New DataTable
          Using reader As IDataReader = statement.ExecuteReader
            schemaData.Load(reader)
          End Using
          For Each schemaRow As DataRow In schemaData.AsEnumerable
            originalSchemaField.Add(New ColumnDescriptor(schemaRow.Field(Of String)("COLUMN_NAME"),
                                                         formatDataType(schemaRow),
                                                         schemaRow.Field(Of String)("NULLABLE").Equals("Y", StringComparison.InvariantCultureIgnoreCase)))
          Next schemaRow
        End Using
      End Using
    End If
    Return originalSchemaField
  End Function

  Private Function formatDataType(schemaData As DataRow) As String
    Dim result As String = String.Empty
    If schemaData.Table.Columns.Contains("DATA_TYPE") Then
      If schemaData.Field(Of String)("DATA_TYPE").Equals("varchar2", StringComparison.InvariantCultureIgnoreCase) Then
        result = String.Format("{0}({1})", schemaData("DATA_TYPE"),
                                           schemaData("DATA_LENGTH").ToString)
      ElseIf schemaData.Field(Of String)("DATA_TYPE").Equals("number", StringComparison.InvariantCultureIgnoreCase) AndAlso
             Not IsDBNull(schemaData("DATA_PRECISION")) Then
        result = String.Format("{0}({1},{2})", schemaData("DATA_TYPE"),
                                               CInt(schemaData("DATA_PRECISION")),
                                               CInt(schemaData("DATA_SCALE")))
      ElseIf schemaData.Field(Of String)("DATA_TYPE").Equals("number", StringComparison.InvariantCultureIgnoreCase) AndAlso
              IsDBNull(schemaData("DATA_PRECISION")) Then
        result = "integer"
      ElseIf schemaData.Field(Of String)("DATA_TYPE").Equals("date", StringComparison.InvariantCultureIgnoreCase) OrElse
             schemaData.Field(Of String)("DATA_TYPE").Equals("long", StringComparison.InvariantCultureIgnoreCase) OrElse
             schemaData.Field(Of String)("DATA_TYPE").Equals("long raw", StringComparison.InvariantCultureIgnoreCase) Then
        result = String.Format("{0}", schemaData("DATA_TYPE"))
      ElseIf schemaData.Field(Of String)("DATA_TYPE").Equals("clob", StringComparison.InvariantCultureIgnoreCase) Then
        result = "varchar(max)"
      ElseIf schemaData.Field(Of String)("DATA_TYPE").Equals("blob", StringComparison.InvariantCultureIgnoreCase) Then
        result = "varbinary(max)"
      Else
        Throw New InvalidOperationException(String.Format("Unexpected data type {0} encountered.", schemaData.Field(Of String)("DATA_TYPE")))
      End If
    Else
      Throw New InvalidOperationException("Schema table does not contain a column named ""DATA_TYPE""")
    End If
    Return result
  End Function

  Protected Overrides Sub DoConvertTable()
    CreateTemporaryDestinationTable()
    Me.RowsConverted = CopyDataToTemporaryDestinationTable()
    RenameSourceToTemporary()
    RenameTemporaryDestinationAsSource()
    DeleteTempolrarySourceTable()
  End Sub

  Private Sub CreateTemporaryDestinationTable()
    Dim sql As New StringBuilder(String.Format("create table {0} " & vbCrLf & "  ( ", TEMPORARY_TABLE_NAME))
    For Each column As ColumnDescriptor In Me.NewSchema
      sql.AppendFormat(vbCrLf & "    ""{0}"" {1} {2}null, ",
                       column.ColumnName,
                       GetOracleDataType(column.DataType),
                       If(column.IsNullable, String.Empty, "not "))
    Next column
    sql.Length = sql.Length - 2
    sql.Append(vbCrLf & "  )")
    Using statement As DbCommand = Me.Connection.CreateCommand
      statement.CommandText = sql.ToString
      statement.ExecuteNonQuery()
    End Using
  End Sub

  Private Function GetOracleDataType(sqlDataType As String) As String
    Dim result As String = sqlDataType
    If sqlDataType.Equals("varchar(max)", StringComparison.InvariantCultureIgnoreCase) Then
      result = "CLOB"
    ElseIf sqlDataType.Equals("varbinary(max)", StringComparison.InvariantCultureIgnoreCase) Then
      result = "BLOB"
    End If
    Return result
  End Function

  Private Function CopyDataToTemporaryDestinationTable() As Integer
    Using statement As DbCommand = Me.Connection.CreateCommand
      statement.CommandText = String.Format("insert into {0} " & vbCrLf & "select{1}from {2}", TEMPORARY_TABLE_NAME, getSelectColumns(), Me.TableName)
      statement.CommandTimeout = 0
      Return statement.ExecuteNonQuery()
    End Using
  End Function

  Private Function getSelectColumns() As String
    Dim result As New StringBuilder
    For Each column As ColumnDescriptor In GetExistingTableSchema()
      result.AppendFormat(If(result.Length > 0,
                             ", " & vbCrLf & "       {0}",
                             " {0}"), If(column.DataType.Equals("long", StringComparison.InvariantCultureIgnoreCase) OrElse
                                         column.DataType.Equals("long raw", StringComparison.InvariantCultureIgnoreCase),
                                         String.Format("TO_LOB({0})", column.ColumnName),
                                         String.Format("""{0}""", column.ColumnName)))
    Next column
    result.AppendLine()
    Return result.ToString
  End Function

  Private Sub RenameSourceToTemporary()
    RenameTable(Me.TableName, TEMPORARY_REPLACED_TABLE_NAME)
  End Sub

  Private Sub RenameTemporaryDestinationAsSource()
    RenameTable(TEMPORARY_TABLE_NAME, Me.TableName)
  End Sub

  Private Sub RenameTable(oldName As String, newName As String)
    Using statement As DbCommand = Me.Connection.CreateCommand
      statement.CommandText = String.Format("rename {0} to {1}", oldName, newName)
      statement.ExecuteNonQuery()
    End Using
  End Sub

  Private Sub DeleteTempolrarySourceTable()
    Using statement As DbCommand = Me.Connection.CreateCommand
      statement.CommandText = String.Format("drop table {0}", TEMPORARY_REPLACED_TABLE_NAME)
      statement.ExecuteNonQuery()
    End Using
  End Sub

End Class
