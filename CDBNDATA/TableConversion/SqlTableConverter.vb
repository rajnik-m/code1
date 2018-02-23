Imports System.Data.Common
Imports CARE.Data
Imports System.IO

Public Class SqlServerTableConverter
  Inherits TableConverter

  Private Const TEMPORARY_TABLE_NAME As String = "__temporary_conversion_table__"
  Private Const TEMPORARY_REPLACED_TABLE_NAME As String = "__temporary_replaced_table__"

  Private Shared SQL As New StringBuilder

  Private indexesField As New Dictionary(Of String, List(Of String))
  
  Shared Sub New()
    SQL.AppendLine("select cols.COLUMN_NAME, ")
    SQL.AppendLine("       cols.DATA_TYPE, ")
    SQL.AppendLine("       cols.CHARACTER_MAXIMUM_LENGTH, ")
    SQL.AppendLine("       cols.NUMERIC_PRECISION, ")
    SQL.AppendLine("       cols.NUMERIC_SCALE, ")
    SQL.AppendLine("       cols.DATETIME_PRECISION, ")
    SQL.AppendLine("       cols.IS_NULLABLE ")
    SQL.AppendLine("from   INFORMATION_SCHEMA.COLUMNS cols ")
    SQL.AppendLine("where  cols.TABLE_NAME = @tableName ")
    SQL.AppendLine("       and cols.TABLE_SCHEMA = 'dbo' ")
    SQL.AppendLine("order by ORDINAL_POSITION")
  End Sub

  Protected Friend Sub New(tableName As String, changedColumns As List(Of ColumnDescriptor), connection As CDBConnection, logfile As StreamWriter)
    MyBase.New(tableName, changedColumns, connection, logfile)
  End Sub

  Protected Overrides Function GetExistingTableSchema() As IList(Of TableConverter.ColumnDescriptor)
    Dim result As New List(Of ColumnDescriptor)
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
          result.Add(New ColumnDescriptor(schemaRow.Field(Of String)("COLUMN_NAME"),
                                          formatDataType(schemaRow),
                                          schemaRow.Field(Of String)("IS_NULLABLE").Equals("YES", StringComparison.InvariantCultureIgnoreCase)))
        Next schemaRow
      End Using
    End Using
    Return result
  End Function

  Private Function formatDataType(schemaData As DataRow) As String
    Dim result As String = String.Empty
    If schemaData.Table.Columns.Contains("DATA_TYPE") Then
      If schemaData.Field(Of String)("DATA_TYPE").Equals("char", StringComparison.InvariantCultureIgnoreCase) OrElse
         schemaData.Field(Of String)("DATA_TYPE").Equals("nchar", StringComparison.InvariantCultureIgnoreCase) OrElse
         schemaData.Field(Of String)("DATA_TYPE").Equals("varchar", StringComparison.InvariantCultureIgnoreCase) OrElse
         schemaData.Field(Of String)("DATA_TYPE").Equals("nvarchar", StringComparison.InvariantCultureIgnoreCase) OrElse
         schemaData.Field(Of String)("DATA_TYPE").Equals("varbinary", StringComparison.InvariantCultureIgnoreCase) OrElse
         schemaData.Field(Of String)("DATA_TYPE").Equals("binary", StringComparison.InvariantCultureIgnoreCase) Then
        result = String.Format("{0}({1})", schemaData("DATA_TYPE"), If(CInt(schemaData("CHARACTER_MAXIMUM_LENGTH")) > 0,
                                                                      CInt(schemaData("CHARACTER_MAXIMUM_LENGTH")).ToString,
                                                                      "max"))
      ElseIf schemaData.Field(Of String)("DATA_TYPE").Equals("bigint", StringComparison.InvariantCultureIgnoreCase) OrElse
             schemaData.Field(Of String)("DATA_TYPE").Equals("Int", StringComparison.InvariantCultureIgnoreCase) OrElse
             schemaData.Field(Of String)("DATA_TYPE").Equals("smallint", StringComparison.InvariantCultureIgnoreCase) OrElse
             schemaData.Field(Of String)("DATA_TYPE").Equals("money", StringComparison.InvariantCultureIgnoreCase) OrElse
             schemaData.Field(Of String)("DATA_TYPE").Equals("date", StringComparison.InvariantCultureIgnoreCase) OrElse
             schemaData.Field(Of String)("DATA_TYPE").Equals("DateTime", StringComparison.InvariantCultureIgnoreCase) OrElse
             schemaData.Field(Of String)("DATA_TYPE").Equals("image", StringComparison.InvariantCultureIgnoreCase) OrElse
             schemaData.Field(Of String)("DATA_TYPE").Equals("Text", StringComparison.InvariantCultureIgnoreCase) OrElse
             schemaData.Field(Of String)("DATA_TYPE").Equals("ntext", StringComparison.InvariantCultureIgnoreCase) Then
        result = String.Format("{0}", schemaData("DATA_TYPE"))
      ElseIf schemaData.Field(Of String)("DATA_TYPE").Equals("decimal", StringComparison.InvariantCultureIgnoreCase) OrElse
             schemaData.Field(Of String)("DATA_TYPE").Equals("numeric", StringComparison.InvariantCultureIgnoreCase) OrElse
             schemaData.Field(Of String)("DATA_TYPE").Equals("float", StringComparison.InvariantCultureIgnoreCase) Then
        result = String.Format("{0}({1},{2})", schemaData("DATA_TYPE"),
                                               CInt(schemaData("NUMERIC_PRECISION")),
                                               CInt(schemaData("NUMERIC_SCALE")))
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
      sql.AppendFormat(vbCrLf & "    [{0}] {1} {2}null, ",
                       column.ColumnName,
                       column.DataType,
                       If(column.IsNullable, String.Empty, "not "))
    Next column
    sql.Length = sql.Length - 2
    sql.Append(vbCrLf & "  )")
    Using statement As DbCommand = Me.Connection.CreateCommand
      statement.CommandText = sql.ToString
      statement.ExecuteNonQuery()
    End Using
  End Sub

  Private Function CopyDataToTemporaryDestinationTable() As Integer
    Using statement As DbCommand = Me.Connection.CreateCommand
      statement.CommandText = String.Format("insert into {0} " & vbCrLf & "select * from {1}", TEMPORARY_TABLE_NAME, Me.TableName)
      statement.CommandTimeout = 0
      Return statement.ExecuteNonQuery()
    End Using
  End Function

  Private Sub RenameSourceToTemporary()
    RenameTable(Me.TableName, TEMPORARY_REPLACED_TABLE_NAME)
  End Sub

  Private Sub RenameTemporaryDestinationAsSource()
    RenameTable(TEMPORARY_TABLE_NAME, Me.TableName.ToLower)
  End Sub

  Private Sub RenameTable(oldName As String, newName As String)
    Using statement As DbCommand = Me.Connection.CreateCommand
      statement.CommandText = String.Format("exec sp_rename '{0}', '{1}'", oldName, newName)
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
