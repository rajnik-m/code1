Imports System.IO
Imports System.Text
Imports System.Array
Imports System.Data.SqlTypes

''' <summary>
''' A data reader that is backed by a CSV file.
''' </summary>
''' <remarks>Column names are taken from the mandatory header record.  All column types are 
''' string.  The Microsoft specific extension where dates are surrounded with hash marks ('#' 
''' characters) is not supported.</remarks>
Public Class CsvReader
  Implements IDataReader

  Private mvClosed As Boolean = False
  Private mvFileReader As FileReader = Nothing
  Private mvSchemaTable As New DataTable
  Private mvDataRow As DataRow = Nothing
  Private mvReferenceTable As New DataTable

  ''' <summary>
  ''' Initializes a new instance of the <see cref="CsvReader" /> class.
  ''' </summary>
  ''' <param name="path">The path.</param>
  Public Sub New(path As String)
    mvSchemaTable.Columns.Add("ColumnName", GetType(String))
    mvSchemaTable.Columns.Add("ColumnOrdinal", GetType(Integer))
    mvSchemaTable.Columns.Add("ColumnSize", GetType(Integer))
    mvSchemaTable.Columns.Add("NumericPrecision", GetType(Short))
    mvSchemaTable.Columns.Add("NumericScale", GetType(Short))
    mvSchemaTable.Columns.Add("IsUnique", GetType(Boolean))
    mvSchemaTable.Columns.Add("IsKey", GetType(Boolean))
    mvSchemaTable.Columns.Add("DataType", GetType(Type))
    mvSchemaTable.Columns.Add("AllowDBNull", GetType(Boolean))
    mvSchemaTable.Columns.Add("ProviderType", GetType(Integer))
    mvSchemaTable.Columns.Add("IsIdentity", GetType(Boolean))
    mvSchemaTable.Columns.Add("IsAutoIncrement", GetType(Boolean))
    mvSchemaTable.Columns.Add("IsRowVersion", GetType(Boolean))
    mvSchemaTable.Columns.Add("IsLong", GetType(Boolean))
    mvSchemaTable.Columns.Add("IsReadOnly", GetType(Boolean))
    mvSchemaTable.Columns.Add("ProviderSpecificDataType", GetType(Type))
    mvSchemaTable.Columns.Add("DataTypeName", GetType(String))
    mvSchemaTable.Columns.Add("IsColumnSet", GetType(Boolean))
    mvFileReader = New FileReader(path, FileReader.FileReaderTypes.rftCharSeparated, New Char() {CChar(",")}, False)
    mvFileReader.ReadLine()
    Dim vOrdinal As Integer = 0
    For Each vField As String In mvFileReader.Fields
      Dim vRow As DataRow = mvSchemaTable.NewRow
      vRow.BeginEdit()
      vRow("ColumnName") = vField
      vRow("ColumnOrdinal") = vOrdinal
      vRow("ColumnSize") = 2048
      vRow("NumericPrecision") = 0
      vRow("NumericScale") = 0
      vRow("IsUnique") = False
      vRow("IsKey") = False
      vRow("DataType") = GetType(String)
      vRow("AllowDBNull") = True
      vRow("ProviderType") = SqlDbType.NVarChar
      vRow("IsIdentity") = False
      vRow("IsAutoIncrement") = False
      vRow("IsRowVersion") = False
      vRow("IsLong") = False
      vRow("IsReadOnly") = False
      vRow("ProviderSpecificDataType") = GetType(SqlString)
      vRow("DataTypeName") = "nvarchar"
      vRow("IsColumnSet") = False
      vRow.EndEdit()
      mvSchemaTable.Rows.Add(vRow)
      mvReferenceTable.Columns.Add(vField, GetType(String))
      vOrdinal += 1
    Next vField
  End Sub

  ''' <summary>
  ''' Closes the input file.
  ''' </summary>
  Public Sub Close() Implements System.Data.IDataReader.Close
    mvFileReader.Dispose()
    mvClosed = True
  End Sub

  ''' <summary>
  ''' Gets a value indicating the depth of nesting for the current row.
  ''' </summary>
  ''' <returns>The level of nesting.</returns>
  ''' <remarks>Nesting is not supported and so 0 is always returned.  This is in line with Microsoft recommendations.</remarks>
  Public ReadOnly Property Depth As Integer Implements System.Data.IDataReader.Depth
    Get
      Return 0
    End Get
  End Property

  ''' <summary>
  ''' Returns a <see cref="T:System.Data.DataTable" /> that has the same schema as the CSV file being read <see cref="T:System.Data.IDataReader" />.
  ''' </summary><returns>
  ''' A <see cref="T:System.Data.DataTable" /> that has the same schema as the CSV file.
  ''' </returns>
  ''' <exception cref="T:System.InvalidOperationException">The <see cref="T:System.Data.IDataReader" /> is closed. </exception>
  Public Function GetReferenceTable() As System.Data.DataTable
    If Me.IsClosed Then
      Throw New InvalidOperationException("The reader is closed.")
    End If
    Return mvReferenceTable
  End Function

  ''' <summary>
  ''' Returns a <see cref="T:System.Data.DataTable" /> that describes the column metadata of the <see cref="T:System.Data.IDataReader" />.
  ''' </summary><returns>
  ''' A <see cref="T:System.Data.DataTable" /> that describes the column metadata.
  ''' </returns>
  ''' <exception cref="T:System.InvalidOperationException">The <see cref="T:System.Data.IDataReader" /> is closed. </exception>
  Public Function GetSchemaTable() As System.Data.DataTable Implements System.Data.IDataReader.GetSchemaTable
    If Me.IsClosed Then
      Throw New InvalidOperationException("The reader is closed.")
    End If
    Return mvSchemaTable
  End Function

  ''' <summary>
  ''' Gets a value indicating whether the data reader is closed.
  ''' </summary>
  ''' <returns>true if the data reader is closed; otherwise, false.</returns>
  Public ReadOnly Property IsClosed As Boolean Implements System.Data.IDataReader.IsClosed
    Get
      Return mvClosed
    End Get
  End Property

  ''' <summary>
  ''' Advances the data reader to the next result, when reading the results of batch SQL statements.
  ''' </summary>
  ''' <returns>true if there are more rows; otherwise, false.</returns>
  ''' <remarks>Multiple result sets are not supported and so this always returns false.</remarks>
  Public Function NextResult() As Boolean Implements System.Data.IDataReader.NextResult
    Return False
  End Function

  ''' <summary>
  ''' Advances the <see cref="T:System.Data.IDataReader" /> to the next record.
  ''' </summary>
  ''' <returns> true if there are more rows; otherwise, false.</returns>
  Public Function Read() As Boolean Implements System.Data.IDataReader.Read
    Dim RowRead As Boolean = False
    If Not mvFileReader.EndOfFile Then
      mvDataRow = ReadRecord()
      RowRead = True
    End If
    Return RowRead
  End Function

  ''' <summary>
  ''' Gets the number of rows changed, inserted, or deleted by execution of the SQL statement.
  ''' </summary>
  ''' <returns>The number of rows changed, inserted, or deleted; 0 if no rows were affected or the 
  ''' statement failed; and -1 for SELECT statements.</returns>
  ''' <remarks>Only select statements are suported and so this always returns -1.</remarks>
  Public ReadOnly Property RecordsAffected As Integer Implements System.Data.IDataReader.RecordsAffected
    Get
      Return -1
    End Get
  End Property

  ''' <summary>
  ''' Gets the number of columns in the current row.
  ''' </summary>
  ''' <returns>When not positioned in a valid recordset, 0; otherwise, the number of columns in the current 
  ''' record. The default is -1.</returns>
  Public ReadOnly Property FieldCount As Integer Implements System.Data.IDataRecord.FieldCount
    Get
      Return mvReferenceTable.Columns.Count
    End Get
  End Property

  ''' <summary>
  ''' Gets the value of the specified column as a Boolean.
  ''' </summary>
  ''' <param name="i">The zero-based column ordinal.</param>
  ''' <returns>The value of the column.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetBoolean(i As Integer) As Boolean Implements System.Data.IDataRecord.GetBoolean
    Return CType(mvDataRow(i), Boolean)
  End Function

  ''' <summary>
  ''' Gets the 8-bit unsigned integer value of the specified column.
  ''' </summary>
  ''' <param name="i">The zero-based column ordinal.</param>
  ''' <returns>The 8-bit unsigned integer value of the specified column.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetByte(i As Integer) As Byte Implements System.Data.IDataRecord.GetByte
    Return CType(mvDataRow(i), Byte)
  End Function

  ''' <summary>
  ''' Reads a stream of bytes from the specified column offset into the buffer as an array, starting at the given buffer offset.
  ''' </summary>
  ''' <param name="i">The zero-based column ordinal.</param>
  ''' <param name="fieldOffset">The index within the field from which to start the read operation.</param>
  ''' <param name="buffer">The buffer into which to read the stream of bytes.</param>
  ''' <param name="bufferoffset">The index for <paramref name="buffer" /> to start the read operation.</param>
  ''' <param name="length">The number of bytes to read.</param>
  ''' <returns>The actual number of bytes read.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetBytes(i As Integer, fieldOffset As Long, buffer() As Byte, bufferoffset As Integer, length As Integer) As Long Implements System.Data.IDataRecord.GetBytes
    ConstrainedCopy(Encoding.UTF8.GetBytes(DirectCast(mvDataRow(i), String).ToCharArray), CInt(fieldOffset), buffer, bufferoffset, If(length > buffer.GetLength(0) - bufferoffset, buffer.GetLength(0) - bufferoffset, length))
    Return If(length > buffer.GetLength(0) - bufferoffset, buffer.GetLength(0) - bufferoffset, length)
  End Function

  ''' <summary>
  ''' Gets the character value of the specified column.
  ''' </summary>
  ''' <param name="i">The zero-based column ordinal.</param>
  ''' <returns>The character value of the specified column.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetChar(i As Integer) As Char Implements System.Data.IDataRecord.GetChar
    Return CType(mvDataRow(i), Char)
  End Function

  ''' <summary>
  ''' Reads a stream of characters from the specified column offset into the buffer as an array, starting at the given buffer offset.
  ''' </summary>
  ''' <param name="i">The zero-based column ordinal.</param>
  ''' <param name="fieldoffset">The index within the row from which to start the read operation.</param>
  ''' <param name="buffer">The buffer into which to read the stream of bytes.</param>
  ''' <param name="bufferoffset">The index for <paramref name="buffer" /> to start the read operation.</param>
  ''' <param name="length">The number of bytes to read.</param>
  ''' <returns>The actual number of characters read.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetChars(i As Integer, fieldoffset As Long, buffer() As Char, bufferoffset As Integer, length As Integer) As Long Implements System.Data.IDataRecord.GetChars
    System.Array.ConstrainedCopy(DirectCast(mvDataRow(i), String).ToCharArray, CInt(fieldoffset), buffer, bufferoffset, If(length > buffer.GetLength(0) - bufferoffset, buffer.GetLength(0) - bufferoffset, length))
    Return If(length > buffer.GetLength(0) - bufferoffset, buffer.GetLength(0) - bufferoffset, length)
  End Function

  ''' <summary>
  ''' Returns an <see cref="T:System.Data.IDataReader" /> for the specified column ordinal.
  ''' </summary>
  ''' <param name="i">The index of the field to find.</param>
  ''' <returns>An <see cref="T:System.Data.IDataReader" />.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  ''' <remarks>We do not suport embedded data and so this always returns nothing.</remarks>
  Public Function GetData(i As Integer) As System.Data.IDataReader Implements System.Data.IDataRecord.GetData
    Return Nothing
  End Function

  ''' <summary>
  ''' Gets the data type information for the specified field.
  ''' </summary>
  ''' <param name="i">The index of the field to find.</param>
  ''' <returns>The data type information for the specified field.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetDataTypeName(i As Integer) As String Implements System.Data.IDataRecord.GetDataTypeName
    Return mvDataRow(i).GetType.Name
  End Function

  ''' <summary>
  ''' Gets the date and time data value of the specified field.
  ''' </summary>
  ''' <param name="i">The index of the field to find.</param>
  ''' <returns>The date and time data value of the specified field.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetDateTime(i As Integer) As Date Implements System.Data.IDataRecord.GetDateTime
    Return CType(mvDataRow(i), Date)
  End Function

  ''' <summary>
  ''' Gets the fixed-position numeric value of the specified field.
  ''' </summary>
  ''' <param name="i">The index of the field to find.</param>
  ''' <returns>The fixed-position numeric value of the specified field.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetDecimal(i As Integer) As Decimal Implements System.Data.IDataRecord.GetDecimal
    Return CType(mvDataRow(i), Decimal)
  End Function

  ''' <summary>
  ''' Gets the double-precision floating point number of the specified field.
  ''' </summary>
  ''' <param name="i">The index of the field to find.</param>
  ''' <returns> The double-precision floating point number of the specified field.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetDouble(i As Integer) As Double Implements System.Data.IDataRecord.GetDouble
    Return CType(mvDataRow(i), Double)
  End Function

  ''' <summary>
  ''' Gets the <see cref="T:System.Type" /> information corresponding to the type of <see cref="T:System.Object" /> 
  ''' that would be returned from <see cref="M:System.Data.IDataRecord.GetValue(System.Int32)" />.
  ''' </summary>
  ''' <param name="i">The index of the field to find.</param>
  ''' <returns>The <see cref="T:System.Type" /> information corresponding to the type of <see cref="T:System.Object" /> 
  ''' that would be returned from <see cref="M:System.Data.IDataRecord.GetValue(System.Int32)" />.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetFieldType(i As Integer) As System.Type Implements System.Data.IDataRecord.GetFieldType
    Return mvDataRow(i).GetType
  End Function

  ''' <summary>
  ''' Gets the single-precision floating point number of the specified field.
  ''' </summary>
  ''' <param name="i">The index of the field to find.</param>
  ''' <returns>The single-precision floating point number of the specified field.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetFloat(i As Integer) As Single Implements System.Data.IDataRecord.GetFloat
    Return CType(mvDataRow(i), Single)
  End Function

  ''' <summary>
  ''' Returns the GUID value of the specified field.
  ''' </summary>
  ''' <param name="i">The index of the field to find.</param><returns>
  ''' The GUID value of the specified field.
  ''' </returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetGuid(i As Integer) As System.Guid Implements System.Data.IDataRecord.GetGuid
    Return CType(mvDataRow(i), Guid)
  End Function

  ''' <summary>
  ''' Gets the 16-bit signed integer value of the specified field.
  ''' </summary>
  ''' <param name="i">The index of the field to find.</param>
  ''' <returns>The 16-bit signed integer value of the specified field.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetInt16(i As Integer) As Short Implements System.Data.IDataRecord.GetInt16
    Return CType(mvDataRow(i), Int16)
  End Function

  ''' <summary>
  ''' Gets the 32-bit signed integer value of the specified field.
  ''' </summary>
  ''' <param name="i">The index of the field to find.</param>
  ''' <returns>The 32-bit signed integer value of the specified field.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetInt32(i As Integer) As Integer Implements System.Data.IDataRecord.GetInt32
    Return CType(mvDataRow(i), Int32)
  End Function

  ''' <summary>
  ''' Gets the 64-bit signed integer value of the specified field.
  ''' </summary>
  ''' <param name="i">The index of the field to find.</param>
  ''' <returns>The 64-bit signed integer value of the specified field.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetInt64(i As Integer) As Long Implements System.Data.IDataRecord.GetInt64
    Return CType(mvDataRow(i), Int64)
  End Function

  ''' <summary>
  ''' Gets the name for the field to find.
  ''' </summary>
  ''' <param name="i">The index of the field to find.</param>
  ''' <returns>The name of the field or the empty string (""), if there is no value to return.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetName(i As Integer) As String Implements System.Data.IDataRecord.GetName
    Return If(i < mvReferenceTable.Columns.Count, mvReferenceTable.Columns(i).ColumnName, "")
  End Function

  ''' <summary>
  ''' Return the index of the named field.
  ''' </summary>
  ''' <param name="name">The name of the field to find.</param>
  ''' <returns>The index of the named field.</returns>
  Public Function GetOrdinal(name As String) As Integer Implements System.Data.IDataRecord.GetOrdinal
    Return mvReferenceTable.Columns(name).Ordinal
  End Function

  ''' <summary>
  ''' Gets the string value of the specified field.
  ''' </summary>
  ''' <param name="i">The index of the field to find.</param>
  ''' <returns>The string value of the specified field.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetString(i As Integer) As String Implements System.Data.IDataRecord.GetString
    Return CType(mvDataRow(i), String)
  End Function

  ''' <summary>
  ''' Return the value of the specified field.
  ''' </summary>
  ''' <param name="i">The index of the field to find.</param>
  ''' <returns>The <see cref="T:System.Object" /> which will contain the field value upon return.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function GetValue(i As Integer) As Object Implements System.Data.IDataRecord.GetValue
    Return CType(mvDataRow(i), Object)
  End Function

  ''' <summary>
  ''' Populates an array of objects with the column values of the current record.
  ''' </summary>
  ''' <param name="values">An array of <see cref="T:System.Object" /> to copy the attribute fields into.</param>
  ''' <returns>The number of instances of <see cref="T:System.Object" /> in the array.</returns>
  Public Function GetValues(values() As Object) As Integer Implements System.Data.IDataRecord.GetValues
    mvDataRow.ItemArray.CopyTo(values, 0)
    Return mvDataRow.ItemArray.GetLength(0)
  End Function

  ''' <summary>
  ''' Return whether the specified field is set to null.
  ''' </summary>
  ''' <param name="i">The index of the field to find.</param>
  ''' <returns>true if the specified field is set to null; otherwise, false.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Public Function IsDBNull(i As Integer) As Boolean Implements System.Data.IDataRecord.IsDBNull
    Return False
  End Function

  ''' <summary>
  ''' Gets the column located at the specified index.
  ''' </summary>
  ''' <returns>The column located at the specified index as an <see cref="T:System.Object" />.</returns>
  '''   <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Default Public Overloads ReadOnly Property Item(i As Integer) As Object Implements System.Data.IDataRecord.Item
    Get
      Return mvDataRow(i)
    End Get
  End Property

  ''' <summary>
  ''' Gets the column located at the specified index.
  ''' </summary>
  ''' <returns>The column located at the specified index as an <see cref="T:System.Object" />.</returns>
  ''' <exception cref="T:System.IndexOutOfRangeException">The index passed was outside the range of 0 through <see cref="P:System.Data.IDataRecord.FieldCount" />. </exception>
  Default Public Overloads ReadOnly Property Item(name As String) As Object Implements System.Data.IDataRecord.Item
    Get
      Return mvDataRow(name)
    End Get
  End Property

  ''' <summary>
  ''' Reads the record.
  ''' </summary>
  ''' <returns>a single <see cref="DataRow" /> representing the next data record in the file.</returns>
  Private Function ReadRecord() As DataRow
    Dim vRecord As DataRow = mvReferenceTable.NewRow
    mvFileReader.ReadLine()
    For index As Integer = 0 To mvFileReader.FieldCount - 1
      vRecord(index) = mvFileReader.Item(index)
    Next index
    Return vRecord
  End Function

#Region "IDisposable Support"
  Private disposedValue As Boolean

  ''' <summary>
  ''' Releases unmanaged and - optionally - managed resources
  ''' </summary>
  ''' <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> 
  ''' to release only unmanaged resources.</param>
  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not Me.disposedValue Then
      If disposing Then
        mvFileReader.Dispose()
        mvSchemaTable.Dispose()
        mvReferenceTable.Dispose()
      End If
    End If
    Me.disposedValue = True
  End Sub

  ''' <summary>
  ''' Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
  ''' </summary>
  Public Sub Dispose() Implements IDisposable.Dispose
    Dispose(True)
    GC.SuppressFinalize(Me)
  End Sub
#End Region

End Class