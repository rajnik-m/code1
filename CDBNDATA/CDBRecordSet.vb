Imports System.Data.Common
Imports System.IO
Imports System.Data.OracleClient

Namespace Data

  Public MustInherit Class CDBRecordSet
    Implements IDisposable

    Protected mvDataTable As DataTable
    Protected mvUseDataTable As Boolean
    Protected mvRowNumber As Integer
    Protected mvDataReader As DbDataReader
    Protected mvFields As CDBFields
    Protected mvConnection As CDBConnection
    Protected mvDBConnection As DbConnection
    Protected mvStatus As Boolean                   'SHOULD ONLY USED BY VB6 MIGRATEDCODE

    'The following flag is set then for oracle if a system.decimal field is found with a numeric precision of zero then two decimal places will be set 
    'This should only be used when there are a number of calculated fields which should all return decimal places
    Protected mvSetDecimalPlaces As Boolean

    Protected MustOverride Sub InitialiseFields()
    Protected MustOverride Function GetDecimalValue(ByVal pIndex As Integer, ByVal pDecimalPlaces As Integer) As String

    Protected MustOverride Function GetBinaryValue(ByVal pIndex As Integer) As String

    Protected MustOverride Function GetBinaryByteValue(ByVal pIndex As Integer) As Byte()

    Public ReadOnly Property Fields() As CDBFields
      Get
        Return mvFields
      End Get
    End Property

    Public Function FetchForXML() As Boolean
      Dim vFound As Boolean = mvDataReader.Read
      If mvFields Is Nothing Then InitialiseFields()
      Dim vFieldIndex As Integer = 1
      mvFields(vFieldIndex).Value = ""
      If vFound AndAlso Not mvDataReader.IsDBNull(0) Then
        Dim vSB As New StringBuilder
        Do
          vSB.Append(mvDataReader.GetString(0))
        Loop While mvDataReader.Read
        mvFields(vFieldIndex).Value = vSB.ToString
      End If
      Return vFound
    End Function

    Public Function Fetch() As Boolean
      Dim vFound As Boolean
      If mvUseDataTable Then
        vFound = mvDataTable.Rows.Count > mvRowNumber
        If mvFields Is Nothing Then InitialiseFieldsFromTable()
        FetchFromTable(vFound)
        mvRowNumber += 1
      Else
        vFound = mvDataReader.Read
        If mvFields Is Nothing Then InitialiseFields()
        FetchFromReader(vFound)
      End If
      mvStatus = vFound
      Return vFound
    End Function

    Private Sub InitialiseFieldsFromTable()
      mvFields = New CDBFields
      For vIndex As Integer = 0 To mvDataTable.Columns.Count - 1
        Select Case mvDataTable.Columns(vIndex).DataType.Name
          Case "Int16", "Int32", "Int64"
            mvFields.Add(mvDataTable.Columns(vIndex).ColumnName, CDBField.FieldTypes.cftLong)
          Case "DateTime"
            mvFields.Add(mvDataTable.Columns(vIndex).ColumnName, CDBField.FieldTypes.cftDate)
          Case "Decimal"
            mvFields.Add(mvDataTable.Columns(vIndex).ColumnName, CDBField.FieldTypes.cftNumeric)
          Case "String"
            mvFields.Add(mvDataTable.Columns(vIndex).ColumnName)
          Case "Byte[]"
            'Debug.Assert(mvUseDataTable = False, "Cannot read bulk attribute when using data table")
            If IsBinaryColumn(mvDataTable.Columns(vIndex).ColumnName) Then
              mvFields.Add(mvDataTable.Columns(vIndex).ColumnName, CDBField.FieldTypes.cftBinary)
            Else
              mvFields.Add(mvDataTable.Columns(vIndex).ColumnName, CDBField.FieldTypes.cftBulk)
            End If
          Case Else
            Debug.Assert(False, String.Format("Unknown field type {0} ", mvDataTable.Columns(vIndex).DataType.Name))
        End Select
      Next
    End Sub

    Private Sub FetchFromTable(ByVal vFound As Boolean)
      Dim vFieldIndex As Integer = 1
      For vIndex As Integer = 0 To mvDataTable.Columns.Count - 1
        mvFields(vFieldIndex).Value = ""
        If vFound Then
          Select Case mvFields(vFieldIndex).FieldType
            Case CDBField.FieldTypes.cftDate
              If mvDataTable.Rows(mvRowNumber)(vIndex).GetType.Name <> "DBNull" Then
                Dim vDate As Date = CDate(mvDataTable.Rows(mvRowNumber)(vIndex))
                Dim vDateString As String = vDate.ToString(CAREDateTimeFormat)
                If vDateString.EndsWith("00:00:00") Then
                  mvFields(vFieldIndex).Value = vDate.ToString(CAREDateFormat)
                Else
                  mvFields(vFieldIndex).SetFieldType(CDBField.FieldTypes.cftTime)
                  mvFields(vFieldIndex).Value = vDateString
                End If
              End If
            Case CDBField.FieldTypes.cftTime
              If mvDataTable.Rows(mvRowNumber)(vIndex).GetType.Name <> "DBNull" Then
                Dim vDate As Date = CDate(mvDataTable.Rows(mvRowNumber)(vIndex))
                Dim vDateString As String = vDate.ToString(CAREDateTimeFormat)
                If vDateString.EndsWith("00:00:00") Then
                  mvFields(vFieldIndex).Value = vDate.ToString(CAREDateFormat)
                Else
                  mvFields(vFieldIndex).Value = vDateString
                End If
              End If
            Case CDBField.FieldTypes.cftBinary
              mvFields(vFieldIndex).Value = Convert.ToBase64String(CType(mvDataTable.Rows(mvRowNumber)(vIndex), Byte()))
              mvFields(vFieldIndex).ByteValue = CType(mvDataTable.Rows(mvRowNumber)(vIndex), Byte())
            Case Else
              mvFields(vFieldIndex).Value = mvDataTable.Rows(mvRowNumber)(vIndex).ToString
          End Select
        End If
        vFieldIndex += 1
      Next
    End Sub

    Private Sub FetchFromReader(ByVal vFound As Boolean)
      Dim vFieldIndex As Integer = 1
      For vIndex As Integer = 0 To mvDataReader.FieldCount - 1
        mvFields(vFieldIndex).Value = ""
        If vFound AndAlso Not mvDataReader.IsDBNull(vIndex) Then
          Select Case mvFields(vFieldIndex).FieldType
            Case CDBField.FieldTypes.cftCharacter, CDBField.FieldTypes.cftMemo, CDBField.FieldTypes.cftUnicode
              mvFields(vFieldIndex).Value = mvDataReader.GetString(vIndex)
            Case CDBField.FieldTypes.cftBit
              If mvDataReader.GetBoolean(vIndex) Then
                mvFields(vFieldIndex).Value = "1"
              Else
                mvFields(vFieldIndex).Value = "0"
              End If
            Case CDBField.FieldTypes.cftInteger
              mvFields(vFieldIndex).Value = mvDataReader.GetInt16(vIndex).ToString
            Case CDBField.FieldTypes.cftLong
              mvFields(vFieldIndex).Value = mvDataReader.GetInt32(vIndex).ToString
            Case CDBField.FieldTypes.cftNumeric
              mvFields(vFieldIndex).Value = GetDecimalValue(vIndex, mvFields(vFieldIndex).DecimalPlaces)
            Case CDBField.FieldTypes.cftDate
              Dim vDate As Date = mvDataReader.GetDateTime(vIndex)
              Dim vDateString As String = vDate.ToString(CAREDateTimeFormat)
              If vDateString.EndsWith("00:00:00") Then
                mvFields(vFieldIndex).Value = vDate.ToString(CAREDateFormat)
              Else
                mvFields(vFieldIndex).SetFieldType(CDBField.FieldTypes.cftTime)
                mvFields(vFieldIndex).Value = vDateString
              End If
            Case CDBField.FieldTypes.cftTime
              Dim vDate As Date = mvDataReader.GetDateTime(vIndex)
              Dim vDateString As String = vDate.ToString(CAREDateTimeFormat)
              If vDateString.EndsWith("00:00:00") Then
                mvFields(vFieldIndex).Value = vDate.ToString(CAREDateFormat)
              Else
                mvFields(vFieldIndex).Value = vDateString
              End If
            Case CDBField.FieldTypes.cftBulk
              Dim vSize As Integer = CInt(mvDataReader.GetBytes(vIndex, 0, Nothing, 0, Int32.MaxValue))
              If vSize > 0 Then
                Dim vImage(vSize - 1) As Byte
                mvDataReader.GetBytes(vIndex, 0, vImage, 0, vImage.Length)
                Dim vFileName As String = Path.GetTempFileName
                Using vFS As FileStream = New FileStream(vFileName, FileMode.Create)
                  vFS.Write(vImage, 0, vImage.Length)
                  vFS.Close()
                End Using
                mvFields(vFieldIndex).Value = vFileName
              End If
            Case CDBField.FieldTypes.cftBinary
              'Oracle and SQL server require different code for processing a binary type
              mvFields(vFieldIndex).Value = GetBinaryValue(vIndex)
              mvFields(vFieldIndex).ByteValue = GetBinaryByteValue(vIndex)
          End Select
        End If
        vFieldIndex += 1
      Next
    End Sub

    Public ReadOnly Property Status() As Boolean
      'SHOULD ONLY USED BY VB6 MIGRATEDCODE
      Get
        Return mvStatus
      End Get
    End Property

    Friend ReadOnly Property Active() As Boolean
      Get
        If mvUseDataTable Then
          Return False
        Else
          Return mvDataReader IsNot Nothing AndAlso mvDataReader.IsClosed = False
        End If
      End Get
    End Property

    Friend ReadOnly Property DbConnection() As DbConnection
      Get
        Return mvDBConnection
      End Get
    End Property

    Public Property SetDecimalPlaces As Boolean
      Get
        Return mvSetDecimalPlaces
      End Get
      Set(ByVal pValue As Boolean)
        mvSetDecimalPlaces = pValue
      End Set
    End Property

    Public Sub CloseRecordSet()
      If mvDataTable IsNot Nothing Then
        mvDataTable.Dispose()
        mvDataTable = Nothing
      End If
      If mvDataReader IsNot Nothing Then
        mvDataReader.Close()
        mvDataReader = Nothing
      End If
      mvUseDataTable = False
      mvFields = Nothing
      If mvConnection IsNot Nothing Then
        mvConnection.RecordSets.Remove(Me)
        mvConnection.NotifyRecordSetClosed(Me)
        mvConnection = Nothing
      End If
    End Sub

    Public Function GetString(ByVal pName As String) As String
      If mvUseDataTable Then
        Return mvDataTable.Rows(mvRowNumber)(pName).ToString
      Else
        Return mvDataReader.GetString(mvDataReader.GetOrdinal(pName))
      End If
    End Function

    Public Function GetString(ByVal pID As Integer) As String
      If mvUseDataTable Then
        Return mvDataTable.Rows(mvRowNumber)(pID).ToString()
      Else
        Return mvDataReader.GetString(pID)
      End If
    End Function

    Private Function IsBinaryColumn(ByVal pColumnName As String) As Boolean
      Dim vIsBinary As Boolean = False
      If pColumnName.Equals("Password", StringComparison.InvariantCultureIgnoreCase) Then
        vIsBinary = True
      ElseIf pColumnName.Equals("authentication_code", StringComparison.InvariantCultureIgnoreCase) Then
        vIsBinary = True
      End If
      Return vIsBinary
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean

    Protected Overridable Sub Dispose(disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          Try
            Me.CloseRecordSet()
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