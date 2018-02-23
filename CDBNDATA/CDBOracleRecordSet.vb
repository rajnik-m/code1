Imports System.Data.OracleClient
Imports System.Data.Common
Imports System.IO

Namespace Data

  Friend Class CDBOracleRecordSet
    Inherits CDBRecordSet

    Protected Overrides Sub InitialiseFields()
      With mvDataReader
        mvFields = New CDBFields
        Dim vDT As DataTable = .GetSchemaTable
        For Each vRow As DataRow In vDT.Rows
          Select Case vRow("DataType").ToString
            Case "System.String"
              If vRow("IsLong").ToString = "True" Then
                mvFields.Add(vRow("ColumnName").ToString.ToLower, CDBField.FieldTypes.cftMemo)
              Else
                mvFields.Add(vRow("ColumnName").ToString.ToLower)
              End If
            Case "System.DateTime"
              mvFields.Add(vRow("ColumnName").ToString.ToLower, CDBField.FieldTypes.cftDate)
            Case "System.Decimal"
              Dim vNumericPrecision As Integer = CInt(vRow("NumericPrecision"))
              If vNumericPrecision = 38 Then
                mvFields.Add(vRow("ColumnName").ToString.ToLower, CDBField.FieldTypes.cftLong)
              Else
                Dim vDecimalPlaces As Integer = 0
                If vNumericPrecision > 0 Then
                  vDecimalPlaces = CInt(vRow("NumericScale"))
                Else
                  Debug.Print(vRow("ColumnName").ToString & " Has no numeric precision")
                  If mvSetDecimalPlaces AndAlso CInt(vRow("NumericScale")) = 0 Then vDecimalPlaces = 2
                End If
                Debug.Print(String.Format("{0} Precision {1} Places {2}", vRow("ColumnName"), vNumericPrecision, vDecimalPlaces))
                mvFields.Add(vRow("ColumnName").ToString.ToLower, CDBField.FieldTypes.cftNumeric).DecimalPlaces = vDecimalPlaces
              End If
            Case "System.Byte[]"
              If vRow("ColumnName").ToString = "PASSWORD" Then
                mvFields.Add(vRow("ColumnName").ToString.ToLower, CDBField.FieldTypes.cftBinary)
              Else
                mvFields.Add(vRow("ColumnName").ToString.ToLower, CDBField.FieldTypes.cftBulk)
              End If
            Case Else
              Debug.Assert(False, String.Format("Unknown field type {0} ", vRow("DataType").ToString))
          End Select
        Next
      End With
    End Sub

    Protected Overrides Function GetDecimalValue(ByVal pIndex As Integer, ByVal pDecimalPlaces As Integer) As String
      If pDecimalPlaces > 0 Then
        Return mvDataReader.GetDecimal(pIndex).ToString(String.Format("F{0}", pDecimalPlaces))
      Else
        'BR14503: Replaced mvDataReader.GetDecimal(pIndex).ToString() with the following to handle more than 38 decimal places
        Return DirectCast(mvDataReader, OracleDataReader).GetOracleValue(pIndex).ToString
      End If
    End Function

    Protected Overrides Function GetBinaryValue(ByVal pIndex As Integer) As String
      'BR19743 Oracle System value of Binary actually referes to blob for password.  We need to extract the blob and return the string value that it contains to compare to input password.
      Return Convert.ToBase64String(GetBinaryByteValue(pIndex))
    End Function

    Protected Overrides Function GetBinaryByteValue(ByVal pIndex As Integer) As Byte()
      'BR19743 Oracle System value of Binary actually referes to blob for password.
      Dim vBlob As System.Data.OracleClient.OracleLob = Nothing
      vBlob = DirectCast(mvDataReader, OracleDataReader).GetOracleLob(pIndex)
      Return CType(vBlob.Value, Byte())
    End Function

    Public Sub New(ByVal pOracleDataReader As OracleDataReader, ByVal pConnection As CDBConnection, ByVal pDBConnection As DbConnection)
      mvDataReader = pOracleDataReader
      mvConnection = pConnection
      mvDBConnection = pDBConnection
    End Sub
  End Class
End Namespace