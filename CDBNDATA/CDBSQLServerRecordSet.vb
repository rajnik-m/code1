Imports System.Data.SqlClient
Imports System.Data.Common

Namespace Data

  Friend Class CDBSQLServerRecordSet
    Inherits CDBRecordSet

    Protected Overrides Sub InitialiseFields()
      With mvDataReader
        mvFields = New CDBFields
        For vIndex As Integer = 0 To .FieldCount - 1
          Select Case .GetDataTypeName(vIndex)
            Case "varchar", "char", "nvarchar", "nchar"
              mvFields.Add(.GetName(vIndex))
            Case "datetime"
              mvFields.Add(.GetName(vIndex), CDBField.FieldTypes.cftDate)
            Case "decimal"
              mvFields.Add(.GetName(vIndex), CDBField.FieldTypes.cftNumeric)
            Case "smallint"
              mvFields.Add(.GetName(vIndex), CDBField.FieldTypes.cftInteger)
            Case "bit"
              mvFields.Add(.GetName(vIndex), CDBField.FieldTypes.cftBit)
            Case "int"
              mvFields.Add(.GetName(vIndex), CDBField.FieldTypes.cftLong)
            Case "text", "ntext", "varchar(max)"
              mvFields.Add(.GetName(vIndex), CDBField.FieldTypes.cftMemo)
            Case "image", "varbinary"
              mvFields.Add(.GetName(vIndex), CDBField.FieldTypes.cftBulk)
            Case "sql_variant"
              mvFields.Add(.GetName(vIndex))
            Case Else

              Debug.Assert(False, String.Format("Unknown field type {0} ", .GetDataTypeName(vIndex)))
              mvFields.Add(.GetName(vIndex), CDBField.FieldTypes.cftUnknown)
          End Select
        Next
      End With
    End Sub

    Public Sub New(ByVal pDataTable As DataTable, ByVal pConnection As CDBConnection, ByVal pDBConnection As DbConnection)
      mvDataReader = Nothing
      mvRowNumber = 0
      mvDataTable = pDataTable
      mvUseDataTable = True
      mvConnection = pConnection
      mvDBConnection = pDBConnection
    End Sub

    Public Sub New(ByVal pSqlDataReader As SqlDataReader, ByVal pConnection As CDBConnection, ByVal pDBConnection As DbConnection)
      mvDataTable = Nothing
      mvRowNumber = 0
      mvUseDataTable = False
      mvDataReader = pSqlDataReader
      mvConnection = pConnection
      mvDBConnection = pDBConnection
    End Sub

    Protected Overrides Function GetDecimalValue(ByVal pIndex As Integer, ByVal pDecimalPlaces As Integer) As String
      If mvUseDataTable Then
        Return mvDataTable.Rows(mvRowNumber)(pIndex).ToString
      Else
        Return mvDataReader.GetDecimal(pIndex).ToString()
      End If
    End Function

    Protected Overrides Function GetBinaryValue(ByVal pIndex As Integer) As String
      Return Convert.ToBase64String(GetBinaryByteValue(pIndex))
    End Function

    Protected Overrides Function GetBinaryByteValue(ByVal pIndex As Integer) As Byte()
      Return CType(mvDataTable.Rows(mvRowNumber)(pIndex), Byte())
    End Function

  End Class

End Namespace