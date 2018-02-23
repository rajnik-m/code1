Namespace Data

  Partial Public Class CDBConnection

    Public Enum cdbDataAccessMode
      damNormal = 0 'Normal operation
      damTest 'Disable all Inserts, Updates and Deletes
      damGenerateSQL 'Generate SQL for Inserts, Updates and Deletes
    End Enum

    Public ReadOnly Property RDBMSType() As RDBMSTypes
      Get
        Return mvRDBMSType
      End Get
    End Property

    Public Property DatabaseAccessMode() As cdbDataAccessMode
      Get
        Select Case mvDataAccessMode
          Case DataAccessModes.damGenerateSQL
            Return cdbDataAccessMode.damGenerateSQL
          Case DataAccessModes.damTest
            Return cdbDataAccessMode.damTest
          Case Else
            Return cdbDataAccessMode.damNormal
        End Select
      End Get
      Set(ByVal value As cdbDataAccessMode)
        Select Case value
          Case cdbDataAccessMode.damGenerateSQL
            mvDataAccessMode = DataAccessModes.damGenerateSQL
          Case cdbDataAccessMode.damTest
            mvDataAccessMode = DataAccessModes.damTest
          Case Else
            mvDataAccessMode = DataAccessModes.damNormal
        End Select
      End Set
    End Property

    Public Sub DeleteRecordsMultiTable(ByRef pTables As String, ByVal pWhereFields As CDBFields)
      Dim vTables() As String = Split(pTables, ",")
      For Each vTable As String In vTables
        DeleteRecords(vTable, pWhereFields, False)
      Next
    End Sub

    Public Function GetCount(ByVal pTable As String, ByVal pWhereFields As CDBFields, ByVal pWhereClause As String) As Integer
      Dim vNoRecords As Boolean
      Dim vCount As Integer

      If pWhereFields IsNot Nothing Then
        Return GetCount(pTable, pWhereFields)
      Else
        If pWhereClause.Length = 0 Then
          Return GetCount(pTable, Nothing)
        Else
          Dim vSQLStatement As New SQLStatement(Me, String.Format("SELECT COUNT(*) AS record_count FROM {0} WHERE {1}", pTable, pWhereClause))
          Dim vRecordSet As CDBRecordSet = vSQLStatement.GetRecordSet
          If vRecordSet.Fetch() Then
            vCount = vRecordSet.Fields(1).IntegerValue
          Else
            vNoRecords = True
          End If
          vRecordSet.CloseRecordSet()
          If vNoRecords Then RaiseError(DataAccessErrors.daeCountNoRecords, vSQLStatement.SQL)
          Return vCount
        End If
      End If
    End Function

    Public Function GetValue(ByVal pSQL As String, Optional ByRef pRecordsFound As Boolean = False) As String
      Dim vValue As String
      Dim vRecordSet As CDBRecordSet = Me.GetRecordSet(pSQL)
      If vRecordSet.Fetch() Then
        pRecordsFound = True
        vValue = vRecordSet.Fields(1).Value
      Else
        pRecordsFound = False
        vValue = ""
      End If
      vRecordSet.CloseRecordSet()
      Return vValue
    End Function

    Public Function GetRecordSetAnsiJoins(ByVal pSQL As String) As CDBRecordSet
      Return GetRecordSetAnsiJoins(pSQL, 0, RecordSetOptions.None)
    End Function

  End Class

End Namespace



