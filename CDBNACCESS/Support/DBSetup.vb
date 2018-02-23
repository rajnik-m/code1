Namespace Access
  Public Class DBSetup
    Public Shared Function CheckMailingTables(ByVal pEnv As CDBEnvironment, ByVal pFileNum As Integer, ByVal pDropTemp As Boolean) As String
      Dim vRecordSet As CDBRecordSet
      Dim vTableName As String
      Dim vLCaseName As String
      Dim vIsTemp As Boolean
      Dim vPossible As Boolean
      Dim vIgnore As Boolean
      Dim vOwner As String
      Dim vUsers As New CDBCollection
      Dim vPos As Integer
      Dim vErrorString As StringBuilder

      vRecordSet = New SQLStatement(pEnv.Connection, "DISTINCT logname", "users", New CDBFields).GetRecordSet
      While vRecordSet.Fetch()
        vOwner = vRecordSet.Fields("logname").Value.ToUpper
        vUsers.Add(vOwner, vOwner)
      End While
      vRecordSet.CloseRecordSet()
      vOwner = String.Empty

      vErrorString = New StringBuilder()
      Dim vDataTable As DataTable = pEnv.Connection.GetTableNames()    'Get all table names - all schemas for oracle
      For Each vRow As DataRow In vDataTable.Rows
        vTableName = vRow.Item("TABLE_NAME").ToString
        vLCaseName = vTableName.ToLower
        vIgnore = False
        vOwner = vRow.Item("OWNER").ToString.ToUpper
        If vOwner <> "DBO" AndAlso vOwner <> "CARE_ADMIN" Then
          vPos = InStr(vOwner, "\")
          If vPos > 0 Then
            vIgnore = Not vUsers.Exists(Mid$(vOwner, vPos + 1))
          Else
            vIgnore = Not vUsers.Exists(vOwner)
          End If
        End If
        vIsTemp = False
        vPossible = False

        If vIgnore OrElse Left$(vLCaseName, 4) = "ext_" OrElse Left$(vLCaseName, 4) = "sys_" OrElse vLCaseName = "db_info" Then
          'Ignore
        ElseIf vLCaseName.Contains("smcam_smapp") Then
          vIsTemp = True 'Selection manager temporary table
        ElseIf Left$(vLCaseName, 8) = "cm_temp_" Then
          vIsTemp = True 'Reports temporary table
        ElseIf Left$(vLCaseName, 9) = "rpt_temp_" Then
          vIsTemp = True 'Reports temporary table
        ElseIf vLCaseName.Contains("_temp_") Then
          vPossible = True 'Could be temp mailing
        ElseIf Len(vTableName) <= 7 AndAlso vTableName.Contains("_") Then
          vPossible = True 'Could be temp mailing
        ElseIf Left$(vLCaseName, 3) = "ca_" Then
          vPossible = True 'Could be temp campaign mailing
        End If

        If vPossible Then
          If pEnv.Connection.AttributeExists(vTableName, "selection_set") Then
            vIsTemp = True
          ElseIf pEnv.Connection.AttributeExists(vTableName, "segment_sequence") Then
            vIsTemp = True
          End If
        End If

        If vIsTemp Then
          If vOwner.Length > 0 Then
            If vOwner <> "DBO" AndAlso vOwner.Contains("\") Then vOwner = """" & vOwner & """"
            vOwner = vOwner & "."
          End If
          If pDropTemp Then
            vErrorString.AppendLine("Dropped Temporary Mailing Table: " & vTableName)
            If vOwner.Length > 0 Then vTableName = vOwner & vTableName
            pEnv.Connection.DropTable(vTableName)
          Else
            vErrorString.AppendLine("Found Temporary Mailing Table: " & vTableName)
          End If
        End If
      Next
      Return vErrorString.ToString
    End Function

    Public Shared Function CreateFixedRenewalLookup(ByVal pEnv As CDBEnvironment, ByVal pConfigValue As String) As Boolean
      Return CreateFixedRenewalLookup(pEnv, pConfigValue, Nothing)
    End Function
    Public Shared Function CreateFixedRenewalLookup(ByVal pEnv As CDBEnvironment, ByVal pConfigValue As String, ByVal pLogFile As LogFile) As Boolean
      Dim vValid As Boolean = True
      Try
        Dim vWhereFields As New CDBFields
        Dim vUpdateFields As New CDBFields
        Dim vValues() As String
        Dim vIndex As Integer
        Dim vDesc As String
        Dim vDate As Date

        'Delete all maintenance lookup data for membership_types.fixed_cycle column
        With vWhereFields
          .Add("table_name", "membership_types")
          .Add("attribute_name", "fixed_cycle")
        End With
        pEnv.Connection.DeleteRecords("maintenance_lookup", vWhereFields, False)
        'Repopulate the maintenance lookup data from the new config value
        With vUpdateFields
          .Add("table_name", "membership_types")
          .Add("attribute_name", "fixed_cycle")
          .Add("lookup_code")
          .Add("lookup_desc")
        End With
        vValues = pConfigValue.Split("|"c)
        If pConfigValue.Length > 0 Then
          For vIndex = 0 To vValues.Length - 1
            Dim vMonth As Integer = 0
            With vUpdateFields
              .Item("lookup_code").Value = vValues(vIndex)
              vDesc = ""
              If vValues(vIndex).Length = 3 AndAlso vValues(vIndex).Substring(vValues(vIndex).Length - 1, 1).ToUpper = "P" Then
                'ddP
                vDesc = " of the Present Membership Term"
              ElseIf vValues(vIndex).Length = 4 Then
                'ddmm
                vMonth = 1
              ElseIf vValues(vIndex).Length = 5 Then
                If vValues(vIndex).Substring(vValues(vIndex).Length - 1, 1).ToUpper = "P" Then
                  'ddmmP
                  vDesc = " of the Present Membership Term"
                  vMonth = 1
                ElseIf vValues(vIndex).Substring(vValues(vIndex).Length - 3, 1).ToUpper = "P" Then
                  'ddPdd
                  vDesc = " of the Present Membership Term"
                ElseIf vValues(vIndex).Substring(vValues(vIndex).Length - 1, 1).ToUpper = "F" Then
                  'ddmmF
                  vMonth = 1
                End If
              ElseIf vValues(vIndex).Length = 9 AndAlso vValues(vIndex).Substring(vValues(vIndex).Length - 5, 1).ToUpper = "P" Then
                'ddmmPddmm
                vDesc = " of the Present Membership Term"
                vMonth = 1
              End If
              If vDesc.Length = 0 Then vDesc = " of the Future Membership Term"
              If vMonth = 1 Then vMonth = IntegerValue(vValues(vIndex).Substring(2, 2)) Else vMonth = Today.Date.Month
              vDate = DateSerial(Date.Now.Year, vMonth, IntegerValue(vValues(vIndex).Substring(0, 2)))
              vDesc = vDate.ToString("dd MMMM") & vDesc
              .Item("lookup_desc").Value = vDesc
            End With
            pEnv.Connection.InsertRecord("maintenance_lookup", vUpdateFields)
          Next
        End If
      Catch vEX As Exception
        If pLogFile IsNot Nothing Then pLogFile.WriteLine(String.Format(ErrorText.DaeDBUpgradeMaintError, vEX.Message))
        vValid = False
      End Try
      Return vValid
    End Function
    Public Shared Sub GetNativeDataType(ByVal pCon As CDBConnection, ByVal pAttrName As String, ByRef pDataType As String, ByRef pParmCount As Integer, Optional ByVal pChangeDataType As Boolean = True, Optional ByVal pTableName As String = "")
      Dim vAttrName As String

      vAttrName = LCase$(pAttrName)
      pDataType = LCase$(pDataType)
      Select Case pCon.RDBMSType
        Case CDBConnection.RDBMSTypes.rdbmsSqlServer
          Select Case pDataType
            Case "char", "character", "nlschar", "nlscharacter", "varchar"
              pDataType = "varchar"
              pParmCount = 1
            Case "integer", "smallint"
              pDataType = "smallint"
            Case "longinteger", "int"
              pDataType = "int"
            Case "decimal"
              pDataType = "decimal"
              pParmCount = 2
            Case "numeric"
              pDataType = "numeric"
              pParmCount = 2
            Case "text", "nlstext", "varchar(max)"
              pDataType = "varchar(max)"
            Case "date", "time", "datetime"
              pDataType = "datetime"
            Case "bulk", "image", "varbinary(max)"
              pDataType = "varbinary(max)"
            Case "int identity"
              If vAttrName = "idem_record_id" Then
                'IDEM Identity columns so ignore
              Else
                Debug.Print("Unknown Data Type " & pDataType)
              End If
            Case "unicharacter"
              pDataType = "nvarchar"
              pParmCount = 1
            Case "binary"
              pParmCount = 1
            Case Else
              If pDataType.Length > 0 Then
                Debug.Print("Unknown Data Type " & pDataType)
              End If
          End Select
        Case CDBConnection.RDBMSTypes.rdbmsOracle
          Select Case pDataType
            Case "char", "character", "nlschar", "nlscharacter", "varchar2"
              pDataType = "varchar2"
              pParmCount = 1
            Case "integer", "smallint", "longinteger"
              pDataType = "integer"
            Case "decimal", "number"
              pDataType = "number"
              pParmCount = 2
            Case "text", "nlstext", "clob"
              pDataType = "clob"
            Case "date", "time"
              pDataType = "date"
            Case "bulk", "blob"
              pDataType = "blob"
            Case "binary"
              pDataType = "blob"
            Case Else
              If pDataType.Length > 0 Then
                Debug.Print("Unknown Data Type " & pDataType)
              End If
          End Select
      End Select
    End Sub
  End Class
End Namespace
