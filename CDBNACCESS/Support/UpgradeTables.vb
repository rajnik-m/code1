Namespace Access
  Public Class UpgradeTables

    'local variable(s) to hold property value(s)
    Private mvCol As Hashtable
    Private mvUpgradeTable As UpgradeTable
    Private mvSQL() As udtSQL
    Private Enum utSQLTypes
      utDrop
      utAlterOrCreate
      utIndices
      utComments
      utData
    End Enum

    Private Structure udtSQL
      Dim TableName As String
      Dim SQLType As utSQLTypes
      Dim SQL As String
    End Structure

    Public ReadOnly Property Count() As Integer
      Get
        Count = mvCol.Count
      End Get
    End Property
    Public ReadOnly Property Exists(ByVal pIndexKey As String) As Boolean
      Get
        Return mvCol.ContainsKey(pIndexKey)
      End Get
    End Property
    Public ReadOnly Property Item(ByVal pIndexKey As String) As UpgradeTable
      Get
        Item = CType(mvCol.Item(pIndexKey), UpgradeTable)
      End Get
    End Property
    Public Function GenerateSQL(ByVal pEnv As CDBEnvironment, ByVal pSQLFile As System.IO.StreamWriter, ByVal pLogFile As System.IO.StreamWriter, ByVal pPrint As Boolean, ByVal pExecute As Boolean, ByVal pMaxChange As Integer, ByVal pSQLGenerated As Boolean, ByVal pInitialise As Boolean, ByRef pIndexErrorOnly As Nullable(Of Boolean), ByVal pOriginalChangeNumber As Integer) As String
      Dim vErrorNumber As Integer
      Dim vDB As String
      Dim vTableCount As Integer
      Dim vErrorExit As Boolean
      Dim vErrorText As String = ""
      Dim vErrMessage As String
      Dim vErrMsg As String = ""
      Dim vUpgradeChange As UpgradeChange
      Dim vIndexError As Boolean
      Dim vIndex As Integer
      Dim vWhereFields As New CDBFields
      Dim vFields As New CDBFields
      Dim vConnection As CDBConnection
      Dim vHistory As Boolean
      Dim vErrorIndex As Integer = 0

      If Not pSQLGenerated Then
        ReDim mvSQL(0)
        mvSQL(0).SQL = ""
        mvSQL(0).TableName = ""
      End If

      vHistory = pEnv.Connection.TableExists("table_version_history")

      Dim vAllIndexErrors As New ParameterList
      For Each vKey As Object In mvCol.Keys
        mvUpgradeTable = CType(mvCol(vKey), UpgradeTable)
        vErrorNumber = 0
        vConnection = mvUpgradeTable.Connection
        vDB = "DATA"

        'DROP TABLE
        If mvUpgradeTable.ToBeDropped Then
          vErrMsg = "Dropping Table " & mvUpgradeTable.Key
          vTableCount = vTableCount + 1
          vIndex = 0
          If Not pSQLGenerated Then
            AddToSQL("DROP TABLE " & mvUpgradeTable.Key, utSQLTypes.utDrop)
            If pEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle Then AddToSQL("DROP PUBLIC SYNONYM " & mvUpgradeTable.Key, utSQLTypes.utDrop)
          End If
          For vIndex = 0 To UBound(mvSQL) - 1
            If mvSQL(vIndex).TableName = mvUpgradeTable.Key And mvSQL(vIndex).SQLType = utSQLTypes.utDrop And mvSQL(vIndex).SQL.Length > 0 Then
              If pPrint Then PrintSQL(pSQLFile, vDB, mvSQL(vIndex).SQL)
              If pExecute Then vErrorNumber = ExecuteSQLStatement(vConnection, mvSQL(vIndex).SQL, vErrorText)
              If vErrorNumber <> 0 Then
                vErrorIndex = vIndex
                Exit For
              End If
            End If
          Next
        End If
        'ALTER OR CREATE TABLE
        If vErrorNumber = 0 And (mvUpgradeTable.StructureModified Or mvUpgradeTable.ToBeCreated) Then
          If mvUpgradeTable.StructureModified Then
            vErrMsg = "Modifying Table " & mvUpgradeTable.Key
          Else
            vErrMsg = "Creating Table " & mvUpgradeTable.Key
          End If
          vTableCount = vTableCount + 1
          If Not pSQLGenerated Then ProcessTableChanges(pEnv)
          Dim vCDBIndexes As New CDBIndexes()
          If mvUpgradeTable.HasUnicode OrElse
             (pEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle AndAlso
              mvUpgradeTable.StructureModified) Then
            vCDBIndexes.Init(pEnv.Connection, mvUpgradeTable.Key)
            vCDBIndexes.DropAll(pEnv.Connection)
          End If
          For vIndex = 0 To UBound(mvSQL) - 1
            If mvSQL(vIndex).TableName = mvUpgradeTable.Key And mvSQL(vIndex).SQLType = utSQLTypes.utAlterOrCreate And mvSQL(vIndex).SQL.Length > 0 Then
              If pPrint Then PrintSQL(pSQLFile, vDB, mvSQL(vIndex).SQL)
              If pExecute Then vErrorNumber = ExecuteSQLStatement(vConnection, mvSQL(vIndex).SQL, vErrorText)
              If vErrorNumber <> 0 Then
                vErrorIndex = vIndex
                Exit For
              End If
            End If
          Next
          If mvUpgradeTable.HasUnicode OrElse
             (pEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle AndAlso
              mvUpgradeTable.StructureModified) Then
            vCDBIndexes.ReCreate(pEnv.Connection)
          End If
        End If

        If vErrorNumber = 0 And mvUpgradeTable.UpgradeIndices.Count > 0 Then
          'TABLE INDICES
          vTableCount = vTableCount + 1
          vErrMsg = "Modifying Indices on Table " & mvUpgradeTable.Key
          If Not pSQLGenerated Then ProcessIndices(pEnv)
          Dim vIndexErrorNumber As Integer = 0
          For vIndex = 0 To UBound(mvSQL) - 1
            If mvSQL(vIndex).TableName = mvUpgradeTable.Key And mvSQL(vIndex).SQLType = utSQLTypes.utIndices And mvSQL(vIndex).SQL.Length > 0 Then
              If pPrint Then PrintSQL(pSQLFile, vDB, mvSQL(vIndex).SQL)
              If pExecute Then vErrorNumber = ExecuteSQLStatement(vConnection, mvSQL(vIndex).SQL, vErrorText)
              If vErrorNumber <> 0 And Left$(mvSQL(vIndex).SQL, 10) = "DROP INDEX" Then
                If vErrorText.IndexOf("does not exist") > 0 Then
                  vErrorNumber = 0
                  vErrorText = String.Empty
                ElseIf vErrorText.IndexOf("nicht im systemkatalog vorhanden", 1) > 0 Then
                  vErrorNumber = 0
                  vErrorText = String.Empty
                End If
              End If
              If vErrorNumber <> 0 Then
                If vAllIndexErrors.ContainsKey(vIndex.ToString) = False Then vAllIndexErrors.Add(vIndex.ToString, vErrorText) 'Store the error and continue with other indices
                vIndexErrorNumber = vErrorNumber
                vErrorIndex = vIndex
                'Exit For
              End If
            End If
          Next
          If vIndexErrorNumber <> 0 Then vErrorNumber = vIndexErrorNumber 'Reset vErrorNumber so that nothing else gets attempted
        End If

        If vErrorNumber = 0 And mvUpgradeTable.DataMods.Count > 0 Then
          'DATA MODIFICATIONS
          vTableCount = vTableCount + 1
          vErrMsg = "Modifying Data in Table " & mvUpgradeTable.Key
          If Not pSQLGenerated Then ProcessDataMods()
          For vIndex = 0 To UBound(mvSQL) - 1
            If mvSQL(vIndex).TableName = mvUpgradeTable.Key And mvSQL(vIndex).SQLType = utSQLTypes.utData And mvSQL(vIndex).SQL.Length > 0 Then
              If pPrint Then PrintSQL(pSQLFile, vDB, mvSQL(vIndex).SQL)
              If pExecute Then vErrorNumber = ExecuteSQLStatement(vConnection, mvSQL(vIndex).SQL, vErrorText)
              If vErrorNumber <> 0 Then
                vErrorIndex = vIndex
                Exit For
              End If
            End If
          Next
        End If

        If vErrorNumber = 0 And vHistory Then
          For Each vChangeKey As Object In mvUpgradeTable.mvChanges.Keys
            vUpgradeChange = CType(mvUpgradeTable.mvChanges.Item(vChangeKey), UpgradeChange)
            If vUpgradeChange.ToBeApplied And vUpgradeChange.VersionHistoryExists Then
              If Not pSQLGenerated Then ProcessVersionHistory(pEnv, mvUpgradeTable, vUpgradeChange)
            End If
          Next
          For vIndex = 0 To UBound(mvSQL) - 1
            If mvSQL(vIndex).TableName = mvUpgradeTable.Key And mvSQL(vIndex).SQLType = utSQLTypes.utComments And mvSQL(vIndex).SQL.Length > 0 Then
              If pPrint Then PrintSQL(pSQLFile, vDB, mvSQL(vIndex).SQL)
              If pExecute Then vErrorNumber = ExecuteSQLStatement(vConnection, mvSQL(vIndex).SQL, vErrorText)
              If vErrorNumber <> 0 Then
                vErrorIndex = vIndex
                Exit For
              End If
            End If
          Next
        End If

        If vErrorNumber = 0 Then
          'loop thru each change and write to log file
          pLogFile.WriteLine("")
          pLogFile.WriteLine("Table : " & mvUpgradeTable.Key)
          For Each vChangeKey As Object In mvUpgradeTable.mvChanges.Keys
            vUpgradeChange = CType(mvUpgradeTable.mvChanges.Item(vChangeKey), UpgradeChange)
            If vUpgradeChange.ToBeApplied Then
              If pExecute Then
                vErrMessage = "Applied Successfully"
              Else
                vErrMessage = "Will Be Applied"
              End If
            Else
              vErrMessage = "Already Applied"
            End If
            pLogFile.WriteLine(vbTab & "Change " & vUpgradeChange.ChangeNumber & " : " & vUpgradeChange.ChangeComment & " : " & vErrMessage)
          Next
        Else
          pLogFile.WriteLine("")
          pLogFile.WriteLine("Table : " & mvUpgradeTable.Key)
          'check that the error is either an index or a synonym problem
          If vAllIndexErrors.Count > 0 Then
            vIndexError = True
            For Each vItem As DictionaryEntry In vAllIndexErrors
              vIndexError = True
              If pIndexErrorOnly.HasValue = False Then pIndexErrorOnly = True
              vErrorNumber = CInt(vItem.Key)
              vErrorText = vItem.Value.ToString
              pLogFile.WriteLine(vbTab & vErrMsg)
              If mvSQL(vErrorNumber).SQL.IndexOf("INDEX ", 1) > 0 Then
                pLogFile.WriteLine(vbTab & "ERROR: INDEX already exists or INDEX name already in use")
              ElseIf mvSQL(vErrorNumber).SQL.IndexOf(" SYNONYM ") > 0 Then
                pLogFile.WriteLine(vbTab & "ERROR: SYNONYM name already in use")
              ElseIf (InStr(vErrorText, "duplicate") > 0) AndAlso InStr(1, mvSQL(vIndex).SQL, "UNIQUE INDEX ", vbTextCompare) > 0 Then
                pLogFile.WriteLine(vbTab & "ERROR: Cannot create UNIQUE INDEX; duplicate keys found")
              End If
              pLogFile.WriteLine(vbTab & String.Format(ErrorText.DaeDBUpgradeActualError, vErrorText))
              pLogFile.WriteLine(vbTab & mvSQL(vErrorNumber).SQL)
            Next
          Else
            If (vErrorText.IndexOf("name is already") > 0 Or vErrorText.IndexOf("already exists") > 0 _
            Or vErrorText.IndexOf("column list already indexed") > 0 Or vErrorText.IndexOf("already an index") > 0) _
            And (mvSQL(vIndex).SQL.IndexOf("INDEX ", 1) > 0 Or mvSQL(vIndex).SQL.IndexOf(" SYNONYM ") > 0) Then
              vIndexError = True
              If pIndexErrorOnly.HasValue = False Then pIndexErrorOnly = True
              If InStr(mvSQL(vIndex).SQL, " INDEX ") > 0 Then
                pLogFile.WriteLine(vbTab & vErrMsg & " : INDEX already exists or INDEX name already in use")
              Else
                pLogFile.WriteLine(vbTab & vErrMsg & " : SYNONYM name already in use")
              End If
              pLogFile.WriteLine(vbTab & String.Format(ErrorText.DaeDBUpgradeActualError, vErrorText))
              pLogFile.WriteLine(vbTab & mvSQL(vIndex).SQL)
            ElseIf InStr(vErrorText, "synonym to be dropped does not exist") > 0 Then
              pLogFile.WriteLine(vbTab & vErrMsg & " : SYNONYM to be dropped does not exist")
              pLogFile.WriteLine(vbTab & String.Format(ErrorText.DaeDBUpgradeActualError, vErrorText))
              pLogFile.WriteLine(vbTab & mvSQL(vIndex).SQL)
              pIndexErrorOnly = False
            ElseIf (InStr(vErrorText, "duplicate") > 0) And InStr(1, mvSQL(vIndex).SQL, "UNIQUE INDEX ", vbTextCompare) > 0 Then
              vIndexError = True
              If pIndexErrorOnly.HasValue = False Then pIndexErrorOnly = True
              pLogFile.WriteLine(vbTab & vErrMsg & " : Cannot create UNIQUE INDEX; duplicate keys found")
              pLogFile.WriteLine(vbTab & String.Format(ErrorText.DaeDBUpgradeActualError, vErrorText))
              pLogFile.WriteLine(vbTab & mvSQL(vIndex).SQL)
            Else
              'if it's not an index or synonym problem then display a dialog indicating what happened
              pLogFile.WriteLine(vbTab & vErrMsg & " Failed : " & vErrorText)
              pLogFile.WriteLine(vbTab & mvSQL(vIndex).SQL)
              pIndexErrorOnly = False
              Return vErrMsg & " Failed : " & vErrorText
            End If
          End If
        End If
      Next

      If Not vErrorExit Then
        If vTableCount = 0 Then

        Else
          If pExecute And pEnv.Connection.TableExists("config") Then
            If pEnv.GetConfig("last_db_structure_change").Length > 0 Then
              'update config
              If pMaxChange > 0 AndAlso pMaxChange > pOriginalChangeNumber Then
                'Only update the change number if it is greater then the start change number
                vWhereFields.Add("config_name", CDBField.FieldTypes.cftCharacter, "last_db_structure_change")
                vFields.AddAmendedOnBy(If(String.IsNullOrWhiteSpace(pEnv.User.UserID), "dbinit", pEnv.User.UserID))
                vFields.Add("config_value", CDBField.FieldTypes.cftCharacter, pMaxChange.ToString)
                pEnv.Connection.UpdateRecords("config", vFields, vWhereFields)
              End If
            Else
              'insert config
              If pEnv.Connection.TableExists("config") Then
                vFields.AddAmendedOnBy(If(String.IsNullOrWhiteSpace(pEnv.User.UserID), "dbinit", pEnv.User.UserID))
                vFields.Add("config_name", CDBField.FieldTypes.cftCharacter, "last_db_structure_change")
                vFields.Add("config_value", CDBField.FieldTypes.cftCharacter, Format$(pMaxChange))
                pEnv.Connection.InsertRecord("config", vFields)
              End If
            End If
            If vIndexError Then
              Return "The database has been upgraded successfully, but some problems with the creation of indices were encountered" & vbCrLf & vbCrLf & "You should review the Upgrade Log file for the details of the database changes and index problems"
            End If
          End If
        End If
      Else

      End If
      If vErrorNumber = 0 Then
        Return ""
      Else
        Return vErrorNumber.ToString
      End If
    End Function
    Private Sub PrintSQL(ByVal pSQLFile As System.IO.StreamWriter, ByVal pDB As String, ByVal pSQL As String)
      pSQLFile.WriteLine(pDB & "," & pSQL)
      pSQLFile.WriteLine("")
    End Sub
    Private Function DetermineDataType(ByVal pUpgradeAttr As UpgradeAttribute) As String
      Dim vSQL As String
      vSQL = pUpgradeAttr.DataType
      If pUpgradeAttr.ParameterCount > 0 Then
        If pUpgradeAttr.Parameter1 <> 0 Then
          vSQL = vSQL & "(" & Format(pUpgradeAttr.Parameter1)
          If pUpgradeAttr.ParameterCount = 2 And pUpgradeAttr.Parameter2 <> 0 Then
            vSQL = vSQL & "," & Format(pUpgradeAttr.Parameter2)
          End If
          vSQL = vSQL & ")"
        End If
      End If
      Return vSQL
    End Function
    Public Sub Remove(ByVal pIndexKey As String)
      mvCol.Remove(pIndexKey)
    End Sub
    Public Function Add(ByVal pEnv As CDBEnvironment, ByVal pKey As String, ByVal pDatabaseName As String, ByVal pStructureModified As Boolean, ByVal pToBeCreated As Boolean, ByVal pToBeDropped As Boolean) As UpgradeTable
      'create a new object
      Dim vNewMember As UpgradeTable
      Dim vAttr As UpgradeAttribute
      Dim vDataType As String
      Dim vDoAdd As Boolean
      Dim vIRS As DataTable
      Dim vConn As CDBConnection

      vNewMember = New UpgradeTable
      vAttr = New UpgradeAttribute

      vDoAdd = True
      vConn = pEnv.Connection
      If vConn.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle Then pKey = pKey.ToUpper()
      pKey = Left$(pKey, 30)
      If pDatabaseName = "HELP" Then vDoAdd = False
      If vDoAdd Then
        'set the properties passed into the method
        vIRS = vConn.GetAttributeNames(pKey)
        If vIRS.Rows.Count > 0 Then
          If vIRS.Rows(0).Item("OWNER").ToString <> "sys" Then pToBeCreated = False
        Else
          If Not pToBeCreated Then vDoAdd = False
        End If
        If vDoAdd Then
          With vNewMember
            .Key = pKey
            .Connection = vConn
            .StructureModified = pStructureModified
            .ToBeDropped = pToBeDropped
            .ToBeCreated = pToBeCreated
            If Not (.ToBeCreated Or .ToBeDropped) Then
              For Each vRow As DataRow In vIRS.Rows
                If Not .UpgradeAttributes.Exists(vRow.Item("Column_Name").ToString) Then
                  vDataType = vRow.Item("DATA_TYPE").ToString
                  DBSetup.GetNativeDataType(pEnv.Connection, vRow.Item("COLUMN_NAME").ToString, vDataType, 0)
                  If vDataType.ToUpper = "C" OrElse Mid(vDataType, 1, 7) = "varchar" OrElse Mid(vDataType, 1, 8) = "varchar2" OrElse _
                    vDataType.ToUpper = "U" OrElse Mid(vDataType, 1, 8) = "nvarchar" Then
                    vAttr = .UpgradeAttributes.Add(pEnv, vRow.Item("COLUMN_NAME").ToString, vDataType, vRow.Item("LENGTH").ToString, vRow.Item("SCALE").ToString, "", CType(IIf(vRow.Item("Nullable").ToString = "Y", 1, 0), UpgradeAttribute.NullOptions))
                  Else
                    vAttr = .UpgradeAttributes.Add(pEnv, vRow.Item("COLUMN_NAME").ToString, vDataType, vRow.Item("PRECISION").ToString, vRow.Item("SCALE").ToString, "", CType(IIf(vRow.Item("Nullable").ToString = "Y", 1, 0), UpgradeAttribute.NullOptions))
                  End If
                End If
              Next
            End If
          End With
          mvCol.Add(pKey, vNewMember)
        End If

      End If
      'return the object created
      Return vNewMember
    End Function
    Private Sub AddToSQL(ByVal pSQL As String, ByVal pSQLType As utSQLTypes)
      If pSQL IsNot Nothing AndAlso pSQL.Length > 0 Then
        mvSQL(UBound(mvSQL)).TableName = mvUpgradeTable.Key
        mvSQL(UBound(mvSQL)).SQLType = pSQLType
        mvSQL(UBound(mvSQL)).SQL = pSQL
        ReDim Preserve mvSQL(UBound(mvSQL) + 1)
      End If
    End Sub
    Private Sub AddAttrToAlter(ByVal pConn As CDBConnection, ByRef pSQL As String, ByVal pUpgradeAttr As UpgradeAttribute, ByVal pType As String)
      Dim vTemp As String
      If pSQL Is Nothing Then pSQL = ""
      If pSQL.Length > 0 Then
        pSQL = pSQL & ", "
      Else
        If pConn.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle Then
          pSQL = "("
        Else
          pSQL = ""
        End If
      End If
      If pConn.IsSpecialColumn(pUpgradeAttr.Key) Then
        vTemp = pConn.DBSpecialCol("", pUpgradeAttr.Key)
      Else
        vTemp = pUpgradeAttr.Key
      End If
      pSQL = pSQL & vTemp
      If Not pUpgradeAttr.ToBeDeleted Then pSQL = pSQL & " " & DetermineDataType(pUpgradeAttr)
    End Sub

    ''' <summary>Executes the current SQL statement.</summary>
    ''' <param name="pConn">Current <see cref="CDBConnection">CDBConnection</see> object.</param>
    ''' <param name="pSQL">The SQL to be executed.</param>
    ''' <param name="pErrorText">If an error is generated by the SQL, this will be set to the error message.</param>
    ''' <returns>0 if the SQL executes successfully, otherwise 1.</returns>
    ''' <remarks></remarks>
    Private Function ExecuteSQLStatement(ByVal pConn As CDBConnection, ByVal pSQL As String, ByRef pErrorText As String) As Integer
      Try
        pConn.ExecuteSQL(pSQL)
        Return 0
      Catch ex As Exception
        pErrorText = ex.Message
        Return 1
      End Try
    End Function
    Public Sub ProcessVersionHistory(ByVal pEnv As CDBEnvironment, ByVal pUpgradeTable As UpgradeTable, ByVal pUpgradeChange As UpgradeChange)
      Dim vTVH As New TableVersionHistory(pEnv)
      Dim vList As CDBParameters
      Dim vValueChanged As Boolean
      With pUpgradeChange
        vTVH.Init(pEnv, pUpgradeTable.Key, .ChangeNumber)
        vList = New CDBParameters
        vList.Add("TableName", pUpgradeTable.Key)
        vList.Add("ChangeNumber", .ChangeNumber)
        vList.Add("ReleaseNumber", .ReleaseNumber)
        vList.Add("VersionNumber", "n/a")
        vList.Add("Logname", .Logname)
        vList.Add("ChangeDate", .ChangeDate)
        vList.Add("ChangeDescription", .ChangeDescription)
        If pUpgradeTable.Key <> vTVH.TableName OrElse .ChangeNumber <> vTVH.ChangeNumber OrElse .ReleaseNumber <> vTVH.ReleaseNumber _
          OrElse "n/a" <> vTVH.VersionNumber OrElse .Logname <> vTVH.Logname OrElse .ChangeDate <> vTVH.ChangeDate OrElse .ChangeDescription <> vTVH.ChangeDescription Then
          vValueChanged = True
        End If
        If vTVH.Existing Then
          vTVH.Update(vList)
        Else
          vTVH.Create(vList)
        End If
      End With
      pUpgradeTable.Connection.DatabaseAccessMode = CDBConnection.cdbDataAccessMode.damGenerateSQL
      vTVH.Save()
      pUpgradeTable.Connection.DatabaseAccessMode = CDBConnection.cdbDataAccessMode.damNormal
      If vValueChanged Then
        AddToSQL(pUpgradeTable.Connection.LastGeneratedSQL, utSQLTypes.utComments)
      End If
    End Sub
    Public Sub ProcessDataMods()
      Dim vDataMod As DataMod
      Dim vFields As New CDBFields
      Dim vWhereFields As New CDBFields

      For vCtr As Integer = 1 To mvUpgradeTable.DataMods.Count
        vDataMod = mvUpgradeTable.DataMods.Item(vCtr)
        vFields = vDataMod.Fields(mvUpgradeTable.Connection)
        vWhereFields = vDataMod.WhereFields(mvUpgradeTable.Connection)
        With mvUpgradeTable.Connection
          .DatabaseAccessMode = CDBConnection.cdbDataAccessMode.damGenerateSQL
          Select Case vDataMod.ChangeType
            Case DataModTypes.dmtDelete
              .DeleteRecords(mvUpgradeTable.Key, vWhereFields)
              AddToSQL(.LastGeneratedSQL, utSQLTypes.utData)
            Case DataModTypes.dmtInsert
              .InsertRecord(mvUpgradeTable.Key, vFields)
              AddToSQL(.LastGeneratedSQL, utSQLTypes.utData)
            Case DataModTypes.dmtUpdate
              .UpdateRecords(mvUpgradeTable.Key, vFields, vWhereFields)
              AddToSQL(.LastGeneratedSQL, utSQLTypes.utData)
          End Select
          .DatabaseAccessMode = CDBConnection.cdbDataAccessMode.damNormal
        End With
      Next
    End Sub
    Public Sub ProcessIndices(ByVal pEnv As CDBEnvironment)
      Dim vUpgradeIndex As UpgradeIndex = Nothing

      For Each vKey As Object In mvUpgradeTable.UpgradeIndices.mvCol.Keys
        vUpgradeIndex = mvUpgradeTable.UpgradeIndices.Item(vKey)
        Dim vSQL As String = String.Format(" ON {0} ({1})",
                                           mvUpgradeTable.Key,
                                           Function(attributes As IList(Of String)) As String
                                             Dim result As String = String.Empty
                                             For Each attribute As String In attributes
                                               result &= String.Format("{0}{1}", If(String.IsNullOrWhiteSpace(result), String.Empty, ", "),
                                                                       attribute)
                                             Next attribute
                                             Return result
                                           End Function(vUpgradeIndex.Attributes))
        vSQL = vSQL & pEnv.GetTableSpaceInfo(mvUpgradeTable.Key, True)

        If vUpgradeIndex.ToBeDeleted Then
          For Each vStatement As String In pEnv.Connection.DropIndexSql(mvUpgradeTable.Key, vUpgradeIndex.Key)
            AddToSQL(vStatement, utSQLTypes.utIndices)
          Next vStatement
        End If

        If vUpgradeIndex.ToBeCreated Then
          If Not mvUpgradeTable.ToBeCreated Then
            For Each vStatement As String In pEnv.Connection.DropIndexSql(mvUpgradeTable.Key, vUpgradeIndex.Key)
              AddToSQL(vStatement, utSQLTypes.utIndices)
            Next vStatement
          End If
          For Each vStatement As String In pEnv.Connection.CreateIndexSql(vUpgradeIndex.UniqueIndex, mvUpgradeTable.Key, vUpgradeIndex.Attributes)
            AddToSQL(vStatement, utSQLTypes.utIndices)
          Next vStatement
        End If
      Next vKey
    End Sub

    Private Sub ProcessTableChanges(ByVal pEnv As CDBEnvironment)

      Dim vSQL As String = ""
      Dim vSQL2() As String
      Dim vSQL3() As String
      Dim vSQL4() As String
      Dim vSQLComments() As String
      Dim vSQLAdd As String = ""
      Dim vSQLAlter As String = ""
      Dim vSQLDrop As String = ""
      Dim vIndex As Integer
      Dim vNotNulls As Boolean
      Dim vAttrNotNull As Integer
      Dim vSynonym As String = ""
      Dim vRights As String = ""
      Dim vUpgradeAttr As UpgradeAttribute
      Dim vUpgradeTable As UpgradeTable
      Dim vIRS As DataTable = New DataTable

      vUpgradeTable = New UpgradeTable
      ReDim vSQL2(0)
      ReDim vSQL3(0)
      ReDim vSQL4(0)
      ReDim vSQLComments(0)

      If mvUpgradeTable.ToBeCreated Then
        vIndex = 1
        vSQL = SQLCreateTable(pEnv, mvUpgradeTable)
        Select Case pEnv.Connection.RDBMSType
          Case CDBConnection.RDBMSTypes.rdbmsOracle
            vSynonym = "CREATE PUBLIC SYNONYM " & mvUpgradeTable.Key & " FOR " & pEnv.User.Logname & "." & mvUpgradeTable.Key
            vRights = "GRANT SELECT, INSERT, UPDATE, DELETE ON " & mvUpgradeTable.Key & " TO care_user"
          Case Else
            'waiting for more databases to be supported
        End Select
      Else
        Select Case pEnv.Connection.RDBMSType
          Case CDBConnection.RDBMSTypes.rdbmsOracle
            For Each vUpgradeAttr In mvUpgradeTable.UpgradeAttributes.mvCol
              If vUpgradeAttr.ToBeDeleted Then  'DROP ATTRIBUTE
                AddAttrToAlter(pEnv.Connection, vSQLDrop, vUpgradeAttr, "DELETE")
              ElseIf vUpgradeAttr.ToBeCreated Then  'CREATE ATTRIBUTE
                AddAttrToAlter(pEnv.Connection, vSQLAdd, vUpgradeAttr, "ADD")
                If vUpgradeAttr.Nullable = UpgradeAttribute.NullOptions.noNullsInvalid Then
                  If mvUpgradeTable.RecordCount = 0 Then
                    vSQLAdd = vSQLAdd & " Not Null"
                  Else
                    vNotNulls = True
                  End If
                End If
              ElseIf vUpgradeAttr.StructureModified Then  'ALTER ATTRIBUTE
                AddAttrToAlter(pEnv.Connection, vSQLAlter, vUpgradeAttr, "CHANGE")
                vAttrNotNull = 1
                vIRS = mvUpgradeTable.Connection.GetAttributeNames(mvUpgradeTable.Key)
                If vIRS.Rows.Count > 0 Then
                  For Each vRow As DataRow In vIRS.Rows
                    If vRow.Item("Column_Name").ToString = vUpgradeAttr.Key Then
                      vAttrNotNull = IntegerValue(IIf(vRow.Item("Nullable").ToString = "Y", 1, 0).ToString)
                    End If
                  Next
                End If
                If vUpgradeAttr.Nullable = UpgradeAttribute.NullOptions.noNullsInvalid Then  'attribute is mandatory
                  If mvUpgradeTable.RecordCount = 0 Or vAttrNotNull = UpgradeAttribute.NullOptions.noNullsInvalid Then 'if no records in table OR attribute is already NOT NULL
                    If vAttrNotNull <> UpgradeAttribute.NullOptions.noNullsInvalid Then vSQLAlter = vSQLAlter & " Not Null"
                  Else
                    vNotNulls = True
                  End If
                Else  'attribute is not mandatory
                  If vAttrNotNull <> UpgradeAttribute.NullOptions.noNullsAllowed Then vSQLAlter = vSQLAlter & " Null"
                End If
              End If
            Next
            If (vSQLAdd.Length > 0 Or vSQLAlter.Length > 0) Then
              SetDefaultValues(pEnv.Connection, vSQL2, vSQL3)
              ReDim Preserve vSQL2(UBound(vSQL2) + 1)
              ReDim Preserve vSQL3(UBound(vSQL3) + 1)
            End If

            If vSQLAdd.Length > 0 Then vSQLAdd = "ALTER TABLE " & mvUpgradeTable.Key & " ADD " & vSQLAdd & ")"
            If vSQLAlter.Length > 0 Then vSQLAlter = "ALTER TABLE " & mvUpgradeTable.Key & " MODIFY " & vSQLAlter & ")"
            If vSQLDrop.Length > 0 Then vSQLDrop = "ALTER TABLE " & mvUpgradeTable.Key & " DELETE " & vSQLDrop & ")"

          Case CDBConnection.RDBMSTypes.rdbmsSqlServer
            For Each vUpgradeAttr In mvUpgradeTable.UpgradeAttributes.mvCol
              If vUpgradeAttr.ToBeDeleted Then  'DROP ATTRIBUTE
                AddAttrToAlter(pEnv.Connection, vSQLDrop, vUpgradeAttr, "DELETE")
              ElseIf vUpgradeAttr.ToBeCreated Then  'CREATE ATTRIBUTE
                AddAttrToAlter(pEnv.Connection, vSQLAdd, vUpgradeAttr, "ADD")
                If vUpgradeAttr.Nullable = UpgradeAttribute.NullOptions.noNullsInvalid Then vNotNulls = True 'SQLServer2000 does not allow new mandatory attributes even if the table is empty
              ElseIf vUpgradeAttr.StructureModified Then  'ALTER ATTRIBUTE
                If vSQLAlter.Length > 0 Then
                  'SQLServer requires separate Alter statements for each attribute to be altered
                  vSQLAlter = "ALTER TABLE " & mvUpgradeTable.Key & " ALTER COLUMN " & vSQLAlter
                  vSQL2(UBound(vSQL2)) = vSQLAlter
                  ReDim Preserve vSQL2(UBound(vSQL2) + 1)
                  vSQLAlter = ""
                End If
                AddAttrToAlter(pEnv.Connection, vSQLAlter, vUpgradeAttr, "CHANGE")
                vAttrNotNull = 1
                vIRS = mvUpgradeTable.Connection.GetAttributeNames(mvUpgradeTable.Key)

                For Each vRow As DataRow In vIRS.Rows
                  If vRow.Item("Column_Name").ToString = vUpgradeAttr.Key Then
                    vAttrNotNull = IntegerValue(IIf(vRow.Item("Nullable").ToString = "Y", 1, 0).ToString)
                  End If
                Next
                If vUpgradeAttr.Nullable = UpgradeAttribute.NullOptions.noNullsInvalid Then  'attribute is mandatory
                  If mvUpgradeTable.RecordCount = 0 Or vAttrNotNull = UpgradeAttribute.NullOptions.noNullsInvalid Then 'if no records in table OR attribute is already NOT NULL
                    vSQLAlter = vSQLAlter & " Not Null"
                  Else
                    vNotNulls = True
                  End If
                Else  'attribute is not mandatory
                  If vAttrNotNull <> UpgradeAttribute.NullOptions.noNullsAllowed Then vSQLAlter = vSQLAlter & " Null"
                End If
              End If
            Next
            If vSQLAdd.Length > 0 OrElse vSQLAlter.Length > 0 Then
              SetDefaultValues(pEnv.Connection, vSQL3, vSQL4)
              ReDim Preserve vSQL2(UBound(vSQL2) + 1)
              ReDim Preserve vSQL3(UBound(vSQL3) + 1)
            End If

            If vSQLAdd.Length > 0 Then vSQLAdd = "ALTER TABLE " & mvUpgradeTable.Key & " ADD " & vSQLAdd
            If vSQLAlter.Length > 0 Then vSQLAlter = "ALTER TABLE " & mvUpgradeTable.Key & " ALTER COLUMN " & vSQLAlter
            If vSQLDrop.Length > 0 Then vSQLDrop = "ALTER TABLE " & mvUpgradeTable.Key & " DROP COLUMN " & vSQLDrop
          Case Else
            'waiting for more databases to be supported
        End Select
      End If
      'Output all the generated SQL
      If vSQL.Length > 0 Then AddToSQL(vSQL, utSQLTypes.utAlterOrCreate)
      If vSynonym.Length > 0 Then AddToSQL(vSynonym, utSQLTypes.utAlterOrCreate)
      If vRights.Length > 0 Then AddToSQL(vRights, utSQLTypes.utAlterOrCreate)
      If vSQLAdd.Length > 0 Then AddToSQL(vSQLAdd, utSQLTypes.utAlterOrCreate)
      If vSQLAlter.Length > 0 Then AddToSQL(vSQLAlter, utSQLTypes.utAlterOrCreate)
      If vSQLDrop.Length > 0 Then AddToSQL(vSQLDrop, utSQLTypes.utAlterOrCreate)
      'ORACLE = SQL to add defaults for new mandatory attributes
      'SQLSERVER = Additional ALTER statements
      For vIndex = 0 To UBound(vSQL2)
        If vSQL2(vIndex) <> "" Then AddToSQL(vSQL2(vIndex), utSQLTypes.utAlterOrCreate)
      Next
      'ORACLE = SQL to make new mandatory attributes not null
      'SQLSERVER = SQL to add defaults for new mandatory attributes
      For vIndex = 0 To UBound(vSQL3)
        If vSQL3(vIndex) <> "" Then AddToSQL(vSQL3(vIndex), utSQLTypes.utAlterOrCreate)
      Next
      'ORACLE = Not used
      'SQL SERVER = SQL to make new mandatory attributes not null
      For vIndex = 0 To UBound(vSQL4)
        If vSQL4(vIndex) <> "" Then AddToSQL(vSQL4(vIndex), utSQLTypes.utAlterOrCreate)
      Next
      'Attribute Comments for all Databases
      For vIndex = 0 To UBound(vSQLComments)
        If vSQLComments(vIndex) <> "" Then AddToSQL(vSQLComments(vIndex), utSQLTypes.utComments) 'utAlterOrCreate
      Next
    End Sub
    Public Function SQLCreateTable(ByVal pEnv As CDBEnvironment, ByVal pUpgradeTable As UpgradeTable) As String
      Dim vSQL As String = ""
      Dim vUpgradeAttr As UpgradeAttribute

      For Each vUpgradeAttr In pUpgradeTable.UpgradeAttributes.mvCol
        If vSQL.Length > 0 Then vSQL = vSQL & ", "
        If pEnv.Connection.IsSpecialColumn(vUpgradeAttr.Key) Then
          vSQL = vSQL & pEnv.Connection.DBSpecialCol("", vUpgradeAttr.Key) & " "
        Else
          vSQL = vSQL & vUpgradeAttr.Key & " "
        End If
        vSQL = vSQL & DetermineDataType(vUpgradeAttr) & " "
        If vUpgradeAttr.Nullable = UpgradeAttribute.NullOptions.noNullsInvalid Then vSQL = vSQL & "Not Null"
      Next
      vSQL = "CREATE TABLE " & IIf(pEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer, "dbo.", "").ToString & pUpgradeTable.Key & " (" & vSQL & ")"
      vSQL = vSQL & pEnv.GetTableSpaceInfo(pUpgradeTable.Key, False)
      SQLCreateTable = vSQL
    End Function

    Private Sub SetDefaultValues(ByVal pConn As CDBConnection, ByRef pSQL1() As String, ByRef pSQL2() As String)
      Dim vSQL As String = ""
      Dim vSet As Boolean
      Dim vUpgradeAttr As UpgradeAttribute
      Dim vAttrNotNull As UpgradeAttribute.NullOptions
      Dim vIRS As DataTable = New DataTable
      Dim vTemp As String

      For Each vUpgradeAttr In mvUpgradeTable.UpgradeAttributes.mvCol
        If (vUpgradeAttr.ToBeCreated Or vUpgradeAttr.StructureModified) And (Len(vUpgradeAttr.DefaultValue) > 0 And mvUpgradeTable.RecordCount > 0) Then
          If vUpgradeAttr.ToBeCreated Then
            If Not vSet Then
              vSQL = "UPDATE " + mvUpgradeTable.Key + " SET "
              vSet = True
            Else
              vSQL = vSQL + ", "
            End If
          Else
            vSQL = "UPDATE " + mvUpgradeTable.Key + " SET "
          End If
          vSQL = vSQL + vUpgradeAttr.Key + " = "
          Select Case vUpgradeAttr.DataType
            Case "integer", "longinteger", "smallint", "int", "decimal", "number", "double"
              vSQL = vSQL + vUpgradeAttr.DefaultValue
            Case "date", "time", "datetime"
              Dim vDate As Date
              If Date.TryParse(vUpgradeAttr.DefaultValue, vDate) Then
                vSQL = vSQL & pConn.SQLLiteral("", vDate)
                'vSQL = vSQL + "{d '" & vDate.ToString("yyyy-MM-dd") & "'}"
              End If
            Case Else
              If Len(vUpgradeAttr.DefaultValue) > 0 Then
                vSQL = vSQL & "'" & vUpgradeAttr.DefaultValue & "'"
              Else
                vSQL = vSQL & "NULL"
              End If
          End Select
          If vUpgradeAttr.StructureModified Then
            vSQL = vSQL & " WHERE " & vUpgradeAttr.Key & " IS NULL"
            pSQL1(UBound(pSQL1)) = vSQL
            ReDim Preserve pSQL1(UBound(pSQL1) + 1)
          End If
        End If
      Next
      If vSet Then pSQL1(UBound(pSQL1)) = vSQL

      vSQL = "ALTER TABLE " & mvUpgradeTable.Key & " "
      vSet = False
      For Each vUpgradeAttr In mvUpgradeTable.UpgradeAttributes.mvCol
        If (vUpgradeAttr.ToBeCreated Or vUpgradeAttr.StructureModified) And vUpgradeAttr.Nullable = UpgradeAttribute.NullOptions.noNullsInvalid And (mvUpgradeTable.RecordCount > 0 Or pConn.RDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer) Then
          If pConn.RDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer And vSet = True Then
            pSQL2(UBound(pSQL2)) = vSQL
            ReDim Preserve pSQL2(UBound(pSQL2) + 1)
            vSQL = "ALTER TABLE " & mvUpgradeTable.Key & " "
            vSet = False
          End If
          If vSet Then
            vSQL = vSQL & ", "
          Else
            If pConn.RDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer Then
              vSQL = vSQL & "ALTER COLUMN "
            Else
              vSQL = vSQL & "MODIFY ("
            End If
            vSet = True
          End If
          'before adding alter attribute, check for special col
          If pConn.IsSpecialColumn(vUpgradeAttr.Key) Then
            vTemp = pConn.DBSpecialCol("", vUpgradeAttr.Key)
          Else
            vTemp = vUpgradeAttr.Key
          End If
          vSQL = vSQL & vTemp & " " & DetermineDataType(vUpgradeAttr)
          vAttrNotNull = UpgradeAttribute.NullOptions.noNullsAllowed
          vIRS = mvUpgradeTable.Connection.GetAttributeNames(mvUpgradeTable.Key)
          If vIRS.Rows.Count > 0 Then
            For Each vRow As DataRow In vIRS.Rows
              If vRow.Item("Column_Name").ToString = vUpgradeAttr.Key Then
                vAttrNotNull = CType(IIf(vRow.Item("Nullable").ToString = "Y", 1, 0), UpgradeAttribute.NullOptions)
              End If
            Next
          End If
          If mvUpgradeTable.RecordCount = 0 Or vAttrNotNull = UpgradeAttribute.NullOptions.noNullsInvalid Then 'if no records in table OR attribute is already NOT NULL
            If pConn.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle Then
              If vAttrNotNull <> UpgradeAttribute.NullOptions.noNullsInvalid Then vSQL = vSQL & " Not Null"
            ElseIf pConn.RDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer Then
              vSQL = vSQL & " Not Null"
            End If
          ElseIf (pConn.RDBMSType <> CDBConnection.RDBMSTypes.rdbmsOracle) OrElse (pConn.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle AndAlso vUpgradeAttr.ToBeCreated = True AndAlso vUpgradeAttr.Nullable = UpgradeAttribute.NullOptions.noNullsInvalid) Then
            vSQL = vSQL & " Not Null"
          End If
        End If
      Next
      If vSet Then
        If pConn.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle Then vSQL = vSQL & ")"
        pSQL2(UBound(pSQL2)) = vSQL
      End If
    End Sub

    Public Sub New()
      mvCol = New Hashtable(StringComparer.InvariantCultureIgnoreCase)
    End Sub
  End Class



End Namespace
