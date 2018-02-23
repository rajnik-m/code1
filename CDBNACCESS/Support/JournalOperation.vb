Public Class JournalOperation

  Private mvActiveDays As Integer
  Private mvJournalType As String
  Private mvJournalTypeDesc As String
  Private mvOperation As String
  Private mvOperationDesc As String

  Public Sub New()

  End Sub

  Public Sub New(ByVal pJournalType As String, ByVal pJournalTypeDesc As String, ByVal pActiveDays As Integer, ByVal pOperation As String, ByVal pOperationDesc As String)
    mvJournalType = pJournalType
    mvJournalTypeDesc = pJournalTypeDesc
    mvActiveDays = pActiveDays
    mvOperation = pOperation
    mvOperationDesc = pOperationDesc
  End Sub

  Public Shared Function GetKey(ByVal pJournalType As String, ByVal pOperation As String) As String
    Return String.Format("{0}-{1}", pJournalType, pOperation)
  End Function

  Public ReadOnly Property Key() As String
    Get
      Return String.Format("{0}-{1}", mvJournalType, mvOperation)
    End Get
  End Property

  Public ReadOnly Property Description() As String
    Get
      Return mvOperationDesc & " " & mvJournalTypeDesc
    End Get
  End Property

  Public ReadOnly Property ActiveDays() As Integer
    Get
      Return mvActiveDays
    End Get
  End Property

  Public Function GetJournalDetailsTable(ByVal pEnv As CDBEnvironment, ByVal pContactType As Contact.ContactTypes, ByVal pJournalType As String, ByVal pUniqueID As String) As CDBDataTable
    Dim vMulti As Boolean
    Dim vValues() As String
    Dim vItems() As String
    Dim vIndex As Integer
    Dim vDetailRS As CDBRecordSet
    Dim vJournalRS As CDBRecordSet
    Dim vTable As New CDBDataTable

    If pUniqueID.Length > 0 And pJournalType.Length > 0 Then
      Dim vTypeRS As CDBRecordSet = pEnv.Connection.GetRecordSet("SELECT select_sql, org_select_sql, multi_selection, select1_attr, select2_attr, select3_attr FROM journal_types WHERE journal_type = '" & pJournalType & "'")
      Dim vRow As CDBDataRow
      If vTypeRS.Fetch Then
        Dim vContactNumber As String = ""
        Dim vAddressNumber As String = ""
        Dim vSelect1 As String
        Dim vSelect2 As String
        Dim vSelect3 As String
        If pJournalType = "MAIL" Or pJournalType = "CMAD" Then
          vSelect1 = pUniqueID
          vSelect2 = ""
          vSelect3 = ""
        Else
          Dim vJournalSql As String = "SELECT contact_number, address_number, select_1, select_2, select_3"
          vJournalRS = pEnv.Connection.GetRecordSet(vJournalSql & " FROM contact_journals WHERE contact_journal_number = " & pUniqueID)
          If vJournalRS.Fetch Then
            vContactNumber = CStr(vJournalRS.Fields(1).LongValue)
            vAddressNumber = CStr(vJournalRS.Fields(2).LongValue)
            vSelect1 = vJournalRS.Fields(3).Value
            vSelect2 = vJournalRS.Fields(4).Value
            vSelect3 = vJournalRS.Fields(5).Value
          Else
            vContactNumber = ""
            vAddressNumber = ""
            vSelect1 = ""
            vSelect2 = ""
            vSelect3 = ""
          End If
          vJournalRS.CloseRecordSet()
        End If

        Dim vSQL As String
        If pContactType = Contact.ContactTypes.ctcOrganisation Then
          'If displaying an organisation journal then use the org select if its there
          vSQL = JournalReplaceValues(vTypeRS.Fields("org_select_sql").Value, vSelect1, vSelect2, vSelect3, vContactNumber, vAddressNumber)
          If vSQL.Length = 0 Then vSQL = JournalReplaceValues(vTypeRS.Fields("select_sql").Value, vSelect1, vSelect2, vSelect3, vContactNumber, vAddressNumber)
        Else
          vSQL = JournalReplaceValues(vTypeRS.Fields("select_sql").Value, vSelect1, vSelect2, vSelect3, vContactNumber, vAddressNumber)
        End If
        vMulti = vTypeRS.Fields("multi_selection").Bool

        Dim vGetJournalDetails As Boolean = Not (vMulti)
        If vMulti Then
          'Attribute per record - ie from audit table
          vDetailRS = pEnv.Connection.GetRecordSet("SELECT data_values FROM amendment_history WHERE contact_journal_number = " & pUniqueID)
          If vDetailRS.Fetch() Then
            vItems = Split(vDetailRS.Fields(1).MultiLine, Chr(22))
            vTable.AddColumnsFromList("Item,OldValues,NewValues")
            vRow = vTable.AddRow
            Dim vRowNumber As Integer = 0
            Dim vColumn As String = ""        'Add fix for compiler warning
            For vIndex = 0 To vItems.Length - 1
              If vItems(vIndex) = "OLD" Then
                vColumn = "OldValues"
              ElseIf vItems(vIndex) = "NEW" Then
                vColumn = "NewValues"
                vRowNumber = 0
              ElseIf Mid(vItems(vIndex), 3) = "NEW" Then
                vColumn = "NewValues"
                vRowNumber = 0
              Else
                vValues = Split(vItems(vIndex), ":")
                If vValues.Length > 1 Then
                  If vRowNumber > vTable.Rows.Count - 1 Then
                    vRow = vTable.AddRow
                  Else
                    vRow = vTable.Rows.Item(vRowNumber)
                  End If
                  vRow.Item("Item") = StrConv(Replace(vValues(0), "_", " "), VbStrConv.ProperCase)
                  vRow.Item(vColumn) = Replace(vValues(1), vbCrLf, ", ")
                  vRowNumber = vRowNumber + 1
                End If
              End If
            Next
          ElseIf pJournalType = "PP" Then
            'Creation of a new Payment Plan (for example) does not create amendment history so get journal details instead
            vGetJournalDetails = True
          Else
            vTable.AddColumnsFromList("Item,Value")
            vRow = vTable.AddRow
            vRow.Item(1) = XLAT("Journal Details not found")
            vRow.Item(2) = XLAT("the information may have been removed")
          End If
          vDetailRS.CloseRecordSet()
        End If

        If vGetJournalDetails Then
          vTable.AddColumn("Item", CDBField.FieldTypes.cftCharacter)
          vTable.AddColumn("Value", CDBField.FieldTypes.cftMemo)          'Changed to memo to handle precis

          vJournalRS = pEnv.GetRecordSetParseSQL(vSQL)
          'Only expect one record
          If vJournalRS.Fetch Then
            GetJournalDetails(pEnv, vTable, vJournalRS, pJournalType)
          Else
            If pContactType = Contact.ContactTypes.ctcOrganisation Then
              'We used Org selection but this could be an employee line so use the contact select now
              vJournalRS.CloseRecordSet()
              vSQL = JournalReplaceValues(vTypeRS.Fields("select_sql").Value, vSelect1, vSelect2, vSelect3, vContactNumber, vAddressNumber)
              vJournalRS = pEnv.GetRecordSetParseSQL(vSQL)
              If vJournalRS.Fetch Then
                GetJournalDetails(pEnv, vTable, vJournalRS, pJournalType)
              Else
                vRow = vTable.AddRow
                vRow.Item(1) = XLAT("Journal Details not found")
                vRow.Item(2) = XLAT("the information may have been removed")
              End If
            Else
              vRow = vTable.AddRow
              vRow.Item(1) = XLAT("Journal Details not found")
              vRow.Item(2) = XLAT("the information may have been removed")
            End If
          End If
          vJournalRS.CloseRecordSet()
        End If
      Else
        vTable.AddColumnsFromList("Item,Value")
        vRow = vTable.AddRow
        vRow.Item(1) = XLATP1("Journal Type '%s' not recognised", pJournalType)
        vRow.Item(2) = XLAT("the information may have been removed")
      End If
      vTypeRS.CloseRecordSet()
    End If
    Return vTable
  End Function

  Private Sub GetJournalDetails(ByVal pEnv As CDBEnvironment, ByVal pTable As CDBDataTable, ByVal pRS As CDBRecordSet, ByVal pJournalType As String)
    Dim vDescSQL As String
    Dim vValue As String
    Dim vPos As Integer
    Dim vPosEnd As Integer
    Dim vDescAttr As String
    Dim vDispRS As CDBRecordSet
    Dim vDescRS As CDBRecordSet
    Dim vRow As CDBDataRow

    'Get the attributes details to display

    vDispRS = pEnv.Connection.GetRecordSet("SELECT attribute_name, control_caption, desc_sql FROM journal_type_controls WHERE journal_type = '" & pJournalType & "' ORDER BY journal_type, sequence_number")
    While vDispRS.Fetch
      vDescSQL = vDispRS.Fields("desc_sql").Value
      vValue = pRS.Fields((vDispRS.Fields("attribute_name").Value)).Value
      If Len(vValue) > 0 And Len(vDescSQL) > 0 Then
        If UCase(Mid(vDescSQL, 1, 6)) = "SELECT" Then
          vDescSQL = Replace(vDescSQL, "?", vValue)
          vDescSQL = Replace(vDescSQL, "#", vValue)
          vPos = InStr(1, vDescSQL, " ") + 1
          vPosEnd = InStr(vPos, vDescSQL, " ")
          vDescAttr = Mid(vDescSQL, vPos, vPosEnd - vPos)
          vDescRS = pEnv.GetRecordSetParseSQL(vDescSQL)
          If vDescRS.Fetch Then
            vValue = vDescRS.Fields(vDescAttr).Value
          Else
            vValue = ""
          End If
          vDescRS.CloseRecordSet()
        Else
          'Swap codes for text
          vPos = InStr(vDescSQL, vValue & "-")
          If vPos > 0 Then
            vPos = vPos + Len(vValue & "-")
            vPosEnd = InStr(vPos, vDescSQL, "|")
            If vPosEnd = 0 Then
              vPosEnd = Len(vDescSQL) + 1
            End If
            vValue = Mid(vDescSQL, vPos, vPosEnd - vPos)
          End If
        End If
      End If
      If Len(vValue) > 0 Then
        vRow = pTable.AddRow
        vRow.Item(1) = vDispRS.Fields("control_caption").Value
        vRow.Item(2) = vValue
      End If
    End While
    vDispRS.CloseRecordSet()
  End Sub

  Private Function JournalReplaceValues(ByVal pSQL As String, ByVal pSelect1 As String, ByVal pSelect2 As String, ByVal pSelect3 As String, ByVal pContactNumber As String, ByVal pAddressNumber As String) As String
    Dim vSQL As String = pSQL
    Dim vItems As New Dictionary(Of String, String)
    vItems.Add("^select1_attr^", pSelect1)
    vItems.Add("^select2_attr^", pSelect2)
    vItems.Add("^select3_attr^", pSelect3)
    vItems.Add("^contact_number^", pContactNumber)
    vItems.Add("^address_number^", pAddressNumber)
    vItems.Add("^organisation_number^", pContactNumber)
    For Each vItem As KeyValuePair(Of String, String) In vItems
      If vSQL.Contains(vItem.Key) Then
        If vItem.Value = "" Then
          vSQL = vSQL.Replace(vItem.Key, "0")
        Else
          vSQL = vSQL.Replace(vItem.Key, vItem.Value)
        End If
      End If
    Next
    Return vSQL
  End Function

End Class
