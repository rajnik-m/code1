Imports System.Data
Imports System.Linq

Namespace Access

  Public Class CDBDataTable

    Private mvRows As New List(Of CDBDataRow)
    Private mvColumns As New CollectionList(Of CDBDataColumn)(1)
    Private mvEnv As CDBEnvironment
    Private mvTimeout As Integer 'Seconds
    Private mvStart As Date
    Private mvMaxRows As Integer
    Private mvStartRow As Integer
    Private mvCheckAccess As Boolean
    Private mvListManagerViewSelection As Boolean

    Public Structure SortSpecification
      Dim Column As String
      Dim Descending As Boolean
    End Structure

    Public ReadOnly Property Debug_DataTable As DataTable
      Get
        Return ConvertToDataTable()
      End Get
    End Property

    Public Function ConvertToDataTable() As DataTable

      Dim vTable As DataTable = New DataTable("DataTable")

      For Each vCol As CARE.Access.CDBDataColumn In Me.Columns
        If vCol.FieldType = CDBField.FieldTypes.cftNumeric Then
          vTable.Columns.Add(vCol.Name, System.Type.GetType("System.Decimal"))
        ElseIf vCol.FieldType = CDBField.FieldTypes.cftDate Then
          vTable.Columns.Add(vCol.Name, System.Type.GetType("System.DateTime"))
        ElseIf vCol.FieldType = CDBField.FieldTypes.cftInteger Or vCol.FieldType = CDBField.FieldTypes.cftIdentity Or vCol.FieldType = CDBField.FieldTypes.cftLong Then
          vTable.Columns.Add(vCol.Name, System.Type.GetType("System.Int32"))
        Else
          vTable.Columns.Add(vCol.Name)
        End If
      Next

      For Each vRow As CARE.Access.CDBDataRow In Me.Rows
        Dim NewRow As DataRow = vTable.NewRow()

        For Each vCol As CARE.Access.CDBDataColumn In Me.Columns
          If vRow.Item(vCol.Name) IsNot Nothing AndAlso vRow.Item(vCol.Name).Length > 0 Then NewRow(vCol.Name) = vRow.Item(vCol.Name)
        Next

        vTable.Rows.Add(NewRow)
      Next

      Return vTable

    End Function
    Public Function RowsAsCommaSeperated(ByVal pDataTable As CDBDataTable, ByVal pColumnName As String) As String
      Dim vResult As String = String.Empty
      vResult = String.Join(",", (From row In pDataTable.ConvertToDataTable().AsEnumerable Select row(pColumnName)).ToArray)
      Return vResult
    End Function
    Public Sub New()

    End Sub

    Public Sub New(ByVal pEnv As CDBEnvironment, ByVal pSQLStatement As SQLStatement)
      mvEnv = pEnv
      FillFromSQL(mvEnv, pSQLStatement, True)
    End Sub

    Public Sub New(ByVal pDataTable As DataTable)
      For Each vColumn As DataColumn In pDataTable.Columns
        Dim vFieldType As CDBField.FieldTypes
        Select Case vColumn.DataType.Name
          Case "Integer", "Int32"
            vFieldType = CDBField.FieldTypes.cftInteger
          Case "DateTime"
            vFieldType = CDBField.FieldTypes.cftTime
          Case "Decimal"
            vFieldType = CDBField.FieldTypes.cftNumeric
          Case Else
            vFieldType = CDBField.FieldTypes.cftCharacter
        End Select
        AddColumn(vColumn.ColumnName, vFieldType)
      Next
      For Each vRow As DataRow In pDataTable.Rows
        Me.AddRowFromItems(vRow.ItemArray)
      Next
    End Sub

    Public Sub SetEnvironment(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
    End Sub

    Public Property MaximumRows As Integer
      Get
        Return mvMaxRows
      End Get
      Set(ByVal pValue As Integer)
        mvMaxRows = pValue
      End Set
    End Property

    Public Property StartRow As Integer
      Get
        Return mvStartRow
      End Get
      Set(ByVal pValue As Integer)
        mvStartRow = pValue
      End Set
    End Property

    Public Property CheckAccess As Boolean
      Get
        Return mvCheckAccess
      End Get
      Set(ByVal pValue As Boolean)
        mvCheckAccess = pValue
      End Set
    End Property

    ''' <summary>Gets or sets boolean flag indicating whether this is selecting data for List Manager</summary>
    Public Property ListManagerViewSelection() As Boolean
      Get
        Return mvListManagerViewSelection
      End Get
      Set(ByVal pValue As Boolean)
        mvListManagerViewSelection = pValue
      End Set
    End Property

    Public Sub RemoveDuplicateRows(ByVal pColumnName As String)
      Dim vRemoved As Boolean
      Dim vLastValue As String = ""
      Do                              'Assumes all rows will have a value in the given column
        vRemoved = False
        For Each vDataRow As CDBDataRow In mvRows
          If vDataRow.Item(pColumnName) = vLastValue Then
            mvRows.Remove(vDataRow)
            vRemoved = True
            Exit For
          Else
            vLastValue = vDataRow.Item(pColumnName)
          End If
        Next
        vLastValue = ""
        CheckTimeout()
      Loop While vRemoved
    End Sub
    Public Sub RemoveFullyDuplicatedRows()
      Dim vSortSpec(0) As SortSpecification
      RemoveFullyDuplicatedRows(vSortSpec)
    End Sub

    'BR20266 Function extended so that a sort specification can be passed in, in order to Remove Rows based on MultipleColumns
    'Function assumes Datatable has already been sorted. Shouldln't be called for 1 column.
    Public Sub RemoveFullyDuplicatedRows(ByVal pspecification() As SortSpecification)
      Dim vRemove As Boolean
      Dim vMatchDataRow As CDBDataRow
      Dim vMatchIndex As Integer = -1
      Dim vMatchNextRow As Boolean = True 'Set to indicate that a new data row is required for comparing other data rows against.
      Dim vHaveMatchRow As Boolean = True 'Set to indicate that there's a data row for other data rows to compare against.
      Dim vLastIndex As Integer = -1
      Dim vResumeAfterIndex As Integer = -1 'Set when a row has been removed to indicate where to resume processing in the next iteration.
      Dim vExit As Boolean = False

      Do
        For Each vDataRow As CDBDataRow In mvRows
          If vResumeAfterIndex = -1 Then  'i.e. process is not resuming after a row removal.
            If vMatchNextRow = True Then
              vMatchDataRow = vDataRow
              vMatchNextRow = False
              vHaveMatchRow = True
            ElseIf vHaveMatchRow Then
              'Now check if all columns from this row have the same values as on VMatchDataRow
              vRemove = True
              Dim vColumnSpecifier As Integer = UBound(pspecification)
              If pspecification.Length > 1 Then
                While vColumnSpecifier >= 0
                  Dim vColIndex As Integer = mvColumns(pspecification(vColumnSpecifier).Column).Index
                  If vDataRow.Item(vColIndex) <> vMatchDataRow.Item(vColIndex) Then
                    vRemove = False
                  End If
                  vColumnSpecifier -= 1
                End While
              Else
                For vColIndex As Integer = 0 To mvColumns.Count - 1
                  If vDataRow.Item(vColIndex) <> vMatchDataRow.Item(vColIndex) Then
                    vRemove = False
                    Exit For
                  End If
                Next vColIndex
              End If
              If vRemove = True Then
                'The rows are the same, remove vDataRow
                mvRows.Remove(vDataRow)
                'The number of CDBDateRow in mvRows has changed, therefore leaving For-loop now before program crash
                ' but first make a note of the previous vDataRow Index. This will be used so that the process can resume from that point. 
                vResumeAfterIndex = vLastIndex
                vLastIndex = -1
                Exit For
              End If
            ElseIf vDataRow.Index = vMatchDataRow.Index Then
              vMatchNextRow = True
            End If
            vLastIndex = vDataRow.Index
          Else
            If vDataRow.Index = vResumeAfterIndex Then
              'Process has reached the point where the previous iteration left. Comparison can continue. 
              vResumeAfterIndex = -1
            End If
          End If
        Next
        CheckTimeout()

        If vResumeAfterIndex = -1 Then
          'There are no more duplicate rows for vMatchDataRow
          'Set indicators to move to next row for other rows to compare with 
          vHaveMatchRow = False
        End If
        If vMatchNextRow Then
          'This flag not set so there is no next row for others to compare with 
          vExit = True
        End If
      Loop Until vExit
    End Sub

    Public Sub RemoveRow(ByVal pRow As CDBDataRow)
      mvRows.Remove(pRow)
    End Sub

    Public Sub ReOrderRowsByColumn(ByVal pIndex As String, Optional ByVal pDescending As Boolean = False)
      Dim SS(0) As SortSpecification
      SS(0).Column = pIndex
      SS(0).Descending = pDescending
      ReOrderRowsByMultipleColumns(SS)
    End Sub

    Public Sub ReOrderRowsByMultipleColumns(ByVal pSpecification() As SortSpecification)
      'Multiple column sort.
      'pSpecification is an array of Column Names and Sort Directions. The lower
      'indicies in the array are the more significant sort keys.
      Dim vColumnSpecifier As Integer = UBound(pSpecification)
      Dim vColl As New List(Of CDBDataRow)
      While vColumnSpecifier >= 0
        Dim vColumnIndex As Integer = mvColumns(pSpecification(vColumnSpecifier).Column).Index
        Dim vDescending As Boolean = pSpecification(vColumnSpecifier).Descending

        Dim vItemValue As String
        Dim vRowValue As String
        Dim vIndex As Integer
        Select Case mvColumns(vColumnIndex).FieldType
          Case CDBField.FieldTypes.cftCharacter, CDBField.FieldTypes.cftMemo
            For Each vRow As CDBDataRow In mvRows
              vRowValue = vRow.Item(vColumnIndex)
              For vIndex = 0 To vColl.Count - 1
                vItemValue = vColl(vIndex).Item(vColumnIndex)
                If vDescending Then
                  If vRowValue > vItemValue Then Exit For
                Else
                  If vRowValue < vItemValue Then Exit For
                End If
              Next
              If vIndex <= vColl.Count Then
                vColl.Insert(vIndex, vRow)
              Else
                vColl.Add(vRow)
              End If
            Next

          Case CDBField.FieldTypes.cftInteger, CDBField.FieldTypes.cftLong, CDBField.FieldTypes.cftNumeric
            For Each vRow As CDBDataRow In mvRows
              vRowValue = vRow.Item(vColumnIndex)
              For vIndex = 0 To vColl.Count - 1
                vItemValue = vColl(vIndex).Item(vColumnIndex)
                If vDescending Then
                  If Val(vRowValue) > Val(vItemValue) Then Exit For
                Else
                  If Val(vRowValue) < Val(vItemValue) Then Exit For
                End If
              Next
              If vIndex <= vColl.Count Then
                vColl.Insert(vIndex, vRow)
              Else
                vColl.Add(vRow)
              End If
            Next

          Case CDBField.FieldTypes.cftDate, CDBField.FieldTypes.cftTime
            Dim vRowDate As Date
            For Each vRow As CDBDataRow In mvRows
              vRowValue = vRow.Item(vColumnIndex)
              Dim vRowIsDate As Boolean
              If vRowValue.Length > 0 Then
                vRowIsDate = True
                vRowDate = CDate(vRowValue)
              Else
                vRowIsDate = False
              End If
              For vIndex = 0 To vColl.Count - 1
                vItemValue = vColl(vIndex).Item(vColumnIndex)
                If vRowIsDate And vItemValue.Length > 0 Then
                  If vDescending Then
                    If vRowDate > CDate(vItemValue) Then Exit For
                  Else
                    If vRowDate < CDate(vItemValue) Then Exit For
                  End If
                End If
              Next
              If vIndex <= vColl.Count Then
                vColl.Insert(vIndex, vRow)
              Else
                vColl.Add(vRow)
              End If
            Next
        End Select
        mvRows = vColl
        vColl = New List(Of CDBDataRow)
        vColumnSpecifier -= 1
      End While
    End Sub

    Public Sub RestrictRows(ByVal pStartRow As Integer, ByVal pNumberOfRows As Integer)
      If pStartRow <= mvRows.Count Then
        Dim vRows As New List(Of CDBDataRow)
        Dim vEndRow As Integer = pStartRow + (pNumberOfRows - 1)
        If vEndRow >= mvRows.Count Then vEndRow = mvRows.Count - 1
        For vIndex As Integer = pStartRow To vEndRow
          vRows.Add(mvRows(vIndex))
        Next
        mvRows = vRows
      End If
    End Sub

    Public ReadOnly Property Rows() As List(Of CDBDataRow)
      Get
        'NOTE In the .NET code Rows start at zero whereas columns start at 1 eek!!!
        Return mvRows
      End Get
    End Property

    Public ReadOnly Property FirstRow() As CDBDataRow
      Get
        If mvRows.Count > 0 Then
          Return mvRows(0)
        Else
          Return Nothing
        End If
      End Get
    End Property

    Public Function PreviousRow(ByVal pCurrentRow As CDBDataRow) As CDBDataRow
      'Return previous item in the collection before the current item
      Dim vPrevRow As CDBDataRow = Nothing
      Dim vFound As Boolean
      'If pCurrentItem Is Nothing then the last item is required
      For Each vRow As CDBDataRow In mvRows
        If Not (pCurrentRow Is Nothing) Then
          If vRow Is pCurrentRow Then vFound = True
        End If
        If vFound Then Exit For
        vPrevRow = vRow
      Next vRow
      Return vPrevRow
    End Function

    Public ReadOnly Property Columns() As CollectionList(Of CDBDataColumn)
      Get
        Return mvColumns
      End Get
    End Property

    Public Function AddColumn(ByVal pName As String, ByVal pType As CDBField.FieldTypes) As CDBDataColumn
      Dim vRow As CDBDataRow

      If mvColumns.ContainsKey(pName) Then RaiseError(DataAccessErrors.daeColumnDefinedAlready, pName)
      Dim vCol As New CDBDataColumn(pName, pType, mvColumns.Count + 1)
      mvColumns.Add(pName, vCol)
      For Each vRow In mvRows
        vRow.ResetColumnsCount()
      Next
      Return vCol
    End Function

    Public Sub AddColumnsFromList(ByVal pList As String)
      Dim vName As String
      Dim vPos As Integer
      Dim vNames() As String = pList.Split(","c)
      For vIndex As Integer = 0 To vNames.Length - 1
        vName = vNames(vIndex)
        vPos = vName.IndexOf(".")
        If vPos >= 0 Then vName = vName.Substring(0, vPos + 1)
        AddColumn(vName, CDBField.FieldTypes.cftCharacter)
      Next
    End Sub

    Public Function ColumnNames() As String
      Dim vNames As New StringBuilder
      Dim vAddSeparator As Boolean

      For Each vCol As CDBDataColumn In mvColumns
        If vAddSeparator Then vNames.Append(",")
        vNames.Append(vCol.Name)
        vAddSeparator = True
      Next
      Return vNames.ToString
    End Function

    Public Function InsertRow(pInsertIDX As Integer) As CDBDataRow
      Dim vRow As CDBDataRow
      vRow = New CDBDataRow(mvColumns, mvRows.Count + 1)
      mvRows.Insert(pInsertIDX, vRow)
      Return vRow
    End Function

    Public Function AddRow() As CDBDataRow
      Dim vRow As CDBDataRow
      vRow = New CDBDataRow(mvColumns, mvRows.Count + 1)
      mvRows.Add(vRow)
      Return vRow
    End Function

    Public Function AddRowFromItems(ByVal ParamArray pItems() As String) As CDBDataRow
      Dim vRow As New CDBDataRow(mvColumns, mvRows.Count + 1)
      Dim vIndex As Integer = 1
      For Each vItem As String In pItems
        vRow.Item(vIndex) = vItem
        vIndex += 1
      Next
      mvRows.Add(vRow)
      Return vRow
    End Function

    Public Function AddRowFromItems(ByVal pItems() As Object) As CDBDataRow
      Dim vRow As New CDBDataRow(mvColumns, mvRows.Count + 1)
      Dim vIndex As Integer = 1
      For Each vItem As Object In pItems
        vRow.Item(vIndex) = vItem.ToString
        vIndex += 1
      Next
      mvRows.Add(vRow)
      Return vRow
    End Function

    Public Sub CheckTimeout()
      If mvTimeout > 0 Then
        If Not mvStart = Nothing Then
          If DateDiff(DateInterval.Second, mvStart, System.DateTime.Now) > mvTimeout Then
            RaiseError(DataAccessErrors.daeProcessTimeout)
          End If
        End If
      End If
    End Sub

    Public Sub FillFromSQLDB(ByVal pEnv As CDBEnvironment, ByVal pDatabase As String, ByVal pSQL As String, Optional ByVal pItems As String = "", Optional ByVal pAdditionalItems As String = "")
      mvStart = System.DateTime.Now
      mvEnv = pEnv
      Dim vRecordSet As CDBRecordSet = pEnv.GetConnection(pDatabase).GetRecordSet(pSQL)
      InitFromRecordSet(pEnv, vRecordSet, pItems, pAdditionalItems, False)
    End Sub

    Public Sub FillFromSQLDONOTUSE(ByVal pEnv As CDBEnvironment, ByVal pSQL As String, Optional ByVal pItems As String = "", Optional ByVal pAdditionalItems As String = "")
      Dim vSQLStatement As New SQLStatement(pEnv.Connection, pSQL)
      FillFromSQL(pEnv, vSQLStatement, pItems, pAdditionalItems)
    End Sub

    Public Sub FillFromSQL(ByVal pEnv As CDBEnvironment, ByVal pSQLStatement As SQLStatement, ByVal pUseStandardNames As Boolean)
      FillFromSQL(pEnv, pSQLStatement, "", "")
      If pUseStandardNames Then SetStandardColumnNames()
    End Sub

    Public Sub FillFromSQL(ByVal pEnv As CDBEnvironment, ByVal pSQLStatement As SQLStatement)
      FillFromSQL(pEnv, pSQLStatement, "", "")
    End Sub

    Public Sub FillFromSQL(ByVal pEnv As CDBEnvironment, ByVal pSQLStatement As SQLStatement, ByVal pItems As String)
      FillFromSQL(pEnv, pSQLStatement, pItems, "")
    End Sub

    Public Sub FillFromSQL(ByVal pEnv As CDBEnvironment, ByVal pSQLStatement As SQLStatement, ByVal pItems As String, ByVal pAdditionalItems As String)
      FillFromSQL(pEnv, pSQLStatement, pItems, pAdditionalItems, False)
    End Sub

    Public Sub FillFromSQL(ByVal pEnv As CDBEnvironment, ByVal pSQLStatement As SQLStatement, ByVal pItems As String, ByVal pAdditionalItems As String, ByVal pPopulateIfNoRecords As Boolean)
      mvStart = System.DateTime.Now
      mvEnv = pEnv
      pSQLStatement.Timeout = mvTimeout
      InitFromRecordSet(pEnv, pSQLStatement.GetRecordSet, pItems, pAdditionalItems, pPopulateIfNoRecords)
    End Sub

    Private Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pItems As String, ByVal pAdditionalItems As String, ByVal pPopulateIfNoRecords As Boolean)
      Dim vHasAccess As Boolean
      Dim vAccessLevel As String = ""
      mvStart = System.DateTime.Now
      mvEnv = pEnv
      If mvStartRow > 0 Then
        While mvStartRow > 0
          pRecordSet.Fetch()
          mvStartRow -= 1
          If pRecordSet.Status = False Then Exit While
        End While
      End If
      If pRecordSet.Fetch() Then
        If pAdditionalItems.Length > 0 Then
          Dim vSBItems As New StringBuilder
          If pItems.Length = 0 Then
            For vIndex As Integer = 1 To pRecordSet.Fields.Count
              vSBItems.Append(pRecordSet.Fields(vIndex).Name)
              vSBItems.Append(",")
            Next
          Else
            vSBItems.Append(pItems)
            vSBItems.Append(",")
          End If
          vSBItems.Append(pAdditionalItems)
          pItems = vSBItems.ToString
        End If
        Dim vItems() As String = Nothing
        Dim vUseItems As Boolean
        If pItems.Length > 0 Then
          vUseItems = True
          vItems = pItems.Split(","c)
        End If

        If mvColumns.Count = 0 Then
          If vUseItems Then
            For vIndex As Integer = 0 To vItems.Length - 1
              AddColumn(vItems(vIndex), CDBField.FieldTypes.cftCharacter)
            Next
          Else
            For Each vField As CDBField In pRecordSet.Fields
              AddColumn(vField.Name, vField.FieldType)
            Next
          End If
        End If

        If pRecordSet.Fields.Exists("AccessLevel") Then
          vHasAccess = True
          vAccessLevel = "AccessLevel"
        ElseIf pRecordSet.Fields.Exists("accesslevel") Then
          vAccessLevel = "accesslevel"
          vHasAccess = True
        End If

        Dim vAttrName As String
        For vIndex As Integer = 1 To mvColumns.Count
          If vUseItems Then
            vAttrName = vItems(vIndex - 1)
          Else
            vAttrName = pRecordSet.Fields(vIndex).Name
          End If
          If vHasAccess And vAttrName = "ownership_access_level AS AccessLevel" Then
            vAttrName = vAccessLevel
            vItems(vIndex - 1) = vAttrName
          End If
          mvColumns(vIndex).AttributeName = vAttrName 'Will strip off any leading prefix
          vAttrName = mvColumns(vIndex).AttributeName
          If pRecordSet.Fields.Exists(vAttrName) Then
            mvColumns(vIndex).FieldType = pRecordSet.Fields(vAttrName).FieldType
          ElseIf pRecordSet.Fields.Exists(vAttrName.ToLower) Then
            mvColumns(vIndex).FieldType = pRecordSet.Fields(vAttrName.ToLower).FieldType
          Else
            Select Case vAttrName
              Case "DISTINCT_DOCUMENT_NUMBER", "DISTINCT_CONTACT_NUMBER", "DISTINCT_ORGANISATION_NUMBER", "DISTINCT_PAYMENT_PLAN_NUMBER", "LINE_TYPE_NUMBER", "DISTINCT_EVENT_BOOKING_NUMBER", "DISTINCT_EVENT_NUMBER"
                mvColumns(vIndex).FieldType = CDBField.FieldTypes.cftLong
            End Select
          End If
        Next

        Dim vValue As String = ""
        Dim vLastValue As String = ""
        Do
          Dim vRow As CDBDataRow = AddRow()
          For vIndex As Integer = 1 To mvColumns.Count
            If vUseItems Then
              Dim vItemName As String = vItems(vIndex - 1)
              Dim vPos As Integer = vItemName.IndexOf(".")
              If vPos >= 0 Then vItemName = vItemName.Substring(vPos + 1)
              vItemName = vItemName.Trim()
              Select Case vItemName
                Case "ACCESS"
                  Dim vAccess As Boolean = True
                  If pRecordSet.Fields("created_by").Value = mvEnv.User.Logname Then
                    If pRecordSet.Fields("creator_header").Bool = False Then vAccess = False
                  ElseIf pRecordSet.Fields("department").Value = mvEnv.User.Department Then
                    If pRecordSet.Fields("department_header").Bool = False Then vAccess = False
                  Else
                    If pRecordSet.Fields("public_header").Bool = False Then vAccess = False
                  End If
                  vValue = BooleanString(vAccess)
                Case "DISTINCT_PRODUCT_LINE"
                  vValue = pRecordSet.Fields("batch_number").Value & "," & pRecordSet.Fields("transaction_number").Value & "," & pRecordSet.Fields("line_number").Value
                  If vValue = vLastValue Then
                    mvRows.Remove(mvRows(mvRows.Count - 1))
                    mvRows(mvRows.Count - 1).Item(vIndex) = "N/A"
                    Exit For
                  Else
                    vLastValue = vValue
                    vValue = pRecordSet.Fields("product_desc").Value
                  End If
                Case "DISTINCT_DESPATCH_TRANSACTION"
                  ' BR 11665 - Need to include a further condition to multiple dispatch notes for a transaction will show
                  vValue = pRecordSet.Fields("transaction_number").Value & "," & pRecordSet.Fields("batch_number").Value & "," & pRecordSet.Fields("picking_list_number").Value & "," & pRecordSet.Fields("despatch_note_number").Value
                  If vValue = vLastValue Then
                    mvRows.Remove(mvRows(mvRows.Count - 1))
                    Exit For
                  Else
                    vLastValue = vValue
                    vValue = pRecordSet.Fields("transaction_number").Value
                  End If
                Case "DISTINCT_CONTACT_NUMBER"
                  vValue = pRecordSet.Fields("contact_number").Value
                  If vValue = vLastValue Then
                    mvRows.Remove(mvRows(mvRows.Count - 1))
                    Exit For
                  Else
                    vLastValue = vValue
                  End If
                Case "DISTINCT_ORGANISATION_NUMBER"
                  vValue = pRecordSet.Fields("organisation_number").Value
                  If vValue = vLastValue Then
                    mvRows.Remove(mvRows(mvRows.Count - 1))
                    Exit For
                  Else
                    vLastValue = vValue
                  End If
                Case "DISTINCT_DOCUMENT_NUMBER"
                  vValue = pRecordSet.Fields("communications_log_number").Value
                  If vValue = vLastValue Then
                    mvRows.Remove(mvRows(mvRows.Count - 1))
                    Exit For
                  Else
                    vLastValue = vValue
                  End If
                Case "DISTINCT_PAYMENT_PLAN_NUMBER"
                  vValue = pRecordSet.Fields("order_number").Value
                  If vValue = vLastValue Then
                    mvRows.Remove(mvRows(mvRows.Count - 1))
                    Exit For
                  Else
                    vLastValue = vValue
                  End If
                Case "DISTINCT_EVENT_BOOKING_NUMBER"
                  vValue = pRecordSet.Fields("booking_number").Value
                  If vValue = vLastValue Then
                    mvRows.Remove(mvRows(mvRows.Count - 1))
                    Exit For
                  Else
                    vLastValue = vValue
                  End If
                Case "DISTINCT_EVENT_NUMBER"
                  vValue = pRecordSet.Fields("event_number").Value
                  If vValue = vLastValue Then
                    mvRows.Remove(mvRows(mvRows.Count - 1))
                    Exit For
                  Else
                    vLastValue = vValue
                  End If
                Case "ENCRYPTED_CC_NUMBER"
                  vValue = String.Empty
                Case "EXPECTED_FRACTION"
                  If pRecordSet.Fields("expected_fraction_quantity").Value.Length > 0 Then
                    vValue = pRecordSet.Fields("expected_fraction_quantity").Value & "/" & pRecordSet.Fields("expected_fraction_division").Value
                  Else
                    vValue = ""
                  End If
                Case "BOOKING_STATUS_DESC"
                  vValue = EventBooking.GetBookingStatusDescription(pRecordSet.Fields("booking_status").Value)
                Case "ADDRESS_LINE"
                  Dim vAddress As New Address(pEnv)
                  vAddress.InitFromRecordSetCountry(pRecordSet)
                  vValue = vAddress.AddressLine
                Case "ADDRESS_MULTI_LINE"
                  Dim vAddress As New Address(pEnv)
                  vAddress.InitFromRecordSetCountry(pRecordSet)
                  vValue = vAddress.AddressMultiLine
                Case "TOWN_ADDRESS_LINE"
                  Dim vAddress As New Address(pEnv)
                  vAddress.InitFromRecordSetCountry(pRecordSet)
                  vValue = vAddress.TownAddressLine
                Case "CONTACT_GROUP"
                  vValue = pRecordSet.Fields("contact_group").Value
                  If vValue.Length = 0 Then vValue = ContactGroup.DefaultGroupCode
                Case "ORGANISATION_GROUP"
                  vValue = pRecordSet.Fields("organisation_group").Value
                  If vValue.Length = 0 Then vValue = OrganisationGroup.DefaultGroupCode
                Case "CONTACT_NAME"
                  Dim vContact As New Contact(pEnv)
                  vContact.InitFromRecordSetName(pRecordSet)
                  vValue = vContact.Name
                Case "PAYEE_CONTACT_NAME"
                  Dim vContact As New Contact(pEnv)
                  vContact.InitFromRecordSetName(pRecordSet, "payee")

                  vValue = vContact.Name
                Case "CONTACT_TELEPHONE"
                  Dim vContact As New Contact(pEnv)
                  vContact.InitFromRecordSetPhone(pRecordSet)
                  vValue = vContact.PhoneNumber
                Case "PHONE_NUMBER"
                  Dim vCommunication As New Communication(mvEnv)
                  vCommunication.InitFromRecordSetPhone(pRecordSet)
                  vValue = vCommunication.PhoneNumber
                Case "CONTACT_TYPE_1", "CONTACT_TYPE_2"
                  vValue = "C"
                Case "EVENT_TYPE"
                  vValue = "E"
                Case "EXAM_UNIT_TYPE"
                  vValue = "U"
                Case "EXAM_CENTRE_TYPE"
                  vValue = "N"
                Case "EXAM_CENTRE_UNIT"
                  vValue = "X"
                Case "MEETING_TYPE"
                  vValue = "M"
                Case "LEGACY_TYPE"
                  vValue = "L"
                Case "FUND_LINK_NAME"
                  vValue = "FR-" & pRecordSet.Fields("fundraising_request_number").Value
                  If pRecordSet.Fields("scheduled_payment_number").IntegerValue > 0 Then
                    vValue = vValue & "/FPS-" & pRecordSet.Fields("scheduled_payment_number").Value
                  End If
                Case "TRANSACTION_TYPE_1"
                  vValue = "T"
                Case "ORGANISATION_TELEPHONE"
                  Dim vOrganisation As New Organisation(pEnv)
                  vOrganisation.InitFromRecordSetPhone(pRecordSet)
                  vValue = vOrganisation.PhoneNumber
                Case "ORGANISATION_TYPE_1", "ORGANISATION_TYPE_2"
                  vValue = "O"
                Case "LINK_TYPE"
                  vValue = CommunicationsLogLink.GetLinkTypeDescription(pRecordSet.Fields("link_type").Value)
                Case "ACTION_LINK_TYPE"
                  vValue = Action.ActionLinkTypeDescription(pRecordSet.Fields("type").Value)
                Case "MEETING_LINK_TYPE"
                  'NYI("MEETING_LINK_TYPE")    'TODO MEETING_LINK_TYPE
                  '  vValue = pEnv.GetMeetingLinkTypeDescription(pRecordSet.Fields("link_type").Value)
                  If pRecordSet.Fields("link_type").Value = "W" Then vValue = "With"
                  If pRecordSet.Fields("link_type").Value = "C" Then vValue = ProjectText.String15805 'Copied To
                  If pRecordSet.Fields("link_type").Value = "R" Then vValue = ProjectText.String15807 'Related To
                Case "DOCUMENT_NAME"
                  Dim vCommsLog As New CommunicationsLog(mvEnv)
                  vCommsLog.InitFromRecordSet(pRecordSet)
                  vValue = vCommsLog.Name
                Case "LINE_TYPE_NUMBER"
                  vValue = pRecordSet.Fields("member_number").Value
                  If vValue.Length = 0 Then vValue = pRecordSet.Fields("covenant_number").Value
                  If vValue.Length = 0 Then vValue = pRecordSet.Fields("order_number").Value
                Case "PAYMENT_PLAN_TYPE" 'P11 form
                  Select Case pRecordSet.Fields("order_type").Value
                    Case "M"
                      vValue = ProjectText.String15808 'Membership
                    Case "O"
                      vValue = ProjectText.String15809 'Other
                    Case "B"
                      vValue = ProjectText.String15810 'Standing Order
                    Case "D"
                      vValue = ProjectText.String15811 'Direct Debit
                    Case "A"
                      vValue = ProjectText.String15812 'Credit Card Authority
                    Case "C"
                      vValue = ProjectText.String15813 'Covenant
                    Case "L"
                      vValue = ProjectText.String33022  'Loan
                    Case Else
                      vValue = pRecordSet.Fields("order_type").Value
                  End Select
                Case "PHONEBOOK_PHONE"
                  vValue = CStr(pRecordSet.Fields("std_code").Value & " ").TrimStart(" "c)
                  vValue = vValue & pRecordSet.Fields("telephone").Value
                Case "PHONEBOOK_ADDRESS"
                  vValue = pRecordSet.Fields("address1").Value
                  If pRecordSet.Fields("address2").Value.Length > 0 Then vValue = vValue & vValue
                  vValue = vValue & ", " & pRecordSet.Fields("postcode").Value & " "
                  vValue = vValue & pRecordSet.Fields("town").Value
                Case "SORT_CODE", "sort_code", "PAYERS_SORT_CODE", "PAYERS_NEW_SORT_CODE", "ORIGINATORS_SORT_CODE"
                  vValue = pRecordSet.Fields(vItemName.ToLower).Value
                  If vValue.Length > 5 Then
                    If pEnv.IsDefaultCountryUK Then vValue = vValue.Substring(0, 2) & "-" & vValue.Substring(2, 2) & "-" & vValue.Substring(4, 2)
                  End If
                Case "iban_number", "payers_iban_number", "originators_iban_number", "recipient_iban_number"
                  vValue = pRecordSet.Fields(vItemName.ToLower).Value
                  If vValue.Length > 4 Then vValue = Text.RegularExpressions.Regex.Replace(vValue, "(.{0,4})", "$1 ")
                Case "TRANSACTION_REFERENCE"
                  vValue = pRecordSet.Fields("batch_number").Value & "/" & pRecordSet.Fields("transaction_number").Value
                Case "EVENT_REFERENCE"
                  vValue = pRecordSet.Fields("event_reference").Value
                  If vValue.Length = 0 Then vValue = pRecordSet.Fields("event_desc").Value
                Case "rate_desc"
                  vValue = pRecordSet.Fields("rate_desc").Value
                  If pRecordSet.Fields.Exists("currency_code") Then
                    Dim vTemp As String = pRecordSet.Fields("currency_code").Value
                    If vTemp.Length > 0 AndAlso vTemp <> pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCurrencyCode) Then
                      vValue = vValue & " (" & vTemp & ")"
                    End If
                  End If
                Case "expiry_date"
                  vValue = pRecordSet.Fields("expiry_date").Value
                  If vValue.Length = 4 Then vValue = vValue.Substring(0, 2) & "/" & vValue.Substring(vValue.Length - 2)
                Case "credit_card_number"
                  vValue = String.Empty
                Case "notes"
                  vValue = pRecordSet.Fields(vItemName).MultiLine.TrimEnd(vbCrLf.ToCharArray)
                Case "surname"
                  vValue = pRecordSet.Fields(vItemName).Value
                  If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) AndAlso pRecordSet.Fields.Exists("surname_prefix") Then
                    If pRecordSet.Fields.Item("surname_prefix").Value.Length > 0 Then
                      vValue = pRecordSet.Fields.Item("surname_prefix").Value & " " & vValue
                    End If
                  End If
                Case "raw_mark AS raw_mark_check"
                  vValue = pRecordSet.Fields("raw_mark").Value
                Case "original_mark AS original_mark_check"
                  vValue = pRecordSet.Fields("original_mark").Value
                Case "original_grade AS original_grade_check"
                  vValue = pRecordSet.Fields("original_grade").Value
                Case "original_result AS original_result_check"
                  vValue = pRecordSet.Fields("original_result").Value
                Case "FUNDRAISING_REQUEST"
                  vValue = "F"
                Case "WORKSTREAM"
                  vValue = "W"
                Case "CPD_PERIOD_DOC"
                  vValue = "CE"
                Case "CPD_POINT_DESC"
                  vValue = pRecordSet.Fields("cpd_category_type_desc").Value & "-" & pRecordSet.Fields("cpd_category_desc").Value
                Case "CPD_POINT_DOC"
                  vValue = "CP"
                Case "POSITION_LINK_DESC"
                  Dim vContact As New Contact(pEnv)
                  vContact.InitFromRecordSetName(pRecordSet)
                  vValue = vContact.Name & "-" & pRecordSet.Fields("position").Value
                Case "POSITION_LINK_TYPE"
                  vValue = "P"
                Case ""
                  vValue = ""
                Case Else
                  If pRecordSet.Fields.Exists(vItemName) Then
                    vValue = pRecordSet.Fields(vItemName).Value
                  ElseIf pRecordSet.Fields.Exists(vItemName.ToLower) Then
                    vValue = pRecordSet.Fields(vItemName.ToLower).Value
                  Else
                    vValue = ""
                  End If
              End Select
              vRow.Item(vIndex) = vValue
            Else
              If pRecordSet.Fields(vIndex).FieldType = CDBField.FieldTypes.cftMemo Then
                vValue = pRecordSet.Fields(vIndex).MultiLine.TrimEnd(vbCrLf.ToCharArray)
                vRow.Item(vIndex) = vValue
              Else
                vValue = pRecordSet.Fields(vIndex).Value
                If pRecordSet.Fields(vIndex).Name = "credit_card_number" Then
                  vValue = String.Empty
                ElseIf mvListManagerViewSelection = True AndAlso pRecordSet.Fields(vIndex).Name = "address" Then
                  'When selecting List Manager data, display address on one line only so remove vbCrLf
                  vValue = vValue.Replace(vbCr, "").Replace(vbLf, " ").Trim(" "c)
                End If
                vRow.Item(vIndex) = vValue
              End If
            End If
          Next
          CheckTimeout()
          If mvCheckAccess AndAlso vRow.Item("Access") = "N" Then mvRows.Remove(vRow)
          If mvMaxRows > 0 AndAlso mvRows.Count >= mvMaxRows Then Exit Do
        Loop While pRecordSet.Fetch()
        'Go through all the columns to see if any of the date fields are actually times
        'This is necessary as the first record retrieved may have had a null value in the database and therefore think it is just a date
        For vIndex As Integer = 1 To mvColumns.Count
          If mvColumns(vIndex).FieldType = CDBField.FieldTypes.cftDate Then
            If pRecordSet.Fields.Exists(mvColumns(vIndex).AttributeName) Then
              mvColumns(vIndex).FieldType = pRecordSet.Fields(mvColumns(vIndex).AttributeName).FieldType
            End If
          End If
        Next
      ElseIf pPopulateIfNoRecords = True AndAlso mvColumns.Count = 0 AndAlso pRecordSet.Fields IsNot Nothing Then
        For Each vField As CDBField In pRecordSet.Fields
          AddColumn(vField.Name, vField.FieldType)
        Next
      End If
      pRecordSet.CloseRecordSet()
    End Sub

    Public Sub SetStandardColumnNames()
      For Each vColumn As CDBDataColumn In mvColumns
        vColumn.SetStandardColumnName()
      Next
    End Sub

    Public Sub SetParameterNames()
      Dim vNames As New List(Of String)
      For Each vRow As CDBDataRow In mvRows
        If vRow.Item("ParameterName").Length = 0 Then
          Dim vBaseName As String = ProperName(vRow.Item("AttributeName"))
          Dim vCount As Integer = 1
          Dim vName As String = vBaseName
          While vNames.Contains(vName)
            vCount += 1
            vName = vBaseName & vCount.ToString
          End While
          vRow.Item("ParameterName") = vName
          vNames.Add(vName)
        End If
      Next
    End Sub

    Public Sub SuppressData()
      'Suppress Data if user only has Browse access
      'Look at the ownership group information to decide which records we cannot show information for
      For Each vDataRow As CDBDataRow In mvRows
        Dim vGroup As String = vDataRow.Item("OwnershipGroup")
        If Not String.IsNullOrEmpty(vGroup) Then
          If mvEnv.User.AccessLevelFromOwnershipGroup(vGroup) = CDBEnvironment.OwnershipAccessLevelTypes.oaltBrowse Then
            For Each vDataCol As CDBDataColumn In mvColumns
              Select Case vDataCol.Name
                Case "Surname", "Initials", "ContactNumber", "LabelName", "Forenames", "Title", "ContactType", _
                     "Name", "Abbreviation", "OwnershipGroup", "OwnershipGroupDesc", "OwnershipAccessLevel", _
                     "OwnershipAccessLevelDesc", "ContactName", "Description", "ItemType"
                  'OK
                Case "AccountNumber", "Number", "CreditCardNumber", "NiNumber"  'BR18231
                  vDataRow.Item(vDataCol.Name) = ""
                Case Else
                  If Not vDataCol.Name.EndsWith("Number") Then vDataRow.Item(vDataCol.Name) = ""
              End Select
            Next
          End If
          CheckTimeout()
        End If
      Next
    End Sub

    Public Sub SuppressDuplicateColumnData(ByVal pColumnName As String, Optional ByVal pColumn2Name As String = "")
      Dim vLastValue As String = ""
      Dim vLastValue2 As String = ""
      For Each vRow As CDBDataRow In mvRows
        If vRow.Item(pColumnName) = vLastValue Then
          If pColumn2Name.Length = 0 OrElse vRow.Item(pColumn2Name) = vLastValue2 Then
            vRow.Item(pColumnName) = ""
          End If
        Else
          vLastValue = vRow.Item(pColumnName)
          If pColumn2Name.Length > 0 Then vLastValue2 = vRow.Item(pColumn2Name)
        End If
      Next vRow
    End Sub

    Public Property Timeout() As Integer
      Get
        Return mvTimeout
      End Get
      Set(ByVal pValue As Integer)
        mvTimeout = pValue
      End Set
    End Property

    Public Sub SetDocumentAccess()
      Dim vAccess As Boolean
      For Each vRow As CDBDataRow In mvRows
        vAccess = True
        If vRow.Item("CreatedBy") = mvEnv.User.Logname Then
          If vRow.BoolItem("CreatorHeader") = False Then vAccess = False
        ElseIf vRow.Item("DepartmentCode") = mvEnv.User.Department Then
          If vRow.BoolItem("DepartmentHeader") = False Then vAccess = False
        Else
          If vRow.BoolItem("PublicHeader") = False Then vAccess = False
        End If
        vRow.BoolItem("Access") = vAccess
      Next
    End Sub

    Public ReadOnly Property XMLContents() As String
      Get
        Dim vStream As New IO.MemoryStream
        Dim vSettings As New Xml.XmlWriterSettings()
        vSettings.NewLineHandling = Xml.NewLineHandling.None
        Dim vWriter As Xml.XmlWriter = Xml.XmlWriter.Create(vStream, vSettings)
        With vWriter
          vWriter.WriteProcessingInstruction("xml", "version=""1.0""")
          .WriteStartElement("CDBDataTable")
          For Each vRow As CDBDataRow In Me.Rows
            .WriteStartElement("Row")
            For Each vCol As CDBDataColumn In Me.Columns
              .WriteAttributeString(vCol.Name, vRow.Item(vCol.Index))
            Next
            .WriteEndElement()
          Next
          .WriteEndElement()
          .Flush()
        End With
        Dim vDoc As New Xml.XmlDocument
        Dim vReader As New System.IO.StreamReader(vStream)
        vStream.Position = 0
        Return vReader.ReadToEnd()
      End Get
    End Property

    Public Sub RemoveRowsWithBlankColumn(ByVal pColumnName As String)
      Dim vRemoved As Boolean
      Do
        vRemoved = False
        For Each vDataRow As CDBDataRow In mvRows
          If vDataRow.Item(pColumnName) = "" Then
            mvRows.Remove(vDataRow)
            vRemoved = True
            Exit For
          End If
        Next vDataRow
      Loop While vRemoved
    End Sub

    Public Sub AddRowFromListWithQuotes(ByVal pList As String)
      Dim vValues() As String
      Dim vIndex As Integer
      Dim vPos As Integer
      Dim vPos2 As Integer
      Dim vFound As Boolean
      Dim vTemp() As String
      Dim vValue As String = ""
      Dim vCombine As Boolean
      Dim vCount As Integer

      Dim vRow As CDBDataRow = Me.AddRow
      Do
        'Look for a portion of pList that contains a comma within a quoted string
        vPos = InStr(vPos2 + 1, pList, "'")
        vPos2 = InStr(vPos + 1, pList, "'")
        If vPos > 0 And vPos2 > 0 Then vFound = InStr(Mid(pList, vPos + 1, (vPos2 - vPos) - 1), ",") > 0
      Loop While Not vFound And InStr(vPos2 + 1, pList, "'") > 0
      If vFound Then
        'Found a quoted string containing at least one comma
        ReDim vValues(0)
        vCount = -1
        vTemp = Split(pList, ",")
        For vIndex = 0 To UBound(vTemp)
          If Left(vTemp(vIndex), 1) = "'" And Right(vTemp(vIndex), 1) <> "'" Then
            'Found an opening quotation mark
            vCombine = True
            vValue = Mid(vTemp(vIndex), 2)
          ElseIf vCombine Then
            vValue = vValue & "," & vTemp(vIndex)
            If Right(vTemp(vIndex), 1) = "'" Then
              'Found the closing quotation mark
              vCombine = False
              vValue = Mid(vValue, 1, Len(vValue) - 1) 'Remove the closing quotation mark
            End If
          ElseIf Left(vTemp(vIndex), 1) = "'" And Right(vTemp(vIndex), 1) = "'" Then
            'This value has both opening and closing quotation marks, so remove both
            vValue = Mid(vTemp(vIndex), 2, Len(vTemp(vIndex)) - 2)
          Else
            'Use this value as is
            vValue = vTemp(vIndex)
          End If
          If Not vCombine Then
            vCount = vCount + 1
            If vCount > UBound(vValues) Then ReDim Preserve vValues(vCount)
            vValues(vCount) = vValue
          End If
        Next
      Else
        vValues = Split(pList, ",")
      End If

      For vIndex = 0 To UBound(vValues)
        If Left(vValues(vIndex), 1) = "'" And Right(vValues(vIndex), 1) = "'" Then
          'This value has both opening and closing quotation marks, so remove both
          'This would be encountered when pList contains one or more quoted string, but none of those strings contain any commas
          vRow.Item(vIndex + 1) = Mid(vValues(vIndex), 2, Len(vValues(vIndex)) - 2)
        Else
          vRow.Item(vIndex + 1) = vValues(vIndex)
        End If
      Next
    End Sub

    Public Function FindRow(ByVal pColumnName As String, ByVal pValue As String) As CDBDataRow
      For Each vRow As CDBDataRow In Me.Rows
        If vRow.Item(pColumnName) = pValue Then Return vRow
      Next
      Return Nothing
    End Function

  End Class
End Namespace
