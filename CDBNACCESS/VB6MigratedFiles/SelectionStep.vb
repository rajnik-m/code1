

Namespace Access
  Public Class SelectionStep

    Public Enum SelectionStepRecordSetTypes 'These are bit values
      sstprtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum SelectionStepFields
      ssfAll = 0
      ssfCriteriaSet
      ssfViewName
      ssfSequenceNumber
      ssfFilterSql
      ssfSelectAction
      ssfRecordCount
      ssfAmendedBy
      ssfAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "selection_steps"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("criteria_set", CDBField.FieldTypes.cftLong)
          .Add("view_name")
          .Add("sequence_number", CDBField.FieldTypes.cftInteger)
          .Add("filter_sql")
          .Add("select_action")
          .Add("record_count", CDBField.FieldTypes.cftLong)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)

          .Item(SelectionStepFields.ssfCriteriaSet).SetPrimaryKeyOnly()
          .Item(SelectionStepFields.ssfSequenceNumber).SetPrimaryKeyOnly()

          .Item(SelectionStepFields.ssfCriteriaSet).PrefixRequired = True
          .Item(SelectionStepFields.ssfViewName).PrefixRequired = True
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As SelectionStepFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(SelectionStepFields.ssfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(SelectionStepFields.ssfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As SelectionStepRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = SelectionStepRecordSetTypes.sstprtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "sst")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As SelectionStepRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And SelectionStepRecordSetTypes.sstprtAll) = SelectionStepRecordSetTypes.sstprtAll Then
          .SetItem(SelectionStepFields.ssfCriteriaSet, vFields)
          .SetItem(SelectionStepFields.ssfViewName, vFields)
          .SetItem(SelectionStepFields.ssfSequenceNumber, vFields)
          .SetItem(SelectionStepFields.ssfFilterSql, vFields)
          .SetItem(SelectionStepFields.ssfSelectAction, vFields)
          .SetItem(SelectionStepFields.ssfRecordCount, vFields)
          .SetItem(SelectionStepFields.ssfAmendedBy, vFields)
          .SetItem(SelectionStepFields.ssfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(SelectionStepFields.ssfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByRef pCriteriaSet As Integer, ByRef pViewName As String, ByRef pSequenceNumber As Integer, ByRef pFilterSQL As String, ByRef pAction As String, ByRef pRecordCount As Integer)
      With mvClassFields
        .Item(SelectionStepFields.ssfCriteriaSet).IntegerValue = pCriteriaSet
        .Item(SelectionStepFields.ssfViewName).Value = pViewName
        .Item(SelectionStepFields.ssfSequenceNumber).IntegerValue = pSequenceNumber
        .Item(SelectionStepFields.ssfFilterSql).Value = pFilterSQL
        .Item(SelectionStepFields.ssfSelectAction).Value = pAction
        .Item(SelectionStepFields.ssfRecordCount).IntegerValue = pRecordCount
      End With
    End Sub

    Public Function GetActionDesc(ByRef pSelectAction As String) As String
      Select Case pSelectAction
        Case "S"
          Return "Selected"
        Case "R"
          Return "Removed"
        Case "P"
          Return "Replaced"
        Case "D"
          Return "Set Default Address"
        Case "U"
          Return "Set Address By Usage"
        Case Else
          Return ""         'Added to fix compiler warning
      End Select
    End Function

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(SelectionStepFields.ssfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(SelectionStepFields.ssfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CriteriaSet() As Integer
      Get
        CriteriaSet = mvClassFields.Item(SelectionStepFields.ssfCriteriaSet).IntegerValue
      End Get
    End Property

    Public ReadOnly Property FilterSql() As String
      Get
        FilterSql = mvClassFields.Item(SelectionStepFields.ssfFilterSql).Value
      End Get
    End Property

    Public ReadOnly Property RecordCount() As Integer
      Get
        RecordCount = mvClassFields.Item(SelectionStepFields.ssfRecordCount).IntegerValue
      End Get
    End Property

    Public ReadOnly Property SelectAction() As String
      Get
        SelectAction = mvClassFields.Item(SelectionStepFields.ssfSelectAction).Value
      End Get
    End Property

    Public ReadOnly Property SequenceNumber() As Integer
      Get
        SequenceNumber = mvClassFields.Item(SelectionStepFields.ssfSequenceNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ViewName() As String
      Get
        ViewName = mvClassFields.Item(SelectionStepFields.ssfViewName).Value
      End Get
    End Property
 
    Public Function GetAdjustedFilterSQL() As String
      'This function returns an adjusted FilterSQL value. The adjustment is to convert ADDRESS field into a VARCHAR so that
      'it can be used easily in a SQL WHERE clause. 

      'e.g. Value of “x.address = 'unknown' may be adjusted to "replace(cast(x.address as varchar(max)),char(10),' ') = 'unknown'"

      '(This Filter_SQL is often defined as a TEXT on the database, but this type of field cannot be used easily in comparison.)

      'This routine should handle:
      '1) qualifiers e.g. addresses.address
      '2) ignore 'address' within quotes e.g. description = 'unknown address'
      '3) square brackets e.g. x.[address] = 'unknown'
      '4) convertion in oracle and ms sql
      '5) upper and lower cases e.g. address = 'n/a' and ADDRES = 'n/a'

      Dim vDBFieldName As String = "address"

      Dim vAdjustedSQL As String
      Dim vTextField As String
      Dim vFieldLength As Integer

      Dim vAdjustedFieldName As String
      Dim vAdjustedFieldLength As Integer
      Dim vAdjustedSQLLength As Integer

      Dim vFoundInPostion As Integer
      Dim vChar As Char
      Dim vCharPosition As Integer
      Dim vInQuote As Boolean
      Dim vCharX As Integer
      Dim vFromPosition As Integer = 1
      Dim vExitAllAdjustments As Boolean = False
      Dim vRejectFind As Boolean = True
      Dim vLoopCount As Integer 'This is used to prevent looping. There should never be 1000 occurence of ADDRESS in the FilterSQL so exit loop if it is this value

      vAdjustedSQL = FilterSql
      vFieldLength = Len(vDBFieldName)

      If vFieldLength > 0 Then

        'NOTE There may be more than one occurence of Text Field name. 

        Do Until vExitAllAdjustments Or vLoopCount = 1000
          vLoopCount += 1

          vAdjustedSQLLength = Len(vAdjustedSQL)
          vTextField = vDBFieldName
          vFieldLength = Len(vTextField)


          If vFromPosition > vAdjustedSQLLength - vFieldLength Then
            'There is no more space for any more occurences there exit loop 
            vExitAllAdjustments = True
          Else
            vFoundInPostion = InStr(vFromPosition, LCase(vAdjustedSQL), vTextField)
            If vFoundInPostion < 1 Then
              'TextField name not found, so exit loop
              vExitAllAdjustments = True
            End If
          End If

          If Not vExitAllAdjustments Then

            vRejectFind = False

            'Is TextField name in quotes?
            vInQuote = False
            vCharX = 0
            For Each vChar In vAdjustedSQL
              vCharX += 1
              If vCharX >= vFoundInPostion Then
                Exit For
              End If
              If vChar = "'" Then
                If vInQuote Then
                  vInQuote = False
                Else
                  vInQuote = True
                End If
              End If
            Next

            If vInQuote Then
              'ignore this occurence of the TextField name and set up for looking for next
              vRejectFind = True
            Else
              If vFoundInPostion > 1 Then
                'Check if preceeding syntax are OK
                vCharPosition = vFoundInPostion - 1
                vChar = GetChar(vAdjustedSQL, vCharPosition)
                If vChar = "." Or vChar = "[" Or vChar = " " Or vChar = " " Or vChar = "=" Then
                  vTextField = Mid(vAdjustedSQL, vFoundInPostion, vFieldLength) 'This keeps the letters in the same case
                  ExtendedFieldName(vAdjustedSQL, vTextField, vFoundInPostion) 'This method add vTestField with any qualifiers and square brackets 
                  vFieldLength = Len(vTextField)
                Else
                  vRejectFind = True
                End If
              End If
            End If

            If Not vRejectFind Then
              If (vFoundInPostion + vFieldLength) <= vAdjustedSQLLength Then
                'check following syntax are OK
                vChar = GetChar(vAdjustedSQL, vFoundInPostion + vFieldLength)
                If Not (vChar = " " Or vChar = " " Or vChar = "=") Then
                  vRejectFind = True
                End If
              End If
            End If

            If vRejectFind Then
              vFromPosition = vFoundInPostion + vFieldLength
            Else
              'Now adjust this TextField name
              Dim vFirstPart As String = ""
              If vFoundInPostion > 1 Then
                vFirstPart = Strings.Left(vAdjustedSQL, vFoundInPostion - 1)
              End If
              vAdjustedFieldName = mvEnv.Connection.DBReplaceLineFeedWithSpace(vTextField)
              vAdjustedFieldLength = Len(vAdjustedFieldName)
              vAdjustedSQL = vFirstPart & Replace(vAdjustedSQL, vTextField, vAdjustedFieldName, vFoundInPostion, 1)
              vFromPosition = vFoundInPostion + vAdjustedFieldLength
            End If
          End If
        Loop
      End If

      If vLoopCount = 1000 Then 'Adjustment failed, return FilterSQL value
        vAdjustedSQL = FilterSql
      End If

      Return vAdjustedSQL
    End Function
    Private Sub ExtendedFieldName(ByVal pSQL As String, ByRef pFieldName As String, ByRef pInPosition As Integer)
      'This Extend Field Name with qualifiers and square brackets if present
      '
      'If it cannot extend it will return pFieldName

      'e.g. for Input:
      ' pSQL = "dbo.addresses.[address] LIKE '%Unknown%'"
      ' pFieldName = "address"
      ' pInPosition = 16
      ' the Ouptut:
      ' pFieldName = "dbo.addresses.[address]"
      ' pInPosition = 1

      Dim vExtendedName As String
      Dim vExtendedInPostion As Integer
      Dim vCannotExtend As Boolean = False
      Dim vCharIndex As Integer 'For getting characters in pSQL before the pInPosition
      Dim vCharIndex2 As Integer 'For getting characters in pSQL before the pInPosition

      vExtendedName = pFieldName
      vExtendedInPostion = pInPosition

      If pInPosition > 1 Then
        vCharIndex = pInPosition - 2 'Index of character before

        'Checking for square brackets
        If pSQL(vCharIndex) = "[" Then
          If Len(pSQL) < pInPosition + Len(pFieldName) Then
            vCannotExtend = True
          Else
            If pSQL(pInPosition + Len(pFieldName) - 1) = "]" Then
              'Putting square brackets around FieldName
              vExtendedName = pSQL(vCharIndex) & vExtendedName & pSQL(pInPosition + Len(pFieldName) - 1)
              vExtendedInPostion = vExtendedInPostion - 1
              vCharIndex = vCharIndex - 1
            Else
              vCannotExtend = True
            End If
          End If
        End If

        If Not vCannotExtend Then
          'Checking for any qualifiers
          If vCharIndex > 0 Then
            If pSQL(vCharIndex) = "." Then  'there is a qualifier 
              'Build qualifier
              For vCharIndex2 = vCharIndex To 0 Step -1
                If pSQL(vCharIndex2) = " " _
                  Or pSQL(vCharIndex2) = " " _
                  Or pSQL(vCharIndex2) = "=" Then
                  Exit For
                Else
                  vExtendedName = pSQL(vCharIndex2) & vExtendedName
                  vExtendedInPostion = vExtendedInPostion - 1
                End If
              Next
            End If
          End If
        End If

      End If
      pFieldName = vExtendedName
      pInPosition = vExtendedInPostion
    End Sub

  End Class
End Namespace
