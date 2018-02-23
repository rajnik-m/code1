

Namespace Access
  Public Class CriteriaDetails

    Public Enum CriteriaSetDetailRecordSetTypes 'These are bit values
      csdrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CriteriaSetDetailFields
      csdfAll = 0
      csdfCriteriaSet
      csdfSequenceNumber
      csdfSearchArea
      csdfIE
      csdfCO
      csdfMainValue
      csdfSubsidiaryValue
      csdfPeriod
      csdfCounted
      csdfAmendedBy
      csdfAmendedOn
      csdfAndOr
      csdfLeftParenthesis
      csdfRightParenthesis
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvValid As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "criteria_set_details"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("criteria_set", CDBField.FieldTypes.cftLong)
          .Add("sequence_number", CDBField.FieldTypes.cftInteger)
          .Add("search_area")
          .Add("i_e")
          .Add("c_o")
          .Add("main_value", CDBField.FieldTypes.cftUnicode)
          .Add("subsidiary_value", CDBField.FieldTypes.cftUnicode)
          .Add("period")
          .Add("counted", CDBField.FieldTypes.cftLong)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("and_or")
          .Add("left_parenthesis")
          .Add("right_parenthesis")
        End With

        mvClassFields.Item(CriteriaSetDetailFields.csdfCriteriaSet).SetPrimaryKeyOnly()
        mvClassFields.Item(CriteriaSetDetailFields.csdfSequenceNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CriteriaSetDetailFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(CriteriaSetDetailFields.csdfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CriteriaSetDetailFields.csdfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CriteriaSetDetailRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CriteriaSetDetailRecordSetTypes.csdrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "csd")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCriteriaSet As Integer = 0, Optional ByRef pSequenceNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCriteriaSet > 0 Then
        If pSequenceNumber > 0 Then
          vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CriteriaSetDetailRecordSetTypes.csdrtAll) & " FROM criteria_set_details csd WHERE criteria_set = " & pCriteriaSet & " AND sequence_number = " & pSequenceNumber)
        Else
          vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CriteriaSetDetailRecordSetTypes.csdrtAll) & " FROM criteria_set_details csd WHERE criteria_set = " & pCriteriaSet)
        End If
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CriteriaSetDetailRecordSetTypes.csdrtAll)
        Else
          InitClassFields()
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CriteriaSetDetailRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CriteriaSetDetailFields.csdfCriteriaSet, vFields)
        .SetItem(CriteriaSetDetailFields.csdfSequenceNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CriteriaSetDetailRecordSetTypes.csdrtAll) = CriteriaSetDetailRecordSetTypes.csdrtAll Then
          .SetItem(CriteriaSetDetailFields.csdfSearchArea, vFields)
          .SetItem(CriteriaSetDetailFields.csdfIE, vFields)
          .SetItem(CriteriaSetDetailFields.csdfCO, vFields)
          .SetItem(CriteriaSetDetailFields.csdfMainValue, vFields)
          .SetItem(CriteriaSetDetailFields.csdfSubsidiaryValue, vFields)
          .SetItem(CriteriaSetDetailFields.csdfPeriod, vFields)
          .SetItem(CriteriaSetDetailFields.csdfCounted, vFields)
          .SetItem(CriteriaSetDetailFields.csdfAmendedBy, vFields)
          .SetItem(CriteriaSetDetailFields.csdfAmendedOn, vFields)
          .SetItem(CriteriaSetDetailFields.csdfAndOr, vFields)
          .SetItem(CriteriaSetDetailFields.csdfLeftParenthesis, vFields)
          .SetItem(CriteriaSetDetailFields.csdfRightParenthesis, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CriteriaSetDetailFields.csdfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByVal pCriteriaSet As Integer, ByVal pSequenceNumber As Integer, ByVal pSearchArea As String, _
                      ByVal pIE As String, ByVal pCO As String, Optional ByRef pMainValue As String = "", _
                      Optional ByVal pSubsidiaryValue As String = "", Optional ByVal pPeriod As String = "", _
                      Optional ByVal pCounted As String = "", Optional ByVal pAndOr As String = "", _
                      Optional ByVal pLeftParentheses As String = "", Optional ByVal pRightParenthesis As String = "")
      With mvClassFields
        .Item(CriteriaSetDetailFields.csdfCriteriaSet).IntegerValue = pCriteriaSet
        .Item(CriteriaSetDetailFields.csdfSequenceNumber).IntegerValue = pSequenceNumber
        .Item(CriteriaSetDetailFields.csdfSearchArea).Value = pSearchArea
        .Item(CriteriaSetDetailFields.csdfIE).Value = pIE
        .Item(CriteriaSetDetailFields.csdfCO).Value = pCO
        If Len(pMainValue) > 0 Then .Item(CriteriaSetDetailFields.csdfMainValue).Value = pMainValue
        If Len(pSubsidiaryValue) > 0 Then .Item(CriteriaSetDetailFields.csdfSubsidiaryValue).Value = pSubsidiaryValue
        If Len(pPeriod) > 0 Then .Item(CriteriaSetDetailFields.csdfPeriod).Value = pPeriod
        If Len(pCounted) > 0 Then .Item(CriteriaSetDetailFields.csdfCounted).IntegerValue = CInt(Val(pCounted))
        If Len(pAndOr) > 0 Then .Item(CriteriaSetDetailFields.csdfAndOr).Value = pAndOr
        If Len(pLeftParentheses) > 0 Then .Item(CriteriaSetDetailFields.csdfLeftParenthesis).Value = pLeftParentheses
        If Len(pRightParenthesis) > 0 Then .Item(CriteriaSetDetailFields.csdfRightParenthesis).Value = pRightParenthesis
      End With
    End Sub

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
        AmendedBy = mvClassFields.Item(CriteriaSetDetailFields.csdfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CriteriaSetDetailFields.csdfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property AndOr() As String
      Get
        AndOr = mvClassFields.Item(CriteriaSetDetailFields.csdfAndOr).Value
      End Get
    End Property

    Public Property ContactOrOrganisation() As String
      Get
        ContactOrOrganisation = mvClassFields.Item(CriteriaSetDetailFields.csdfCO).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(CriteriaSetDetailFields.csdfCO).Value = Value
      End Set
    End Property

    Public ReadOnly Property Counted() As Integer
      Get
        Counted = mvClassFields.Item(CriteriaSetDetailFields.csdfCounted).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CriteriaSetNumber() As Integer
      Get
        CriteriaSetNumber = mvClassFields.Item(CriteriaSetDetailFields.csdfCriteriaSet).IntegerValue
      End Get
    End Property

    Public Property IncludeOrExclude() As String
      Get
        IncludeOrExclude = mvClassFields.Item(CriteriaSetDetailFields.csdfIE).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(CriteriaSetDetailFields.csdfIE).Value = Value
      End Set
    End Property

    Public ReadOnly Property LeftParenthesis() As String
      Get
        LeftParenthesis = mvClassFields.Item(CriteriaSetDetailFields.csdfLeftParenthesis).Value
      End Get
    End Property

    Public Property MainValue() As String
      Get
        MainValue = mvClassFields.Item(CriteriaSetDetailFields.csdfMainValue).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(CriteriaSetDetailFields.csdfMainValue).Value = Value
      End Set
    End Property

    Public Property Period() As String
      Get
        Period = mvClassFields.Item(CriteriaSetDetailFields.csdfPeriod).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(CriteriaSetDetailFields.csdfPeriod).Value = Value
      End Set
    End Property

    Public ReadOnly Property RightParenthesis() As String
      Get
        RightParenthesis = mvClassFields.Item(CriteriaSetDetailFields.csdfRightParenthesis).Value
      End Get
    End Property

    Public Property SearchArea() As String
      Get
        SearchArea = mvClassFields.Item(CriteriaSetDetailFields.csdfSearchArea).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(CriteriaSetDetailFields.csdfSearchArea).Value = Value
      End Set
    End Property

    Public ReadOnly Property SequenceNumber() As Integer
      Get
        SequenceNumber = mvClassFields.Item(CriteriaSetDetailFields.csdfSequenceNumber).IntegerValue
      End Get
    End Property

    Public Property SubsidiaryValue() As String
      Get
        SubsidiaryValue = mvClassFields.Item(CriteriaSetDetailFields.csdfSubsidiaryValue).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(CriteriaSetDetailFields.csdfSubsidiaryValue).Value = Value
      End Set
    End Property

    Public Property Valid() As Boolean
      Get
        'This is used as a flag to indicated that the class' property values have been set by the user.
        Valid = mvValid
      End Get
      Set(ByVal Value As Boolean)
        'This is used as a flag to indicated that the class' property values have been set by the user.
        mvValid = Value
      End Set
    End Property

    Public Sub Delete(Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
    End Sub

    Public Sub DeleteAllDetails(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("criteria_set", CriteriaSetNumber)
      mvEnv.Connection.StartTransaction()
      mvEnv.Connection.DeleteRecords("criteria_set_details", vWhereFields, False)
      mvEnv.Connection.CommitTransaction()
    End Sub
  End Class
End Namespace
