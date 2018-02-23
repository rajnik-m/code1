

Namespace Access
  Public Class Source

    Public Enum SourceRecordSetTypes 'These are bit values
      surtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum SourceFields
      sofAll = 0
      sofSource
      sofSourceDesc
      sofIncentiveTriggerLevel
      sofThankYouLetter
      sofIncentiveScheme
      sofHistoryOnly
      sofAmendedBy
      sofAmendedOn
      sofDistributionCode
      sofDiscountPercentage
      sofSourceNumber
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvSegment As Segment
    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "sources"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("source")
          .Add("source_desc")
          .Add("incentive_trigger_level", CDBField.FieldTypes.cftNumeric)
          .Add("thank_you_letter")
          .Add("incentive_scheme")
          .Add("history_only")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("distribution_code")
          .Add("discount_percentage", CDBField.FieldTypes.cftNumeric)
          .Add("source_number", CDBField.FieldTypes.cftLong)
        End With

        mvClassFields.Item(SourceFields.sofSource).SetPrimaryKeyOnly()
        mvClassFields.Item(SourceFields.sofSourceNumber).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport)
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
      mvClassFields.Item(SourceFields.sofHistoryOnly).Value = "N"
    End Sub

    Private Sub SetValid(ByVal pField As SourceFields)
      'Add code here to ensure all values are valid before saving
      If mvExisting = False And mvClassFields.Item(SourceFields.sofSourceNumber).IntegerValue = 0 Then mvClassFields.Item(SourceFields.sofSourceNumber).Value = CStr(mvEnv.GetControlNumber("SR"))
      mvClassFields.Item(SourceFields.sofAmendedOn).Value = TodaysDate()
      mvClassFields.Item(SourceFields.sofAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As SourceRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = SourceRecordSetTypes.surtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "s")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pSource As String = "", Optional ByRef pSourceNumber As Integer = 0, Optional ByRef pInitSegment As Boolean = False)
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String

      mvEnv = pEnv
      If Len(pSource) > 0 Or pSourceNumber > 0 Then
        vSQL = "SELECT " & GetRecordSetFields(SourceRecordSetTypes.surtAll) & " FROM sources s WHERE "
        If Len(pSource) > 0 Then
          vSQL = vSQL & "source = '" & pSource & "'"
        Else
          vSQL = vSQL & "source_number = " & pSourceNumber
        End If
        vRecordSet = pEnv.Connection.GetRecordSet(vSQL)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, SourceRecordSetTypes.surtAll, pInitSegment)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As SourceRecordSetTypes, Optional ByRef pInitSegment As Boolean = False)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(SourceFields.sofSource, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And SourceRecordSetTypes.surtAll) = SourceRecordSetTypes.surtAll Then
          .SetItem(SourceFields.sofSourceDesc, vFields)
          .SetItem(SourceFields.sofIncentiveTriggerLevel, vFields)
          .SetItem(SourceFields.sofThankYouLetter, vFields)
          .SetItem(SourceFields.sofIncentiveScheme, vFields)
          .SetItem(SourceFields.sofHistoryOnly, vFields)
          .SetItem(SourceFields.sofAmendedBy, vFields)
          .SetItem(SourceFields.sofAmendedOn, vFields)
          .SetItem(SourceFields.sofDistributionCode, vFields)
          .SetItem(SourceFields.sofDiscountPercentage, vFields)
          .SetOptionalItem(SourceFields.sofSourceNumber, vFields)
        End If
      End With
      If pInitSegment Then InitSegment()
    End Sub
    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(SourceFields.sofAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
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
        AmendedBy = mvClassFields.Item(SourceFields.sofAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(SourceFields.sofAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property DiscountPercentage() As Double
      Get
        DiscountPercentage = mvClassFields.Item(SourceFields.sofDiscountPercentage).DoubleValue
      End Get
    End Property

    Public ReadOnly Property DistributionCode() As String
      Get
        DistributionCode = mvClassFields.Item(SourceFields.sofDistributionCode).Value
      End Get
    End Property

    Public ReadOnly Property HistoryOnly() As Boolean
      Get
        HistoryOnly = mvClassFields.Item(SourceFields.sofHistoryOnly).Bool
      End Get
    End Property

    Public ReadOnly Property IncentiveScheme() As String
      Get
        IncentiveScheme = mvClassFields.Item(SourceFields.sofIncentiveScheme).Value
      End Get
    End Property

    Public ReadOnly Property IncentiveTriggerLevel() As Double
      Get
        IncentiveTriggerLevel = CDbl(mvClassFields.Item(SourceFields.sofIncentiveTriggerLevel).Value)
      End Get
    End Property

    Public ReadOnly Property SourceCode() As String
      Get
        SourceCode = mvClassFields.Item(SourceFields.sofSource).Value
      End Get
    End Property

    Public ReadOnly Property SourceDesc() As String
      Get
        SourceDesc = mvClassFields.Item(SourceFields.sofSourceDesc).Value
      End Get
    End Property

    Public ReadOnly Property SourceNumber() As Integer
      Get
        SourceNumber = mvClassFields.Item(SourceFields.sofSourceNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ThankYouLetter() As String
      Get
        ThankYouLetter = mvClassFields.Item(SourceFields.sofThankYouLetter).Value
      End Get
    End Property
    Public ReadOnly Property Segment() As Segment
      Get
        If mvSegment Is Nothing Then InitSegment()
        Segment = mvSegment
      End Get
    End Property

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      Init(pEnv)
      With mvClassFields
        mvClassFields(SourceFields.sofSource).Value = pParams("Source").Value
      End With
      Update(pParams)
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      With mvClassFields
        If pParams.Exists("SourceDesc") Then mvClassFields(SourceFields.sofSourceDesc).Value = pParams("SourceDesc").Value
        If pParams.Exists("IncentiveScheme") Then mvClassFields(SourceFields.sofIncentiveScheme).Value = pParams("IncentiveScheme").Value
        If pParams.Exists("IncentiveTriggerLevel") Then mvClassFields(SourceFields.sofIncentiveTriggerLevel).Value = pParams("IncentiveTriggerLevel").Value
        If pParams.Exists("ThankYouLetter") Then mvClassFields(SourceFields.sofThankYouLetter).Value = pParams("ThankYouLetter").Value
        If pParams.Exists("DistributionCode") Then mvClassFields(SourceFields.sofDistributionCode).Value = pParams("DistributionCode").Value
        If pParams.Exists("DiscountPercentage") Then mvClassFields(SourceFields.sofDiscountPercentage).Value = pParams("DiscountPercentage").Value
        If pParams.Exists("HistoryOnly") Then mvClassFields(SourceFields.sofHistoryOnly).Value = pParams("HistoryOnly").Value
      End With
    End Sub

    Public Function DataTable() As CDBDataTable
      DataTable = mvClassFields.DataTable
    End Function

    Private Sub InitSegment()
      Dim vRS As CDBRecordSet
      Dim vSQL As String
      Dim vWhereFields As New CDBFields

      mvSegment = New Segment
      mvSegment.Init(mvEnv, "", "", "", True)
      vWhereFields.Add("source", CDBField.FieldTypes.cftCharacter, SourceCode)
      vSQL = "SELECT " & mvSegment.GetRecordSetFields(Segment.SegmentRecordSetTypes.srtAll) & " FROM segments sg%1 WHERE %2 %3"
      If mvEnv.Connection.GetCount("segments", vWhereFields, "") > 1 Then
        vWhereFields.Add("sg.mailing", CDBField.FieldTypes.cftInteger, "mh.mailing")
        vSQL = Replace(vSQL, "%1", ", mailing_history mh")
        vSQL = Replace(vSQL, "%3", " ORDER BY mailing_date DESC")
      Else
        vSQL = Replace(vSQL, "%1", "")
        vSQL = Replace(vSQL, "%3", "")
      End If
      vSQL = Replace(vSQL, "%2", mvEnv.Connection.WhereClause(vWhereFields))
      'Set vRS = mvEnv.Connection.GetRecordSet("SELECT " & mvSegment.GetRecordSetFields(srtAll) & " FROM segments s WHERE s.source = '" & SourceCode & "'")
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      With vRS
        If .Fetch() = True Then
          mvSegment = New Segment
          mvSegment.Init(mvEnv, (.Fields.Item("campaign").Value), (.Fields.Item("appeal").Value), (.Fields.Item("segment").Value))
        End If
        .CloseRecordSet()
      End With
    End Sub
  End Class
End Namespace
