

Namespace Access
  Public Class SegmentProductAllocation

    Public Enum SegmentProductAllocationRecordSetTypes 'These are bit values
      spartAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum SegmentProductAllocationFields
      spafAll = 0
      spafCampaign
      spafAppeal
      spafSegment
      spafAmountNumber
      spafProduct
      spafRate
      spafAmendedBy
      spafAmendedOn
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
          .DatabaseTableName = "segment_product_allocation"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("campaign")
          .Add("appeal")
          .Add("segment")
          .Add("amount_number", CDBField.FieldTypes.cftInteger)
          .Add("product")
          .Add("rate")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)

          .Item(SegmentProductAllocationFields.spafCampaign).SetPrimaryKeyOnly()
          .Item(SegmentProductAllocationFields.spafAppeal).SetPrimaryKeyOnly()
          .Item(SegmentProductAllocationFields.spafSegment).SetPrimaryKeyOnly()
          .Item(SegmentProductAllocationFields.spafAmountNumber).SetPrimaryKeyOnly()

          .SetUniqueFieldsFromPrimaryKeys()
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As SegmentProductAllocationFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(SegmentProductAllocationFields.spafAmendedOn).Value = TodaysDate()
      mvClassFields.Item(SegmentProductAllocationFields.spafAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByRef pRSType As SegmentProductAllocationRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = SegmentProductAllocationRecordSetTypes.spartAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "spa")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByRef pEnv As CDBEnvironment, Optional ByRef pCampaign As String = "", Optional ByRef pAppeal As String = "", Optional ByRef pSegment As String = "", Optional ByRef pAmountNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pCampaign) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(SegmentProductAllocationRecordSetTypes.spartAll) & " FROM segment_product_allocation WHERE campaign = '" & pCampaign & "' AND appeal = '" & pAppeal & "' AND segment = '" & pSegment & "' AND amount_number = " & pAmountNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, SegmentProductAllocationRecordSetTypes.spartAll)
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

    Public Sub InitFromRecordSet(ByRef pEnv As CDBEnvironment, ByRef pRecordSet As CDBRecordSet, ByRef pRSType As SegmentProductAllocationRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(SegmentProductAllocationFields.spafCampaign, vFields)
        .SetItem(SegmentProductAllocationFields.spafAppeal, vFields)
        .SetItem(SegmentProductAllocationFields.spafSegment, vFields)
        .SetItem(SegmentProductAllocationFields.spafAmountNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And SegmentProductAllocationRecordSetTypes.spartAll) = SegmentProductAllocationRecordSetTypes.spartAll Then
          .SetItem(SegmentProductAllocationFields.spafProduct, vFields)
          .SetItem(SegmentProductAllocationFields.spafRate, vFields)
          .SetItem(SegmentProductAllocationFields.spafAmendedBy, vFields)
          .SetItem(SegmentProductAllocationFields.spafAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(SegmentProductAllocationFields.spafAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByVal pCampaign As String, ByVal pAppeal As String, ByVal pSegment As String, ByVal pAmountNumber As String, ByVal pProduct As String, ByVal pRate As String)
      Init(pEnv)
      With mvClassFields
        .Item(SegmentProductAllocationFields.spafCampaign).Value = pCampaign
        .Item(SegmentProductAllocationFields.spafAppeal).Value = pAppeal
        .Item(SegmentProductAllocationFields.spafSegment).Value = pSegment
        .Item(SegmentProductAllocationFields.spafAmountNumber).Value = pAmountNumber
        .Item(SegmentProductAllocationFields.spafProduct).Value = pProduct
        .Item(SegmentProductAllocationFields.spafRate).Value = pRate
      End With
    End Sub

    Public Sub Update(ByVal pProduct As String, ByRef pRate As String)
      With mvClassFields
        .Item(SegmentProductAllocationFields.spafProduct).Value = pProduct
        .Item(SegmentProductAllocationFields.spafRate).Value = pRate
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
        AmendedBy = mvClassFields.Item(SegmentProductAllocationFields.spafAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(SegmentProductAllocationFields.spafAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property AmountNumber() As Integer
      Get
        AmountNumber = mvClassFields.Item(SegmentProductAllocationFields.spafAmountNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Appeal() As String
      Get
        Appeal = mvClassFields.Item(SegmentProductAllocationFields.spafAppeal).Value
      End Get
    End Property

    Public ReadOnly Property Campaign() As String
      Get
        Campaign = mvClassFields.Item(SegmentProductAllocationFields.spafCampaign).Value
      End Get
    End Property

    Public ReadOnly Property Product() As String
      Get
        Product = mvClassFields.Item(SegmentProductAllocationFields.spafProduct).Value
      End Get
    End Property

    Public ReadOnly Property RateCode() As String
      Get
        RateCode = mvClassFields.Item(SegmentProductAllocationFields.spafRate).Value
      End Get
    End Property

    Public ReadOnly Property Segment() As String
      Get
        Segment = mvClassFields.Item(SegmentProductAllocationFields.spafSegment).Value
      End Get
    End Property

  End Class
End Namespace
