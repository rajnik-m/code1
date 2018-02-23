

Namespace Access
  Public Class SegmentCostCentre

    Public Enum SegmentCostCentreRecordSetTypes 'These are bit values
      sccrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum SegmentCostCentreFields
      sccfAll = 0
      sccfCampaign
      sccfAppeal
      sccfSegment
      sccfCostCentre
      sccfCostCentrePercentage
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
          .DatabaseTableName = "segment_cost_centres"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("campaign")
          .Add("appeal")
          .Add("segment")
          .Add("cost_centre")
          .Add("cost_centre_percentage", CDBField.FieldTypes.cftNumeric)
        End With

      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As SegmentCostCentreFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As SegmentCostCentreRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = SegmentCostCentreRecordSetTypes.sccrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "scc")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As SegmentCostCentreRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And SegmentCostCentreRecordSetTypes.sccrtAll) = SegmentCostCentreRecordSetTypes.sccrtAll Then
          .SetItem(SegmentCostCentreFields.sccfCampaign, vFields)
          .SetItem(SegmentCostCentreFields.sccfAppeal, vFields)
          .SetItem(SegmentCostCentreFields.sccfSegment, vFields)
          .SetItem(SegmentCostCentreFields.sccfCostCentre, vFields)
          .SetItem(SegmentCostCentreFields.sccfCostCentrePercentage, vFields)
        End If
      End With
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(SegmentCostCentreFields.sccfAll)
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

    Public ReadOnly Property Appeal() As String
      Get
        Appeal = mvClassFields.Item(SegmentCostCentreFields.sccfAppeal).Value
      End Get
    End Property

    Public ReadOnly Property Campaign() As String
      Get
        Campaign = mvClassFields.Item(SegmentCostCentreFields.sccfCampaign).Value
      End Get
    End Property

    Public ReadOnly Property CostCentre() As String
      Get
        CostCentre = mvClassFields.Item(SegmentCostCentreFields.sccfCostCentre).Value
      End Get
    End Property

    Public ReadOnly Property CostCentrePercentage() As Double
      Get
        CostCentrePercentage = CDbl(mvClassFields.Item(SegmentCostCentreFields.sccfCostCentrePercentage).Value)
      End Get
    End Property

    Public ReadOnly Property Segment() As String
      Get
        Segment = mvClassFields.Item(SegmentCostCentreFields.sccfSegment).Value
      End Get
    End Property

    Public Sub Create(ByVal pCampaign As String, ByVal pAppeal As String, ByVal pSegment As String, ByVal pCostCentre As String, ByVal pCostCentrePercentage As String)
      With mvClassFields
        .Item(SegmentCostCentreFields.sccfCampaign).Value = pCampaign
        .Item(SegmentCostCentreFields.sccfAppeal).Value = pAppeal
        .Item(SegmentCostCentreFields.sccfSegment).Value = pSegment
        .Item(SegmentCostCentreFields.sccfCostCentre).Value = pCostCentre
        .Item(SegmentCostCentreFields.sccfCostCentrePercentage).Value = pCostCentrePercentage
      End With
    End Sub
  End Class
End Namespace
