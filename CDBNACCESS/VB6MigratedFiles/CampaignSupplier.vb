

Namespace Access
  Public Class CampaignSupplier

    Public Enum CampaignSupplierRecordSetTypes 'These are bit values
      csrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CampaignSupplierFields
      csfAll = 0
      csfCampaign
      csfAppeal
      csfSegment
      csfContactNumber
      csfSupplierRole
      csfNotes
      csfAmendedBy
      csfAmendedOn
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
          .DatabaseTableName = "campaign_suppliers"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("campaign")
          .Add("appeal")
          .Add("segment")
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("supplier_role")
          .Add("notes", CDBField.FieldTypes.cftMemo)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Item(CampaignSupplierFields.csfCampaign).SetPrimaryKeyOnly()
          .Item(CampaignSupplierFields.csfAppeal).SetPrimaryKeyOnly()
          .Item(CampaignSupplierFields.csfSegment).SetPrimaryKeyOnly()
          .Item(CampaignSupplierFields.csfContactNumber).SetPrimaryKeyOnly()
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CampaignSupplierFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(CampaignSupplierFields.csfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CampaignSupplierFields.csfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CampaignSupplierRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CampaignSupplierRecordSetTypes.csrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cs")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pSupplier As Integer = 0, Optional ByVal pCampaign As String = "", Optional ByVal pAppeal As String = "", Optional ByVal pSegment As String = "")
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      mvEnv = pEnv
      If pCampaign.Length > 0 And pSupplier > 0 Then
        With vWhereFields
          .Add("campaign", CDBField.FieldTypes.cftCharacter, pCampaign)
          .Add("contact_number", CDBField.FieldTypes.cftLong, pSupplier)
          If pAppeal.Length > 0 Then .Add("appeal", CDBField.FieldTypes.cftCharacter, pAppeal)
          If pSegment.Length > 0 Then .Add("segment", CDBField.FieldTypes.cftCharacter, pSegment)
        End With
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CampaignSupplierRecordSetTypes.csrtAll) & " FROM campaign_suppliers c WHERE " & pEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CampaignSupplierRecordSetTypes.csrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CampaignSupplierRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And CampaignSupplierRecordSetTypes.csrtAll) = CampaignSupplierRecordSetTypes.csrtAll Then
          .SetItem(CampaignSupplierFields.csfCampaign, vFields)
          .SetItem(CampaignSupplierFields.csfAppeal, vFields)
          .SetItem(CampaignSupplierFields.csfSegment, vFields)
          .SetItem(CampaignSupplierFields.csfContactNumber, vFields)
          .SetItem(CampaignSupplierFields.csfSupplierRole, vFields)
          .SetItem(CampaignSupplierFields.csfNotes, vFields)
          .SetItem(CampaignSupplierFields.csfAmendedBy, vFields)
          .SetItem(CampaignSupplierFields.csfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pCampaign As String, ByRef pAppeal As String, ByRef pSegment As String, ByRef pContactNumber As Integer, ByRef pRole As String, ByRef pNotes As String)
      Dim vRecordSet As CDBRecordSet
      Dim vWhere As New CDBFields
      vWhere.Add("campaign", CDBField.FieldTypes.cftCharacter, pCampaign, CDBField.FieldWhereOperators.fwoEqual)
      vWhere.Add("appeal", CDBField.FieldTypes.cftCharacter, pAppeal, CDBField.FieldWhereOperators.fwoEqual)
      vWhere.Add("segment", CDBField.FieldTypes.cftCharacter, pSegment, CDBField.FieldWhereOperators.fwoEqual)
      vWhere.Add("contact_number", pContactNumber, CDBField.FieldWhereOperators.fwoEqual)
      If mvClassFields Is Nothing Then Init(pEnv)
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CampaignSupplierRecordSetTypes.csrtAll) & " FROM campaign_suppliers WHERE " & mvEnv.Connection.WhereClause(vWhere))
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(pEnv, vRecordSet, CampaignSupplierRecordSetTypes.csrtAll)
      Else
        With mvClassFields
          .Item(CampaignSupplierFields.csfAppeal).Value = pAppeal
          .Item(CampaignSupplierFields.csfCampaign).Value = pCampaign
          .Item(CampaignSupplierFields.csfSegment).Value = pSegment
          .Item(CampaignSupplierFields.csfContactNumber).IntegerValue = pContactNumber
          .Item(CampaignSupplierFields.csfSupplierRole).Value = pRole
          .Item(CampaignSupplierFields.csfNotes).Value = pNotes
        End With
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub Update(Optional ByVal pSupplierRole As String = "", Optional ByVal pNotes As String = "")
      With mvClassFields
        .Item(CampaignSupplierFields.csfSupplierRole).Value = pSupplierRole
        .Item(CampaignSupplierFields.csfNotes).Value = pNotes
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CampaignSupplierFields.csfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
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
        AmendedBy = mvClassFields.Item(CampaignSupplierFields.csfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CampaignSupplierFields.csfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Appeal() As String
      Get
        Appeal = mvClassFields.Item(CampaignSupplierFields.csfAppeal).Value
      End Get
    End Property

    Public ReadOnly Property Campaign() As String
      Get
        Campaign = mvClassFields.Item(CampaignSupplierFields.csfCampaign).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(CampaignSupplierFields.csfContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(CampaignSupplierFields.csfNotes).MultiLineValue
      End Get
    End Property

    Public ReadOnly Property Segment() As String
      Get
        Segment = mvClassFields.Item(CampaignSupplierFields.csfSegment).Value
      End Get
    End Property

    Public ReadOnly Property SupplierRole() As String
      Get
        SupplierRole = mvClassFields.Item(CampaignSupplierFields.csfSupplierRole).Value
      End Get
    End Property

  End Class
End Namespace
