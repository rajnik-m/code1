

Namespace Access
  Public Class TickBox

    Public Enum TickBoxRecordSetTypes 'These are bit values
      tbrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum TickBoxFields
      tbfAll = 0
      tbfCampaign
      tbfAppeal
      tbfSegment
      tbfTickBoxNumber
      tbfActivity
      tbfActivityValue
      tbfMailingSuppression
      tbfAmendedBy
      tbfAmendedOn
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
          .DatabaseTableName = "tick_boxes"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("campaign")
          .Add("appeal")
          .Add("segment")
          .Add("tick_box_number", CDBField.FieldTypes.cftInteger)
          .Add("activity")
          .Add("activity_value")
          .Add("mailing_suppression")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(TickBoxFields.tbfCampaign).SetPrimaryKeyOnly()
        mvClassFields.Item(TickBoxFields.tbfAppeal).SetPrimaryKeyOnly()
        mvClassFields.Item(TickBoxFields.tbfSegment).SetPrimaryKeyOnly()
        mvClassFields.Item(TickBoxFields.tbfTickBoxNumber).SetPrimaryKeyOnly()
        mvClassFields.SetUniqueFieldsFromPrimaryKeys()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As TickBoxFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(TickBoxFields.tbfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(TickBoxFields.tbfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As TickBoxRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = TickBoxRecordSetTypes.tbrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "tb")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCampaign As String = "", Optional ByRef pAppeal As String = "", Optional ByRef pSegment As String = "", Optional ByRef pTickBoxNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pCampaign) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(TickBoxRecordSetTypes.tbrtAll) & " FROM tick_boxes tb WHERE campaign = '" & pCampaign & "' AND appeal = '" & pAppeal & "' AND segment = '" & pSegment & "' AND tick_box_number = " & pTickBoxNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, TickBoxRecordSetTypes.tbrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As TickBoxRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(TickBoxFields.tbfCampaign, vFields)
        .SetItem(TickBoxFields.tbfAppeal, vFields)
        .SetItem(TickBoxFields.tbfSegment, vFields)
        .SetItem(TickBoxFields.tbfTickBoxNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And TickBoxRecordSetTypes.tbrtAll) = TickBoxRecordSetTypes.tbrtAll Then
          .SetItem(TickBoxFields.tbfActivity, vFields)
          .SetItem(TickBoxFields.tbfActivityValue, vFields)
          .SetItem(TickBoxFields.tbfMailingSuppression, vFields)
          .SetItem(TickBoxFields.tbfAmendedBy, vFields)
          .SetItem(TickBoxFields.tbfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(TickBoxFields.tbfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub Create(ByRef pEnv As CDBEnvironment, ByRef pCampaign As String, ByRef pAppeal As String, ByRef pSegment As String, ByRef pNumber As Integer, ByRef pActivity As String, ByRef pActivityValue As String, ByRef pSuppression As String)
      Init(pEnv)
      With mvClassFields
        .Item(TickBoxFields.tbfCampaign).Value = pCampaign
        .Item(TickBoxFields.tbfAppeal).Value = pAppeal
        .Item(TickBoxFields.tbfSegment).Value = pSegment
        .Item(TickBoxFields.tbfTickBoxNumber).Value = CStr(pNumber)
        Update(pActivity, pActivityValue, pSuppression)
      End With
    End Sub

    Public Sub Update(ByRef pActivity As String, ByRef pActivityValue As String, ByRef pSuppression As String)
      With mvClassFields
        .Item(TickBoxFields.tbfActivity).Value = pActivity
        .Item(TickBoxFields.tbfActivityValue).Value = pActivityValue
        .Item(TickBoxFields.tbfMailingSuppression).Value = pSuppression
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

    Public ReadOnly Property Activity() As String
      Get
        Activity = mvClassFields.Item(TickBoxFields.tbfActivity).Value
      End Get
    End Property

    Public ReadOnly Property ActivityValue() As String
      Get
        ActivityValue = mvClassFields.Item(TickBoxFields.tbfActivityValue).Value
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(TickBoxFields.tbfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(TickBoxFields.tbfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Appeal() As String
      Get
        Appeal = mvClassFields.Item(TickBoxFields.tbfAppeal).Value
      End Get
    End Property

    Public ReadOnly Property Campaign() As String
      Get
        Campaign = mvClassFields.Item(TickBoxFields.tbfCampaign).Value
      End Get
    End Property

    Public ReadOnly Property MailingSuppression() As String
      Get
        MailingSuppression = mvClassFields.Item(TickBoxFields.tbfMailingSuppression).Value
      End Get
    End Property

    Public ReadOnly Property Segment() As String
      Get
        Segment = mvClassFields.Item(TickBoxFields.tbfSegment).Value
      End Get
    End Property

    Public ReadOnly Property TickBoxNumber() As Integer
      Get
        TickBoxNumber = mvClassFields.Item(TickBoxFields.tbfTickBoxNumber).IntegerValue
      End Get
    End Property
  End Class
End Namespace
