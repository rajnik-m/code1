

Namespace Access
  Public Class CommunicationsLogHistory

    Public Enum CommunicationsLogHistoryRecordSetTypes 'These are bit values
      clhrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CommunicationsLogHistoryFields
      clhfAll = 0
      clhfCommunicationsLogNumber
      clhfActionDate
      clhfActionTime
      clhfAction
      clhfUserName
      clhfNotes
    End Enum

    Public Enum CommunicationsLogHistoryActions
      clhaCreated 'created
      clhaViewed
      clhaPrinted
      clhaEMailed
      clhaImported
      clhaExported
      clhaEdited
      clhaUpdated
      clhaDistributed
      clhaTransferred
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
          .DatabaseTableName = "communications_log_history"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("communications_log_number", CDBField.FieldTypes.cftLong)
          .Add("action_date", CDBField.FieldTypes.cftDate)
          .Add("action_time")
          .Add("action")
          .Add("user_name")
          .Add("notes", CDBField.FieldTypes.cftMemo)

          .Item(CommunicationsLogHistoryFields.clhfCommunicationsLogNumber).SetPrimaryKeyOnly()
          .Item(CommunicationsLogHistoryFields.clhfActionDate).SetPrimaryKeyOnly()
          .Item(CommunicationsLogHistoryFields.clhfActionTime).SetPrimaryKeyOnly()
          .Item(CommunicationsLogHistoryFields.clhfAction).SetPrimaryKeyOnly()
        End With

      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CommunicationsLogHistoryFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CommunicationsLogHistoryRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CommunicationsLogHistoryRecordSetTypes.clhrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "clh")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CommunicationsLogHistoryRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And CommunicationsLogHistoryRecordSetTypes.clhrtAll) = CommunicationsLogHistoryRecordSetTypes.clhrtAll Then
          .SetItem(CommunicationsLogHistoryFields.clhfCommunicationsLogNumber, vFields)
          .SetItem(CommunicationsLogHistoryFields.clhfActionDate, vFields)
          .SetItem(CommunicationsLogHistoryFields.clhfActionTime, vFields)
          .SetItem(CommunicationsLogHistoryFields.clhfAction, vFields)
          .SetItem(CommunicationsLogHistoryFields.clhfUserName, vFields)
          .SetItem(CommunicationsLogHistoryFields.clhfNotes, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CommunicationsLogHistoryFields.clhfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByRef pCommsLogNumber As Integer, ByRef pAction As CommunicationsLogHistoryActions, Optional ByRef pNotes As String = "")
      Dim vAction As String = ""

      With mvClassFields
        .Item(CommunicationsLogHistoryFields.clhfCommunicationsLogNumber).IntegerValue = pCommsLogNumber
        .Item(CommunicationsLogHistoryFields.clhfActionDate).Value = TodaysDate()
        .Item(CommunicationsLogHistoryFields.clhfActionTime).Value = TimeOfDay.ToString("HHmmss")
        .Item(CommunicationsLogHistoryFields.clhfUserName).Value = mvEnv.User.Logname
        Select Case pAction
          Case CommunicationsLogHistoryActions.clhaCreated
            vAction = "created"
          Case CommunicationsLogHistoryActions.clhaViewed
            vAction = "viewed"
          Case CommunicationsLogHistoryActions.clhaPrinted
            vAction = "printed"
          Case CommunicationsLogHistoryActions.clhaEMailed
            vAction = "e-mailed"
          Case CommunicationsLogHistoryActions.clhaImported
            vAction = "imported"
          Case CommunicationsLogHistoryActions.clhaExported
            vAction = "exported"
          Case CommunicationsLogHistoryActions.clhaEdited
            vAction = "edited"
          Case CommunicationsLogHistoryActions.clhaUpdated
            vAction = "updated"
          Case CommunicationsLogHistoryActions.clhaDistributed
            vAction = "distributed"
          Case CommunicationsLogHistoryActions.clhaTransferred
            vAction = "transferred"
        End Select
        .Item(CommunicationsLogHistoryFields.clhfAction).Value = vAction
        .Item(CommunicationsLogHistoryFields.clhfNotes).Value = pNotes
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

    Public ReadOnly Property Action() As String
      Get
        Action = mvClassFields.Item(CommunicationsLogHistoryFields.clhfAction).Value
      End Get
    End Property

    Public ReadOnly Property ActionDate() As String
      Get
        ActionDate = mvClassFields.Item(CommunicationsLogHistoryFields.clhfActionDate).Value
      End Get
    End Property

    Public ReadOnly Property ActionTime() As String
      Get
        ActionTime = mvClassFields.Item(CommunicationsLogHistoryFields.clhfActionTime).Value
      End Get
    End Property

    Public ReadOnly Property CommunicationsLogNumber() As Integer
      Get
        CommunicationsLogNumber = mvClassFields.Item(CommunicationsLogHistoryFields.clhfCommunicationsLogNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(CommunicationsLogHistoryFields.clhfNotes).MultiLineValue
      End Get
    End Property

    Public ReadOnly Property UserName() As String
      Get
        UserName = mvClassFields.Item(CommunicationsLogHistoryFields.clhfUserName).Value
      End Get
    End Property
  End Class
End Namespace
