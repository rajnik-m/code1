

Namespace Access
  Public Class SubTopic

    Public Enum SubTopicRecordSetTypes 'These are bit values
      strtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum SubTopicFields
      stfAll = 0
      stfTopic
      stfSubTopic
      stfSubTopicDesc
      stfActivity
      stfActivityValue
      stfActionNumber
      stfHistoryOnly
      stfAmendedBy
      stfAmendedOn
      stfActivityDuration
      stfSetCallCompleted
      stfCallBackMinutes
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvQuantity As Double
    Private mvParagraphText As String
    Private mvParagraphRetrieved As Boolean
    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "sub_topics"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("topic")
          .Add("sub_topic")
          .Add("sub_topic_desc")
          .Add("activity")
          .Add("activity_value")
          .Add("action_number", CDBField.FieldTypes.cftLong)
          .Add("history_only")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("activity_duration", CDBField.FieldTypes.cftLong)
          .Add("set_call_completed").InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbTelemarketing)
          .Add("call_back_minutes", CDBField.FieldTypes.cftInteger).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbTelemarketing)
        End With

        mvClassFields.Item(SubTopicFields.stfTopic).SetPrimaryKeyOnly()
        mvClassFields.Item(SubTopicFields.stfSubTopic).SetPrimaryKeyOnly()

        mvClassFields.Item(SubTopicFields.stfActivityDuration).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataSubTopicActivityDuration)
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
      mvParagraphRetrieved = False
    End Sub
    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As SubTopicFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(SubTopicFields.stfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(SubTopicFields.stfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As SubTopicRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = SubTopicRecordSetTypes.strtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "st")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pTopic As String = "", Optional ByRef pSubTopic As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pTopic) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(SubTopicRecordSetTypes.strtAll) & " FROM sub_topics st WHERE topic = '" & pTopic & "' AND sub_topic = '" & pSubTopic & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, SubTopicRecordSetTypes.strtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As SubTopicRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(SubTopicFields.stfTopic, vFields)
        .SetItem(SubTopicFields.stfSubTopic, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And SubTopicRecordSetTypes.strtAll) = SubTopicRecordSetTypes.strtAll Then
          .SetItem(SubTopicFields.stfSubTopicDesc, vFields)
          .SetItem(SubTopicFields.stfActivity, vFields)
          .SetItem(SubTopicFields.stfActivityValue, vFields)
          .SetItem(SubTopicFields.stfActionNumber, vFields)
          .SetItem(SubTopicFields.stfHistoryOnly, vFields)
          .SetItem(SubTopicFields.stfAmendedBy, vFields)
          .SetItem(SubTopicFields.stfAmendedOn, vFields)
          .SetOptionalItem(SubTopicFields.stfActivityDuration, vFields)
          .SetOptionalItem(SubTopicFields.stfSetCallCompleted, vFields)
          .SetOptionalItem(SubTopicFields.stfCallBackMinutes, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(SubTopicFields.stfAll)
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

    Public ReadOnly Property ActionNumber() As Integer
      Get
        ActionNumber = mvClassFields.Item(SubTopicFields.stfActionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Activity() As String
      Get
        Activity = mvClassFields.Item(SubTopicFields.stfActivity).Value
      End Get
    End Property

    Public ReadOnly Property ActivityValue() As String
      Get
        ActivityValue = mvClassFields.Item(SubTopicFields.stfActivityValue).Value
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(SubTopicFields.stfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(SubTopicFields.stfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property HistoryOnly() As Boolean
      Get
        HistoryOnly = mvClassFields.Item(SubTopicFields.stfHistoryOnly).Bool
      End Get
    End Property

    Public ReadOnly Property SubTopicCode() As String
      Get
        SubTopicCode = mvClassFields.Item(SubTopicFields.stfSubTopic).Value
      End Get
    End Property

    Public ReadOnly Property SubTopicDesc() As String
      Get
        SubTopicDesc = mvClassFields.Item(SubTopicFields.stfSubTopicDesc).Value
      End Get
    End Property

    Public ReadOnly Property Topic() As String
      Get
        Topic = mvClassFields.Item(SubTopicFields.stfTopic).Value
      End Get
    End Property

    Public ReadOnly Property ActivityDuration() As String
      Get
        'Could be null
        ActivityDuration = mvClassFields.Item(SubTopicFields.stfActivityDuration).Value
      End Get
    End Property

    Public ReadOnly Property SetCallCompleted() As Boolean
      Get
        Return mvClassFields.Item(SubTopicFields.stfSetCallCompleted).Bool
      End Get
    End Property

    Public ReadOnly Property CallBackMinutes() As String
      Get
        Return mvClassFields.Item(SubTopicFields.stfCallBackMinutes).Value
      End Get
    End Property

    Public ReadOnly Property Quantity() As Double
      Get
        Quantity = mvQuantity
      End Get
    End Property

    Public ReadOnly Property ParagraphText() As String
      Get
        Dim vRecordSet As CDBRecordSet

        If Not mvParagraphRetrieved Then
          mvParagraphRetrieved = True
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT paragraph_text FROM sub_topic_paragraphs WHERE topic = '" & Topic & "' AND sub_topic = '" & SubTopicCode & "'")
          With vRecordSet
            If .Fetch() = True Then mvParagraphText = .Fields(1).MultiLine
            .CloseRecordSet()
          End With
        End If
        ParagraphText = mvParagraphText
      End Get
    End Property

    Public Sub SetQuantity(ByVal pData As Double)
      mvQuantity = pData
    End Sub

    Public Function AddActivity() As Boolean
      AddActivity = Len(mvClassFields(SubTopicFields.stfActivity).Value) > 0 And Len(mvClassFields(SubTopicFields.stfActivityValue).Value) > 0
    End Function
  End Class
End Namespace
