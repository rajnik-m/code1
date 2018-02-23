Public Class WorkstreamActionLink : Inherits ActionLink

  Private Enum WorkstreamActionLinkFields
    AllFields = 0
    ActionLinkId
    WorkstreamId
    ActionNumber
    Type
    CreatedBy
    CreatedOn
    AmendedBy
    AmendedOn
  End Enum

  Public Sub New(ByVal pEnv As CDBEnvironment)
    MyBase.New(pEnv)
  End Sub

  Protected Overrides Sub AddFields()
    mvClassFields.Add("action_link_id", CDBField.FieldTypes.cftInteger)
    mvClassFields.Add("workstream_id", CDBField.FieldTypes.cftInteger)
    mvClassFields.Add("action_number", CDBField.FieldTypes.cftLong)
    mvClassFields.Add("type")
    mvClassFields.Add("created_by")
    mvClassFields.Add("created_on", CDBField.FieldTypes.cftDate)

    mvClassFields.Item(WorkstreamActionLinkFields.ActionLinkId).PrimaryKey = True
    mvClassFields.Item(WorkstreamActionLinkFields.ActionLinkId).PrefixRequired = True
    mvClassFields.Item(WorkstreamActionLinkFields.WorkstreamId).PrefixRequired = True
    mvClassFields.Item(WorkstreamActionLinkFields.ActionNumber).PrefixRequired = True
    mvClassFields.Item(WorkstreamActionLinkFields.Type).PrefixRequired = True
    mvClassFields.Item(WorkstreamActionLinkFields.CreatedBy).PrefixRequired = True
    mvClassFields.Item(WorkstreamActionLinkFields.CreatedOn).PrefixRequired = True
    mvClassFields.SetControlNumberField(WorkstreamActionLinkFields.ActionLinkId, "ALI")
  End Sub

  Public Overloads Overrides Sub Init(ByVal pObjectType As IActionLink.ActionLinkObjectTypes, ByVal pActionNumber As Integer, ByVal pContactNumber As Integer, ByVal pLinkType As IActionLink.ActionLinkTypes) 'Implements IActionLink.Init
    InitFromObjectType(pObjectType)
    If pActionNumber > 0 AndAlso pContactNumber > 0 Then
      Dim vWhereFields As New CDBFields
      vWhereFields.Add((mvClassFields.Item(WorkstreamActionLinkFields.ActionNumber).Name), pActionNumber)
      vWhereFields.Add((mvClassFields.Item(WorkstreamActionLinkFields.WorkstreamId).Name), pContactNumber)
      vWhereFields.Add((mvClassFields.Item(WorkstreamActionLinkFields.Type).Name), LinkTypeCode(pLinkType))
      InitWithPrimaryKey(vWhereFields)
    Else
      MyBase.Init()
    End If
  End Sub

  Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
    Get
      Return True
    End Get
  End Property

  Protected Overrides ReadOnly Property TableAlias() As String
    Get
      Return "wal"
    End Get
  End Property

  Protected Overrides ReadOnly Property DatabaseTableName() As String
    Get
      Return "action_links"
    End Get
  End Property

  Public Shadows Sub InitFromParams(ByVal pEnv As CDBEnvironment,
                                        ByRef pObjectType As IActionLink.ActionLinkObjectTypes,
                                        ByRef pActionNumber As Integer,
                                        ByRef pNumber As Integer,
                                        ByRef pLinkType As IActionLink.ActionLinkTypes)
    mvEnv = pEnv
    Init()
    ObjectType = pObjectType
    SetValid()
    mvClassFields.Item(WorkstreamActionLinkFields.ActionNumber).IntegerValue = pActionNumber
    mvClassFields.Item(WorkstreamActionLinkFields.WorkstreamId).IntegerValue = pNumber
    mvClassFields.Item(WorkstreamActionLinkFields.Type).Value = LinkTypeCode(pLinkType)
  End Sub

  Private Shadows Sub InitFromParams(ByVal pEnv As CDBEnvironment,
                                     ByRef pObjectType As IActionLink.ActionLinkObjectTypes,
                                     ByRef pActionNumber As Integer,
                                     ByRef pNumber As Integer,
                                     ByRef pLinkType As IActionLink.ActionLinkTypes,
                                     ByRef pNotified As String,
                                     ByVal pAdditionalNumber As Integer)
  End Sub

  Public Overrides Sub Create(ByVal pEnv As CDBEnvironment, ByVal pActionNumber As Integer, ByVal pWorkstreamId As Integer, ByVal pLinkType As IActionLink.ActionLinkTypes)
    mvEnv = pEnv
    InitClassFields()
    mvClassFields.Item(WorkstreamActionLinkFields.ActionNumber).LongValue = pActionNumber
    mvClassFields.Item(WorkstreamActionLinkFields.WorkstreamId).LongValue = pWorkstreamId
    mvClassFields.Item(WorkstreamActionLinkFields.Type).Value = LinkTypeCode(pLinkType)
  End Sub

  Public Overrides Function CreateInstance(ByVal pEnv As CDBEnvironment) As CARERecord
    Return New WorkstreamActionLink(pEnv)
  End Function

  Public ReadOnly Property ActionLinkId() As Integer
    Get
      Return mvClassFields.Item(WorkstreamActionLinkFields.ActionLinkId).IntegerValue
    End Get
  End Property

  Public ReadOnly Property WorkstreamId() As Integer
    Get
      Return mvClassFields.Item(WorkstreamActionLinkFields.WorkstreamId).IntegerValue
    End Get
  End Property

  Public Overloads ReadOnly Property ContactNumber() As Integer
    'Overload the base ContactNumber property to return the WorkstreamID
    Get
      Return WorkstreamId
    End Get
  End Property

End Class
