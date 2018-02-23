Namespace Access

  Public Class ContactPositionAction
    Inherits ActionLink

    Private Enum ContactPositionActionFields
      AllFields = 0
      ActionLinkId
      ContactPositionNumber
      ActionNumber
      Type
      CreatedBy
      CreatedOn
      AmendedBy
      AmendedOn
    End Enum

    Protected Overrides Sub AddFields()
      mvClassFields.Add("action_link_id", CDBField.FieldTypes.cftInteger)
      mvClassFields.Add("contact_position_number", CDBField.FieldTypes.cftInteger)
      mvClassFields.Add("action_number", CDBField.FieldTypes.cftLong)
      mvClassFields.Add("type")
      mvClassFields.Add("created_by")
      mvClassFields.Add("created_on", CDBField.FieldTypes.cftDate)

      mvClassFields.Item(ContactPositionActionFields.ActionLinkId).PrimaryKey = True
      mvClassFields.Item(ContactPositionActionFields.ActionLinkId).PrefixRequired = True
      mvClassFields.Item(ContactPositionActionFields.ContactPositionNumber).PrefixRequired = True
      mvClassFields.Item(ContactPositionActionFields.ActionNumber).PrefixRequired = True
      mvClassFields.Item(ContactPositionActionFields.Type).PrefixRequired = True
      mvClassFields.Item(ContactPositionActionFields.CreatedBy).PrefixRequired = True
      mvClassFields.Item(ContactPositionActionFields.CreatedOn).PrefixRequired = True
      mvClassFields.SetControlNumberField(ContactPositionActionFields.ActionLinkId, "ALI")
    End Sub

    Public Overloads Overrides Sub Init(pObjectType As IActionLink.ActionLinkObjectTypes, pActionNumber As Integer, pContactPositionNumber As Integer, pLinkType As IActionLink.ActionLinkTypes) 'Implements IActionLink.Init
      InitFromObjectType(pObjectType)
      If (pActionNumber > 0) And (pContactPositionNumber > 0) Then
        Dim vWhereFields As New CDBFields
        vWhereFields.Add((mvClassFields.Item(ContactPositionActionFields.ActionNumber).Name), pActionNumber)
        vWhereFields.Add((mvClassFields.Item(ContactPositionActionFields.ContactPositionNumber).Name), pContactPositionNumber)
        vWhereFields.Add((mvClassFields.Item(ContactPositionActionFields.Type).Name), LinkTypeCode(pLinkType))
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
        Return "cpa"
      End Get
    End Property

    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "action_links"
      End Get
    End Property

    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    Public Shadows Sub InitFromParams(ByVal pEnv As CDBEnvironment,
                                        ByRef pObjectType As IActionLink.ActionLinkObjectTypes,
                                        ByRef pActionNumber As Integer,
                                        ByRef pNumber As Integer,
                                        ByRef pLinkType As IActionLink.ActionLinkTypes)
      mvEnv = pEnv
      Init()
      ObjectType = pObjectType
      SetValid()
      mvClassFields.Item(ContactPositionActionFields.ActionNumber).IntegerValue = pActionNumber
      mvClassFields.Item(ContactPositionActionFields.ContactPositionNumber).IntegerValue = pNumber
      mvClassFields.Item(ContactPositionActionFields.Type).Value = LinkTypeCode(pLinkType)
    End Sub

    Private Shadows Sub InitFromParams(ByVal pEnv As CDBEnvironment,
                                       ByRef pObjectType As IActionLink.ActionLinkObjectTypes,
                                       ByRef pActionNumber As Integer,
                                       ByRef pNumber As Integer,
                                       ByRef pLinkType As IActionLink.ActionLinkTypes,
                                       ByRef pNotified As String,
                                       ByVal pAdditionalNumber As Integer)
    End Sub


    Public Overrides Sub Create(ByVal pEnv As CDBEnvironment, ByVal pActionNumber As Integer, ByVal pContactPositionNumber As Integer, ByVal pLinkType As IActionLink.ActionLinkTypes)
      mvEnv = pEnv
      InitClassFields()
      mvClassFields.Item(ContactPositionActionFields.ActionNumber).LongValue = pActionNumber
      mvClassFields.Item(ContactPositionActionFields.ContactPositionNumber).LongValue = pContactPositionNumber
      mvClassFields.Item(ContactPositionActionFields.Type).Value = LinkTypeCode(pLinkType)
    End Sub

    Public Overrides Function CreateInstance(ByVal pEnv As CDBEnvironment) As CARERecord
      Return New ContactPositionAction(mvEnv)
    End Function

    Public ReadOnly Property ActionLinkId() As Integer
      Get
        Return mvClassFields.Item(ContactPositionActionFields.ActionLinkId).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ContactPositionNumber() As Integer
      Get
        Return mvClassFields.Item(ContactPositionActionFields.ContactPositionNumber).IntegerValue
      End Get
    End Property

    Public Overloads ReadOnly Property ContactNumber() As Integer
      'Overload the base ContactNumber property to return the ExamCentreID
      Get
        Return ContactPositionNumber
      End Get
    End Property

  End Class
End Namespace
