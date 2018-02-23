Namespace Access

  Public Class ExamCentreAction
    Inherits ActionLink

    Private Enum ExamCentreActionFields
      AllFields = 0
      ActionLinkId
      ExamCentreId
      ActionNumber
      Type
      CreatedBy
      CreatedOn
      AmendedBy
      AmendedOn
    End Enum

    Protected Overrides Sub AddFields()
      mvClassFields.Add("action_link_id", CDBField.FieldTypes.cftInteger)
      mvClassFields.Add("exam_centre_id", CDBField.FieldTypes.cftInteger)
      mvClassFields.Add("action_number", CDBField.FieldTypes.cftLong)
      mvClassFields.Add("type")
      mvClassFields.Add("created_by")
      mvClassFields.Add("created_on", CDBField.FieldTypes.cftDate)

      mvClassFields.Item(ExamCentreActionFields.ActionLinkId).PrimaryKey = True
      mvClassFields.Item(ExamCentreActionFields.ActionLinkId).PrefixRequired = True
      mvClassFields.Item(ExamCentreActionFields.ExamCentreId).PrefixRequired = True
      mvClassFields.Item(ExamCentreActionFields.ActionNumber).PrefixRequired = True
      mvClassFields.Item(ExamCentreActionFields.Type).PrefixRequired = True
      mvClassFields.Item(ExamCentreActionFields.CreatedBy).PrefixRequired = True
      mvClassFields.Item(ExamCentreActionFields.CreatedOn).PrefixRequired = True
      mvClassFields.SetControlNumberField(ExamCentreActionFields.ActionLinkId, "ALI")
    End Sub

    Public Overloads Overrides Sub Init(pObjectType As IActionLink.ActionLinkObjectTypes, pActionNumber As Integer, pContactNumber As Integer, pLinkType As IActionLink.ActionLinkTypes) 'Implements IActionLink.Init
      InitFromObjectType(pObjectType)
      If (pActionNumber > 0) And (pContactNumber > 0) Then
        Dim vWhereFields As New CDBFields
        vWhereFields.Add((mvClassFields.Item(ExamCentreActionFields.ActionNumber).Name), pActionNumber)
        vWhereFields.Add((mvClassFields.Item(ExamCentreActionFields.ExamCentreId).Name), pContactNumber)
        vWhereFields.Add((mvClassFields.Item(ExamCentreActionFields.Type).Name), LinkTypeCode(pLinkType))
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
        Return "eca"
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
      mvClassFields.Item(ExamCentreActionFields.ActionNumber).IntegerValue = pActionNumber
      mvClassFields.Item(ExamCentreActionFields.ExamCentreId).IntegerValue = pNumber
      mvClassFields.Item(ExamCentreActionFields.Type).Value = LinkTypeCode(pLinkType)
    End Sub

    Private Shadows Sub InitFromParams(ByVal pEnv As CDBEnvironment,
                                       ByRef pObjectType As IActionLink.ActionLinkObjectTypes,
                                       ByRef pActionNumber As Integer,
                                       ByRef pNumber As Integer,
                                       ByRef pLinkType As IActionLink.ActionLinkTypes,
                                       ByRef pNotified As String,
                                       ByVal pAdditionalNumber As Integer)
    End Sub


    Public Overrides Sub Create(ByVal pEnv As CDBEnvironment, ByVal pActionNumber As Integer, ByVal pExamCentreId As Integer, ByVal pLinkType As IActionLink.ActionLinkTypes)
      mvEnv = pEnv
      InitClassFields()
      mvClassFields.Item(ExamCentreActionFields.ActionNumber).LongValue = pActionNumber
      mvClassFields.Item(ExamCentreActionFields.ExamCentreId).LongValue = pExamCentreId
      mvClassFields.Item(ExamCentreActionFields.Type).Value = LinkTypeCode(pLinkType)
    End Sub

    Public Overrides Function CreateInstance(ByVal pEnv As CDBEnvironment) As CARERecord
      Return New ExamCentreAction(mvEnv)
    End Function

    Public ReadOnly Property ActionLinkId() As Integer
      Get
        Return mvClassFields.Item(ExamCentreActionFields.ActionLinkId).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ExamCentreId() As Integer
      Get
        Return mvClassFields.Item(ExamCentreActionFields.ExamCentreId).IntegerValue
      End Get
    End Property

    Public Overloads ReadOnly Property ContactNumber() As Integer
      'Overload the base ContactNumber property to return the ExamCentreID
      Get
        Return ExamCentreId
      End Get
    End Property

  End Class
End Namespace
