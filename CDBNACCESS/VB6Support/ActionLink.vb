Namespace Access

  Partial Public Class ActionLink
    Implements IActionLink

    Private mvObjectType As IActionLink.ActionLinkObjectTypes
    Private mvLinkName As String
    Private mvAddressNumber As Integer

    Protected Property ObjectType As IActionLink.ActionLinkObjectTypes
      Get
        Return mvObjectType
      End Get
      Set(value As IActionLink.ActionLinkObjectTypes)
        mvObjectType = value
      End Set
    End Property
    Public Overloads Sub ClearFields()
      mvLinkName = ""
      mvAddressNumber = 0
    End Sub

    Public Overridable Overloads Sub Init(pObjectType As IActionLink.ActionLinkObjectTypes, pActionNumber As Integer, pContactNumber As Integer, pLinkType As IActionLink.ActionLinkTypes) Implements IActionLink.Init
      InitFromObjectType(pObjectType)
      If (pActionNumber > 0) And (pContactNumber > 0) Then
        Dim vWhereFields As New CDBFields
        vWhereFields.Add((mvClassFields.Item(ContactActionFields.ActionNumber).Name), pActionNumber)
        'setLinkedItemId(pContactNumber)
        vWhereFields.Add((mvClassFields.Item(ContactActionFields.ContactNumber).Name), pContactNumber)
        'only add ContactActionFields.Type if exists in database table
        If mvClassFields.Item(ContactActionFields.Type).InDatabase Then
          vWhereFields.Add((mvClassFields.Item(ContactActionFields.Type).Name), LinkTypeCode(pLinkType))
        End If
        InitWithPrimaryKey(vWhereFields)
      Else
        MyBase.Init()
      End If
    End Sub

    Friend Overridable Sub InitFromObjectType(ByVal pObjectType As IActionLink.ActionLinkObjectTypes) Implements IActionLink.InitFromObjectType
      Init()
      Select Case pObjectType
        Case IActionLink.ActionLinkObjectTypes.alotOrganisation
          mvClassFields.Item(ContactActionFields.ContactNumber).SetName("organisation_number")
          mvClassFields.DatabaseTableName = "organisation_actions"
          mvClassFields.Item(ContactActionFields.Notified).InDatabase = False
        Case IActionLink.ActionLinkObjectTypes.alotDocument
          mvClassFields.Item(ContactActionFields.ContactNumber).SetName("document_number")
          mvClassFields.DatabaseTableName = "document_actions"
          mvClassFields.Item(ContactActionFields.Notified).InDatabase = False
        Case IActionLink.ActionLinkObjectTypes.alotFundraising
          mvClassFields.Item(ContactActionFields.ContactNumber).SetName("fundraising_request_number")
          mvClassFields.DatabaseTableName = "fundraising_actions"
          mvClassFields.Item(ContactActionFields.Notified).InDatabase = False
          mvClassFields.Item(ContactActionFields.Type).InDatabase = False
          mvClassFields.Item(ContactActionFields.AdditionalNumber).InDatabase = True
          mvClassFields.Item(ContactActionFields.AdditionalNumber).SetName("scheduled_payment_number")
      End Select
      mvObjectType = pObjectType
    End Sub

    Public ReadOnly Property TableName() As String
      Get
        Return mvClassFields.DatabaseTableName
      End Get
    End Property

    Public ReadOnly Property AliasValue() As String
      Get
        Return mvClassFields.TableAlias
      End Get
    End Property

    Friend ReadOnly Property ObjectLinkType() As IActionLink.ActionLinkObjectTypes Implements IActionLink.ObjectLinkType
      Get
        Return mvObjectType
      End Get
    End Property

    Public Sub InitFromParams(ByVal pEnv As CDBEnvironment,
                              ByRef pObjectType As IActionLink.ActionLinkObjectTypes,
                              ByRef pActionNumber As Integer,
                              ByRef pNumber As Integer,
                              ByRef pLinkType As IActionLink.ActionLinkTypes,
                              ByRef pNotified As String,
                              ByVal pAdditionalNumber As Integer)
      mvEnv = pEnv
      InitFromObjectType(pObjectType)
      SetValid()
      mvClassFields.Item(ContactActionFields.ActionNumber).IntegerValue = pActionNumber
      mvClassFields.Item(ContactActionFields.ContactNumber).IntegerValue = pNumber
      mvClassFields.Item(ContactActionFields.Type).Value = LinkTypeCode(pLinkType)
      If pObjectType = IActionLink.ActionLinkObjectTypes.alotContact Then
        mvClassFields.Item(ContactActionFields.Notified).Value = pNotified
      End If
      If pAdditionalNumber > 0 Then
        mvClassFields.Item(ContactActionFields.AdditionalNumber).IntegerValue = pAdditionalNumber
      End If
    End Sub

    Public Overrides Sub InitFromRecordSet(ByVal pRecordSet As Data.CDBRecordSet)
      MyBase.InitFromRecordSet(pRecordSet)
      Select Case mvObjectType
        Case IActionLink.ActionLinkObjectTypes.alotContact
          mvLinkName = pRecordSet.Fields.FieldExists("label_name").Value
          mvAddressNumber = pRecordSet.Fields.FieldExists("address_number").IntegerValue
        Case IActionLink.ActionLinkObjectTypes.alotOrganisation
          mvLinkName = pRecordSet.Fields.FieldExists("name").Value
          mvAddressNumber = pRecordSet.Fields.FieldExists("address_number").IntegerValue
      End Select
    End Sub

    Public Overridable Overloads Sub Create(ByVal pEnv As CDBEnvironment,
                                            ByVal pActionNumber As Integer,
                                            ByVal pContactNumber As Integer,
                                            ByVal pLinkType As IActionLink.ActionLinkTypes) Implements IActionLink.Create
      mvEnv = pEnv
      InitClassFields()
      With mvClassFields
        .Item(ContactActionFields.ActionNumber).LongValue = pActionNumber
        .Item(ContactActionFields.ContactNumber).LongValue = pContactNumber
        .Item(ContactActionFields.Type).Value = LinkTypeCode(pLinkType)
      End With
    End Sub

    Public ReadOnly Property AddressNumber() As Integer
      Get
        AddressNumber = mvAddressNumber
      End Get
    End Property

  End Class

End Namespace
