Namespace Access

  Partial Public Class ActionLink
    Inherits CARERecord
    Implements IActionLink
    Implements IRecordCreate

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum ContactActionFields
      AllFields = 0
      ActionNumber
      ContactNumber
      Type
      Notified
      AdditionalNumber
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("action_number", CDBField.FieldTypes.cftLong)
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("type")
        .Add("notified")
        .Add("additional_number", CDBField.FieldTypes.cftLong)

        .Item(ContactActionFields.ContactNumber).PrefixRequired = True
        .Item(ContactActionFields.AdditionalNumber).InDatabase = False
      End With
    End Sub

    Protected Friend Overridable Sub setLinkedItemId(pId As Integer)
      mvClassFields(ContactActionFields.ContactNumber).SetValue = pId.ToString
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "ca"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "contact_actions"
      End Get
    End Property

    '--------------------------------------------------
    'Default constructor
    '--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public ReadOnly Property ActionNumber() As Integer Implements IActionLink.ActionNumber
      Get
        Return mvClassFields("action_number").IntegerValue
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer Implements IActionLink.LinkedItemId
      Get
        Return mvClassFields(ContactActionFields.ContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Type() As String Implements IActionLink.Type
      Get
        Return mvClassFields("type").Value
      End Get
    End Property

    Public Property Notified() As Boolean
      Get
        Return mvClassFields(ContactActionFields.Notified).Bool
      End Get
      Set(ByVal value As Boolean)
        mvClassFields(ContactActionFields.Notified).Bool = value
      End Set
    End Property

    Public Overridable ReadOnly Property IsNotifySupported As Boolean
      Get
        Return mvClassFields(ContactActionFields.Notified).InDatabase
      End Get
    End Property

    Public Overridable ReadOnly Property IsAdditionalNumberSupported As Boolean
      Get
        Return mvClassFields(ContactActionFields.AdditionalNumber).InDatabase
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ContactActionFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ContactActionFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non AutoGenerated Code"

    Public ReadOnly Property LinkType() As IActionLink.ActionLinkTypes Implements IActionLink.LinkType
      Get
        Select Case Type
          Case "A"
            Return IActionLink.ActionLinkTypes.altActioner
          Case "M"
            Return IActionLink.ActionLinkTypes.altManager
          Case "R"
            Return IActionLink.ActionLinkTypes.altRelated
        End Select
      End Get
    End Property

    Public Shared Function LinkTypeCode(ByVal pLinkType As IActionLink.ActionLinkTypes) As String
      Select Case pLinkType
        Case IActionLink.ActionLinkTypes.altActioner
          Return "A"
        Case IActionLink.ActionLinkTypes.altManager
          Return "M"
        Case Else
          Return "R"
      End Select
    End Function

    Public Overridable Function CreateInstance(ByVal pEnv As CDBEnvironment) As CARERecord Implements IRecordCreate.CreateInstance
      Return New ActionLink(mvEnv)
    End Function


#End Region

  End Class
End Namespace
