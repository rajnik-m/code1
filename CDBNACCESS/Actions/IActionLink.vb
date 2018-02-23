Public Interface IActionLink
  Inherits ICARERecord

  Enum ActionLinkObjectTypes
    alotContact
    alotOrganisation
    alotDocument
    alotFundraising
    alotExamCentre
    alotWorkstream
    alotContactPosition
  End Enum

  Enum ActionLinkTypes
    altActioner
    altManager
    altRelated
  End Enum

  ReadOnly Property ActionNumber As Integer
  ReadOnly Property LinkedItemId As Integer
  ReadOnly Property Type As String
  ReadOnly Property LinkType() As IActionLink.ActionLinkTypes
  Sub InitFromObjectType(ByVal pObjectType As IActionLink.ActionLinkObjectTypes)
  Overloads Sub Init(ByVal pObjectType As ActionLinkObjectTypes,
                     ByVal pActionNumber As Integer,
                     ByVal pContactNumber As Integer,
                     ByVal pLinkType As ActionLinkTypes)
  Overloads Sub Create(ByVal pEnv As CDBEnvironment,
                       ByVal pActionNumber As Integer,
                       ByVal pContactNumber As Integer,
                       ByVal pLinkType As IActionLink.ActionLinkTypes)
  ReadOnly Property ObjectLinkType() As IActionLink.ActionLinkObjectTypes
End Interface
