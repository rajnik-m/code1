

Namespace Access
  Public Class PostcodeProximityOrg

    Dim mvEnv As CDBEnvironment

    Dim mvOrganisationNumber As Integer
    Dim mvEasting As Integer
    Dim mvNorthing As Integer
    Dim mvPostCode As String
    Public ReadOnly Property OrganisationNumber() As Integer
      Get
        OrganisationNumber = mvOrganisationNumber
      End Get
    End Property
    Public ReadOnly Property Easting() As Integer
      Get
        Easting = mvEasting
      End Get
    End Property
    Public ReadOnly Property Northing() As Integer
      Get
        Northing = mvNorthing
      End Get
    End Property
    Public ReadOnly Property PostCode() As String
      Get
        PostCode = mvPostCode
      End Get
    End Property
    Public Sub Init(ByVal pEnv As CDBEnvironment, ByVal pOrganisationNumber As Integer, ByVal pEasting As Integer, ByVal pNorthing As Integer, ByVal pPostCode As String)
      mvEnv = pEnv
      mvOrganisationNumber = pOrganisationNumber
      mvEasting = pEasting
      mvNorthing = pNorthing
      mvPostCode = pPostCode
    End Sub
  End Class
End Namespace
