<Serializable()> _
Public Class AppealCollectionInfo
  Private mvCampaign As String
  Private mvAppeal As String
  Private mvCollection As String

  Public Sub New(ByVal pCampaign As String, ByVal pAppeal As String, ByVal pCollection As String)
    mvCampaign = pCampaign
    mvAppeal = pAppeal
    mvCollection = pCollection
  End Sub

  Public ReadOnly Property Campaign() As String
    Get
      Return mvCampaign
    End Get
  End Property
  Public ReadOnly Property Appeal() As String
    Get
      Return mvAppeal
    End Get
  End Property

  Public ReadOnly Property Collection() As String
    Get
      Return mvCollection
    End Get
  End Property
End Class
