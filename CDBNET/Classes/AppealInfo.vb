<Serializable()> _
Public Class AppealInfo
  Private mvCampaign As String
  Private mvAppeal As String
  Private mvAppealDesc As String
  Private mvSegment As String
  Private mvCollectionNumber As Integer
  'Private mvAppealType As String
  'Private mvCampaignItem As CampaignItem
  Private mvCampaignCopyType As CampaignCopyTypes

  Public Enum CampaignCopyTypes
    cctAppeal
    cctH2HCollection
    cctMannedCollection
    cctUnMannedCollection
    cctSegment
  End Enum

  'Public Sub New(ByVal pCampaign As String, ByVal pAppeal As String, ByVal pSegment As String, ByVal pCollectionNumber As Integer, ByVal pAppealDesc As String) ', ByVal pAppealType As String)

  '  mvCampaign = pCampaign
  '  mvAppeal = pAppeal
  '  mvCampaignCopyType = CampaignCopyTypes.cctAppeal
  '  mvAppealDesc = pAppealDesc
  '  If pSegment.Length > 0 Then
  '    mvSegment = pSegment
  '    mvCampaignCopyType = CampaignCopyTypes.cctSegment
  '  End If

  '  If pCollectionNumber > 0 Then
  '    mvCollectionNumber = pCollectionNumber
  '    mvCampaignCopyType = CampaignCopyTypes.cctCollection
  '  End If
  'End Sub

  Public Sub New(ByVal pCampaignItem As CampaignItem, ByVal pAppealDesc As String)

    mvCampaign = pCampaignItem.Campaign
    mvAppeal = pCampaignItem.Appeal
    mvCampaignCopyType = CampaignCopyTypes.cctAppeal
    mvAppealDesc = pAppealDesc
    If pCampaignItem.ItemType = CampaignItem.CampaignItemTypes.citSegment Then
      mvSegment = pCampaignItem.Segment
      mvCampaignCopyType = CampaignCopyTypes.cctSegment
    End If

    If pCampaignItem.ItemType = CampaignItem.CampaignItemTypes.citCollection Then
      mvCollectionNumber = pCampaignItem.CollectionNumber
      Select Case pCampaignItem.AppealType
        Case CampaignItem.AppealTypes.atH2HCollection
          mvCampaignCopyType = CampaignCopyTypes.cctH2HCollection
        Case CampaignItem.AppealTypes.atMannedCollection
          mvCampaignCopyType = CampaignCopyTypes.cctMannedCollection
        Case CampaignItem.AppealTypes.atUnMannedCollection
          mvCampaignCopyType = CampaignCopyTypes.cctUnMannedCollection
      End Select
    End If
  End Sub


  'Public Sub New(ByVal pCampaignItem As CampaignItem, ByVal pAppealDesc As String)
  '  'MyBase.New(pCampaignItem.Code, pCampaignItem.StartDate, pCampaignItem.EndDate)
  '  mvCampaignItem = pCampaignItem
  '  mvAppealDesc = pAppealDesc
  'End Sub

  'Public Sub New()
  '	mvCampaign = ""
  '	mvAppeal = ""
  'End Sub

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

  Public ReadOnly Property Segment() As String
    Get
      Return mvSegment
    End Get
  End Property

  Public ReadOnly Property CollectionNumber() As Integer
    Get
      Return mvCollectionNumber
    End Get
  End Property

  Public ReadOnly Property CampaignCopyType() As CampaignCopyTypes
    Get
      Return mvCampaignCopyType
    End Get
  End Property

  Public Sub FillParameterList(ByRef pList As ParameterList)
    pList("Campaign") = mvCampaign
    pList("Appeal") = mvAppeal
    Select Case mvCampaignCopyType
      Case CampaignCopyTypes.cctAppeal
        pList("AppealDesc") = mvAppealDesc
      Case CampaignCopyTypes.cctSegment
        pList("Segment") = mvSegment
      Case CampaignCopyTypes.cctH2HCollection, _
           CampaignCopyTypes.cctMannedCollection, _
           CampaignCopyTypes.cctUnMannedCollection
        pList.IntegerValue("CollectionNumber") = mvCollectionNumber
    End Select
  End Sub
End Class
