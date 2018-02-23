<Serializable()> _
Public Class CampaignCopyInfo
  Private mvCampaign As String
  Private mvAppeal As String
  Private mvAppealDesc As String
  Private mvSegment As String
  Private mvCollectionNumber As Integer
  'Private mvAppealType As String
  'Private mvCampaignItem As CampaignItem
  Private mvCampaignCopyType As CampaignCopyTypes
  Private mvCollection As String
  Private mvCollectionDesc As String
  Private mvSegmentDesc As String
  Private mvAppealType As String
  Private mvCriteriaSet As Integer
  Private mvHasCriteriaSetDetails As Boolean
  Private mvHasCriteriaSetSelectionSteps As Boolean

  Public Enum CampaignCopyTypes
    cctAppeal
    cctH2HCollection
    cctMannedCollection
    cctUnMannedCollection
    cctSegment
    cctSegmentCriteria
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
  Public Sub New(ByVal pCampaignItem As CampaignItem, ByVal pAppealDesc As String, ByVal pCollection As String, ByVal pCollectionDesc As String, ByVal pSegmentDesc As String)
    Me.New(pCampaignItem, pAppealDesc, pCollection, pCollectionDesc, pSegmentDesc, False)
  End Sub

  Public Sub New(ByVal pCampaignItem As CampaignItem, ByVal pAppealDesc As String, ByVal pCollection As String, ByVal pCollectionDesc As String, ByVal pSegmentDesc As String, ByVal pCopySegmentCriteria As Boolean)

    mvCampaign = pCampaignItem.Campaign
    mvAppeal = pCampaignItem.Appeal
    mvAppealType = pCampaignItem.AppealTypeCode
    mvCampaignCopyType = CampaignCopyTypes.cctAppeal
    mvAppealDesc = pAppealDesc
    mvCollection = pCollection
    mvCollectionDesc = pCollectionDesc
    mvSegmentDesc = pSegmentDesc
    If pCampaignItem.ItemType = CampaignItem.CampaignItemTypes.citSegment Then
      mvSegment = pCampaignItem.Segment
      mvCampaignCopyType = CampaignCopyTypes.cctSegment
      If pCopySegmentCriteria Then
        mvCriteriaSet = pCampaignItem.CriteriaSet
        mvHasCriteriaSetDetails = pCampaignItem.HasCriteriaSetDetails
        mvHasCriteriaSetSelectionSteps = pCampaignItem.HasCriteriaSelectionSteps
        mvCampaignCopyType = CampaignCopyTypes.cctSegmentCriteria
      End If
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
  Public ReadOnly Property AppealDesc() As String
    Get
      Return mvAppealDesc
    End Get
  End Property
  Public ReadOnly Property AppealType() As String
    Get
      Return mvAppealType
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

  Public ReadOnly Property Collection() As String
    Get
      Return mvCollection
    End Get
  End Property

  Public ReadOnly Property CollectionDesc() As String
    Get
      Return mvCollectionDesc
    End Get
  End Property

  Public ReadOnly Property CriteriaSet() As Integer
    Get
      Return mvCriteriaSet
    End Get
  End Property

  Public ReadOnly Property HasCriteriaSetDetails() As Boolean
    Get
      Return mvHasCriteriaSetDetails
    End Get
  End Property

  Public ReadOnly Property HasCriteriaSelectionSteps() As Boolean
    Get
      Return mvHasCriteriaSetSelectionSteps
    End Get
  End Property

  Public ReadOnly Property SegmentDesc() As String
    Get
      Return mvSegmentDesc
    End Get
  End Property

  Public Sub FillParameterList(ByRef pList As ParameterList)
    pList("Campaign") = mvCampaign
    pList("Appeal") = mvAppeal
    Select Case mvCampaignCopyType
      Case CampaignCopyTypes.cctAppeal
        pList("AppealDesc") = mvAppealDesc
      Case CampaignCopyTypes.cctSegment, CampaignCopyTypes.cctSegmentCriteria
        pList("Segment") = mvSegment
        If mvCampaignCopyType = CampaignCopyTypes.cctSegmentCriteria Then
          pList.IntegerValue("CriteriaSet") = mvCriteriaSet
        End If
      Case CampaignCopyTypes.cctH2HCollection, _
           CampaignCopyTypes.cctMannedCollection, _
           CampaignCopyTypes.cctUnMannedCollection
        pList.IntegerValue("CollectionNumber") = mvCollectionNumber
    End Select
  End Sub
End Class
