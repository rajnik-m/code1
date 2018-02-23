Public Class ParentItem
  'Public Enum ParentItemTypes
  '  pitNone
  '  pitCampaign
  '  pitAppeal
  '  pitCollection
  'End Enum

  'Private mvParentItemType As ParentItemTypes
  'Private mvCampaign As String
  'Private mvAppeal As String
  'Private mvSegment As String
  'Private mvCollectionNumber As Integer
  'Private mvStartDate As String
  'Private mvEndDate As String
  'Public Sub New(ByVal pParentItemType As ParentItemTypes, ByVal plist As ParameterList)
  '  mvParentItemType = pParentItemType
  '  With vlist
  '    mvAppeal = plist("Appeal")
  '    mvCampaign = plist("Campaign")
  '    mvCollectionNumber = plist("CollectionNumber")
  '    mvCampaign = plist("Campaign")
  '    mvCampaign = plist("Campaign")
  '  End With
  '  Dim vItems() As String = pCode.Split("_"c)
  '  mvCampaign = vItems(0)
  '  mvItemType = CampaignItemTypes.citCampaign
  '  If vItems.Length > 1 Then
  '    mvAppeal = vItems(1)
  '    mvItemType = CampaignItemTypes.citAppeal
  '    Dim vAppealTypeCode As String = ""
  '    If vItems.Length > 2 Then vAppealTypeCode = vItems(2)
  '    mvAppealTypeCode = vAppealTypeCode
  '    Select Case vAppealTypeCode
  '      Case "M"
  '        mvAppealType = AppealTypes.atMannedCollection
  '      Case "U"
  '        mvAppealType = AppealTypes.atUnMannedCollection
  '      Case "H"
  '        mvAppealType = AppealTypes.atH2HCollection
  '      Case Else
  '        mvAppealType = AppealTypes.atSegment
  '    End Select
  '    If vItems.Length > 3 Then
  '      If mvAppealType = AppealTypes.atSegment Then
  '        mvSegment = vItems(3)
  '        mvItemType = CampaignItemTypes.citSegment
  '      Else
  '        mvCollectionNumber = FormHelper.IntegerValue(vItems(3))
  '        mvItemType = CampaignItemTypes.citCollection
  '      End If
  '    End If
  '  End If
  'End Sub

  'Public Sub FillParameterList(ByRef pList As ParameterList)
  '  pList("Campaign") = mvCampaign
  '  Select Case mvItemType
  '    Case CampaignItemTypes.citAppeal
  '      pList("Appeal") = mvAppeal
  '      pList("AppealType") = mvAppealTypeCode
  '    Case CampaignItemTypes.citSegment
  '      pList("Appeal") = mvAppeal
  '      pList("AppealType") = mvAppealTypeCode
  '      pList("Segment") = mvSegment
  '    Case CampaignItemTypes.citCollection
  '      pList("Appeal") = mvAppeal
  '      pList("AppealType") = mvAppealTypeCode
  '      If mvCollectionNumber > 0 Then pList.IntegerValue("CollectionNumber") = mvCollectionNumber
  '  End Select
  'End Sub

  'Public ReadOnly Property Appeal() As String
  '  Get
  '    Return mvAppeal
  '  End Get
  'End Property
  'Public ReadOnly Property AppealType() As AppealTypes
  '  Get
  '    Return mvAppealType
  '  End Get
  'End Property
  'Public ReadOnly Property Campaign() As String
  '  Get
  '    Return mvCampaign
  '  End Get
  'End Property
  'Public ReadOnly Property Code() As String
  '  Get
  '    Return mvCode
  '  End Get
  'End Property
  'Public ReadOnly Property CollectionNumber() As Integer
  '  Get
  '    Return mvCollectionNumber
  '  End Get
  'End Property
  'Public ReadOnly Property Existing() As Boolean
  '  Get
  '    Select Case mvItemType
  '      Case CampaignItemTypes.citCampaign
  '        Return mvCampaign.Length > 0
  '      Case CampaignItemTypes.citAppeal
  '        Return mvAppeal.Length > 0
  '      Case CampaignItemTypes.citSegment
  '        Return mvSegment.Length > 0
  '      Case CampaignItemTypes.citCollection
  '        Return mvCollectionNumber > 0
  '    End Select
  '  End Get
  'End Property
  'Public ReadOnly Property ItemType() As CampaignItemTypes
  '  Get
  '    Return mvItemType
  '  End Get
  'End Property
  'Public ReadOnly Property Segment() As String
  '  Get
  '    Return mvSegment
  '  End Get
  'End Property

End Class
