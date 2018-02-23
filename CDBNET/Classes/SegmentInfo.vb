Public Class SegmentInfo
	Private mvCampaign As String
	Private mvAppeal As String
	Private mvSegment As String

  Public Sub New(ByVal pCampaign As String, ByVal pAppeal As String, ByVal pSegment As String)
    mvCampaign = pCampaign
    mvAppeal = pAppeal
    mvSegment = pSegment
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

	Public ReadOnly Property Segment() As String
		Get
			Return mvSegment
		End Get
	End Property



End Class
