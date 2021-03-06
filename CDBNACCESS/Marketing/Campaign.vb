Namespace Access

  Partial Public Class Campaign
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum CampaignFields
      AllFields = 0
      Campaign
      CampaignDesc
      StartDate
      EndDate
      Notes
      Manager
      CampaignBusinessType
      CampaignStatus
      CampaignStatusDate
      CampaignStatusReason
      ActualIncome
      ActualIncomeDate
      TotalItemisedCost
      Topic
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("campaign")
        .Add("campaign_desc")
        .Add("start_date", CDBField.FieldTypes.cftDate)
        .Add("end_date", CDBField.FieldTypes.cftDate)
        .Add("notes", CDBField.FieldTypes.cftMemo)
        .Add("manager")
        .Add("campaign_business_type")
        .Add("campaign_status")
        .Add("campaign_status_date", CDBField.FieldTypes.cftDate)
        .Add("campaign_status_reason", CDBField.FieldTypes.cftMemo)
        .Add("actual_income", CDBField.FieldTypes.cftNumeric)
        .Add("actual_income_date", CDBField.FieldTypes.cftDate)
        .Add("total_itemised_cost", CDBField.FieldTypes.cftNumeric)
        .Add("topic").InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbTelemarketing)

        .Item(CampaignFields.Campaign).PrimaryKey = True
        .SetUniqueField(CampaignFields.Campaign)
        mvClassFields.Item(CampaignFields.TotalItemisedCost).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCampaignItemisedCosts)
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "c"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "campaigns"
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
    Public ReadOnly Property CampaignCode() As String
      Get
        Return mvClassFields(CampaignFields.Campaign).Value
      End Get
    End Property
    Public ReadOnly Property CampaignDesc() As String
      Get
        Return mvClassFields(CampaignFields.CampaignDesc).Value
      End Get
    End Property
    Public ReadOnly Property StartDate() As String
      Get
        Return mvClassFields(CampaignFields.StartDate).Value
      End Get
    End Property
    Public ReadOnly Property EndDate() As String
      Get
        Return mvClassFields(CampaignFields.EndDate).Value
      End Get
    End Property
    Public ReadOnly Property Notes() As String
      Get
        Return mvClassFields(CampaignFields.Notes).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(CampaignFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property Manager() As String
      Get
        Return mvClassFields(CampaignFields.Manager).Value
      End Get
    End Property
    Public ReadOnly Property CampaignBusinessType() As String
      Get
        Return mvClassFields(CampaignFields.CampaignBusinessType).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(CampaignFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property CampaignStatus() As String
      Get
        Return mvClassFields(CampaignFields.CampaignStatus).Value
      End Get
    End Property
    Public ReadOnly Property CampaignStatusDate() As String
      Get
        Return mvClassFields(CampaignFields.CampaignStatusDate).Value
      End Get
    End Property
    Public ReadOnly Property CampaignStatusReason() As String
      Get
        Return mvClassFields(CampaignFields.CampaignStatusReason).Value
      End Get
    End Property

    Public ReadOnly Property ActualIncome() As Double
      Get
        Return mvClassFields(CampaignFields.ActualIncome).DoubleValue
      End Get
    End Property
    Public ReadOnly Property ActualIncomeDate() As String
      Get
        Return mvClassFields(CampaignFields.ActualIncomeDate).Value
      End Get
    End Property
    Public ReadOnly Property TotalItemisedCost() As Double
      Get
        Return mvClassFields(CampaignFields.TotalItemisedCost).DoubleValue
      End Get
    End Property
    Public ReadOnly Property Topic() As String
      Get
        Return mvClassFields(CampaignFields.Topic).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"

    Dim mvMarkHistorical As String

    Public Overloads Sub Init(ByVal pCampaign As String)
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("campaign", pCampaign)
      MyBase.InitWithPrimaryKey(vWhereFields)
    End Sub
    Public Sub InitFromRecordSetCampaign(ByVal pRecordSet As CDBRecordSet)
      MyBase.InitFromRecordSetFields(pRecordSet, GetRecordSetFields)
    End Sub
    Public ReadOnly Property MarkHistorical() As String
      Get
        If mvMarkHistorical = "" And mvClassFields(CampaignFields.CampaignStatus).Value = "" Then
          Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, "mark_historical", "campaign_statuses", New CDBField("campaign_statuses", mvClassFields(CampaignFields.CampaignStatus).Value)).GetRecordSet
          If vRecordSet.Fetch() Then
            mvMarkHistorical = vRecordSet.Fields(1).Value
          End If
          vRecordSet.CloseRecordSet()
        End If
        Return mvMarkHistorical
      End Get
    End Property

#End Region
  End Class
End Namespace
