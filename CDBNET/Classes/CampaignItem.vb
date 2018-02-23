Public Class CampaignItem

  Public Enum AppealTypes
    atSegment
    atMannedCollection
    atUnMannedCollection
    atH2HCollection
  End Enum

  Public Enum CampaignItemTypes
    citCampaign
    citAppeal
    citSegment
    citCollection
  End Enum

  Private mvItemType As CampaignItemTypes
  Private mvParentItemType As CampaignItemTypes
  Private mvCampaign As String
  Private mvAppeal As String = ""
  Private mvAppealType As AppealTypes
  Private mvSegment As String = ""
  Private mvCollectionNumber As Integer
  Private mvCode As String
  Private mvAppealTypeCode As String
  Private mvStartDate As String
  Private mvEndDate As String
  Private mvResourcesProducedOn As String
  Private mvResourcesProducedOnSet As Boolean
  Private mvSource As String
  Private mvProduct As String
  Private mvRate As String
  Private mvBankAccount As String
  Private mvTYL As String
  Private mvAdditionalValues As ParameterList = New ParameterList
  Private mvStartTime As String
  Private mvEndTime As String
  Private mvTimesSet As Boolean
  Private mvIncomeFieldsSet As Boolean
  Private mvCollBankAccount As String
  Private mvCollBankAccSet As Boolean
  Private mvAppealActionNumber As Integer
  Private mvFirstAppealResource As String
  Private mvFirstAppealResourceSet As Boolean

  Public Sub New(ByVal pCode As String, ByVal pStartDate As String, ByVal pEndDate As String)
    mvCode = pCode
    Dim vItems() As String = pCode.Split("_"c)
    mvCampaign = vItems(0)
    mvItemType = CampaignItemTypes.citCampaign
    mvParentItemType = Nothing
    If vItems.Length > 1 Then
      mvAppeal = vItems(1)
      mvItemType = CampaignItemTypes.citAppeal
      mvParentItemType = CampaignItemTypes.citCampaign
      Dim vAppealTypeCode As String = ""
      If vItems.Length > 2 Then vAppealTypeCode = vItems(2)
      mvAppealTypeCode = vAppealTypeCode
      Select Case vAppealTypeCode
        Case "M"
          mvAppealType = AppealTypes.atMannedCollection
        Case "U"
          mvAppealType = AppealTypes.atUnMannedCollection
        Case "H"
          mvAppealType = AppealTypes.atH2HCollection
        Case Else
          mvAppealType = AppealTypes.atSegment
      End Select
      If vItems.Length > 3 Then
        If mvAppealType = AppealTypes.atSegment Then
          mvSegment = vItems(3)
          mvItemType = CampaignItemTypes.citSegment
        Else
          mvCollectionNumber = IntegerValue(vItems(3))
          mvItemType = CampaignItemTypes.citCollection
        End If
        mvParentItemType = CampaignItemTypes.citAppeal
      End If
    End If
    mvStartDate = pStartDate
    mvEndDate = pEndDate
  End Sub

  Public Sub FillParameterList(ByRef pList As ParameterList)
    pList("Campaign") = mvCampaign
    Select Case mvItemType
      Case CampaignItemTypes.citAppeal
        pList("Appeal") = mvAppeal
        pList("AppealType") = mvAppealTypeCode
      Case CampaignItemTypes.citSegment
        pList("Appeal") = mvAppeal
        pList("AppealType") = mvAppealTypeCode
        pList("Segment") = mvSegment
      Case CampaignItemTypes.citCollection
        pList("Appeal") = mvAppeal
        pList("AppealType") = mvAppealTypeCode
        If mvCollectionNumber > 0 Then pList.IntegerValue("CollectionNumber") = mvCollectionNumber
    End Select
  End Sub

  Public ReadOnly Property Appeal() As String
    Get
      Return mvAppeal
    End Get
  End Property
  Public ReadOnly Property AppealType() As AppealTypes
    Get
      Return mvAppealType
    End Get
  End Property

  Public ReadOnly Property AppealTypeCode() As String
    Get
      Return mvAppealTypeCode
    End Get
  End Property

  Public ReadOnly Property Campaign() As String
    Get
      Return mvCampaign
    End Get
  End Property
  Public ReadOnly Property Code() As String
    Get
      Return mvCode
    End Get
  End Property
  Public ReadOnly Property CollectionNumber() As Integer
    Get
      Return mvCollectionNumber
    End Get
  End Property
  Public ReadOnly Property Existing() As Boolean
    Get
      Select Case mvItemType
        Case CampaignItemTypes.citCampaign
          Return mvCampaign.Length > 0
        Case CampaignItemTypes.citAppeal
          Return mvAppeal.Length > 0
        Case CampaignItemTypes.citSegment
          Return mvSegment.Length > 0
        Case CampaignItemTypes.citCollection
          Return mvCollectionNumber > 0
      End Select
    End Get
  End Property
  Public ReadOnly Property ItemType() As CampaignItemTypes
    Get
      Return mvItemType
    End Get
  End Property
  Public ReadOnly Property ParentItemType() As CampaignItemTypes
    Get
      Return mvParentItemType
    End Get
  End Property
  Public ReadOnly Property Segment() As String
    Get
      Return mvSegment
    End Get
  End Property
  Public ReadOnly Property StartDate() As String
    Get
      Return mvStartDate
    End Get
  End Property

  Public ReadOnly Property EndDate() As String
    Get
      Return mvEndDate
    End Get
  End Property

  Public Property ResourcesProducedOn() As String
    Get
      If Not mvResourcesProducedOnSet Then
        If mvItemType = CampaignItemTypes.citCollection Then
          Dim vList As New ParameterList
          vList("CollectionNumber") = CollectionNumber.ToString
          Dim vDataRow As DataRow = DataHelper.GetCampaignItem(CareServices.XMLCampaignDataSelectionTypes.xcadtCollection, vList)
          If vDataRow IsNot Nothing AndAlso vDataRow.Item("CollectionType").ToString = "U" Then
            mvResourcesProducedOn = vDataRow.Item("ResourcesProducedOn").ToString
            mvResourcesProducedOnSet = True
          Else
            mvResourcesProducedOn = Nothing
          End If
        End If
      End If
      Return mvResourcesProducedOn
    End Get
    Set(ByVal value As String)
      mvResourcesProducedOn = value
      mvResourcesProducedOnSet = True
    End Set
  End Property
  Public Sub SetIncomeFields(ByVal pSource As String, ByVal pProduct As String, ByVal pRate As String, ByVal pBankAccount As String, ByVal pTYL As String)
    mvSource = pSource
    mvProduct = pProduct
    mvRate = pRate
    mvBankAccount = pBankAccount
    mvTYL = pTYL
    mvIncomeFieldsSet = True
  End Sub

  Public ReadOnly Property Source() As String
    Get
      If Not mvIncomeFieldsSet Then
        GetAppealIncomeFields()
      End If
      Return mvSource
    End Get
  End Property

  Public ReadOnly Property Product() As String
    Get
      If Not mvIncomeFieldsSet Then
        GetAppealIncomeFields()
      End If
      Return mvProduct
    End Get
  End Property

  Public ReadOnly Property Rate() As String
    Get
      If Not mvIncomeFieldsSet Then
        GetAppealIncomeFields()
      End If
      Return mvRate
    End Get
  End Property

  Public ReadOnly Property BankAccount() As String
    Get
      If Not mvIncomeFieldsSet Then
        GetAppealIncomeFields()
      End If
      Return mvBankAccount
    End Get
  End Property

  Public ReadOnly Property TYL() As String
    Get
      If Not mvIncomeFieldsSet Then
        GetAppealIncomeFields()
      End If
      Return mvTYL
    End Get
  End Property

  Public Property AdditionalValues() As ParameterList
    Get
      Return mvAdditionalValues
    End Get
    Set(ByVal value As ParameterList)
      mvAdditionalValues = value
    End Set
  End Property

  Public Sub SetCollectionTimes(ByVal pStartTime As String, ByVal pEndTime As String)
    mvStartTime = pStartTime
    mvEndTime = pEndTime
    mvTimesSet = True
  End Sub

  Public ReadOnly Property StartTime() As String
    Get
      If Not mvTimesSet Then
        GetCollectionTimes()
      End If
      Return mvStartTime
    End Get
  End Property

  Public ReadOnly Property EndTime() As String
    Get
      If Not mvTimesSet Then
        GetCollectionTimes()
      End If
      Return mvEndTime
    End Get
  End Property

  Public Property CollectionBankAccount() As String
    Get
      If Not mvCollBankAccSet Then
        GetCollectionBankAccount()
      End If
      Return mvCollBankAccount
    End Get
    Set(ByVal value As String)
      mvCollBankAccount = value
      mvCollBankAccSet = True
    End Set
  End Property

  Private Sub GetCollectionTimes()
    If mvItemType = CampaignItemTypes.citCollection Then
      Dim vList As New ParameterList(True)
      vList("CollectionNumber") = CollectionNumber.ToString
      Dim vDataRow As DataRow = DataHelper.GetCampaignItem(CareServices.XMLCampaignDataSelectionTypes.xcadtCollection, vList)
      If vDataRow IsNot Nothing AndAlso vDataRow.Item("CollectionType").ToString = "M" Then
        mvStartTime = vDataRow.Item("StartTime").ToString
        mvEndTime = vDataRow.Item("EndTime").ToString
        mvTimesSet = True
      Else
        mvStartTime = Nothing
        mvEndTime = Nothing
      End If
    End If
  End Sub

  Private Sub GetCollectionBankAccount()
    If mvItemType = CampaignItemTypes.citCollection Then
      Dim vList As New ParameterList(True)
      vList("CollectionNumber") = CollectionNumber.ToString
      Dim vDataRow As DataRow = DataHelper.GetCampaignItem(CareServices.XMLCampaignDataSelectionTypes.xcadtCollection, vList)
      If vDataRow IsNot Nothing Then
        mvCollBankAccount = vDataRow.Item("BankAccount").ToString
        mvCollBankAccSet = True
      Else
        mvCollBankAccount = Nothing
      End If
    End If
  End Sub

  Private Sub GetAppealIncomeFields()
    If mvItemType = CampaignItemTypes.citAppeal Then
      Dim vList As New ParameterList(True)
      FillParameterList(vList)
      Dim vDataRow As DataRow = DataHelper.GetCampaignItem(CareServices.XMLCampaignDataSelectionTypes.xcadtAppeal, vList)
      If vDataRow IsNot Nothing Then
        mvSource = vDataRow.Item("Source").ToString
        mvProduct = vDataRow.Item("Product").ToString
        mvRate = vDataRow.Item("Rate").ToString
        mvBankAccount = vDataRow.Item("BankAccount").ToString
        mvTYL = vDataRow.Item("ThankYouLetter").ToString
        mvIncomeFieldsSet = True
      Else
        mvSource = Nothing
        mvProduct = Nothing
        mvRate = Nothing
        mvBankAccount = Nothing
        mvTYL = Nothing
      End If
    End If
  End Sub

  Friend Property AppealActionNumber() As Integer
    Get
      Return mvAppealActionNumber
    End Get
    Set(ByVal pValue As Integer)
      mvAppealActionNumber = pValue
    End Set
  End Property

  Public Property FirstAppealResource() As String
    Get
      If Not mvFirstAppealResourceSet Then
        GetFirstAppealResource()
      End If
      Return mvFirstAppealResource
    End Get
    Set(ByVal value As String)
      mvFirstAppealResource = value
      mvFirstAppealResourceSet = True
    End Set
  End Property

  Private Sub GetFirstAppealResource()
    If mvItemType = CampaignItemTypes.citAppeal Then
      Dim vList As New ParameterList(True)
      FillParameterList(vList)
      Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtAppealResources, vList))
      If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
        Dim vDataRow As DataRow = vTable.Rows(0)
        mvFirstAppealResource = "DAFBOX" 'vDataRow.Item("Product").ToString
      Else
        mvFirstAppealResource = ""
      End If
      mvFirstAppealResourceSet = True
    End If
  End Sub
End Class
