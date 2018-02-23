Namespace Access

  Public Class MembershipPrice
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum MembershipPriceFields
      AllFields = 0
      MembershipType
      PaymentMethod
      PaymentFrequency
      Rate
      Overseas
      Activity
      ActivityValue
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("membership_type")
        .Add("payment_method")
        .Add("payment_frequency")
        .Add("rate")
        .Add("overseas")
        .Add("activity")
        .Add("activity_value")

        .Item(MembershipPriceFields.MembershipType).PrimaryKey = True
        .Item(MembershipPriceFields.PaymentMethod).PrimaryKey = True
        .Item(MembershipPriceFields.PaymentFrequency).PrimaryKey = True
        .Item(MembershipPriceFields.Rate).PrimaryKey = True

        If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbMembershipPricesOverseas) Then
          .Item(MembershipPriceFields.Overseas).InDatabase = False
          .Item(MembershipPriceFields.Activity).InDatabase = False
          .Item(MembershipPriceFields.ActivityValue).InDatabase = False
        End If
        .Item(MembershipPriceFields.Activity).PrefixRequired = True
        .Item(MembershipPriceFields.ActivityValue).PrefixRequired = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "mp"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "membership_prices"
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
    Public ReadOnly Property MembershipTypeCode() As String
      Get
        Return mvClassFields(MembershipPriceFields.MembershipType).Value
      End Get
    End Property
    Public ReadOnly Property PaymentMethod() As String
      Get
        Return mvClassFields(MembershipPriceFields.PaymentMethod).Value
      End Get
    End Property
    Public ReadOnly Property PaymentFrequency() As String
      Get
        Return mvClassFields(MembershipPriceFields.PaymentFrequency).Value
      End Get
    End Property
    Public ReadOnly Property RateCode() As String
      Get
        Return mvClassFields(MembershipPriceFields.Rate).Value
      End Get
    End Property
    Public ReadOnly Property Overseas() As Boolean
      Get
        Return mvClassFields(MembershipPriceFields.Overseas).Bool
      End Get
    End Property
    Public ReadOnly Property Activity() As String
      Get
        Return mvClassFields(MembershipPriceFields.Activity).Value
      End Get
    End Property
    Public ReadOnly Property ActivityValue() As String
      Get
        Return mvClassFields(MembershipPriceFields.ActivityValue).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(MembershipPriceFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(MembershipPriceFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non AutoGenerated Code"

    Private mvConcessionary As Boolean
    Private mvFirstPeriodsProduct As String
    Private mvRateDesc As String
    Private mvCurrentPrice As Double
    Private mvAvailable As Boolean

    Protected Overrides Sub ClearFields()
      MyBase.ClearFields()
      mvConcessionary = False
      mvFirstPeriodsProduct = ""
      mvRateDesc = ""
      mvAvailable = False
    End Sub

    Public ReadOnly Property Concessionary() As Boolean
      Get
        Return mvConcessionary
      End Get
    End Property
    Public ReadOnly Property RateDesc() As String
      Get
        Return mvRateDesc
      End Get
    End Property
    Public ReadOnly Property FirstPeriodsProduct() As String
      Get
        Return mvFirstPeriodsProduct
      End Get
    End Property
    Public ReadOnly Property CurrentPrice() As Double
      Get
        Return mvCurrentPrice
      End Get
    End Property

    Public Property Available As Boolean
      Get
        Return mvAvailable
      End Get
      Set(ByVal pValue As Boolean)
        mvAvailable = pValue
      End Set
    End Property

    Public Overrides Sub InitFromRecordSet(ByVal pRecordSet As Data.CDBRecordSet)
      MyBase.InitFromRecordSet(pRecordSet)
      'concessionary is not an attribute of membership_prices table. It is read from the rates table.
      If pRecordSet.Fields.ContainsKey("concessionary") Then mvConcessionary = pRecordSet.Fields("concessionary").Bool
      If pRecordSet.Fields.ContainsKey("first_periods_product") Then mvFirstPeriodsProduct = pRecordSet.Fields("first_periods_product").Value
      If pRecordSet.Fields.ContainsKey("rate_desc") Then mvRateDesc = pRecordSet.Fields("rate_desc").Value
      If pRecordSet.Fields.ContainsKey("current_price") Then mvCurrentPrice = pRecordSet.Fields("current_price").DoubleValue
    End Sub

    Public Shared Function GetMembershipPrices(ByVal pEnv As CDBEnvironment, Optional ByVal pMembershipTypeCode As String = "") As CollectionList(Of MembershipPrice)
      Dim vMembershipPrices As New CollectionList(Of MembershipPrice)
      If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMembershipPrices) Then
        Dim vMembershipPrice As New MembershipPrice(pEnv)
        vMembershipPrice.Init()
        Dim vFields As String = vMembershipPrice.GetRecordSetFields() & ", concessionary,first_periods_product,rate_desc,current_price"
        Dim vWhereFields As New CDBFields
        If pMembershipTypeCode.Length > 0 Then vWhereFields.Add("mp.membership_type", pMembershipTypeCode)
        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.Add("membership_types mt", "mp.membership_type", "mt.membership_type")
        vAnsiJoins.Add("rates r", "mt.first_periods_product", "r.product", "mp.rate", "r.rate")
        Dim vSQL As New SQLStatement(pEnv.Connection, vFields, vMembershipPrice.DatabaseTableName & " " & vMembershipPrice.TableAlias, vWhereFields, "mp.membership_type, payment_method, payment_frequency, concessionary", vAnsiJoins)
        Dim vRS As CDBRecordSet = vSQL.GetRecordSet()
        While vRS.Fetch()
          Dim vMP As New MembershipPrice(pEnv)
          vMP.InitFromRecordSet(vRS)
          vMembershipPrices.Add(vMP.MembershipTypeCode & "-" & vMP.PaymentMethod & "-" & vMP.PaymentFrequency & "-" & vMP.RateCode, vMP)
        End While
        vRS.CloseRecordSet()
      End If
      Return vMembershipPrices
    End Function

    Public Shared Function HasPrimaryMembershipPrice(ByVal pMembershipPrices As CollectionList(Of MembershipPrice), ByVal pMembershipTypeCode As String, ByVal pPaymentMethod As String, ByVal pPaymentFrequency As String) As Boolean
      For Each vMembershipPrice As MembershipPrice In pMembershipPrices
        If vMembershipPrice.MembershipTypeCode = pMembershipTypeCode AndAlso _
           vMembershipPrice.PaymentMethod = pPaymentMethod AndAlso _
           vMembershipPrice.PaymentFrequency = pPaymentFrequency AndAlso _
           vMembershipPrice.Concessionary = False Then
          Return True
        End If
      Next
    End Function

    Public Shared Function GetValidMembershipPrice(ByVal pEnv As CDBEnvironment, ByVal pMembershipPrices As CollectionList(Of MembershipPrice), ByVal pMembershipTypeCode As String, ByVal pPaymentMethod As String, ByVal pPaymentFrequency As String, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, Optional ByVal pPrimaryOnly As Boolean = True, Optional ByVal pRate As String = "") As MembershipPrice
      SetAvailableMembershipPrices(pEnv, pMembershipPrices, pMembershipTypeCode, pPaymentMethod, pPaymentFrequency, pContactNumber, pAddressNumber, pPrimaryOnly, pRate)
      For Each vMembershipPrice As MembershipPrice In pMembershipPrices
        If vMembershipPrice.Available Then Return vMembershipPrice
      Next
      Return Nothing
    End Function

    Public Shared Function GetAvailableMembershipPrices(ByVal pEnv As CDBEnvironment, ByVal pMembershipTypeCode As String, ByVal pPaymentMethod As String, ByVal pPaymentFrequency As String, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer) As CDBDataTable
      Dim vMembershipPrices As CollectionList(Of MembershipPrice) = GetMembershipPrices(pEnv, pMembershipTypeCode)
      SetAvailableMembershipPrices(pEnv, vMembershipPrices, pMembershipTypeCode, pPaymentMethod, pPaymentFrequency, pContactNumber, pAddressNumber, False, "")
      Dim vDT As New CDBDataTable
      vDT.AddColumnsFromList("MembershipType,PaymentMethod,PaymentFrequency,FirstPeriodsProduct,Rate,RateDesc,CurrentPrice,Concessionary")
      For Each vMembershipPrice As MembershipPrice In vMembershipPrices
        If vMembershipPrice.Available Then
          Dim vRow As CDBDataRow = vDT.AddRow
          vRow.Item("MembershipType") = vMembershipPrice.MembershipTypeCode
          vRow.Item("PaymentMethod") = vMembershipPrice.PaymentMethod
          vRow.Item("PaymentFrequency") = vMembershipPrice.PaymentFrequency
          vRow.Item("FirstPeriodsProduct") = vMembershipPrice.FirstPeriodsProduct
          vRow.Item("Rate") = vMembershipPrice.RateCode
          vRow.Item("RateDesc") = vMembershipPrice.RateDesc
          vRow.Item("CurrentPrice") = vMembershipPrice.CurrentPrice.ToString
          vRow.Item("Concessionary") = BooleanString(vMembershipPrice.Concessionary)
        End If
      Next
      Return vDT
    End Function

    Private Shared Sub SetAvailableMembershipPrices(ByVal pEnv As CDBEnvironment, ByVal pMembershipPrices As CollectionList(Of MembershipPrice), ByVal pMembershipTypeCode As String, ByVal pPaymentMethod As String, ByVal pPaymentFrequency As String, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pPrimaryOnly As Boolean, ByVal pRate As String)
      Dim vCheckOverseas As Boolean
      Dim vOverseasAddress As Boolean
      Dim vActivityDataTable As DataTable = Nothing
      Dim vCheckActivities As Boolean
      Dim vFoundActivityPrice As Boolean

      'First check if any prices have the overseas flag - If so check the overseas flag on the contact and remember to filter by overseas
      For Each vMembershipPrice As MembershipPrice In pMembershipPrices
        If vMembershipPrice.Overseas Then
          'Check if contact has overseas address if not then don't add this membership price
          Dim vAnsiJoins As New AnsiJoins
          vAnsiJoins.Add("countries c", "a.country", "c.country")
          Dim vWhereFields As New CDBFields
          vWhereFields.Add("address_number", pAddressNumber)
          Dim vOverseasSQL As New SQLStatement(pEnv.Connection, "overseas", "addresses a", vWhereFields, "", vAnsiJoins)
          vOverseasAddress = BooleanValue(vOverseasSQL.GetValue)
          vCheckOverseas = True
          Exit For
        End If
      Next
      'Next check if any prices have the activities set - If so get the activities for the contact and remember to filter by it
      For Each vMembershipPrice As MembershipPrice In pMembershipPrices
        If vMembershipPrice.Activity.Length > 0 AndAlso vMembershipPrice.ActivityValue.Length > 0 Then
          'Check if contact this as a valid activity if not then don't add this membership price
            Dim vWhereFields As New CDBFields
            vWhereFields.Add("contact_number", pContactNumber)
            'TODO maybe should not be todays date but the start/renewal date
            vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
            vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
            Dim vActivitySQL As New SQLStatement(pEnv.Connection, "activity,activity_value", "contact_categories", vWhereFields)
            vActivityDataTable = vActivitySQL.GetDataTable()
          vCheckActivities = True
        End If
      Next
      For Each vMembershipPrice As MembershipPrice In pMembershipPrices
        vMembershipPrice.Available = False
        If pMembershipTypeCode.Length > 0 AndAlso vMembershipPrice.MembershipTypeCode <> pMembershipTypeCode Then Continue For
        If pPaymentMethod.Length > 0 AndAlso vMembershipPrice.PaymentMethod <> pPaymentMethod Then Continue For
        If pPaymentFrequency.Length > 0 AndAlso vMembershipPrice.PaymentFrequency <> pPaymentFrequency Then Continue For
        If pRate.Length > 0 AndAlso vMembershipPrice.RateCode <> pRate Then Continue For
        If pPrimaryOnly AndAlso vMembershipPrice.Concessionary Then Continue For
        'If any prices were overseas then the overseas flag must match
        If vCheckOverseas AndAlso vMembershipPrice.Overseas <> vOverseasAddress Then Continue For
        If vCheckActivities Then
          If vMembershipPrice.Activity.Length > 0 AndAlso vMembershipPrice.ActivityValue.Length > 0 Then
            'Check if contact this as a valid activity if not then don't add this membership price
            Dim vActivityFound As Boolean = False
            If vActivityDataTable IsNot Nothing Then
              For Each vRow As DataRow In vActivityDataTable.Rows
                If vRow("activity").ToString = vMembershipPrice.Activity AndAlso vRow("activity_value").ToString = vMembershipPrice.ActivityValue Then
                  vActivityFound = True
                  vFoundActivityPrice = True
                  Exit For
                End If
              Next
            End If
            If vActivityFound = False Then Continue For
          End If
        End If
        vMembershipPrice.Available = True
      Next
      'If any prices are available based on an activity then set any prices that don't have an activity as not availalbe
      If vCheckActivities And vFoundActivityPrice Then
        For Each vMembershipPrice As MembershipPrice In pMembershipPrices
          If vMembershipPrice.Available AndAlso vMembershipPrice.Activity.Length = 0 Then vMembershipPrice.Available = False
        Next
      End If
    End Sub

    Public Shared Function GetMembershipPrice(ByVal pMembershipPrices As CollectionList(Of MembershipPrice), ByVal pMembershipTypeCode As String, ByVal pPaymentMethod As String, ByVal pPaymentFrequency As String, ByVal pRateCode As String) As MembershipPrice
      For Each vMembershipPrice As MembershipPrice In pMembershipPrices
        If vMembershipPrice.MembershipTypeCode = pMembershipTypeCode AndAlso _
           vMembershipPrice.PaymentMethod = pPaymentMethod AndAlso _
           vMembershipPrice.PaymentFrequency = pPaymentFrequency AndAlso _
           vMembershipPrice.RateCode = pRateCode Then
          Return vMembershipPrice
        End If
      Next
      Return Nothing
    End Function

#End Region


  End Class
End Namespace
