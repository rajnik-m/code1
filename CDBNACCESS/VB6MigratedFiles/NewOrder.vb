

Namespace Access
  Public Class NewOrder

    Public Enum NewOrderRecordSetTypes 'These are bit values
      nortAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum NewOrderFields
      nofAll = 0
      nofOrderNumber
      nofReasonForDespatch
      nofContactNumber
      nofAddressNumber
      nofGiftCardStatus
      nofPackToDonor
      nofMailing
      nofDateCreated
      nofDateFulfilled
      nofGiftFrom
      nofGiftTo
      nofGiftMessage
    End Enum

    Public Enum GiftCardStatusTypes
      gcstNone
      gcstBlank
      gcstWritten
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvEnclosures As Collection
    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "new_orders"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("order_number", CDBField.FieldTypes.cftLong)
          .Add("reason_for_despatch")
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("gift_card_status")
          .Add("pack_to_donor")
          .Add("mailing")
          .Add("date_created", CDBField.FieldTypes.cftDate)
          .Add("date_fulfilled", CDBField.FieldTypes.cftDate)
          .Add("gift_from", CDBField.FieldTypes.cftCharacter)
          .Add("gift_to", CDBField.FieldTypes.cftCharacter)
          .Add("gift_message", CDBField.FieldTypes.cftMemo)
        End With

        mvClassFields.Item(NewOrderFields.nofGiftFrom).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataGiftMessage)
        mvClassFields.Item(NewOrderFields.nofGiftTo).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataGiftMessage)
        mvClassFields.Item(NewOrderFields.nofGiftMessage).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataGiftMessage)

      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
      mvEnclosures = Nothing
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As NewOrderFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As NewOrderRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = NewOrderRecordSetTypes.nortAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "no")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pPaymentPlanNumber As Integer = 0, Optional ByVal pExcludeFulfilled As Boolean = True)
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      If pPaymentPlanNumber > 0 Then
        vWhereFields = New CDBFields
        With vWhereFields
          .Add("order_number", CDBField.FieldTypes.cftLong, pPaymentPlanNumber)
          If pExcludeFulfilled Then
            If pEnv.GetConfigOption("me_mandatory_new_orders") Then
              If pEnv.Connection.GetCount("enclosures", vWhereFields) = 0 Then
                'Possibly a "dummy" new orders record exists; prevent null date_fulfilled condition
                pExcludeFulfilled = False
              End If
            End If
          End If
          If pExcludeFulfilled Then .Add("date_fulfilled", CDBField.FieldTypes.cftDate, "")
        End With
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(NewOrderRecordSetTypes.nortAll) & " FROM new_orders no WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, NewOrderRecordSetTypes.nortAll)
        Else
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As NewOrderRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And NewOrderRecordSetTypes.nortAll) = NewOrderRecordSetTypes.nortAll Then
          .SetItem(NewOrderFields.nofOrderNumber, vFields)
          .SetItem(NewOrderFields.nofReasonForDespatch, vFields)
          .SetItem(NewOrderFields.nofContactNumber, vFields)
          .SetItem(NewOrderFields.nofAddressNumber, vFields)
          .SetItem(NewOrderFields.nofGiftCardStatus, vFields)
          .SetItem(NewOrderFields.nofPackToDonor, vFields)
          .SetItem(NewOrderFields.nofMailing, vFields)
          .SetItem(NewOrderFields.nofDateCreated, vFields)
          .SetItem(NewOrderFields.nofDateFulfilled, vFields)
          .SetOptionalItem(NewOrderFields.nofGiftFrom, vFields)
          .SetOptionalItem(NewOrderFields.nofGiftTo, vFields)
          .SetOptionalItem(NewOrderFields.nofGiftMessage, vFields)
        End If
      End With
    End Sub
    Public Sub InitForNewRecord(ByVal pEnv As CDBEnvironment, ByVal pPlanNumber As Integer, ByVal pReasonForDespatch As String, ByVal pContactNumber As Integer, ByRef pAddressNumber As Integer, ByVal pGiftCardStatus As GiftCardStatusTypes, ByVal pPackToDonor As Boolean, ByVal pMailingCode As String, ByVal pDateCreated As String, ByVal pDateFulfilled As String)
      mvEnv = pEnv
      InitClassFields()
      With mvClassFields
        'Always include the primary key attributes
        .Item(NewOrderFields.nofOrderNumber).IntegerValue = pPlanNumber
        .Item(NewOrderFields.nofReasonForDespatch).Value = pReasonForDespatch
        .Item(NewOrderFields.nofContactNumber).IntegerValue = pContactNumber
        .Item(NewOrderFields.nofAddressNumber).IntegerValue = pAddressNumber
        GiftCardStatus = pGiftCardStatus
        .Item(NewOrderFields.nofPackToDonor).Bool = pPackToDonor
        .Item(NewOrderFields.nofMailing).Value = pMailingCode
        .Item(NewOrderFields.nofDateCreated).Value = pDateCreated
        .Item(NewOrderFields.nofDateFulfilled).Value = pDateFulfilled
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(NewOrderFields.nofAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AddressNumber() As Integer
      Get
        Return mvClassFields.Item(NewOrderFields.nofAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields.Item(NewOrderFields.nofContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property DateCreated() As String
      Get
        DateCreated = mvClassFields.Item(NewOrderFields.nofDateCreated).Value
      End Get
    End Property

    Public ReadOnly Property DateFulfilled() As String
      Get
        DateFulfilled = mvClassFields.Item(NewOrderFields.nofDateFulfilled).Value
      End Get
    End Property

    Public Property GiftCardStatus() As GiftCardStatusTypes
      Get
        Select Case mvClassFields.Item(NewOrderFields.nofGiftCardStatus).Value
          Case "B"
            GiftCardStatus = GiftCardStatusTypes.gcstBlank
          Case "W"
            GiftCardStatus = GiftCardStatusTypes.gcstWritten
          Case Else
            GiftCardStatus = GiftCardStatusTypes.gcstNone
        End Select
      End Get
      Set(ByVal Value As GiftCardStatusTypes)
        Select Case Value
          Case GiftCardStatusTypes.gcstNone
            mvClassFields.Item(NewOrderFields.nofGiftCardStatus).Value = "N"
          Case GiftCardStatusTypes.gcstBlank
            mvClassFields.Item(NewOrderFields.nofGiftCardStatus).Value = "B"
          Case GiftCardStatusTypes.gcstWritten
            mvClassFields.Item(NewOrderFields.nofGiftCardStatus).Value = "W"
        End Select
      End Set
    End Property
    Public ReadOnly Property Mailing() As String
      Get
        Mailing = mvClassFields.Item(NewOrderFields.nofMailing).Value
      End Get
    End Property

    Public ReadOnly Property OrderNumber() As Integer
      Get
        Return mvClassFields.Item(NewOrderFields.nofOrderNumber).IntegerValue
      End Get
    End Property
    Public Property PackToDonor() As Boolean
      Get
        PackToDonor = mvClassFields.Item(NewOrderFields.nofPackToDonor).Bool
      End Get
      Set(ByVal Value As Boolean)
        mvClassFields.Item(NewOrderFields.nofPackToDonor).Bool = Value
      End Set
    End Property
    Public ReadOnly Property ReasonForDespatch() As String
      Get
        ReasonForDespatch = mvClassFields.Item(NewOrderFields.nofReasonForDespatch).Value
      End Get
    End Property

    Public ReadOnly Property GiftFrom() As String
      Get
        GiftFrom = mvClassFields.Item(NewOrderFields.nofGiftFrom).Value
      End Get
    End Property

    Public ReadOnly Property GiftTo() As String
      Get
        GiftTo = mvClassFields.Item(NewOrderFields.nofGiftTo).Value
      End Get
    End Property

    Public ReadOnly Property GiftMessage() As String
      Get
        GiftMessage = mvClassFields.Item(NewOrderFields.nofGiftMessage).Value
      End Get
    End Property

    Public ReadOnly Property Enclosures(Optional ByVal pExcludeFulfilled As Boolean = True) As Collection
      Get
        Dim vRS As CDBRecordSet
        Dim vEnclosure As Enclosure
        Dim vWhereFields As CDBFields

        If mvEnclosures Is Nothing Then
          vWhereFields = New CDBFields
          With vWhereFields
            .Add("order_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(NewOrderFields.nofOrderNumber).IntegerValue)
            If pExcludeFulfilled Then .Add("date_fulfilled", CDBField.FieldTypes.cftDate)
          End With

          mvEnclosures = New Collection

          vEnclosure = New Enclosure
          vEnclosure.Init(mvEnv)

          vRS = mvEnv.Connection.GetRecordSet("SELECT " & vEnclosure.GetRecordSetFields(Enclosure.EnclosureRecordSetTypes.encrtAll) & " FROM enclosures WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
          With vRS
            While .Fetch() = True
              vEnclosure = New Enclosure
              vEnclosure.InitFromRecordSet(mvEnv, vRS, Enclosure.EnclosureRecordSetTypes.encrtAll)
              mvEnclosures.Add(vEnclosure)
            End While
            .CloseRecordSet()
          End With
        End If
        Enclosures = mvEnclosures
      End Get
    End Property
    Public Sub Create(ByVal pParams As CDBParameters)
      'Used by Web Services and Smart Client only

      With mvClassFields
        .Item(NewOrderFields.nofOrderNumber).Value = pParams("OrderNumber").Value
        .Item(NewOrderFields.nofReasonForDespatch).Value = pParams("ReasonForDespatch").Value
        .Item(NewOrderFields.nofContactNumber).Value = pParams("PayerContactNumber").Value
        .Item(NewOrderFields.nofAddressNumber).Value = pParams("PayerAddressNumber").Value
        .Item(NewOrderFields.nofGiftCardStatus).Value = pParams("GiftCardStatus").Value
        .Item(NewOrderFields.nofPackToDonor).Value = pParams("PackToDonor").Value
        .Item(NewOrderFields.nofMailing).Value = pParams("TRD_Mailing").Value
        .Item(NewOrderFields.nofDateCreated).Value = pParams.OptionalValue("DateCreated", (TodaysDate()))
        .Item(NewOrderFields.nofDateFulfilled).Value = pParams.ParameterExists("DateFulfilled").Value
        .Item(NewOrderFields.nofGiftFrom).Value = pParams.ParameterExists("GiftFrom").Value
        .Item(NewOrderFields.nofGiftTo).Value = pParams.ParameterExists("GiftTo").Value
        .Item(NewOrderFields.nofGiftMessage).Value = pParams.ParameterExists("GiftMessage").Value
      End With
    End Sub
  End Class
End Namespace
