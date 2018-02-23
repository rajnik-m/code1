

Namespace Access
  Public Class EventOrganiser

    Public Enum EventOrganiserRecordSetTypes 'These are bit values
      eortAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EventOrganiserFields
      eofAll = 0
      eofEventNumber
      eofOrganiser
      eofReference
      eofPriceToAttendees
      eofProduct
      eofRate
      eofOrderDate
      eofLiaisonDate
      eofInvoiceDate
      eofNotes
      eofAmendedBy
      eofAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvOrganiser As Organiser
    Private mvProduct As Product

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "event_organisers"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("event_number", CDBField.FieldTypes.cftInteger)
          .Add("organiser")
          .Add("reference")
          .Add("price_to_attendees", CDBField.FieldTypes.cftNumeric)
          .Add("product")
          .Add("rate")
          .Add("order_date", CDBField.FieldTypes.cftDate)
          .Add("liaison_date", CDBField.FieldTypes.cftDate)
          .Add("invoice_date", CDBField.FieldTypes.cftDate)
          .Add("notes", CDBField.FieldTypes.cftMemo)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With
        mvClassFields.Item(EventOrganiserFields.eofEventNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(EventOrganiserFields.eofEventNumber).PrefixRequired = True
        mvClassFields.Item(EventOrganiserFields.eofReference).SpecialColumn = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As EventOrganiserFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(EventOrganiserFields.eofAmendedOn).Value = TodaysDate()
      mvClassFields.Item(EventOrganiserFields.eofAmendedBy).Value = mvEnv.User.Logname
    End Sub
    Public ReadOnly Property OrganiserDetails() As Organiser
      Get
        If mvOrganiser Is Nothing Then
          mvOrganiser = New Organiser
          mvOrganiser.Init(mvEnv, Organiser)
        End If
        OrganiserDetails = mvOrganiser
      End Get
    End Property

    Public ReadOnly Property Product() As Product
      Get
        If mvProduct Is Nothing Then
          If ProductCode.Length > 0 Then
            mvProduct = New Product(mvEnv)
            mvProduct.Init(ProductCode)
          End If
        End If
        Product = mvProduct
      End Get
    End Property

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property DataTable() As CDBDataTable
      Get
        Dim vTable As New CDBDataTable
        Dim vRow As CDBDataRow

        vTable.AddColumnsFromList("EventNumber,Organiser,OrganiserDesc,OrganiserContactNumber,OrganiserContactName,OrganiserContactAddressNumber,OrganiserContactAddressLine")
        vTable.AddColumnsFromList("InvoiceContactNumber,InvoiceContactName,InvoiceContactAddressNumber,InvoiceContactAddressLine")
        vTable.AddColumnsFromList("OrganiserReference,PriceToAttendees,Product,ProductDesc,Rate,RateDesc,OrderDate,LiaisonDate,InvoiceDate,Notes,AmendedBy,AmendedOn")

        If EventNumber > 0 And Len(Organiser) > 0 Then
          vRow = vTable.AddRow
          With vRow
            .Item(1) = CStr(EventNumber)
            .Item(2) = Organiser
            .Item(3) = OrganiserDetails.OrganiserDesc
            If Not OrganiserDetails.ArrangementContact Is Nothing Then
              .Item(4) = CStr(OrganiserDetails.ArrangementContact.ContactNumber)
              .Item(5) = OrganiserDetails.ArrangementContact.Name
              .Item(6) = CStr(OrganiserDetails.ArrangementAddress.AddressNumber)
              .Item(7) = OrganiserDetails.ArrangementAddress.AddressLine
            End If
            If OrganiserDetails.InvoiceContactNumber > 0 Then
              .Item(8) = CStr(OrganiserDetails.InvoiceContact.ContactNumber)
              .Item(9) = OrganiserDetails.InvoiceContact.Name
              .Item(10) = CStr(OrganiserDetails.InvoiceAddress.AddressNumber)
              .Item(11) = OrganiserDetails.InvoiceAddress.AddressLine
            End If
            .Item(12) = Reference
            .Item(13) = PriceToAttendees
            .Item(14) = ProductCode
            .Item(15) = mvEnv.GetDescription("products", "product", ProductCode)
            .Item(16) = RateCode
            .Item(17) = mvEnv.GetDescription("rates", "rate", RateCode, New CDBFields(New CDBField("product", ProductCode)))
            .Item(18) = OrderDate
            .Item(19) = LiaisonDate
            .Item(20) = InvoiceDate
            .Item(21) = Notes
            .Item(22) = AmendedBy
            .Item(23) = AmendedOn
          End With
        End If
        DataTable = vTable
      End Get
    End Property

    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(EventOrganiserFields.eofAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(EventOrganiserFields.eofAmendedOn).Value
      End Get
    End Property

    Public Property EventNumber() As Integer
      Get
        EventNumber = mvClassFields.Item(EventOrganiserFields.eofEventNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(EventOrganiserFields.eofEventNumber).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property InvoiceDate() As String
      Get
        InvoiceDate = mvClassFields.Item(EventOrganiserFields.eofInvoiceDate).Value
      End Get
    End Property

    Public ReadOnly Property LiaisonDate() As String
      Get
        LiaisonDate = mvClassFields.Item(EventOrganiserFields.eofLiaisonDate).Value
      End Get
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(EventOrganiserFields.eofNotes).MultiLineValue
      End Get
    End Property

    Public ReadOnly Property OrderDate() As String
      Get
        OrderDate = mvClassFields.Item(EventOrganiserFields.eofOrderDate).Value
      End Get
    End Property

    Public ReadOnly Property Organiser() As String
      Get
        Organiser = mvClassFields.Item(EventOrganiserFields.eofOrganiser).Value
      End Get
    End Property

    Public ReadOnly Property PriceToAttendees() As String
      Get
        PriceToAttendees = mvClassFields.Item(EventOrganiserFields.eofPriceToAttendees).Value
      End Get
    End Property

    Public ReadOnly Property ProductCode() As String
      Get
        ProductCode = mvClassFields.Item(EventOrganiserFields.eofProduct).Value
      End Get
    End Property

    'UPGRADE_NOTE: Rate was upgraded to RateCode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public ReadOnly Property RateCode() As String
      Get
        RateCode = mvClassFields.Item(EventOrganiserFields.eofRate).Value
      End Get
    End Property

    Public ReadOnly Property Reference() As String
      Get
        Reference = mvClassFields.Item(EventOrganiserFields.eofReference).Value
      End Get
    End Property

     '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As EventOrganiserRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = EventOrganiserRecordSetTypes.eortAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "eo")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pEventNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pEventNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventOrganiserRecordSetTypes.eortAll) & " FROM event_organisers eo WHERE event_number = " & pEventNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventOrganiserRecordSetTypes.eortAll)
        Else
          InitClassFields()
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Friend Sub InitFromOrganiser(ByRef pOrganiser As EventOrganiser, ByRef pNewEvent As CDBEvent)
      With pOrganiser
        mvClassFields.Item(EventOrganiserFields.eofEventNumber).Value = CStr(pNewEvent.EventNumber)
        mvClassFields.Item(EventOrganiserFields.eofOrganiser).Value = .Organiser
        mvClassFields.Item(EventOrganiserFields.eofReference).Value = .Reference
        mvClassFields.Item(EventOrganiserFields.eofProduct).Value = .ProductCode
        mvClassFields.Item(EventOrganiserFields.eofRate).Value = .RateCode
        mvClassFields.Item(EventOrganiserFields.eofNotes).Value = .Notes
        mvClassFields.Item(EventOrganiserFields.eofPriceToAttendees).Value = .PriceToAttendees
      End With
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventOrganiserRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(EventOrganiserFields.eofEventNumber, vFields)
        .SetItem(EventOrganiserFields.eofOrganiser, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And EventOrganiserRecordSetTypes.eortAll) = EventOrganiserRecordSetTypes.eortAll Then
          .SetItem(EventOrganiserFields.eofReference, vFields)
          .SetItem(EventOrganiserFields.eofPriceToAttendees, vFields)
          .SetItem(EventOrganiserFields.eofProduct, vFields)
          .SetItem(EventOrganiserFields.eofRate, vFields)
          .SetItem(EventOrganiserFields.eofOrderDate, vFields)
          .SetItem(EventOrganiserFields.eofLiaisonDate, vFields)
          .SetItem(EventOrganiserFields.eofInvoiceDate, vFields)
          .SetItem(EventOrganiserFields.eofNotes, vFields)
          .SetItem(EventOrganiserFields.eofAmendedBy, vFields)
          .SetItem(EventOrganiserFields.eofAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub SetValuesFromEvent(ByVal pEvent As CDBEvent)
      If Not mvExisting Then
        mvClassFields.Item(EventOrganiserFields.eofEventNumber).Value = CStr(pEvent.EventNumber)
      End If
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      SetValid(EventOrganiserFields.eofAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
      'BR 10563: If Organiser has changed, reset to nothing to force reinitialisation when next queried
      If Not mvOrganiser Is Nothing Then
        If mvClassFields(EventOrganiserFields.eofOrganiser).Value <> mvOrganiser.OrganiserCode Then
          mvOrganiser = Nothing
        End If
      End If
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      Init(pEnv)
      mvClassFields.Item(EventOrganiserFields.eofEventNumber).Value = pParams("EventNumber").Value
      Update(pParams)
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      With mvClassFields
        If pParams.Exists("Organiser") Then .Item(EventOrganiserFields.eofOrganiser).Value = pParams("Organiser").Value
        If pParams.Exists("OrganiserReference") Then .Item(EventOrganiserFields.eofReference).Value = pParams("OrganiserReference").Value
        If pParams.Exists("PriceToAttendees") Then .Item(EventOrganiserFields.eofPriceToAttendees).Value = pParams("PriceToAttendees").Value
        If pParams.Exists("Product") Then .Item(EventOrganiserFields.eofProduct).Value = pParams("Product").Value
        If pParams.Exists("Rate") Then .Item(EventOrganiserFields.eofRate).Value = pParams("Rate").Value
        If pParams.Exists("OrderDate") Then .Item(EventOrganiserFields.eofOrderDate).Value = pParams("OrderDate").Value
        If pParams.Exists("LiaisonDate") Then .Item(EventOrganiserFields.eofLiaisonDate).Value = pParams("LiaisonDate").Value
        If pParams.Exists("InvoiceDate") Then .Item(EventOrganiserFields.eofInvoiceDate).Value = pParams("InvoiceDate").Value
        If pParams.Exists("Notes") Then .Item(EventOrganiserFields.eofNotes).Value = pParams("Notes").Value
      End With
    End Sub
  End Class
End Namespace
