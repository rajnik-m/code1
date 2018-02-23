Namespace Access

  Partial Public Class ContactAddress

    Private mvUsages As List(Of String)

    Public Enum ContactAddresssLinkTypes
      caltContact
      caltOrganisation
    End Enum

    Public Sub InitFromContactAndAddress(ByVal pEnv As CDBEnvironment, ByVal pLinkType As ContactAddresssLinkTypes, ByRef pContactNumber As Integer, ByRef pAddressNumber As Integer)
      mvEnv = pEnv
      SetAddressLinkType(pLinkType)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add((mvClassFields.Item(ContactAddressFields.ContactNumber).Name), pContactNumber)
      vWhereFields.Add((mvClassFields.Item(ContactAddressFields.AddressNumber).Name), pAddressNumber)
      MyBase.InitWithPrimaryKey(vWhereFields)
      If Not mvExisting Then
        mvClassFields.Item(ContactAddressFields.AddressLinkNumber).IntegerValue = mvEnv.GetControlNumber("AL")
        mvClassFields.Item(ContactAddressFields.ContactNumber).IntegerValue = pContactNumber
        mvClassFields.Item(ContactAddressFields.AddressNumber).IntegerValue = pAddressNumber
      End If
    End Sub

    Public Overloads Sub Create(ByVal pEnv As CDBEnvironment, ByVal pLinkType As ContactAddresssLinkTypes, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pHistorical As String, ByVal pValidFrom As String, ByVal pValidTo As String, Optional ByRef pAmendedBy As String = "", Optional ByRef pAmendedOn As String = "", Optional ByRef pAddressLinkNumber As Integer = 0)
      mvEnv = pEnv
      SetAddressLinkType(pLinkType)
      InitClassFields()
      mvClassFields(ContactAddressFields.ContactNumber).IntegerValue = pContactNumber
      mvClassFields(ContactAddressFields.AddressNumber).IntegerValue = pAddressNumber
      mvClassFields(ContactAddressFields.Historical).Value = pHistorical
      mvClassFields(ContactAddressFields.ValidFrom).Value = pValidFrom
      mvClassFields(ContactAddressFields.ValidTo).Value = pValidTo
      mvClassFields(ContactAddressFields.AmendedBy).Value = pAmendedBy
      mvClassFields(ContactAddressFields.AmendedOn).Value = pAmendedOn
      mvClassFields(ContactAddressFields.AddressLinkNumber).IntegerValue = pAddressLinkNumber
      'If the address is historical we force it to be historical. Otherwise we leave the value as passed in by the parameter
      If IsDate(pValidTo) Then
        If CDate(pValidTo) < Today Then
          mvClassFields(ContactAddressFields.Historical).Bool = True
        End If
      End If
    End Sub

    Private Sub SetAddressLinkType(ByVal pLinkType As ContactAddresssLinkTypes)
      Init()
      If pLinkType = ContactAddresssLinkTypes.caltOrganisation Then
        mvClassFields(ContactAddressFields.ContactNumber).SetName("organisation_number")
        mvClassFields.DatabaseTableName = "organisation_addresses"
      End If
    End Sub

    Public Function HistoricalChanged() As Boolean
      HistoricalChanged = mvHistoricalChanged
    End Function

    Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      mvHistoricalChanged = mvClassFields(ContactAddressFields.Historical).ValueChanged
      MyBase.Save(pAmendedBy, pAudit, pJournalNumber)
    End Sub

    Public ReadOnly Property Usages() As List(Of String)
      Get
        If mvUsages Is Nothing Then
          mvUsages = New List(Of String)
          Dim vRS As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT address_usage FROM contact_address_usages cau WHERE contact_number = " & ContactNumber & " AND address_number = " & AddressNumber)
          While vRS.Fetch
            mvUsages.Add(vRS.Fields(1).Value)
          End While
          vRS.CloseRecordSet()
        End If
        Return mvUsages
      End Get
    End Property

    Public Function ValidateChangeOfDates(Optional ByVal pNewValidFromDate As String = "", Optional ByVal pNewValidToDate As String = "") As Boolean
      'Check that change of dates has not invalidated any Positions
      'Used when updating an Organisation address
      Dim vWhereFields As New CDBFields
      Dim vValidFromChanged As Boolean
      Dim vValidToChanged As Boolean
      
      Dim vChangeValid As Boolean = True
      If mvExisting And (pNewValidFromDate.Length > 0 Or pNewValidToDate.Length > 0) Then
        If pNewValidFromDate.Length > 0 Then
          If ValidFrom.Length > 0 Then
            If CDate(pNewValidFromDate) > CDate(ValidFrom) Then vValidFromChanged = True
          Else
            vValidFromChanged = True
          End If
        End If
        If pNewValidToDate.Length > 0 Then
          If ValidTo.Length > 0 Then
            If CDate(ValidTo) > CDate(pNewValidToDate) Then vValidToChanged = True
          Else
            vValidToChanged = True
          End If
        End If

        If vValidFromChanged Or vValidToChanged Then
          With vWhereFields
            .Add("address_number", AddressNumber)
            .Add("organisation_number", ContactNumber)
            .Add("contact_number", ContactNumber, CDBField.FieldWhereOperators.fwoNotEqual)
            .Add("current", "Y")
            .Item("current").SpecialColumn = True
          End With
          Dim vSQL As String = mvEnv.Connection.WhereClause(vWhereFields)
          If vValidFromChanged Then
            vSQL = vSQL & " AND "
            If vValidToChanged Then vSQL = vSQL & "("
            vSQL = vSQL & "(started" & mvEnv.Connection.SQLLiteral("<", CDBField.FieldTypes.cftDate, pNewValidFromDate) & " OR started IS NULL)"
          End If
          If vValidToChanged Then
            vSQL = vSQL & If(vValidFromChanged, " OR ", " AND ")
            vSQL = vSQL & "(finished " & mvEnv.Connection.SQLLiteral(">", CDBField.FieldTypes.cftDate, pNewValidToDate) & " OR finished IS NULL)"
            If vValidFromChanged Then vSQL = vSQL & ")"
          End If
          If mvEnv.Connection.GetCount("contact_positions", Nothing, vSQL) > 0 Then vChangeValid = False
        End If
      End If
      Return vChangeValid
    End Function

  End Class

End Namespace
