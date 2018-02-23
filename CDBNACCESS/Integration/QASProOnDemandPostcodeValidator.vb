Imports CDBNServices
Imports CDBNServices.QASProOnDemand
Imports CARE.Access.Interfaces
Imports CARE.Config

Imports System.Linq


Namespace Access.PostcodeValidation

  Public Class QASProOnDemandPostcodeValidator
    Implements IPostcodeValidator

    'Required for Properties
    Private mvCheckAddress As Boolean = False

    'Module Variables
    Private mvQASDeliveryPointSuffix As String
    Private mvQASGridReferences As String
    Private mvQASLEAData As String
    Private Shared mvNewInstance As QASProOnDemandPostcodeValidator
    Private mvAuthenticationUsername As String
    Private mvAuthenticationPassword As String

    Dim mvQAS As QASProOnDemandInterface
    Dim mvCanSearchOk As Boolean


    Private Sub New(ByVal uri As Uri, ByVal qasDeliveryPointSuffix As String, ByVal qasGridReferences As String, ByVal qasLEAData As String, ByVal iso3DefaultCountryCode As String)
      If uri IsNot Nothing Then

        mvQAS = New QASProOnDemandInterface
        Dim vQAOK As New QASearchOk
        Dim qasCanSearchValues As New QACanSearch With {.Country = iso3DefaultCountryCode,
                                                        .Engine = GetEngine(),
                                                        .Layout = GetLayout()}

        If NfpConfigrationManager.QAAuthenticationValues IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(NfpConfigrationManager.QAAuthenticationValues.UsernameValue) AndAlso Not String.IsNullOrWhiteSpace(NfpConfigrationManager.QAAuthenticationValues.PasswordValue) Then
          AuthenticationUsername = NfpConfigrationManager.QAAuthenticationValues.UsernameValue
          AuthenticationPassword = NfpConfigrationManager.QAAuthenticationValues.PasswordValue
        Else
          Throw New ArgumentException("QAS ProOnDemand Username/Password values have not been set-up in the Web.config file.")
        End If

        mvQAS.WS.DoCanSearch(QueryHeader, qasCanSearchValues, vQAOK)
        If Not vQAOK.IsOk Then
          Throw New Exception(vQAOK.ErrorMessage)
        Else
          mvCanSearchOk = True
        End If
        Me.QASDeliveryPointSuffix = qasDeliveryPointSuffix
        Me.QASGridReferences = qasGridReferences
        Me.QASLEAData = qasLEAData
      Else
        Throw New ArgumentException("Pro On Demand URL has not been provided.", "uri")
      End If
    End Sub
    
    Public Function GetLicencedCountries() As String
      Dim datamapping(5) As QADataSet
      Dim country As QAGetData = Nothing
      Dim countries As String = String.Empty
      Dim information As QAInformation = mvQAS.WS.DoGetData(QueryHeader, country, datamapping)
      If datamapping IsNot Nothing Then
        If datamapping.Count > 0 Then
          For Each licencedCountry As QADataSet In datamapping
            If String.IsNullOrWhiteSpace(countries) Then
              countries = licencedCountry.ID
            Else
              countries = countries + "|" + licencedCountry.ID.ToString
            End If
          Next
        End If
      End If
      Return countries
    End Function
    Private Property QASDeliveryPointSuffix() As String
      Get
        Return mvQASDeliveryPointSuffix
      End Get
      Set(value As String)
        mvQASDeliveryPointSuffix = value
      End Set
    End Property
    Private Property QASLEAData() As String
      Get
        Return mvQASLEAData
      End Get
      Set(value As String)
        mvQASLEAData = value
      End Set
    End Property
    Private Property QASGridReferences() As String
      Get
        Return mvQASGridReferences
      End Get
      Set(value As String)
        mvQASGridReferences = value
      End Set
    End Property
    Private Property AuthenticationUsername() As String
      Get
        Return mvAuthenticationUsername
      End Get
      Set(value As String)
        mvAuthenticationUsername = value
      End Set
    End Property
    Private Property AuthenticationPassword() As String
      Get
        Return mvAuthenticationPassword
      End Get
      Set(value As String)
        mvAuthenticationPassword = value
      End Set
    End Property
    Private ReadOnly Property AuthenticationValues() As QAAuthentication
      Get
        Return New QAAuthentication With {.Username = AuthenticationUsername,
                                          .Password = AuthenticationPassword}
      End Get
    End Property
    Private ReadOnly Property QueryHeader As QAQueryHeader
      Get
        Return New QAQueryHeader With {.QAAuthentication = AuthenticationValues}
      End Get
    End Property
    Shared Function GetInstance(ByVal uri As Uri, ByVal QASDeliveryPointSuffix As String, ByVal QASGridReferences As String, ByVal QASLEAData As String, ByVal iso3DefaultCountryCode As String) As QASProOnDemandPostcodeValidator
      If mvNewInstance Is Nothing Then
        mvNewInstance = New QASProOnDemandPostcodeValidator(uri, QASDeliveryPointSuffix, QASGridReferences, QASLEAData, iso3DefaultCountryCode)
      End If
      Return mvNewInstance
    End Function

    Public Function GetAddresses(ByVal addresses As IEnumerable(Of IAddress), ByRef qaVerifyLevel As Boolean) As IEnumerable(Of IAddress) Implements IPostcodeValidator.GetAddresses
      If addresses Is Nothing OrElse
        addresses.Count < 1 Then
        Throw New ArgumentException("At least one address must be passed", "addresses")
      End If

      Dim results As New List(Of IAddress)

      If addresses.Count = 1 AndAlso Not String.IsNullOrWhiteSpace(addresses(0).PostcoderID) Then
        results = GetAddressByID(addresses(0).PostcoderID)    'Moniker Search
      Else
        If Not String.IsNullOrWhiteSpace(addresses(0).BuildingNumber) Then
          Dim addressSearch As IAddress = addresses(0)
          Return PostcodeAddress(addressSearch)               'Building Search
        Else
          results = PostcodeSearch(addresses, qaVerifyLevel)                   'Postcode Search
        End If
      End If

      Return results

    End Function
    Public Function PostcodeSearch(ByVal addresses As IEnumerable(Of IAddress), ByRef qaVerifyLevel As Boolean) As List(Of IAddress)
      Dim results As New List(Of IAddress)
      Dim qasSearchValues As New QASearch With {.Country = addresses(0).Iso3166Alpha3CountryCode,
                                              .Engine = GetEngine(),
                                              .Layout = GetLayout()}
      For Each Address As IAddress In addresses
        qasSearchValues.Search = Address.Postcode
        If Not String.IsNullOrWhiteSpace(qasSearchValues.Search) Then
          Dim searchResult As New QASearchResult
          Dim information As QAInformation = mvQAS.WS.DoSearch(QueryHeader, qasSearchValues, searchResult)
          If searchResult.QAPicklist IsNot Nothing Then
            If searchResult.QAPicklist.Total = "0" Then HandleNoData(searchResult)
            If searchResult.VerifyLevel = VerifyLevelType.Verified Then qaVerifyLevel = True
            If searchResult.QAPicklist.PicklistEntry.Length > 1 Then
              For Each vItem As PicklistEntryType In searchResult.QAPicklist.PicklistEntry
                Dim picklist As IAddress
                picklist = PostcoderAddress.ExtractAddress(vItem.Picklist)
                results.Add(New PostcoderAddress With {.PostcoderID = vItem.Moniker,
                                                       .Address = picklist.Address,
                                                       .Town = picklist.Town,
                                                       .County = picklist.County,
                                                       .Postcode = vItem.Postcode,
                                                       .AddressVerified = qaVerifyLevel})
              Next
            ElseIf searchResult.QAPicklist.PicklistEntry.Length = 1 AndAlso searchResult.QAPicklist.Total <> "0" Then
              Dim vItem As PicklistEntryType = searchResult.QAPicklist.PicklistEntry(0)
              results = GetAddressByID(vItem.Moniker)
            End If
          End If
        End If
      Next

      Return results

    End Function
    Public Function GetAddressByID(ByVal postcoderID As String) As List(Of IAddress)
      Dim results As New List(Of IAddress)
      If Not String.IsNullOrWhiteSpace(postcoderID) Then
        Dim getAddress As New QAGetAddress
        getAddress.Moniker = postcoderID
        getAddress.Layout = "CSG" 'This is a bespoke layout for DoGetAddress, our code will not work if the Layout is changed see CreateAddressLine for Layout Details
        Dim finalAddress As New CDBNServices.QASProOnDemand.Address
        Dim information As QAInformation = mvQAS.WS.DoGetAddress(QueryHeader, getAddress, finalAddress)
        If finalAddress IsNot Nothing Then
          Dim addressLine As String = CreateAddressLine(finalAddress)
          results.Add(New PostcoderAddress With {.Address = addressLine,
                                                 .OrganisationName = finalAddress.QAAddress.AddressLine(0).Line,
                                                 .BuildingName = finalAddress.QAAddress.AddressLine(5).Line,
                                                 .BuildingNumber = finalAddress.QAAddress.AddressLine(6).Line,
                                                 .Town = finalAddress.QAAddress.AddressLine(11).Line,
                                                 .County = finalAddress.QAAddress.AddressLine(12).Line,
                                                 .Postcode = finalAddress.QAAddress.AddressLine(13).Line
                                                 })

          If QASDeliveryPointSuffix = "Y" Then  'Only Set PAF Additional Data if DPS config set
            results(0).DPS = finalAddress.QAAddress.AddressLine(14).Line
          End If
          If QASGridReferences = "Y" Then
            If Not String.IsNullOrWhiteSpace(finalAddress.QAAddress.AddressLine(15).Line) Then results(0).Easting = CInt(finalAddress.QAAddress.AddressLine(15).Line)
            If Not String.IsNullOrWhiteSpace(finalAddress.QAAddress.AddressLine(16).Line) Then results(0).Northing = CInt(finalAddress.QAAddress.AddressLine(16).Line)
          End If
          If QASLEAData = "Y" Then
            results(0).LEACode = finalAddress.QAAddress.AddressLine(17).Line
            results(0).LEAName = finalAddress.QAAddress.AddressLine(18).Line
          End If

        End If
      Else
        Throw New ArgumentException("Unable to verify selected Address without PostcoderID", "PostcoderID")
      End If

      Return results

    End Function
    Public Function PostcodeAddress(ByVal address As IAddress) As IEnumerable(Of Interfaces.IAddress) Implements IPostcodeValidator.PostcodeAddress
      If address Is Nothing Then
        Throw New ArgumentException("At least one address must be passed", "addresses")
      End If

      Dim results As New List(Of IAddress)
      Dim qasSearchValues As New QASearch With {.Country = address.Iso3166Alpha3CountryCode,
                                              .Engine = GetEngine(),
                                              .Layout = GetLayout()}
      Dim searchString As String
      Replace(address.Address, vbCrLf, "")
      searchString = address.Address + "|" + address.Town + "|" + address.County + "|" + address.Postcode
      If Not String.IsNullOrWhiteSpace(address.BuildingNumber) Then 'Building Number search
        searchString = address.BuildingNumber + "|" + searchString
      End If
      qasSearchValues.Search = searchString
      If Not String.IsNullOrWhiteSpace(qasSearchValues.Search) Then
        Dim searchResult As New QASearchResult
        Dim information As QAInformation = mvQAS.WS.DoSearch(QueryHeader, qasSearchValues, searchResult)
        If searchResult.QAPicklist IsNot Nothing Then
          If searchResult.QAPicklist.PicklistEntry.Length > 1 Then
            If searchResult.QAPicklist.Total = "0" Then HandleNoData(searchResult)
            For Each vItem As PicklistEntryType In searchResult.QAPicklist.PicklistEntry
              Dim picklist As IAddress
              picklist = PostcoderAddress.ExtractAddress(vItem.Picklist)
              results.Add(New PostcoderAddress With {.PostcoderID = vItem.Moniker,
                                                     .Address = picklist.Address,
                                                     .Town = picklist.Town,
                                                     .County = picklist.County,
                                                     .Postcode = vItem.Postcode})
            Next
          ElseIf searchResult.QAPicklist.PicklistEntry.Length = 1 AndAlso searchResult.QAPicklist.Total <> "0" Then
            Dim vItem As PicklistEntryType = searchResult.QAPicklist.PicklistEntry(0)
            results = GetAddressByID(vItem.Moniker)
          End If
        End If
      End If

      Return results

    End Function
    Public Function ValidateBuilding(ByRef buildingAddress As PostcoderAddress) As Postcoder.ValidatePostcodeStatuses Implements IPostcodeValidator.ValidateBuilding
      Dim searchBuildingAddress As IAddress = buildingAddress
      Dim results As IEnumerable(Of IAddress)
      Dim postcodeStatus As Postcoder.ValidatePostcodeStatuses = Postcoder.ValidatePostcodeStatuses.vpsNone
      results = PostcodeAddress(searchBuildingAddress)
      If results.Count > 0 Then
        searchBuildingAddress.Address = Replace(searchBuildingAddress.Address, "|", ",")
        For Each resultsAddress As PostcoderAddress In results
          If PostcoderAddress.MatchAddress(resultsAddress, searchBuildingAddress) Then
            postcodeStatus = Postcoder.ValidatePostcodeStatuses.vpsBuildingValidated
            buildingAddress = resultsAddress
            Exit For
          End If
        Next
        If postcodeStatus = Postcoder.ValidatePostcodeStatuses.vpsBuildingValidated Then
          If Not String.IsNullOrWhiteSpace(buildingAddress.PostcoderID) AndAlso results.Count > 1 Then
            Dim finalAddress As New List(Of IAddress)
            finalAddress = GetAddressByID(buildingAddress.PostcoderID)
            buildingAddress.Address = finalAddress(0).Address
            buildingAddress.OrganisationName = finalAddress(0).OrganisationName
            buildingAddress.BuildingName = finalAddress(0).BuildingName
            buildingAddress.BuildingNumber = finalAddress(0).BuildingNumber
            buildingAddress.Town = finalAddress(0).Town
            buildingAddress.County = finalAddress(0).County
            buildingAddress.Postcode = finalAddress(0).Postcode
            If QASDeliveryPointSuffix = "Y" Then  'Only Set PAF Additional Data if DPS config set
              buildingAddress.DPS = finalAddress(0).DPS
            End If
            If QASGridReferences = "Y" Then
              If Not String.IsNullOrWhiteSpace(CStr(finalAddress(0).Easting)) Then buildingAddress.Easting = CInt(finalAddress(0).Easting)
              If Not String.IsNullOrWhiteSpace(CStr(finalAddress(0).Northing)) Then buildingAddress.Northing = CInt(finalAddress(0).Northing)
            End If
            If QASLEAData = "Y" Then
              buildingAddress.LEACode = finalAddress(0).LEACode
              buildingAddress.LEAName = finalAddress(0).LEAName
            End If
          End If
          If String.IsNullOrWhiteSpace(searchBuildingAddress.Postcode) AndAlso Not String.IsNullOrWhiteSpace(buildingAddress.Postcode) Then
            'No Search Postcode so Address was postcoded
            postcodeStatus = Postcoder.ValidatePostcodeStatuses.vpsAddressPostcoded
          ElseIf Not String.IsNullOrWhiteSpace(searchBuildingAddress.Postcode) Then
            If String.Compare(searchBuildingAddress.Postcode, buildingAddress.Postcode, StringComparison.InvariantCultureIgnoreCase) <> 0 Then
              postcodeStatus = Postcoder.ValidatePostcodeStatuses.vpsAddressRePostcoded
            End If
          End If
          buildingAddress.AddressVerified = True
        Else
          postcodeStatus = Postcoder.ValidatePostcodeStatuses.vpsBuildingNotValidated
        End If
      Else
        postcodeStatus = Postcoder.ValidatePostcodeStatuses.vpsBuildingNotValidated
      End If

      Return postcodeStatus
    End Function
    ''' <summary>
    ''' This is the CSG Layout specification for QAAddress.AddressLine
    ''' Address line 0-	address Line 1 as output by QADefault
    ''' Address line 1-	address Line 2 as output by QADefault
    ''' Address line 2-	Organisation Name-	O11
    ''' Address line 3-	Whole PO Box-	B11
    ''' Address line 4-	British Forces Post Office-	B12
    ''' Address line 5-	Sub-Building Name-	P22
    ''' Address line 6-	Sub-Building Number-	P21
    ''' Address line 7-	Building Name-	P12
    ''' Address line 8-	Whole Building Number-	P11
    ''' Address line 9-	Whole Dependent thoroughfare-	S21
    ''' Address line 10-Whole thoroughfare-	S11
    ''' Address line 11-Double dependent locality-	L41
    ''' Address line 12-Dependent Locality-	L31
    ''' Address line 13-Town-	L21
    ''' Address line 14-County-	L11
    ''' Address Line 15-Postcode-	C11
    ''' Address Line 16-Delivery Point Suffix-	A11
    ''' Address Line 17-Code-Point Easting
    ''' Address Line 18-Code-Point Northing
    ''' Address Line 19-Local Education Authority Code
    ''' Address Line 20-Local Education Authority Name
    ''' </summary>
    <CLSCompliant(False)>
    Public Function CreateAddressLine(ByVal finalAddress As CDBNServices.QASProOnDemand.Address) As String
      Dim addressLine As String = String.Empty
      Dim index As Integer = 0

      Do
        If Not String.IsNullOrWhiteSpace(finalAddress.QAAddress.AddressLine(index).Line) Then
          If String.IsNullOrWhiteSpace(addressLine) Then
            addressLine = finalAddress.QAAddress.AddressLine(index).Line
          Else
            Select Case index
              Case 1, 2, 3, 5, 6
                If Not String.IsNullOrWhiteSpace(addressLine) Then addressLine = addressLine + vbCrLf
              Case 7
                If Not String.IsNullOrWhiteSpace(finalAddress.QAAddress.AddressLine(6).Line) Then ' If there was a BuildingNumber don't add comma
                  addressLine = addressLine + " "
                End If
              Case 8
                If Not String.IsNullOrWhiteSpace(finalAddress.QAAddress.AddressLine(7).Line) Then
                  addressLine = addressLine + ","   'there was a Dependent throughfare - add coma before throughfare
                ElseIf (String.IsNullOrWhiteSpace(finalAddress.QAAddress.AddressLine(7).Line) AndAlso String.IsNullOrWhiteSpace(finalAddress.QAAddress.AddressLine(6).Line)) Then
                  addressLine = addressLine + vbCrLf
                Else
                  addressLine = addressLine + " "
                End If
              Case 9
                If Not String.IsNullOrWhiteSpace(addressLine) Then addressLine = addressLine + vbCrLf
              Case 10
                If (Not String.IsNullOrWhiteSpace(finalAddress.QAAddress.AddressLine(9).Line)) Then
                  addressLine = addressLine + ","   'there was a Double Dependent Locality 
                Else
                  addressLine = addressLine + vbCrLf
                End If
            End Select

            If Not addressLine.EndsWith(vbCrLf) AndAlso Not addressLine.EndsWith(",") AndAlso Not addressLine.EndsWith(" ") Then addressLine = addressLine + ","
            addressLine = addressLine + finalAddress.QAAddress.AddressLine(index).Line
          End If
        End If

        index = index + 1
        'No need to loop after 12 as those Address Elements are not part of the Address Line and are handled in the calling rountine
      Loop While index < 11

      Return addressLine
    End Function
    <CLSCompliant(False)>
    Public Sub HandleNoData(ByVal searchResult As QASearchResult)
      Dim errorMessage As String = Nothing
      Dim vItem As PicklistEntryType = searchResult.QAPicklist.PicklistEntry(0)
      If vItem.WarnInformation Then
        If InStr(vItem.Picklist, "too many matches") > 0 Then
          errorMessage = vItem.Picklist + "- Maximum Results returned from the ProOnDemand Console is 100. Please refine search criteria."
        Else
          errorMessage = vItem.Picklist
        End If
        Throw New ArgumentException(errorMessage)
      End If
    End Sub
    Private Function GetEngine() As EngineType
      Dim vEngineType As New EngineType
      vEngineType.Flatten = True
      vEngineType.FlattenSpecified = True
      vEngineType.Value = EngineEnumType.Singleline
      Return vEngineType
    End Function

    Private Function GetLayout() As String
      Return "QADefault"
    End Function

    Public ReadOnly Property CanBulkAddressSearch As Boolean Implements IPostcodeValidator.CanBulkAddressSearch
      Get
        Return True
      End Get
    End Property

    Public Property CheckAddress As Boolean Implements IPostcodeValidator.CheckAddress
      Get
        Return mvCheckAddress
      End Get
      Set(value As Boolean)
        mvCheckAddress = value
      End Set
    End Property

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          ' TODO: dispose managed state (managed objects).
        End If

        ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
        ' TODO: set large fields to null.
      End If
      Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub
#End Region

  End Class
End Namespace

