Imports CARE.Access.Interfaces
Imports System.Linq


Namespace Access.PostcodeValidation

  Public Class PostcoderAddress
    Implements Interfaces.IAddress

    Public Property Address As String Implements Interfaces.IAddress.Address

    Public Property AddressID As Integer Implements Interfaces.IAddress.AddressID

    Public Property AddressVerified As Boolean Implements Interfaces.IAddress.AddressVerified

    Public Property BuildingNumber As String Implements Interfaces.IAddress.BuildingNumber

    Public Property BuildingName As String Implements Interfaces.IAddress.BuildingName
    Public Property County As String Implements Interfaces.IAddress.County

    Public Property DPS As String Implements Interfaces.IAddress.DPS

    Public Property Easting As Integer Implements Interfaces.IAddress.Easting

    Public Property LEACode As String Implements Interfaces.IAddress.LEACode

    Public Property LEAName As String Implements Interfaces.IAddress.LEAName

    Public Property Northing As Integer Implements Interfaces.IAddress.Northing

    Public Property OrganisationName As String Implements Interfaces.IAddress.OrganisationName

    Public Property Postcode As String Implements Interfaces.IAddress.Postcode

    Public Property PostcoderID As String Implements Interfaces.IAddress.PostcoderID

    Public Property Town As String Implements Interfaces.IAddress.Town
    Public Property Iso3166Alpha3CountryCode As String Implements Interfaces.IAddress.Iso3166Alpha3CountryCode

    Public Shared Sub ConvertIAddressToDataTable(ByVal results As IEnumerable(Of IAddress), ByRef resultsDT As CDBDataTable)
      If results.Count > 0 Then
        If resultsDT.Rows.Count > 0 Then  'Required to handle Empty DataRow from GetPAFAddress
          resultsDT.Rows.Clear()
        End If
        For Each Address As IAddress In results
          Dim vRow As CDBDataRow
          vRow = resultsDT.AddRow
          Dim vAddressLine As String = Address.Address
          If vAddressLine.Length > 0 AndAlso Left(vAddressLine, 2).Contains(vbCrLf) Then
            vAddressLine = Replace(vAddressLine, vbCrLf, "", 1, 1)
          End If
          vAddressLine = Replace(vAddressLine, vbCrLf, ", ")
          vRow.Item("AddressLine") = vAddressLine
          vRow.Item("Town") = Address.Town
          vRow.Item("County") = Address.County
          vRow.Item("Postcode") = Address.Postcode
          vRow.Item("Address") = Address.Address
          vRow.Item("OrganisationName") = Address.OrganisationName
          If resultsDT.Columns.ContainsKey("BuildingNumber") Then
            vRow.Item("BuildingNumber") = Address.BuildingNumber
          End If
          vRow.Item("DeliveryPointSuffix") = Address.DPS
          If Address.Easting > 0 Then vRow.Item("Easting") = CStr(Address.Easting)
          If Address.Northing > 0 Then vRow.Item("Northing") = CStr(Address.Northing)
          vRow.Item("LeaCode") = Address.LEACode
          vRow.Item("LeaName") = Address.LEAName
          vRow.Item("PostcoderID") = Address.PostcoderID
        Next
      End If
    End Sub
    Public Shared Function MatchAddress(ByVal addressDetails As IAddress, ByVal searchString As IAddress) As Boolean
      MatchAddress = False
      Dim originalAddress As String = addressDetails.Address
      Dim addressLine As String
      addressLine = Replace(addressDetails.Address, vbCrLf, ",")
      If String.IsNullOrWhiteSpace(searchString.BuildingNumber) Then GetBuildingNumber(Split(searchString.Address, ","), searchString.BuildingNumber)
      GetBuildingNumber(Split(addressDetails.Address, ","), addressDetails.BuildingNumber)
      If Not String.IsNullOrWhiteSpace(searchString.Address) AndAlso String.Compare(searchString.Address, addressDetails.Address, StringComparison.InvariantCultureIgnoreCase) = 0 Then
        MatchAddress = True
      Else
        If InStr(searchString.Address.ToUpper, addressDetails.Address.ToUpper) >= 0 AndAlso (Not String.IsNullOrWhiteSpace(searchString.BuildingNumber) AndAlso String.Compare(searchString.BuildingNumber, addressDetails.BuildingNumber, StringComparison.InvariantCultureIgnoreCase) = 0) Then
          If String.Compare(Substring(searchString.Address, Len(searchString.BuildingNumber) + 1), Substring(addressDetails.Address, Len(addressDetails.BuildingNumber) + 1), StringComparison.InvariantCultureIgnoreCase) = 0 Then
            MatchAddress = True
          End If
        Else

          Dim searchBuildingNumber As String = searchString.BuildingNumber
          Dim addressBuildingNumber As String = addressDetails.BuildingNumber
          Dim searchAddress As String
          For vInt As Integer = 1 To 2
            If Not MatchAddress Then
              Dim CheckBuildingNumberExact As Boolean = (vInt = 1)
              If Not String.IsNullOrWhiteSpace(searchBuildingNumber) Then
                searchBuildingNumber = searchString.BuildingNumber.ToUpper
                If CheckBuildingNumberExact = False Then searchBuildingNumber = UCase(Replace(Replace(searchString.BuildingNumber, " ", ""), "-", ""))
              End If
              If Not String.IsNullOrWhiteSpace(addressDetails.BuildingNumber) Then
                addressBuildingNumber = addressDetails.BuildingNumber.ToUpper
                If CheckBuildingNumberExact = False Then addressBuildingNumber = UCase(Replace(Replace(addressDetails.BuildingNumber, " ", ""), "-", ""))
              End If
              searchAddress = UCase(Replace(Replace(searchString.Address, "-", ""), " ", ""))
              addressLine = UCase(Replace(Replace(addressDetails.Address, "-", ""), " ", ""))

              If searchString.Iso3166Alpha3CountryCode = "NLD" Then 'Use Building Number 
                If searchBuildingNumber = addressBuildingNumber Then
                  MatchAddress = True
                Else
                  If Not String.IsNullOrWhiteSpace(searchBuildingNumber) And String.IsNullOrWhiteSpace(addressBuildingNumber) Then
                    If Right(addressLine, Len(searchBuildingNumber)) = searchBuildingNumber Then
                      MatchAddress = True
                    End If
                  End If
                End If
              Else
                If Not String.IsNullOrWhiteSpace(addressBuildingNumber) Then
                  'We got a building number back from QAS
                  If Right(addressLine, Len(addressBuildingNumber)) = searchBuildingNumber Then
                    'Here we found the building number at the end of the given address so we should assume that this is a match
                    MatchAddress = True
                  ElseIf Not String.IsNullOrWhiteSpace(searchBuildingNumber) AndAlso (searchAddress.ToUpper.StartsWith("FLAT") OrElse searchString.BuildingNumber.Contains("-")) Then ' eg Flat11-11a, 
                    If Left(addressLine, Len(searchBuildingNumber)) = searchBuildingNumber Then  'returned address starts with search building number
                      MatchAddress = True
                    End If
                  End If
                Else
                  'No building number from QAS so compare the first line of the addresses
                  If searchAddress.Length > 0 Then
                    'BR13924: Special case for addresses having the word 'Flat' as first line e.g. importing 19 Redhall Close where the actual address is Flat vbCrLf 19 Redhall Close
                    If addressLine.ToUpper.StartsWith("FLAT") Then
                      Dim vPos As Integer = originalAddress.IndexOfAny(vbCrLf.ToCharArray)
                      'Try to get the second line
                      If vPos >= 0 Then addressLine = UCase(Replace(Replace(FirstLine(originalAddress.Substring(vPos + 2)), "-", ""), " ", ""))
                    End If
                    If searchAddress = addressLine Then
                      MatchAddress = True
                    End If
                  End If
                End If
              End If
            End If
          Next
        End If
      End If

      Return MatchAddress
    End Function
    Public Shared Function ExtractAddress(ByVal addressLine As String) As IAddress
      Dim vAddress(5) As String
      Dim vPos As Integer = 0
      Dim vIndex As Integer = 0
      Dim vStart As Integer = 1
      Do
        vPos = InStr(vStart, addressLine, ",")
        If vPos > 0 Then
          vAddress(vIndex) = Trim(Mid(addressLine, vStart, vPos - vStart))
        Else
          vAddress(vIndex) = Trim(Mid(addressLine, vStart))
        End If
        If Len(vAddress(vIndex)) > 0 Then vIndex = vIndex + 1
        vStart = vPos + 1
      Loop While vPos > 0 And vIndex < 6

      Dim addressDetails As New PostcoderAddress
      If Not String.IsNullOrWhiteSpace(vAddress(5)) Then
        addressDetails.County = vAddress(5)
        addressDetails.Town = vAddress(4)
        addressDetails.Address = vAddress(0).ToString + ", " + vAddress(1).ToString + ", " + vAddress(2).ToString + ", " + vAddress(3).ToString
      ElseIf Not String.IsNullOrWhiteSpace(vAddress(4)) Then
        addressDetails.County = vAddress(4)
        addressDetails.Town = vAddress(3)
        addressDetails.Address = vAddress(0).ToString + ", " + vAddress(1).ToString + ", " + vAddress(2).ToString
      ElseIf Not String.IsNullOrWhiteSpace(vAddress(3)) Then
        addressDetails.County = vAddress(3)
        addressDetails.Town = vAddress(2)
        addressDetails.Address = vAddress(0).ToString + ", " + vAddress(1).ToString
      ElseIf Not String.IsNullOrWhiteSpace(vAddress(2)) Then
        addressDetails.County = vAddress(2)
        addressDetails.Town = vAddress(1)
        addressDetails.Address = vAddress(0).ToString
      ElseIf Not String.IsNullOrWhiteSpace(vAddress(1)) Then
        addressDetails.Town = vAddress(1)
        addressDetails.Address = vAddress(0).ToString
      End If

      Return addressDetails
    End Function
    Public Shared Sub GetBuildingNumber(ByVal pAddress() As String, ByRef pBuilding As String, Optional ByRef pBuildingNumber As Boolean = False)
      Dim vchar As String
      vchar = Left(pAddress(0), 1)
      If vchar >= "0" And vchar <= "9" Then
        'We have a number at the start so use as building
        pBuilding = FirstWord(Replace(pAddress(0), ",", " "))
        pBuildingNumber = True
      Else
        'Use the whole first line as building
        pBuilding = pAddress(0)
        pBuildingNumber = False
      End If
    End Sub
  End Class

End Namespace

