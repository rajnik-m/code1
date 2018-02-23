Public Class AddressInformation

  Private mvAddress As String
  Private mvTown As String
  Private mvCounty As String
  Private mvPostCode As String
  Private mvAddressLine1 As String

  Private mvSubBuilding As String
  Private mvBuilding As String
  Private mvOrganisation As String
  Private mvThorofare As String
  Private mvDDLocality As String
  Private mvDLocality As String
  Private mvBuildingNumber As String
  Private mvBuildingName As String
  Private mvDeliveryPointSuffix As String

  Friend Sub ExtractData(ByRef pString As String, Optional ByVal pStartPos As Short = 1, Optional ByRef pSeparator As String = vbLf, Optional ByRef pExtraFields As Boolean = True)
    Dim vAdd2 As String
    Dim vAdd3 As String
    Dim vAdd4 As String
    Dim vPos As Integer

    mvAddress = Trim(Mid(pString, pStartPos, 35))
    vAdd2 = Trim(Mid(pString, 35 + pStartPos, 35))
    vAdd3 = Trim(Mid(pString, 70 + pStartPos, 35))
    vAdd4 = Trim(Mid(pString, 105 + pStartPos, 35))
    mvTown = Trim(Mid(pString, 140 + pStartPos, 35))
    mvCounty = Trim(Mid(pString, 175 + pStartPos, 35))
    mvPostCode = Trim(Mid(pString, 210 + pStartPos, 8))
    vPos = 220 + pStartPos
    mvSubBuilding = Trim(Mid(pString, vPos, 35))
    mvBuilding = Trim(Mid(pString, vPos + 35, 35))
    mvOrganisation = Trim(Mid(pString, vPos + 70, 35))
    mvThorofare = Trim(Mid(pString, vPos + 105, 35))
    mvDDLocality = Trim(Mid(pString, vPos + 140, 35))
    mvDLocality = Trim(Mid(pString, vPos + 175, 35))
    mvBuildingName = Trim(Mid(pString, vPos + 210, 35))
    mvBuildingNumber = Trim(Mid(pString, vPos + 245, 35))
    mvDeliveryPointSuffix = Trim(Mid(pString, vPos + 280, 5))

    If mvAddress = mvOrganisation Then
      mvAddressLine1 = vAdd2
    Else
      mvAddressLine1 = mvAddress
    End If
    If vAdd2.Length > 0 Then
      If mvAddress.Length > 0 Then mvAddress = mvAddress & pSeparator
        mvAddress = mvAddress & vAdd2
      End If
      If vAdd3.Length > 0 Then
      If mvAddress.Length > 0 Then mvAddress = mvAddress & pSeparator
        mvAddress = mvAddress & vAdd3
      End If
      If vAdd4.Length > 0 Then
      If mvAddress.Length > 0 Then mvAddress = mvAddress & pSeparator
        mvAddress = mvAddress & vAdd4
      End If
  End Sub

  Friend Sub PrintDebug()
    Debug.Print(Replace(mvAddress, vbLf, ",") & "," & mvTown & "," & mvCounty & "," & mvPostCode)
  End Sub

  Friend Function ValidBuilding(ByRef pBuildingNumber As Boolean, ByRef pBuilding As String) As Boolean

    'Debug.Print
    'Debug.Print "ValidBuilding No: " & pBuildingNumber & ", Building: " & pBuilding
    'Debug.Print "Number:           " & mvBuildingNumber
    'Debug.Print "Name:             " & mvBuildingName
    'Debug.Print "Building:         " & mvBuilding
    'Debug.Print "Building String:  " & BuildingString
    'Debug.Print "Address1:         " & mvAddressLine1

    If pBuildingNumber Then
      If StrComp(pBuilding, mvBuildingNumber, CompareMethod.Text) = 0 Or StrComp(pBuilding, mvSubBuilding, CompareMethod.Text) = 0 Then
        If StrComp(Left(mvAddressLine1, Len(pBuilding)), pBuilding, CompareMethod.Text) = 0 Then
          ValidBuilding = True
        Else
          'Stop
        End If
      Else
        'The number does not match
      End If
    Else
      If pBuilding = mvBuildingName Then
        ValidBuilding = True
      ElseIf pBuilding = mvBuilding Then
        ValidBuilding = True
      ElseIf pBuilding = BuildingString Then
        ValidBuilding = True
      ElseIf pBuilding = mvAddressLine1 Then
        ValidBuilding = True
      End If
    End If
  End Function

  Friend Function AddressMatch(ByRef pBuilding As String, ByRef pAddressLine As String) As Boolean
    Dim vLen As Integer

    If StrComp(pAddressLine, mvAddressLine1, CompareMethod.Text) = 0 Then
      AddressMatch = True
    Else
      'Just look at the first word after the building because of comparing Road and Rd. etc..
      vLen = Len(pBuilding) + 1
      If StrComp(FirstWord(Trim(Mid(pAddressLine, vLen))), FirstWord(Trim(Mid(mvAddressLine1, vLen))), CompareMethod.Text) = 0 Then AddressMatch = True
    End If
  End Function

  Public ReadOnly Property ThorofareString() As String
    Get
      Dim vResult As String

      vResult = mvThorofare
      If Len(vResult) > 0 Then
        If Len(mvDDLocality) > 0 Then vResult = vResult & ", " & mvDDLocality
        If Len(mvDLocality) > 0 Then vResult = vResult & ", " & mvDLocality
      Else
        vResult = mvDDLocality
        If Len(vResult) > 0 Then
          If Len(mvDLocality) > 0 Then vResult = vResult & ", " & mvDLocality
        Else
          vResult = mvDLocality
        End If
      End If
      vResult = vResult & ", " & mvTown
      If mvCounty.Length > 0 And InStr(mvAddress, mvCounty) > 0 Then vResult = vResult & ", " & mvCounty
      ThorofareString = vResult
    End Get
  End Property

  Public ReadOnly Property BuildingString() As String
    Get
      Dim vResult As String

      vResult = mvOrganisation
      If Len(mvSubBuilding) > 0 Then
        If Len(vResult) > 0 Then
          vResult = vResult & ", " & mvSubBuilding
        Else
          vResult = mvSubBuilding
        End If
      End If
      If mvBuilding.Length > 0 Then
        If Len(vResult) > 0 Then
          vResult = vResult & ", " & mvBuilding
        Else
          vResult = mvBuilding
        End If
      End If
      BuildingString = vResult
    End Get
  End Property

  Public ReadOnly Property Thorofare() As String
    Get
      Thorofare = mvThorofare
    End Get
  End Property
  Public ReadOnly Property DependantLocality() As String
    Get
      DependantLocality = mvDLocality
    End Get
  End Property
  Public ReadOnly Property DoubleDependantLocality() As String
    Get
      DoubleDependantLocality = mvDDLocality
    End Get
  End Property
  Public ReadOnly Property SubBuilding() As String
    Get
      SubBuilding = mvSubBuilding
    End Get
  End Property
  Public ReadOnly Property Building() As String
    Get
      Building = mvBuilding
    End Get
  End Property
  Public ReadOnly Property Organisation() As String
    Get
      Organisation = mvOrganisation
    End Get
  End Property
  Public ReadOnly Property BuildingNumber() As String
    Get
      BuildingNumber = mvBuildingNumber
    End Get
  End Property
  Public ReadOnly Property BuildingName() As String
    Get
      BuildingName = mvBuildingName
    End Get
  End Property
  Public ReadOnly Property DeliveryPointSuffix() As String
    Get
      DeliveryPointSuffix = mvDeliveryPointSuffix
    End Get
  End Property
End Class