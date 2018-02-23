Imports System.Linq
Imports CARE.Access.PostcodeValidation
Imports Advanced.LanguageExtensions
Namespace Access

  Partial Public Class Address

    Public Event DedupError(ByVal pText As String)
    Public Event DedupLogMessage(ByVal pText As String)

    Private mvPositionsAtAddress As Collection
    Private mvContactsAtAddress As Contacts
    Private mvMergeInfo As List(Of AddressMergeInfo)
    Private mvMergeInfoValid As Boolean
    Private mvAddressLinesLength As Integer
    Private mvAddressBlank As Boolean = False

    Private Class AddressMergeInfo
      Public TableName As String = ""
      Public ContactAttr As String = ""
      Public AddressAttr As String = ""
    End Class

    <Flags()> _
    Public Enum AddressRecordSetTypes As Integer 'These are bit values
      artAll = &HFFFFS
      'ADD additional recordset types here
      artNumber = 1
      artDetails = 2
      artSortcode = 4
      artPaf = 8
      artCountrySortCode = 16
      artAddressLines = 32
      artDPS = 64
    End Enum

    Public Overloads Function GetRecordSetFields(ByVal pRSType As AddressRecordSetTypes) As String
      CheckClassFields()
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = AddressRecordSetTypes.artAll Then
        vFields = mvClassFields.FieldNames(mvEnv, "a")
        vFields = vFields & ",uk,country_desc"
      Else
        If (pRSType And AddressRecordSetTypes.artNumber) > 0 Then vFields = "a.address_number,"
        If (pRSType And AddressRecordSetTypes.artDetails) > 0 Then vFields = vFields & "address,house_name,town,county,postcode,a.branch,a.country,address_type," & If(mvClassFields(AddressFields.BuildingNumber).InDatabase, "building_number,", "")
        If (pRSType And AddressRecordSetTypes.artCountrySortCode) > 0 Then
          vFields = vFields & "sortcode,uk,country_desc,"
        ElseIf (pRSType And AddressRecordSetTypes.artSortcode) > 0 Then
          vFields = vFields & "sortcode,"
        End If
        If (pRSType And AddressRecordSetTypes.artPaf) > 0 Then vFields = vFields & "mosaic_code,paf,a.amended_by,a.amended_on,"
        If (pRSType And AddressRecordSetTypes.artAddressLines) > 0 Then vFields = vFields & "address_line1,address_line2,address_line3,address_line4,"
        If (pRSType And AddressRecordSetTypes.artDPS) > 0 Then vFields = vFields & "delivery_point_suffix,"
      End If
      If Right(vFields, 1) = "," Then vFields = Left(vFields, Len(vFields) - 1)
      Return vFields
    End Function

    Public Function GetRecordSetFieldsDetailCountrySortCode() As String
      CheckClassFields()
      Return "address,house_name,town,county,postcode,a.branch,a.country,address_type," & _
             If(mvClassFields(AddressFields.BuildingNumber).InDatabase, "building_number,", "") & _
             "sortcode,uk,country_desc"
    End Function

    Protected Overrides Sub SetValid()
      MyBase.SetValid()
      CheckValidity()
    End Sub

    Public Overloads Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As AddressRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(AddressFields.AddressNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And AddressRecordSetTypes.artDetails) > 0 Then
          .SetItem(AddressFields.AddressType, vFields)
          .SetItem(AddressFields.HouseName, vFields)
          .SetItem(AddressFields.Address, vFields)
          .SetItem(AddressFields.Town, vFields)
          .SetItem(AddressFields.County, vFields)
          .SetItem(AddressFields.Country, vFields)
          .SetItem(AddressFields.Postcode, vFields)
          .SetItem(AddressFields.Branch, vFields)
          If vFields.Exists("building_number") Then .SetOptionalItem(AddressFields.BuildingNumber, vFields)
        End If
        If (pRSType And AddressRecordSetTypes.artCountrySortCode) > 0 Then
          mvUK = vFields("uk").Bool
          mvCountryDescription = vFields("country_desc").Value
          GetCountryDescription() ' BR17231
          mvCountryValid = True
          .SetItem(AddressFields.Sortcode, vFields)
        ElseIf (pRSType And AddressRecordSetTypes.artSortcode) > 0 Then
          .SetItem(AddressFields.Sortcode, vFields)
        End If
        If (pRSType And AddressRecordSetTypes.artPaf) > 0 Then
          .SetItem(AddressFields.MosaicCode, vFields)
          .SetItem(AddressFields.Paf, vFields)
          .SetItem(AddressFields.AmendedBy, vFields)
          .SetItem(AddressFields.AmendedOn, vFields)
        End If
        If (pRSType And AddressRecordSetTypes.artAddressLines) > 0 Then
          .SetOptionalItem(AddressFields.AddressLine1, vFields)
          .SetOptionalItem(AddressFields.AddressLine2, vFields)
          .SetOptionalItem(AddressFields.AddressLine3, vFields)
          .SetOptionalItem(AddressFields.AddressLine4, vFields)
        Else
          DefaultAddressLines()
        End If
        If (pRSType And AddressRecordSetTypes.artDPS) > 0 Then
          .SetOptionalItem(AddressFields.DeliveryPointSuffix, vFields)
        End If
        If vFields.Exists("address_confirmed") Then .SetOptionalItem(AddressFields.AddressConfirmed, vFields)
      End With
      mvAddressBlank = mvClassFields.Item(AddressFields.Address).Value = " "
    End Sub

    Private Sub DefaultAddressLines()
      Dim vAddressLines() As String
      Dim vIndex As Integer
      Dim vValue As String
      vValue = mvClassFields.Item(AddressFields.Address).Value
      mvAddressBlank = vValue = " "
      vAddressLines = Split(Replace$(vValue, vbCr, "") & vbLf & vbLf & vbLf, vbLf)
      For vIndex = 0 To 3
        If Len(vAddressLines(vIndex)) > mvAddressLinesLength Then AdjustLineLengths(vAddressLines, vIndex)
      Next
      mvClassFields(AddressFields.AddressLine1).SetValue = vAddressLines(0)
      mvClassFields(AddressFields.AddressLine2).SetValue = vAddressLines(1)
      mvClassFields(AddressFields.AddressLine3).SetValue = vAddressLines(2)
      mvClassFields(AddressFields.AddressLine4).SetValue = vAddressLines(3)
    End Sub

    Private Sub AdjustLineLengths(ByVal pAddress() As String, ByVal pItem As Integer)
      Dim vValue As String
      Dim vPos As Integer

      vValue = pAddress(pItem)
      If Len(vValue) > mvAddressLinesLength Then
        vPos = mvAddressLinesLength + 1
        Do
          If Mid$(vValue, vPos, 1) = " " Then Exit Do
          vPos = vPos - 1
        Loop While vPos > 0
        If vPos = 0 Then
          pAddress(pItem) = Trim$(Left$(vValue, mvAddressLinesLength))
          vPos = mvAddressLinesLength
        Else
          pAddress(pItem) = Trim$(Left$(vValue, vPos))
        End If
        If pItem < 3 Then
          If Len(pAddress(pItem + 1)) > 0 Then
            pAddress(pItem + 1) = Mid$(vValue, vPos + 1) & ", " & pAddress(pItem + 1)
          Else
            pAddress(pItem + 1) = Mid$(vValue, vPos + 1)
          End If
        End If
      End If
    End Sub

    Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Dim vExisting As Boolean = mvExisting
      SetValid()
      If Not mvExisting Then CreateAddressRegions()
      If Not mvExisting OrElse mvClassFields(AddressFields.Postcode).ValueChanged Then CreateGridReference(mvEasting, mvNorthing)
      If Postcode.Length = 0 OrElse Paf.Length = 0 OrElse Paf = "NV" Then
        mvClassFields(AddressFields.DeliveryPointSuffix).Value = ""
        mvLEACode = ""
        mvLEAName = ""
      End If
      If Town = "#" AndAlso Not Country = "UK" Then
        mvClassFields.Item(AddressFields.Town).Value = " "
      End If
      MyBase.Save(pAmendedBy, pAudit, pJournalNumber)
      UpdateAdditionalData(vExisting, mvLEACode, mvLEAName)
    End Sub

    Public Sub CreateAddressRegions()
      If mvEnv.IsCountryUK(Country) Then mvUK = True
      If Postcode.Length > 0 AndAlso UK Then
        'Create Address Geographical Regions Data if it does not already exist
        'for this Postcode
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("postcode", Postcode)
        If mvEnv.Connection.GetCount("address_geographical_regions", vWhereFields) = 0 Then
          Dim vPostcode As New Postcode(Postcode)
          Dim vRegion As String = ""
          vWhereFields.Add("geographical_region_type")
          Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, "geographical_region_type", "geographical_region_types").GetRecordSet
          While vRecordSet.Fetch
            'For each region type, do iterative search on region postcodes to attempt
            'to assign a Region in same manner as GetBranchPostcode function.
            'eg. if the postcode entered is XX99 4ZZ the searches will be on
            'XX994ZZ, XX994, XX99, XX in that sequence, unless or until a valid record is found.
            vWhereFields("geographical_region_type").Value = vRecordSet.Fields(1).Value
            If vPostcode.FullPC.Length > 0 Then
              vWhereFields("postcode").Value = vPostcode.FullPC
              vRegion = New SQLStatement(mvEnv.Connection, "geographical_region", "geographical_region_postcodes", vWhereFields).GetValue
            End If
            If vRegion.Length = 0 AndAlso vPostcode.AfterSpace.Length > 0 Then
              If vPostcode.AfterSpace.Length <= vPostcode.MaxLength Then
                vWhereFields("postcode").Value = vPostcode.AfterSpace
                vRegion = New SQLStatement(mvEnv.Connection, "geographical_region", "geographical_region_postcodes", vWhereFields).GetValue
              End If
            End If
            If vRegion.Length = 0 Then
              If vPostcode.BeforeSpace.Length <= vPostcode.MaxLength Then
                vWhereFields("postcode").Value = vPostcode.BeforeSpace
                vRegion = New SQLStatement(mvEnv.Connection, "geographical_region", "geographical_region_postcodes", vWhereFields).GetValue
              End If
              If vRegion.Length = 0 AndAlso vPostcode.Alpha.Length <= vPostcode.MaxLength AndAlso vPostcode.Alpha.Length > 0 Then
                vWhereFields("postcode").Value = vPostcode.Alpha
                vRegion = New SQLStatement(mvEnv.Connection, "geographical_region", "geographical_region_postcodes", vWhereFields).GetValue
              End If
            End If
            If vRegion.Length > 0 Then
              'Have found a Region, create an address_geographical_regions record.
              Dim vInsertFields As New CDBFields
              vInsertFields.AddAmendedOnBy(mvEnv.User.UserID)
              vInsertFields.Add("geographical_region_type", vRecordSet.Fields(1).Value)
              vInsertFields.Add("geographical_region", vRegion)
              vInsertFields.Add("postcode", Postcode)
              mvEnv.Connection.InsertRecord("address_geographical_regions", vInsertFields)
              vRegion = ""
            End If
          End While
          vRecordSet.CloseRecordSet()
        End If
      End If
    End Sub

    Sub SaveBranchHistory(ByVal pAddressNumber As Integer, ByVal pOldBranch As String, ByVal pNewBranch As String, ByRef pInformationMessage As String)
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields

      'Create or update branch history records if the branch is different
      If pNewBranch <> pOldBranch Then
        If mvEnv.GetConfigOption("me_maintain_branch_history", False) Then
          If pOldBranch.Length > 0 Then
            'Try to close off old branch history record
            vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, pAddressNumber)
            vWhereFields.Add("branch", CDBField.FieldTypes.cftCharacter, pOldBranch)
            vUpdateFields.Add("historical", CDBField.FieldTypes.cftCharacter, "Y")
            vUpdateFields.Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate)
            vUpdateFields.AddAmendedOnBy(mvEnv.User.UserID)
            If mvEnv.Connection.UpdateRecords("branch_history", vUpdateFields, vWhereFields, False) = 0 Then
              pInformationMessage = ProjectText.String16007 'Failed to set branch history record as Historic. The branch history record may have been missing
            End If
          End If
          If pNewBranch.Length > 0 Then
            'Create new branch history record
            vUpdateFields.Clear()
            vUpdateFields.Add("address_number", CDBField.FieldTypes.cftLong, pAddressNumber)
            vUpdateFields.Add("branch", CDBField.FieldTypes.cftCharacter, pNewBranch)
            vUpdateFields.Add("valid_from", CDBField.FieldTypes.cftDate, TodaysDate)
            vUpdateFields.Add("historical", CDBField.FieldTypes.cftCharacter, "N")
            vUpdateFields.AddAmendedOnBy(mvEnv.User.UserID)
            mvEnv.Connection.InsertRecord("branch_history", vUpdateFields)
          End If
        End If
      End If
    End Sub

    Public ReadOnly Property MergeAddressLine1() As String
      Get
        Dim vResult As String = ""
        Dim vOverflow As String = ""

        If mvMailMergeAddress.Length = 0 Then FormatMergeAddress()
        DoOverflow(vResult, mvMailMergeAddress, vOverflow)
        Return vResult
      End Get
    End Property
    Public ReadOnly Property MergeAddressLine2() As String
      Get
        Dim vResult As String = ""
        Dim vOverflow As String = ""

        If mvMailMergeAddress.Length = 0 Then FormatMergeAddress()
        DoOverflow(vResult, mvMailMergeAddress, vOverflow)
        DoOverflow(vResult, vOverflow, vOverflow)
        Return vResult
      End Get
    End Property
    Public ReadOnly Property MergeAddressLine3() As String
      Get
        Dim vResult As String = ""
        Dim vOverflow As String = ""

        If mvMailMergeAddress.Length = 0 Then FormatMergeAddress()
        DoOverflow(vResult, mvMailMergeAddress, vOverflow)
        DoOverflow(vResult, vOverflow, vOverflow)
        DoOverflow(vResult, vOverflow, vOverflow)
        Return vResult
      End Get
    End Property
    Public ReadOnly Property MergeAddressLine4() As String
      Get
        Dim vResult As String = ""
        Dim vOverflow As String = ""

        If mvMailMergeAddress.Length = 0 Then FormatMergeAddress()
        DoOverflow(vResult, mvMailMergeAddress, vOverflow)
        DoOverflow(vResult, vOverflow, vOverflow)
        DoOverflow(vResult, vOverflow, vOverflow)
        DoOverflow(vResult, vOverflow, vOverflow)
        Return vResult
      End Get
    End Property

    Private Sub FormatMergeAddress()
      'Format an address for mailmerge, tapemerge or a report
      'This is equivalent to the mwriter include file 'format_merge_address'
      Dim vAddress As String
      Dim vRequiredLines As Integer
      Dim vLineWidth As Integer
      Dim vLines As Integer
      Dim vLen1 As Integer
      Dim vLen2 As Integer
      Dim vAddressLines As Integer
      Dim vSaveLines As Integer
      Dim vTempAddress As String = ""
      Dim vAddress2 As String = ""
      Dim vAddress3 As String

      vRequiredLines = 3
      vLineWidth = 35
      vLines = 0
      vAddress = mvClassFields.Item(AddressFields.Address).Value
      If HouseName.Length > 0 Then
        If Country = "CH" And Len(mvEnv.GetConfig("uniserv_post_ch")) > 0 Then
          vAddress = vAddress & vbLf & mvClassFields.Item(AddressFields.HouseName).Value
        Else
          vAddress = mvClassFields.Item(AddressFields.HouseName).Value & vbLf & vAddress
        End If
      End If
      'Get the first line and count the total lines
      DoOverflow(vTempAddress, vAddress, vAddress2)
      vAddressLines = 1
      While vAddress2 <> ""
        vAddressLines = vAddressLines + 1
        DoOverflow(vTempAddress, vAddress2, vAddress2)
        If vTempAddress = "" Then 'ignore blank lines
          vAddressLines = vAddressLines - 1
        End If
      End While

      'let check how many lines are available and see if we need to batch any up
      vAddress2 = vAddress
      vAddress = ""
      vSaveLines = vAddressLines - vRequiredLines
      If vSaveLines > 0 Then
        While vSaveLines > 0 And vAddress2 <> "" And vLines < (vRequiredLines + 1)
          DoOverflow(vTempAddress, vAddress2, vAddress2)
          vLen1 = Len(vTempAddress)
          While vAddress2 <> "" And vLen1 <> 0 And vLines < (vRequiredLines + 1)
            vAddress3 = vTempAddress
            DoOverflow(vTempAddress, vAddress2, vAddress2)
            If vTempAddress <> "" Then
              vLen2 = Len(vTempAddress)
              If (vLen1 + 2 + vLen2) > vLineWidth Then
                If vLines = 0 Then
                  vAddress = vAddress3
                Else
                  vAddress = vAddress & vbLf & vAddress3
                End If
                vLen1 = vLen2
              Else
                If vLines = 0 Then
                  vAddress = vAddress3 & ", " & vTempAddress
                Else
                  vAddress = vAddress & vbLf & vAddress3 & ", " & vTempAddress
                End If
                vSaveLines = vSaveLines - 1
                vLen1 = 0
              End If
              vLines = vLines + 1
            Else
              vTempAddress = vAddress3
            End If
          End While
        End While
      End If
      If vSaveLines <= 0 And vAddress2 <> "" Then
        DoOverflow(vTempAddress, vAddress2, vAddress2)
        If vTempAddress <> "" Then
          If vLines = 0 Then
            vAddress = vTempAddress
          Else
            vAddress = vAddress & vbLf & vTempAddress
          End If
          vLines = vLines + 1
        End If
        While vAddress2 <> ""
          DoOverflow(vTempAddress, vAddress2, vAddress2)
          If vTempAddress <> "" Then
            If vLines = 0 Then
              vAddress = vTempAddress
            Else
              vAddress = vAddress & vbLf & vTempAddress
            End If
            vLines = vLines + 1
          End If
        End While
      End If
      mvMailMergeAddress = vAddress
    End Sub

    Private Sub DoOverflow(ByRef pResult As String, ByRef pSource As String, ByRef pOverflow As String)
      'This routine acts like the mwriter overflow function
      'Given a source string it will return data up to the first newline
      'in the result and the rest of the string in the overflow string
      Dim vPos As Integer
      Dim vResultPos As Integer

      vPos = InStr(pSource, vbLf)
      If vPos > 0 Then
        vResultPos = vPos - 1
        If vResultPos > 0 Then
          If Mid(pSource, vResultPos, 1) = vbCr Then
            vResultPos = vResultPos - 1
          End If
        End If
        If vResultPos > 0 Then
          pResult = Left(pSource, vResultPos)
        Else
          pResult = ""
        End If
        pOverflow = Mid(pSource, vPos + 1)
      Else
        pResult = pSource
        pOverflow = ""
      End If
    End Sub

    Public Function AddressNumbersOrName() As String
      'Used for Credit Card Authorisation as part of the security checks
      'This will return either:
      ' - The numbers extracted from the Address attribute
      ' - OR, the House Name attribute
      ' - OR, the first line of the Address attribute
      'Only a maximum of 20 characters is returned with no more than 6 numerics
      Dim vChar As String
      Dim vIndex As Integer
      Dim vReturnString As String = ""

      'First, try and return just numerics (maximum 6 numerics)
      Dim vAddressString As String = Replace(Replace(mvClassFields.Item(AddressFields.Address).Value, vbCr, ""), vbLf, "")
      For vIndex = 1 To vAddressString.Length
        vChar = Mid(vAddressString, vIndex, 1)
        If (vChar >= "0" And vChar <= "9") Then
          vReturnString = vReturnString & vChar
        End If
      Next
      If vReturnString.Length > 6 Then vReturnString = Left(vReturnString, 6)

      'Second, if no numerics then return the address without carriage-returns or line-feeds
      If vReturnString.Length = 0 Then vReturnString = HouseName
      If vReturnString.Length = 0 Then
        vReturnString = vAddressString
        vIndex = InStr(1, vReturnString, vbCrLf)
        If vIndex > 0 Then vReturnString = Left(vReturnString, vIndex - 1)
      End If

      If vReturnString.Length > 20 Then vReturnString = Left(vReturnString, 20)
      If InStr(1, vReturnString, " ") > 0 Then
        'If it contains spaces then need to put quotes around it
        vReturnString = """" & vReturnString & """"
      End If
      Return vReturnString
    End Function

    Friend Function CCAuthorisationPostcode() As String
      'Used by On-line CC Authorisation
      'The postcode output must have a maximum of 20 characters with no more than 5 numerics
      Dim vChar As String
      Dim vCount As Integer
      Dim vIndex As Integer
      Dim vReturnString As String = ""

      Dim vPostCodeString As String = Postcode
      'Read the PostCode and ensure that there are no more than 5 numerics (probably only foreign addresses)
      'If there are more than 5 numerics then the return string will only be up to the 6th numeric
      If vPostCodeString.Length > 0 Then
        For vIndex = 1 To vPostCodeString.Length
          vChar = Mid(vPostCodeString, vIndex, 1)
          If (vChar >= "0" And vChar <= "9") Then vCount = vCount + 1
          If vCount <= 5 Then vReturnString = vReturnString & vChar
          If vCount >= 6 Then Exit For
        Next
        vReturnString = vReturnString.Trim
      End If

      If InStr(1, vReturnString, " ") > 0 Then
        'If it contains spaces then need to put quotes around it
        vReturnString = """" & vReturnString & """"
      End If
      Return vReturnString
    End Function

    Public ReadOnly Property AddressText() As String
      Get
        Return mvClassFields.Item(AddressFields.Address).MultiLineValue
      End Get
    End Property

    Public ReadOnly Property BranchName() As String
      Get
        Return mvEnv.GetDescription("branches", "branch", Branch)
      End Get
    End Property

    Public ReadOnly Property AddressMultiLine() As String
      Get
        Return FormatAddress(True)
      End Get
    End Property

    Public Sub CloseSite(ByRef pOrganisationNumber As Integer, ByRef pNewAddressNumber As Integer)
      Dim vContactPosition As ContactPosition
      Dim vContact As Contact
      Dim vOldContactAddress As ContactAddress
      Dim vNewContactAddress As ContactAddress
      Dim vWhereFields As CDBFields
      Dim vUpdateFields As CDBFields
      Dim vUsage As String
      Dim vUpdateOldFields As CDBFields
      Dim vOldContactPosition As ContactPosition
      Dim vNewContactPosition As ContactPosition
      Dim vUpdateContactAddress As CDBFields
      Dim vNewContactAddress1 As ContactAddress
      Dim vOverrideConfig As Boolean


      'Set up the new address object
      Dim vNewAddress As Address = New Address(mvEnv)
      vNewAddress.Init(pNewAddressNumber)
      'Set up the old address object
      'Loop thru the positions at the site being closed
      Dim vMail As Boolean
      Dim vNewMail As String
      For Each vContactPosition In PositionsAtAddress
        With vContactPosition
          vOldContactAddress = New ContactAddress(mvEnv)
          vOldContactAddress.InitFromContactAndAddress(mvEnv, ContactAddress.ContactAddresssLinkTypes.caltContact, .ContactNumber, .AddressNumber)
          vWhereFields = New CDBFields
          With vWhereFields
            .Add("contact_number", CDBField.FieldTypes.cftLong, vOldContactAddress.ContactNumber)
            .Add("address_number", CDBField.FieldTypes.cftLong, vOldContactAddress.AddressNumber)
          End With
          vUpdateFields = New CDBFields
          With vUpdateFields
            .Add("address_number", CDBField.FieldTypes.cftLong, vNewAddress.AddressNumber)
            .Add("amended_by", CDBField.FieldTypes.cftCharacter, mvEnv.User.UserID)
            .Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate())
          End With
          If vNewAddress.ContactsAtAddress(True).Exists(.ContactNumber.ToString) Then 'Contact already has a link to the new address
            vNewContactAddress = New ContactAddress(mvEnv)
            vNewContactAddress.InitFromContactAndAddress(mvEnv, ContactAddress.ContactAddresssLinkTypes.caltContact, .ContactNumber, (vNewAddress.AddressNumber))
            'Does the contact have the same usages between the two addresses?
            vWhereFields.Add("address_usage")
            For Each vUsage In vOldContactAddress.Usages
              vWhereFields("address_usage").Value = vUsage
              If vNewContactAddress.Usages.Contains(vUsage) Then
                'Delete the usages for the old address
                mvEnv.Connection.DeleteRecords("contact_address_usages", vWhereFields)
              Else
                'Update the usages for old address to point to the new address
                mvEnv.Connection.UpdateRecords("contact_address_usages", vUpdateFields, vWhereFields)
              End If
            Next vUsage
            'Delete contact_addresses record for contact at the old address
            vWhereFields.Remove("address_usage")
            If mvEnv.GetConfigOption("retain_closed_site_links", False) Then
              vUpdateContactAddress = New CDBFields
              With vUpdateContactAddress
                .Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate())
                .Add("historical", CDBField.FieldTypes.cftCharacter, "Y")
                .Add("amended_by", CDBField.FieldTypes.cftCharacter, mvEnv.User.UserID)
                .Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate())
              End With
              mvEnv.Connection.UpdateRecords("contact_addresses", vUpdateContactAddress, vWhereFields)
            Else
              mvEnv.Connection.DeleteRecords("contact_addresses", vWhereFields)
            End If
          Else
            'Contact is not already linked to the new address
            'Change the usages for old address to point to the new address
            mvEnv.Connection.UpdateRecords("contact_address_usages", vUpdateFields, vWhereFields, False)
            'don't retain old address, but keep default address
            If Not mvEnv.GetConfigOption("retain_closed_site_links", False) And _
              mvEnv.GetConfigOption("keep_contacts_default_address", False) Then
              'need to keep old address as default and, therefore, add new record for new address
              vContact = New Contact(mvEnv)
              vContact.Init(.ContactNumber)
              If vContact.AddressNumber = vOldContactAddress.AddressNumber Then
                If (Me.GetCurrentMemberships(vOldContactAddress.ContactNumber)) > 0 Then
                  vUpdateContactAddress = New CDBFields
                  With vUpdateContactAddress
                    .Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate())
                    .Add("historical", CDBField.FieldTypes.cftCharacter, "Y")
                    .Add("amended_by", CDBField.FieldTypes.cftCharacter, mvEnv.User.UserID)
                    .Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate())
                  End With
                  mvEnv.Connection.UpdateRecords("contact_addresses", vUpdateContactAddress, vWhereFields)
                  'create a new record
                  vNewContactAddress1 = New ContactAddress(mvEnv)
                  vNewContactAddress1.Init()
                  vNewContactAddress1.Create(mvEnv, ContactAddress.ContactAddresssLinkTypes.caltContact, vOldContactAddress.ContactNumber, vNewAddress.AddressNumber, "N", TodaysDate(), vOldContactAddress.ValidTo, mvEnv.User.UserID, TodaysDate())
                  vNewContactAddress1.Save(mvEnv.User.UserID, True)
                  'also, keep old position record at this site
                  'update record to show as historic
                  vOldContactPosition = New ContactPosition(mvEnv)
                  vOldContactPosition.Init(vOldContactAddress.ContactNumber, vContactPosition.AddressNumber, pOrganisationNumber, , , , ContactPosition.CurrentSettingTypes.cstCurrent)
                  Dim vMail1 As Boolean
                  vMail1 = vOldContactPosition.Mail
                  Dim vNewMail1 As String
                  If vMail1 Then
                    vNewMail1 = "Y"
                  Else
                    vNewMail1 = "N"
                  End If
                  vUpdateOldFields = New CDBFields
                  vUpdateOldFields.Add("finished", CDBField.FieldTypes.cftDate, TodaysDate())
                  vUpdateOldFields.Add("amended_by", CDBField.FieldTypes.cftCharacter, mvEnv.User.UserID)
                  vUpdateOldFields.Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate())
                  vUpdateOldFields.Add("mail", CDBField.FieldTypes.cftCharacter, "N")
                  vUpdateOldFields.Add("current", CDBField.FieldTypes.cftCharacter, "N")
                  vUpdateOldFields("current").SpecialColumn = True '.Connection.DBSpecialCol("cp", "current")
                  mvEnv.Connection.UpdateRecords("contact_positions", vUpdateOldFields, vWhereFields)
                  'and insert a new record
                  vNewContactPosition = New ContactPosition(mvEnv)
                  vNewContactPosition.Init()
                  vNewContactPosition.Create(vOldContactAddress.ContactNumber, pNewAddressNumber, pOrganisationNumber, vNewMail1, "True", vOldContactPosition.Position, TodaysDate(), vOldContactPosition.Finished, vOldContactPosition.PositionLocation, vOldContactPosition.PositionFunction, vOldContactPosition.PositionSeniority)
                  vNewContactPosition.Save(mvEnv.User.UserID, True)

                  vOverrideConfig = True

                End If
              End If
            End If
            'Update the contact_addresses record to point to the new address
            If vOverrideConfig = False Then
              If mvEnv.GetConfigOption("retain_closed_site_links", False) Then
                'update the valid to and historical
                vUpdateContactAddress = New CDBFields
                With vUpdateContactAddress
                  .Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate())
                  .Add("historical", CDBField.FieldTypes.cftCharacter, "Y")
                  .Add("amended_by", CDBField.FieldTypes.cftCharacter, mvEnv.User.UserID)
                  .Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate())
                End With
                mvEnv.Connection.UpdateRecords("contact_addresses", vUpdateContactAddress, vWhereFields)
                'create a new record
                vNewContactAddress1 = New ContactAddress(mvEnv)
                vNewContactAddress1.Init()
                vNewContactAddress1.Create(mvEnv, ContactAddress.ContactAddresssLinkTypes.caltContact, vOldContactAddress.ContactNumber, vNewAddress.AddressNumber, "N", TodaysDate(), vOldContactAddress.ValidTo, mvEnv.User.UserID, TodaysDate())
                vNewContactAddress1.Save(mvEnv.User.UserID, True)
              Else
                mvEnv.Connection.UpdateRecords("contact_addresses", vUpdateFields, vWhereFields)
              End If
            End If
          End If
          'Move the position to the new site
          vWhereFields.Add("organisation_number", CDBField.FieldTypes.cftLong, .OrganisationNumber)

          If vOverrideConfig = False Then

            If mvEnv.GetConfigOption("retain_closed_site_links", False) Then
              ' update record to show as historic
              vOldContactPosition = New ContactPosition(mvEnv)
              vOldContactPosition.Init(vOldContactAddress.ContactNumber, vContactPosition.AddressNumber, pOrganisationNumber, , , , ContactPosition.CurrentSettingTypes.cstCurrent)
              vMail = vOldContactPosition.Mail
              If vMail Then
                vNewMail = "Y"
              Else
                vNewMail = "N"
              End If
              vUpdateOldFields = New CDBFields
              vUpdateOldFields.Add("finished", CDBField.FieldTypes.cftDate, TodaysDate())
              vUpdateOldFields.Add("amended_by", CDBField.FieldTypes.cftCharacter, mvEnv.User.UserID)
              vUpdateOldFields.Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate())
              vUpdateOldFields.Add("mail", CDBField.FieldTypes.cftCharacter, "N")
              vUpdateOldFields.Add("current", CDBField.FieldTypes.cftCharacter, "N")
              vUpdateOldFields("current").SpecialColumn = True '.Connection.DBSpecialCol("cp", "current")
              mvEnv.Connection.UpdateRecords("contact_positions", vUpdateOldFields, vWhereFields)
              'and insert a new record
              vNewContactPosition = New ContactPosition(mvEnv)
              vNewContactPosition.Init()
              vNewContactPosition.Create(vOldContactAddress.ContactNumber, pNewAddressNumber, pOrganisationNumber, vNewMail, "True", vOldContactPosition.Position, TodaysDate(), vOldContactPosition.Finished, vOldContactPosition.PositionLocation, vOldContactPosition.PositionFunction, vOldContactPosition.PositionSeniority)
              vNewContactPosition.Save(mvEnv.User.UserID, True)
            Else
              'Here I think we are just updating the address on the current position which is vContactPosition
              Dim vParams As New CDBParameters
              vParams.Add("AddressNumber", pNewAddressNumber)
              vContactPosition.Update(vParams)
              vContactPosition.Save(mvEnv.User.UserID, True)
            End If
          End If
          'Is the old address the contact's default address?
          vContact = New Contact(mvEnv)
          vContact.Init(.ContactNumber)
          If vContact.AddressNumber = AddressNumber Then ' if contact default is same as old address number
            If mvEnv.GetConfigOption("keep_contacts_default_address", False) Then
              If (Me.GetCurrentMemberships(vOldContactAddress.ContactNumber)) > 0 Then
                vContact.SwitchCurrentAddress(Me, vNewAddress, True)
              Else
                vContact.SetDefaultAddress(vNewAddress, True, True, True)
              End If
            Else
              vContact.SetDefaultAddress(vNewAddress, True, True, True)
            End If
          Else
            vContact.SwitchCurrentAddress(Me, vNewAddress, True)
          End If
        End With
      Next vContactPosition
      'Set the closed site address to historic
      Dim vOrganisationAddress As New OrganisationAddress(mvEnv)
      vOrganisationAddress.Init(pOrganisationNumber, AddressNumber)
      vOrganisationAddress.ValidTo = TodaysDate()
      vOrganisationAddress.Historical = True
      vOrganisationAddress.Save(mvEnv.User.UserID, True)
    End Sub

    Public ReadOnly Property PositionsAtAddress() As Collection
      Get
        Dim vContactPosition As New ContactPosition(mvEnv)

        If mvPositionsAtAddress Is Nothing And AddressType = AddressTypes.ataOrganisation Then
          mvPositionsAtAddress = New Collection
          vContactPosition = New ContactPosition(mvEnv)
          vContactPosition.Init()
          Dim vSQL As String = "SELECT " & vContactPosition.GetRecordSetFields()
          vSQL = vSQL & " FROM contact_positions cp, contact_addresses ca, contacts c"
          vSQL = vSQL & " WHERE cp.address_number = " & AddressNumber
          vSQL = vSQL & " AND " & mvEnv.Connection.DBSpecialCol("cp", "current") & " = 'Y'"
          vSQL = vSQL & " AND cp.contact_number = ca.contact_number"
          vSQL = vSQL & " AND cp.address_number = ca.address_number"
          vSQL = vSQL & " AND ca.contact_number = c.contact_number"
          vSQL = vSQL & " AND c.contact_type <> 'O'"
          Dim vRS As CDBRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
          While vRS.Fetch
            vContactPosition = New ContactPosition(mvEnv)
            vContactPosition.InitFromRecordSet(vRS)
            mvPositionsAtAddress.Add(vContactPosition, vRS.Fields("contact_number").Value)
          End While
          vRS.CloseRecordSet()
        End If
        Return mvPositionsAtAddress
      End Get
    End Property

    Public ReadOnly Property ContactsAtAddress(Optional ByVal pPopulateCollection As Boolean = False) As Contacts
      Get
        If mvContactsAtAddress Is Nothing Then
          mvContactsAtAddress = New Contacts(mvEnv)
          If pPopulateCollection Then SetContactsAtAddress(AddressNumber)
        End If
        Return mvContactsAtAddress
      End Get
    End Property

    Private Sub SetContactsAtAddress(ByVal pAddressNumber As Integer) ', ByVal pAddressNo As Long)
      Dim vContact As New Contact(mvEnv)
      Dim vRS As CDBRecordSet
      Dim vSQL As String

      vContact.Init()
      mvContactsAtAddress = New Contacts(mvEnv)
      vSQL = "SELECT " & vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtName) & " FROM contact_addresses ca, contacts c"
      vSQL = vSQL & " WHERE ca.address_number = " & pAddressNumber & " AND c.contact_number = ca.contact_number"
      vSQL = vSQL & " ORDER BY c.contact_number"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      While vRS.Fetch
        vContact = mvContactsAtAddress.Add(vRS.Fields("contact_number").Value)
        vContact.InitFromRecordSet(mvEnv, vRS, Contact.ContactRecordSetTypes.crtName)
      End While
      vRS.CloseRecordSet()
    End Sub

    Public Function GetCurrentMemberships(ByVal pContactNumber As Integer) As Integer
      Dim vCount As Integer = 0
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("contact_number", pContactNumber)
      vWhereFields.Add("cancelled_on", CDBField.FieldTypes.cftDate)
      If pContactNumber > 0 Then
        vCount = mvEnv.Connection.GetCount("members", vWhereFields)
      End If
      Return vCount
    End Function

    Public Sub SetGovernmentRegion()
      mvClassFields(AddressFields.MosaicCode).Value = GetGovernmentRegionFromPostcode(Postcode)
    End Sub

    Public Function GetGovernmentRegionFromPostcode(ByVal pPostCode As String) As String
      'Does iterative searches on database to broaden postcode branch link
      'eg. if the postcode entered is XX99 4ZZ the searches will be on
      'XX994ZZ, XX994, XX99, XX in that sequence, unless or until a valid record is found.
      'Modified from XX994, XX99, XX

      Dim vGovtRegion As String = ""
      Dim vFullPC As String = ""       'Full postcode
      If pPostCode.Length > 0 Then
        Dim vSpaceReqPos As Integer = 6 'If the space is before position 6 then keep it now
        Dim vMaxLength As Integer = 10
        Dim vSpacePos As Integer = InStr(pPostCode, " ")
        Dim vBeforeSpace As String        'lhs of postcode prior to space
        Dim vAfterSpace As String = ""    'lhs of postcode including first character after space
        If vSpacePos = 0 Then
          vBeforeSpace = pPostCode
        Else
          vFullPC = pPostCode
          Dim vOutwardPC As String = Left(pPostCode, vSpacePos - 1) 'All before the space
          vBeforeSpace = Left(pPostCode, vSpacePos - 1)
          Dim vInwardPC As String = Left(LTrim(Mid(pPostCode, vSpacePos)), 1) '1st character after space
          If vInwardPC.Length > 0 Then
            If vSpacePos < vSpaceReqPos Then
              vOutwardPC = vOutwardPC & " " & vInwardPC
            Else
              vOutwardPC = vOutwardPC & vInwardPC
            End If
            vAfterSpace = vOutwardPC
          End If
        End If
        Dim vAlpha As String = RTrim(vBeforeSpace) 'lhs of postcode - up to first numeric
        Dim vLength As Integer = vAlpha.Length
        Do While IsNumeric(Mid(vAlpha, vLength, 1))
          vLength = vLength - 1
          If vLength = 0 Then Exit Do
        Loop
        vAlpha = Left(vAlpha, vLength)

        If vFullPC.Length > 0 Then
          vGovtRegion = mvEnv.Connection.GetValue("SELECT government_region FROM government_region_postcodes WHERE postcode = '" & vFullPC & "'")
        End If
        If vGovtRegion.Length = 0 And vAfterSpace.Length > 0 Then
          If vAfterSpace.Length <= vMaxLength Then vGovtRegion = mvEnv.Connection.GetValue("SELECT government_region FROM government_region_postcodes WHERE postcode = '" & vAfterSpace & "'")
        End If
        If vGovtRegion.Length = 0 Then
          If Len(vBeforeSpace) <= vMaxLength Then vGovtRegion = mvEnv.Connection.GetValue("SELECT government_region FROM government_region_postcodes WHERE postcode = '" & vBeforeSpace & "'")
          If vGovtRegion.Length = 0 Then If vAlpha.Length <= vMaxLength And vAlpha.Length > 0 Then vGovtRegion = mvEnv.Connection.GetValue("SELECT government_region FROM government_region_postcodes WHERE postcode = '" & vAlpha & "'")
        End If
      End If
      If vGovtRegion = "" Then vGovtRegion = mvEnv.GetConfig("cd_unknown_region")
      Return vGovtRegion
    End Function

    Public Overloads Sub Update(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters, pAddressBlank As Boolean)
      mvAddressBlank = pAddressBlank
      Update(pEnv, pParams)
    End Sub

    Public Overloads Sub Update(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'BR18541 Added NoAddressCapitalisation parameter to be used by AddContact web service
      Dim vNoAddressCapitalisation As Boolean = (pParams.OptionalValue("NoCapitalisation", "") = "Y") OrElse (pParams.OptionalValue("NoAddressCapitalisation", "") = "Y")
      Dim vNoCapitalisation As Boolean = vNoAddressCapitalisation
      Dim vUniservMail As Boolean = pEnv.GetConfig("uniserv_mail").Length > 0
      Dim vCountry As String
      If pParams.ParameterExists("Country").Value.Length = 0 Then
        If Country.Length = 0 Then
          vCountry = mvEnv.DefaultCountry 'Only set default country if no country given and there isn't one already
        Else
          vCountry = Country
        End If
      Else
        vCountry = pParams("Country").Value
      End If
      Dim vUK As Boolean = pEnv.IsCountryUK(vCountry)
      Dim vAddress As String
      If mvAddressBlank AndAlso pParams.OptionalValue("Address", mvClassFields(AddressFields.Address).Value) = " " Then
        vAddress = " "
      Else
        If vNoCapitalisation OrElse (vUniservMail AndAlso Not vUK) Then
          vAddress = RemoveTextNoise(pParams.OptionalValue("Address", mvClassFields(AddressFields.Address).Value))
        Else
          vAddress = CapitaliseWords(RemoveTextNoise(pParams.OptionalValue("Address", mvClassFields(AddressFields.Address).Value)))
        End If
      End If
      mvAddressBlank = vAddress = " "
      If pParams.Exists("HouseName") Then mvClassFields.Item(AddressFields.HouseName).Value = pParams("HouseName").Value
      If Not vCountry = "UK" AndAlso pParams.Exists("Town") AndAlso pParams("Town").Value = "#" Then
        pParams("Town").Value = " "
      End If
      SetAddressField(vAddress)
      If vUniservMail AndAlso Not vUK Then
        If pParams.Exists("Town") Then mvClassFields.Item(AddressFields.Town).Value = pParams("Town").Value
        If pParams.Exists("County") Then mvClassFields.Item(AddressFields.County).Value = pParams("County").Value
      Else
        If pParams.Exists("Town") Then mvClassFields.Item(AddressFields.Town).Value = UCase(pParams("Town").Value)
        If pParams.Exists("County") Then
          If vNoCapitalisation OrElse pParams.Exists("SmartClient") OrElse pParams.Exists("CarePortal") Then
            mvClassFields.Item(AddressFields.County).Value = pParams("County").Value
          Else
            mvClassFields.Item(AddressFields.County).Value = CapitaliseWords(pParams("County").Value)
          End If
        End If
      End If
      mvClassFields.Item(AddressFields.Country).Value = vCountry
      If pParams.Exists("Postcode") Then mvClassFields.Item(AddressFields.Postcode).Value = pParams("Postcode").Value
      If pParams.Exists("PafStatus") Then mvClassFields.Item(AddressFields.Paf).Value = pParams("PafStatus").Value
      If pParams.Exists("Branch") Then
        mvClassFields(AddressFields.Branch).Value = pParams("Branch").Value
      ElseIf mvClassFields.Item(AddressFields.Postcode).ValueChanged Then
        mvClassFields(AddressFields.Branch).Value = pEnv.GetBranchFromPostcode(pParams.ParameterExists("Postcode").Value)
      End If
      If vUK And (mvClassFields.Item(AddressFields.Postcode).ValueChanged Or mvClassFields.Item(AddressFields.Town).ValueChanged) Then
        mvClassFields.Item(AddressFields.Sortcode).Value = GetMailsortCode((pParams.ParameterExists("Postcode").Value), Nothing, Nothing, Nothing, pParams.ParameterExists("Town").Value)
      End If
      If pParams.Exists("BuildingNumber") Then mvClassFields(AddressFields.BuildingNumber).Value = pParams("BuildingNumber").Value
      If pParams.Exists("DeliveryPointSuffix") Then mvClassFields(AddressFields.DeliveryPointSuffix).Value = pParams("DeliveryPointSuffix").Value

      If pParams.Exists("Easting") Then mvEasting = pParams("Easting").IntegerValue
      If pParams.Exists("Northing") Then mvNorthing = pParams("Northing").IntegerValue
      If pParams.Exists("LeaCode") Then mvLEACode = pParams("LeaCode").Value
      If pParams.Exists("LeaName") Then mvLEAName = pParams("LeaName").Value
      If pParams.Exists("AddressConfirmed") Then mvClassFields.Item(AddressFields.AddressConfirmed).Value = pParams("AddressConfirmed").Value
    End Sub

    Public Sub UpdateFields(ByVal pAddress As String, ByVal pTown As String, ByVal pPostCode As String, ByVal pHouseName As String, ByVal pCounty As String, ByVal pSortCode As String, ByVal pPaf As String)
      SetAddressField(CapitaliseWords(pAddress))
      Dim vUk As Boolean = (New SQLStatement(mvEnv.Connection,
                                             "uk",
                                             "countries",
                                             New CDBFields({New CDBField("country",
                                                                         If(String.IsNullOrWhiteSpace(mvClassFields(AddressFields.Country).Value),
                                                                            mvEnv.GetConfig("option_country", "UK"),
                                                                            mvClassFields(AddressFields.Country).Value))})).GetValue = "Y")
      If Not vUk AndAlso String.IsNullOrWhiteSpace(mvEnv.GetConfig("uniserv_mail")) Then
        mvClassFields.Item(AddressFields.Town).Value = CapitaliseWords(pTown)
      Else
        mvClassFields.Item(AddressFields.Town).Value = pTown.ToUpper
      End If
      mvClassFields.Item(AddressFields.Postcode).Value = pPostCode
      mvClassFields.Item(AddressFields.County).Value = CapitaliseWords(pCounty)
      mvClassFields.Item(AddressFields.HouseName).Value = CapitaliseWords(pHouseName)
      mvClassFields.Item(AddressFields.Sortcode).Value = pSortCode
      mvClassFields.Item(AddressFields.Paf).Value = pPaf
    End Sub

    Private Sub SetAddressField(ByVal pValue As String)
      Dim vAddressLines() As String
      Dim vIndex As Integer

      mvClassFields.Item(AddressFields.Address).Value = pValue
      'Debug.Print(pValue)
      vAddressLines = Split(pValue.Replace(vbCr, "") & vbLf & vbLf & vbLf, vbLf)
      For vIndex = 0 To 3
        If vAddressLines(vIndex).Length > mvAddressLinesLength Then AdjustLineLengths(vAddressLines, vIndex)
      Next
      mvClassFields(AddressFields.AddressLine1).Value = vAddressLines(0)
      mvClassFields(AddressFields.AddressLine2).Value = vAddressLines(1)
      mvClassFields(AddressFields.AddressLine3).Value = vAddressLines(2)
      mvClassFields(AddressFields.AddressLine4).Value = vAddressLines(3)
    End Sub

    Public Function GetMailsortCode(ByVal pPostCode As String, ByVal pSectorList As SortedList(Of String, String), ByVal pDistrictList As SortedList(Of String, String), Optional ByVal pTownList As SortedList(Of String, String) = Nothing, Optional ByVal pTown As String = "") As String
      Dim vSortCode As String = ""
      If pPostCode.Length > 0 Then
        Dim vPCode As String = pPostCode
        Dim vPos As Integer = InStr(vPCode, " ")
        If vPos = 0 Then 'Outward code only
          vPos = vPCode.Length
          If vPos > 4 Then
            vPCode = vPCode.Substring(0, 4)
          End If
          'Check the outward code for a sortcode
          If vPCode.Length <= 4 Then
            If pDistrictList Is Nothing Then
              vSortCode = mvEnv.Connection.GetValue("SELECT standard_selection FROM pc_district WHERE dist_postcode = '" & vPCode & "'")
            Else
              If pDistrictList.ContainsKey(vPCode) Then vSortCode = pDistrictList.Item(vPCode)
            End If
          End If
        Else 'Outward + Inward code
          'Check sector file which contains combination of outward plus space plus 1st char of inward
          Dim vOutCode As String = vPCode.Substring(0, vPos - 1)
          Dim vInCode As String = CStr(vPCode.Substring(vPos - 1).TrimStart(" "c)).Substring(0, 1)
          Dim vSearchCode As String
          If vPos < 5 Then
            vSearchCode = vOutCode & " " & vInCode
          Else
            vSearchCode = vOutCode & vInCode
          End If
          'Check the sectors file (about 90 records)
          If vSearchCode.Length <= 5 Then
            If pSectorList Is Nothing Then
              vSortCode = mvEnv.Connection.GetValue("SELECT standard_selection FROM pc_sector WHERE sect_postcode = '" & vSearchCode & "'")
            Else
              If pSectorList.ContainsKey(vSearchCode) Then vSortCode = pSectorList.Item(vSearchCode)
            End If
          End If
          If vSortCode.Length = 0 Then
            'Not specified by sector so check district i.e. outward (about 3000 records)
            If vOutCode.Length <= 4 Then
              If pDistrictList Is Nothing Then
                vSortCode = mvEnv.Connection.GetValue("SELECT standard_selection FROM pc_district WHERE dist_postcode = '" & vOutCode & "'")
              Else
                If pDistrictList.ContainsKey(vOutCode) Then vSortCode = pDistrictList.Item(vOutCode)
              End If
            End If
          End If
        End If
      End If
      'Null Postcode or Sortcode not found
      If (pPostCode.Length = 0 Or vSortCode.Length = 0) And pTown.Length > 0 Then
        If pTown.Length > 10 Then pTown = pTown.Substring(0, 10) 'Can only use 10-char town names max
        pTown = pTown.Trim(" "c) 'Remove any trailing spaces
        If pTownList Is Nothing Then
          vSortCode = mvEnv.Connection.GetValue("SELECT residue_selection FROM mail_towns WHERE mail_town = '" & pTown.Replace("'", "''") & "'")
        Else
          If pTownList.ContainsKey(pTown) Then vSortCode = pTownList.Item(pTown)
        End If
      End If
      Return vSortCode
    End Function

    Public Overloads Sub Create(ByVal pEnv As CDBEnvironment, ByRef pContact As Contact, ByRef pParams As CDBParameters, pBlankAddress As Boolean)
      mvAddressBlank = True
      Create(pEnv, pContact, pParams)
    End Sub

    Public Overloads Sub Create(ByVal pEnv As CDBEnvironment, ByRef pContact As Contact, ByRef pParams As CDBParameters) ' ByVal pHouseName As String, ByVal pAddress As String, ByVal pTown As String, ByVal pCounty As String, ByVal pPostCode As String, ByVal pCountry As String, pPaf As String, Optional ByVal pBranch As String = "", Optional pBuildingNumber As String = "")
      Init()
      Update(pEnv, pParams) 'pHouseName, pAddress, pTown, pCounty, pPostCode, pCountry, pPaf, pBranch, pBuildingNumber
      If pContact.ContactType = AddressTypes.ataOrganisation Then mvClassFields.Item(AddressFields.AddressType).Value = "O" Else mvClassFields.Item(AddressFields.AddressType).Value = "C"
      If pEnv.GetConfigOption("cd_use_government_regions") Then SetGovernmentRegion()
      SetValid()
    End Sub

    Private Sub CheckValidity()
      ' Code for building number check moved to IsBuildingNumberValid function. This function can
      ' now be extended accordingly for other checks
      If Not IsBuildingNumberValid(mvClassFields(AddressFields.BuildingNumber).Value) Then
        RaiseError(DataAccessErrors.daeBuildingNumberInvalidFormat)
      End If
    End Sub

    Public Function IsBuildingNumberValid(ByVal pBuildingNumber As String) As Boolean
      Dim vValid As Boolean
      Dim vIndex As Integer
      Dim vChar As String
      Dim vFirstWord As String
      Dim vPos As Integer

      vValid = True
      If pBuildingNumber.Length > 0 And (InStr(pBuildingNumber, " ") > 0) Then
        vValid = False
      Else
        For vIndex = InStr(pBuildingNumber, "-") + 1 To Len(pBuildingNumber)
          vChar = Mid(pBuildingNumber, vIndex, 1)
          If Not vChar Like "[a-zA-Z0-9]" Then
            vValid = False
          End If
        Next
        If vValid Then
          vPos = InStr(pBuildingNumber, "-")
          If vPos > 0 Then
            vFirstWord = Left(pBuildingNumber, vPos - 1)
          Else
            vFirstWord = pBuildingNumber
          End If

          For vIndex = 1 To Len(vFirstWord)
            vChar = Mid(vFirstWord, vIndex, 1)
            If Not vChar Like "[0-9]" Then
              vValid = False
            End If
          Next
        End If
      End If
      Return vValid
    End Function

    Public ReadOnly Property AddressLineByCountry() As String
      Get
        ' BR 11347
        ' Allows us to specify the address format for individual countries
        ' Expand as future requirements dictate
        Select Case mvClassFields.Item(AddressFields.Country).Value
          Case "NL"
            Return FormatAddressNL(False)
          Case Else
            Return FormatAddress(False)
        End Select
      End Get
    End Property

    Private Function FormatAddressNL(ByVal pMultiLine As Boolean) As String
      Dim vEuroFormat As Boolean
      Dim vSeparator As String

      vSeparator = If(pMultiLine, vbCrLf, ", ")

      Dim vHouseName As String = mvClassFields.Item(AddressFields.HouseName).Value
      Dim vTown As String = mvClassFields.Item(AddressFields.Town).Value
      Dim vCounty As String = mvClassFields.Item(AddressFields.County).Value
      Dim vCountry As String = mvClassFields.Item(AddressFields.Country).Value
      Dim vPostcode As String = mvClassFields.Item(AddressFields.Postcode).Value
      Dim vAddress As String
      With mvClassFields
        If InStr(.Item(AddressFields.Address).MultiLineValue, vbCrLf) > 0 Then
          vAddress = Replace(.Item(AddressFields.Address).MultiLineValue, vbLf, " " & BuildingNumber & vbLf, 1, 1)
        Else
          vAddress = .Item(AddressFields.Address).Value & " " & BuildingNumber
        End If
        vAddress = If(pMultiLine, Trim(vAddress), Trim(Replace(Replace(vAddress, vbCr, ""), vbLf, " ")))
      End With

      'Add the house name if present
      If vHouseName.Length > 0 Then
        If vAddress.Length > 0 Then
          vAddress = vHouseName & vSeparator & vAddress
        Else
          vAddress = vHouseName
        End If
      End If

      If vPostcode.Length > 0 Then
        If vAddress.Length > 0 Then vAddress = vAddress & vSeparator
        vAddress = vAddress & vPostcode
      End If
      If vTown.Length > 0 Then vAddress = vAddress & " " & vTown
      vEuroFormat = True

      If vTown.Length > 0 And Not vEuroFormat Then
        If vAddress.Length > 0 Then vAddress = vAddress & vSeparator
        vAddress = vAddress & vTown
      End If
      If vCounty.Length > 0 Then vAddress = vAddress & vSeparator & vCounty
      If vPostcode.Length > 0 And Not vEuroFormat Then vAddress = vAddress & vSeparator & vPostcode

      If vCountry.Length > 0 And vCountry <> mvEnv.DefaultCountry Then
        'There is a country code and it is not the default country
        GetCountryDescription()
        If Not (mvUK And mvEnv.DefaultCountry = "UK") Then
          vAddress = vAddress & vSeparator & mvCountryDescription
        End If
      End If
      Return vAddress
    End Function

    Public ReadOnly Property NonDefaultCountryDescription() As String
      Get
        Dim vCountry As String = mvClassFields.Item(AddressFields.Country).Value
        If vCountry.Length > 0 And vCountry <> mvEnv.DefaultCountry Then
          'There is a country code and it is not the default country
          GetCountryDescription()
          If Not (mvUK And mvEnv.DefaultCountry = "UK") Then
            Return mvCountryDescription
          End If
        End If
        Return ""
      End Get
    End Property

    Public ReadOnly Property OrganisationAddressLine() As String
      Get
        Dim vOrganisationName As String = OrganisationName()
        If vOrganisationName.Length > 0 Then
          Return vOrganisationName & vbCrLf & AddressLine()
        Else
          Return AddressLine()
        End If
      End Get
    End Property

    Public ReadOnly Property OrganisationOrAddressLine() As String
      Get
        Dim vOrganisationName As String = OrganisationName()
        If vOrganisationName.Length > 0 Then
          Return vOrganisationName
        Else
          Return AddressLine()
        End If
      End Get
    End Property

    Public Sub InitForReport(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByRef pPrefix As String)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        '("address_line1") 
        If pRecordSet.Fields.Exists(pPrefix & .Item(AddressFields.AddressLine1).Name) Then .Item(AddressFields.AddressLine1).SetValue = vFields(pPrefix & .Item(AddressFields.AddressLine1).Name).Value
        If pRecordSet.Fields.Exists(pPrefix & .Item(AddressFields.AddressLine2).Name) Then .Item(AddressFields.AddressLine2).SetValue = vFields(pPrefix & .Item(AddressFields.AddressLine2).Name).Value
        If pRecordSet.Fields.Exists(pPrefix & .Item(AddressFields.AddressLine3).Name) Then .Item(AddressFields.AddressLine3).SetValue = vFields(pPrefix & .Item(AddressFields.AddressLine3).Name).Value
        .Item(AddressFields.AddressNumber).SetValue = vFields(pPrefix & .Item(AddressFields.AddressNumber).Name).Value
        .Item(AddressFields.AddressType).SetValue = vFields(pPrefix & .Item(AddressFields.AddressType).Name).Value
        .Item(AddressFields.HouseName).SetValue = vFields(pPrefix & .Item(AddressFields.HouseName).Name).Value
        .Item(AddressFields.Address).SetValue = vFields(pPrefix & .Item(AddressFields.Address).Name).Value
        .Item(AddressFields.Town).SetValue = vFields(pPrefix & .Item(AddressFields.Town).Name).Value
        .Item(AddressFields.County).SetValue = vFields(pPrefix & .Item(AddressFields.County).Name).Value
        .Item(AddressFields.Country).SetValue = vFields(pPrefix & .Item(AddressFields.Country).Name).Value
        .Item(AddressFields.Postcode).SetValue = vFields(pPrefix & .Item(AddressFields.Postcode).Name).Value
        .Item(AddressFields.Branch).SetValue = vFields.FieldExists(pPrefix & .Item(AddressFields.Branch).Name).Value
        .Item(AddressFields.Sortcode).SetValue = vFields(pPrefix & .Item(AddressFields.Sortcode).Name).Value
        If .Item(AddressFields.AddressConfirmed).InDatabase And pRecordSet.Fields.Exists("address_confirmed") Then .Item(AddressFields.AddressConfirmed).SetValue = vFields(pPrefix & .Item(AddressFields.AddressConfirmed).Name).Value
        If .Item(AddressFields.BuildingNumber).InDatabase And pRecordSet.Fields.Exists("building_number") Then .Item(AddressFields.BuildingNumber).SetValue = vFields(pPrefix & .Item(AddressFields.BuildingNumber).Name).Value
        mvUK = vFields(pPrefix & "uk").Bool
        mvCountryDescription = vFields(pPrefix & "country_desc").Value
        '  If vFields(pPrefix & .Item(AddressFields.Country).Name).Value.Length > 0 AndAlso String.Compare(vFields(.Item(AddressFields.Country).Name).Value, "uk", True) <> 0 Then GetCountryDescription()
        mvCountryValid = True
      End With
    End Sub

    Public Overrides ReadOnly Property DataTable() As CDBDataTable
      Get
        'This function is only used by WEB Services at present
        'Please let me know if you want to change it (SDT)
        Dim vTable As New CDBDataTable
        GetCountryDescription()
        vTable.AddColumnsFromList("AddressNumber,HouseName,Address,Town,County,Postcode,CountryCode,CountryDesc,UK,Branch,AddressLine")
        vTable.AddColumnsFromList("AddressType,AmendedOn,AmendedBy,BuildingNumber")
        Dim vRow As CDBDataRow = vTable.AddRow
        With vRow
          .Item(1) = CStr(AddressNumber)
          .Item(2) = HouseName
          .Item(3) = AddressText
          .Item(4) = Town
          .Item(5) = County
          .Item(6) = Postcode
          .Item(7) = Country
          .Item(8) = CountryDescription
          .Item(9) = CStr(UK)
          .Item(10) = BranchName
          .Item(11) = AddressLine()
          .Item(12) = mvClassFields.Item(AddressFields.AddressType).Value
          .Item(13) = AmendedOn
          .Item(14) = AmendedBy
          .Item(15) = BuildingNumber
        End With
        Return vTable
      End Get
    End Property

    Public Function IsPotentialDuplicate(ByVal pAddress As Address) As Boolean
      Dim vFirstAddressLine1 As String
      Dim vFirstAddressLine2 As String
      Dim vHouseName1 As String
      Dim vHouseName2 As String
      ' BR 11347
      Dim vAddress1Country As String
      Dim vAddress2Country As String
      Dim vPostCode1 As String
      Dim vPostCode2 As String

      ' BR 11347
      ' Need to decide what country the addresses are from
      vAddress1Country = mvClassFields.Item(AddressFields.Country).Value
      vAddress2Country = pAddress.Country
      If vAddress1Country = "NL" And vAddress2Country = "NL" Then
        ' We are testing against postcode and building number
        vPostCode1 = Trim(mvClassFields.Item(AddressFields.Postcode).Value)
        vPostCode2 = Trim(pAddress.Postcode)
        If StrComp(vPostCode1, vPostCode2, CompareMethod.Text) = 0 Then
          ' First bit OK, what about the building number
          If StrComp(Trim(mvClassFields.Item(AddressFields.BuildingNumber).Value), pAddress.BuildingNumber, CompareMethod.Text) = 0 Then
            IsPotentialDuplicate = True
          End If
        End If
      Else
        ' As normal
        vFirstAddressLine1 = Trim(FirstLine(mvClassFields.Item(AddressFields.Address).Value))
        vFirstAddressLine2 = Trim(FirstLine(pAddress.AddressText))
        If StrComp(vFirstAddressLine1, vFirstAddressLine2, CompareMethod.Text) = 0 Then
          IsPotentialDuplicate = True
        Else
          vHouseName1 = Trim(mvClassFields.Item(AddressFields.HouseName).Value)
          If vHouseName1.Length > 0 Then
            vHouseName2 = pAddress.HouseName.Trim
            If InStr(1, vHouseName2, vHouseName1, CompareMethod.Text) > 0 Then IsPotentialDuplicate = True
          End If
        End If
      End If
    End Function

    Public Sub InitFromValues(ByVal pAddressNumber As Integer, ByVal pAddressType As AddressTypes, ByVal pHouseName As String, ByVal pAddress As String, ByVal pTown As String, ByVal pCounty As String, ByVal pCountry As String, ByVal pPostCode As String, ByVal pSortCode As String, ByVal pMosaicCode As String, ByVal pPaf As String, Optional ByVal pBuildingNumber As String = "")
      InitClassFields()
      If pAddressType = AddressTypes.ataContact Then mvClassFields.Item(AddressFields.AddressType).Value = "C" Else mvClassFields.Item(AddressFields.AddressType).Value = "O"
      mvClassFields.Item(AddressFields.AddressNumber).IntegerValue = pAddressNumber
      mvClassFields.Item(AddressFields.HouseName).Value = pHouseName
      SetAddressField(pAddress)
      mvClassFields.Item(AddressFields.Town).Value = pTown
      mvClassFields.Item(AddressFields.County).Value = pCounty
      mvClassFields.Item(AddressFields.Country).Value = pCountry
      mvClassFields.Item(AddressFields.Postcode).Value = pPostCode
      mvClassFields.Item(AddressFields.Sortcode).Value = pSortCode
      mvClassFields.Item(AddressFields.MosaicCode).Value = pMosaicCode
      mvClassFields.Item(AddressFields.Paf).Value = pPaf

      If pBuildingNumber.Length > 0 Then mvClassFields.Item(AddressFields.BuildingNumber).Value = pBuildingNumber
      mvAddressBlank = mvClassFields.Item(AddressFields.Address).Value = " "
    End Sub

    Public Enum AddressFoundStatus
      afsExactMatch = 1
      afsPotentialMatch
      afsNoMatch
    End Enum

    Public Sub DedupAddressByPostCode(ByVal pPostCode As String, ByVal pInitials As String, ByVal pForenames As String, ByVal pSurname As String, ByVal pAddressLine1 As String, ByVal pDedupForenameInitials As Boolean, ByRef pContact As Contact) ', pContact2 As Contact, pJointCont As Contact, pOrg As Organisation)
      'Dedup address just using the postcode
      Dim vContact As Contact
      Dim vFound As Boolean
      Dim vLine As String
      Dim vPos As Integer

      mvClassFields.Item(AddressFields.Postcode).Value = pPostCode
      Dim vAddressesForPostcode As IEnumerable(Of Address) = Me.GetRelatedList(Of Address)({AddressFields.Postcode})
      Dim vAddress As Address = Nothing
      For Each vAddress In vAddressesForPostcode
        SetContactsAtAddress(vAddress.AddressNumber)
        For Each vContact In ContactsAtAddress
          If StrComp(vContact.Surname, pSurname, CompareMethod.Text) = 0 Then
            'Surname matches
            If pForenames.HasValue Then
              If pDedupForenameInitials And (StrComp(vContact.Forenames, pForenames, CompareMethod.Text) = 0 And StrComp(vContact.Initials, pInitials, CompareMethod.Text) = 0) Then
                'Forenames & initials match
                pContact = vContact
                vFound = True
              ElseIf StrComp(vContact.Forenames, pForenames, CompareMethod.Text) = 0 Then
                'Forenames match
                pContact = vContact
                vFound = True
              ElseIf StrComp(vContact.Initials, pInitials, CompareMethod.Text) = 0 Then
                'Initials match
                pContact = vContact
                vFound = True
              End If
            Else
              If StrComp(vContact.Initials, pInitials, CompareMethod.Text) = 0 Then
                'Initials match
                pContact = vContact
                vFound = True
              End If
            End If
          End If
          If vFound Then Exit For
        Next vContact 'vContact
        If vFound Then Exit For
      Next vAddress 'vAddress

      If Not vFound Then
        'No contacts found so see if any addresses match
        For Each vAddress In vAddressesForPostcode
          vPos = InStr(vAddress.AddressText, vbLf)
          If vPos > 0 Then
            vLine = Left(vAddress.AddressText, vPos - 1)
          Else
            vLine = vAddress.AddressText
          End If
          If StrComp(vLine, pAddressLine1, CompareMethod.Text) = 0 Then
            'First line of address matches
            vFound = True
          End If
          If vFound Then Exit For
        Next vAddress
      End If
      If vFound And vAddress IsNot Nothing Then
        InitFromValues(vAddress.AddressNumber,
                       vAddress.AddressType,
                       vAddress.HouseName,
                       vAddress.AddressText,
                       vAddress.Town,
                       vAddress.County,
                       vAddress.Country,
                       vAddress.Postcode,
                       vAddress.Sortcode,
                       vAddress.MosaicCode,
                       vAddress.Paf)
      End If
    End Sub

    Public Function DedupAddress(ByVal pCapitals As Boolean, ByVal pDedupTitle As Boolean, ByVal pDedupForenameInitials As Boolean, ByRef pAddLines() As String, ByVal pContactEntry As Boolean, ByVal pJoint As Boolean, ByRef pContact As Contact, ByRef pContact2 As Contact, ByRef pJointCont As Contact, ByRef pOrg As Organisation, ByVal pMaxLenTown As Integer, ByVal pMaxLenCounty As Integer, Optional ByRef pOrgAddressPotDup As Boolean = False, Optional ByRef pEmployeeLoad As Boolean = False, Optional ByRef pDedupAddressOnly As Boolean = False, Optional ByVal pTown As String = "", Optional ByVal pCounty As String = "", Optional ByVal pCountry As String = "", Optional ByVal pCountryDesc As String = "", Optional ByVal pPostCode As String = "") As AddressFoundStatus
      'Deduplicate the address where the postcode is missing
      Dim vRS As CDBRecordSet
      Dim vRS2 As CDBRecordSet
      Dim vHouseName As String = ""
      Dim vTown As String
      Dim vCounty As String
      Dim vCountry As String
      Dim vPostcode As String
      Dim vFound As Integer
      Dim vCountryDesc As String
      Dim vAddress As String
      Dim vUniservMail As Boolean
      Dim vSortCode As String = ""
      Dim vAddFound As AddressFoundStatus
      Dim vConFound As AddressFoundStatus
      Dim vOrgFound As AddressFoundStatus
      Dim vLen As Integer
      Dim vAddressNo As Integer
      Dim vUseAddressNo As Integer
      Dim vPotentialAddNo As Integer
      Dim vBranch As String
      Dim vDone As Boolean
      Dim vItem As String

      Dim vAddLines(6) As String

      vUniservMail = (Len(mvEnv.GetConfig("uniserv_mail")) > 0)
      SetupDictionaries()
      pAddLines.CopyTo(vAddLines, 0)
      For vIndex As Integer = 0 To 6
        If vAddLines(vIndex) Is Nothing Then vAddLines(vIndex) = ""
      Next
      vTown = pTown
      vCounty = pCounty
      vCountry = pCountry
      vCountryDesc = pCountryDesc
      vPostcode = Postcode
      If pPostCode.Length > 0 Then vPostcode = pPostCode

      If pContact2 Is Nothing Then pContact2 = New Contact(mvEnv)
      If pJointCont Is Nothing Then pJointCont = New Contact(mvEnv)

      '---------------------------------------------------------------------------------------------
      'If the country is not yet found look for a country description
      'If none found and there is an address 5/6 line try getting the country desc from the address
      '---------------------------------------------------------------------------------------------
      If Len(vCountry) = 0 Then
        vFound = 0
        If vCountryDesc.Length > 0 Then
          If pCapitals Then vCountryDesc = UCase(vCountryDesc)
          vCountryDesc = RemoveTextNoise(vCountryDesc)
        End If
        If vCountryDesc.Length = 0 Then
          If vCounty.Length > 0 Then 'If there is a county it might be a country
            If mvCountriesDescDict.ContainsKey(vCounty) Then
              vCountryDesc = mvCountriesDescDict(vCounty).ToString
              vCountry = vCounty
              vCounty = ""
              vFound = -1
            End If
          End If
          If vFound = 0 Then
            If vAddLines(6).Length > 0 Then
              vCountryDesc = vAddLines(6)
              vFound = 6
            ElseIf vAddLines(5).Length > 0 Then
              vCountryDesc = vAddLines(5)
              vFound = 5
            ElseIf vAddLines(4).Length > 0 Then
              vCountryDesc = vAddLines(4)
              vFound = 4
            ElseIf vAddLines(3).Length > 0 Then
              vCountryDesc = vAddLines(3)
              vFound = 3
            ElseIf vAddLines(2).Length > 0 Then
              vCountryDesc = vAddLines(2)
              vFound = 2
            End If
          End If
        End If

        If vCountryDesc.Length > 0 Then 'Handle common country entries
          Select Case UCase(vCountryDesc)
            Case "UK", "U K", "GB", "G B"
              vCountry = "UK"
            Case "USA", "U S A"
              vCountry = "USA"
            Case "THE NETHERLANDS", "HOLLAND"
              vCountry = "NL"
            Case "EIRE"
              vCountry = "IRL"
            Case "WEST GERMANY"
              vCountry = "D"
            Case Else
              If vFound > 0 Then 'ie its from the address line fields
                If vPostcode.Length = 0 Then
                  If vCountryDesc.Length < 9 And ContainsNumbers(vCountryDesc) Then
                    If (InStr(vCountryDesc, " ") > 0 And (InStr(vCountryDesc, " ") < 6 And Mid(vCountryDesc, 1, 4) <> "BFPO")) Or Mid(vCountryDesc, 1, 4) = "BFPO" Then
                      vPostcode = UCase(Trim(vCountryDesc))
                      vAddLines(vFound) = "" ' clear the address line used
                      vFound = vFound - 1
                      vCountryDesc = vAddLines(vFound)
                    End If
                  End If
                End If
              End If
              If mvCountriesDict.ContainsKey(vCountryDesc) Then
                vCountry = mvCountriesDict(vCountryDesc).ToString
              Else
                If InStr(vCountryDesc, " ") > 0 Then
                  vItem = FirstWord(vCountryDesc)
                  If mvCountriesDict.ContainsKey(vItem) Then
                    vCountry = mvCountriesDict(vItem).ToString
                    If vFound > 0 Then
                      vAddLines(vFound) = Trim(Mid(vCountryDesc, Len(vItem) + 1))
                      vFound = 0
                    End If
                  Else
                    If vFound > 0 Then 'It's from the address fields just use the default country
                      vFound = 0
                    Else
                      RaiseEvent DedupLogMessage("Invalid Country Description '" & vCountryDesc & "' Default '" & mvEnv.DefaultCountry & "' Used")
                      vFound = 0
                    End If
                    vCountryDesc = ""
                  End If
                Else
                  'Invalid country
                  RaiseEvent DedupLogMessage("Invalid Country Description '" & vCountryDesc & "' Default '" & mvEnv.DefaultCountry & "' Used")
                  vFound = 0
                  vCountryDesc = ""
                End If
              End If
          End Select
          If vFound > 0 Then vAddLines(vFound) = "" 'Clear the address line used
        End If
        If Len(vCountry) = 0 Then vCountry = mvEnv.DefaultCountry
      End If

      '-----------------------------------
      ' Now deal with the town and county
      '-----------------------------------
      If Len(vTown) = 0 Then
        'No Town given - It must be extracted from the address
        'Look for UK counties first
        If vCountry = "UK" And Len(vCounty) = 0 Then
          If vAddLines(6).Length > 0 Then
            If IsCounty(vAddLines(6)) Then
              vCounty = vAddLines(6)
              vAddLines(6) = ""
            End If
          ElseIf vAddLines(5).Length > 0 Then
            If IsCounty(vAddLines(5)) Then
              vCounty = vAddLines(5)
              vAddLines(5) = ""
            End If
          ElseIf vAddLines(4).Length > 0 Then
            If IsCounty(vAddLines(4)) Then
              vCounty = vAddLines(4)
              vAddLines(4) = ""
            End If
          ElseIf vAddLines(3).Length > 0 Then
            If IsCounty(vAddLines(3)) Then
              vCounty = vAddLines(3)
              vAddLines(3) = ""
            End If
          ElseIf vAddLines(2).Length > 0 Then
            If IsCounty(vAddLines(2)) Then
              vCounty = vAddLines(2)
              vAddLines(2) = ""
            End If
          End If
        End If
        If vTown.Length = 0 Then 'Town still not found so get it
          If vAddLines(6).Length > 0 Then
            vTown = vAddLines(6)
            vAddLines(6) = ""
          ElseIf vAddLines(5).Length > 0 Then
            vTown = vAddLines(5)
            vAddLines(5) = ""
          ElseIf vAddLines(4).Length > 0 Then
            vTown = vAddLines(4)
            vAddLines(4) = ""
          ElseIf vAddLines(3).Length > 0 Then
            vTown = vAddLines(3)
            vAddLines(3) = ""
          ElseIf vAddLines(2).Length > 0 Then
            vTown = vAddLines(2)
            vAddLines(2) = ""
          End If
        End If
      End If
      vAddress = vAddLines(1)
      If vAddLines(2).Length > 0 Then vAddress = vAddress & Chr(10) & vAddLines(2)
      If vAddLines(3).Length > 0 Then vAddress = vAddress & Chr(10) & vAddLines(3)
      If vAddLines(4).Length > 0 Then vAddress = vAddress & Chr(10) & vAddLines(4)
      If vAddLines(5).Length > 0 Then vAddress = vAddress & Chr(10) & vAddLines(5)
      If vAddLines(6).Length > 0 Then vAddress = vAddress & Chr(10) & vAddLines(6)

      If Len(vTown) = 0 And Len(vCounty) > 0 Then
        vTown = vCounty
        vCounty = ""
      End If

      If Len(vTown) = 0 And Len(vAddress) = 0 Then
        RaiseEvent DedupError("Town and/or Address Line 1")
      ElseIf Len(vAddress) = 0 Then
        RaiseEvent DedupError("Address Line 1")
      ElseIf Len(vTown) = 0 Then
        RaiseEvent DedupError("Town")
      Else
        If vUniservMail And vCountry <> "UK" Then
          'UniServ active and not UK address - Town stays as is
          If pCapitals Then vTown = UCase(vTown)
        Else
          'Capitalise Town
          vTown = UCase(vTown)
        End If
        If Len(vTown) > pMaxLenTown Then
          RaiseEvent DedupLogMessage("Town: Field truncated from '" & vTown & "' to '" & Left(vTown, pMaxLenTown) & "'")
          vTown = Left(vTown, pMaxLenTown)
        End If
        If Len(vCounty) > pMaxLenCounty Then
          RaiseEvent DedupLogMessage("County: Field truncated from '" & vCounty & "' to '" & Left(vCounty, pMaxLenCounty) & "'")
          vCounty = Left(vCounty, pMaxLenCounty)
        End If
      End If
      If vCountry = "UK" And Len(vPostcode) > 0 Then vSortCode = GetMailsortCode(vPostcode, Nothing, Nothing, Nothing, vTown)
      'Populate the classfields
      InitFromValues(0, AddressTypes.ataContact, vHouseName, vAddress, vTown, vCounty, vCountry, vPostcode, vSortCode, "", "")

      'Find an existing address
      If Country = "UK" And Len(Postcode) > 0 Then
        'If UK and postcode search by country and postcode
        If Not pContactEntry Then
          vRS = mvEnv.Connection.GetRecordSet("SELECT address_type,address_number,house_name,address,branch FROM addresses WHERE country = 'UK' AND postcode = '" & Postcode & "' AND address_type = 'O'")
        Else
          vRS = mvEnv.Connection.GetRecordSet("SELECT address_type,address_number,house_name,address,branch FROM addresses WHERE country = 'UK' AND postcode = '" & Postcode & "'")
        End If
      Else
        'Search by country, town and address
        If Not pContactEntry Then
          vRS = mvEnv.Connection.GetRecordSet("SELECT address_type,address_number,house_name,address,branch FROM addresses WHERE country = '" & Country & "' AND address_type = 'O' AND town" & mvEnv.Connection.SQLLiteral("=", CDBField.FieldTypes.cftCharacter, Town))
        Else
          vRS = mvEnv.Connection.GetRecordSet("SELECT address_type,address_number,house_name,address,branch FROM addresses WHERE country = '" & Country & "' AND town" & mvEnv.Connection.SQLLiteral("=", CDBField.FieldTypes.cftCharacter, Town))
        End If
      End If

      'Now, lets find contact(s) / organisations for this address
      vLen = vAddLines(1).Length
      vAddFound = AddressFoundStatus.afsNoMatch
      vConFound = AddressFoundStatus.afsNoMatch
      vOrgFound = AddressFoundStatus.afsNoMatch
      While vRS.Fetch And vAddFound = AddressFoundStatus.afsNoMatch
        If StrComp(Left(vRS.Fields("address").Value, vLen), vAddLines(1), CompareMethod.Text) = 0 Then
          'First Line of address matches
          vAddressNo = vRS.Fields("address_number").LongValue
          If vRS.Fields("address_type").Value = "C" Then 'Changed from IIF to resolve .NET issue!
            mvClassFields(AddressFields.AddressType).Value = "C"
          Else
            mvClassFields(AddressFields.AddressType).Value = "O"
          End If
          vBranch = vRS.Fields("branch").Value
          If (Len(vRS.Fields("house_name").Value) = 0 And Len(HouseName) = 0) Or (StrComp(vRS.Fields("house_name").Value, HouseName, CompareMethod.Text) = 0) Then
            'If housenames match then use this address
            If vAddFound <> AddressFoundStatus.afsExactMatch Then
              vUseAddressNo = vAddressNo
              vAddFound = AddressFoundStatus.afsExactMatch
            End If
          End If
          If AddressType = AddressTypes.ataOrganisation Then
            'If organisation name check for duplicate - if duplicate flag existing org
            If Len(pOrg.Name) > 0 And vOrgFound = AddressFoundStatus.afsNoMatch Then
              vRS2 = mvEnv.Connection.GetRecordSet("SELECT name,o.organisation_number FROM organisation_addresses oa, organisations o WHERE oa.address_number  = " & CStr(vAddressNo) & " AND oa.organisation_number = o.organisation_number")
              vDone = False
              While vRS2.Fetch And vDone = False
                If StrComp(pOrg.Name, vRS2.Fields("name").Value, CompareMethod.Text) = 0 Then
                  vDone = True
                  pOrg.Init(vRS2.Fields("organisation_number").LongValue)
                  vOrgFound = AddressFoundStatus.afsExactMatch
                  vUseAddressNo = vAddressNo
                  vAddFound = AddressFoundStatus.afsExactMatch
                End If
              End While
              vRS2.CloseRecordSet()
              If vOrgFound = AddressFoundStatus.afsNoMatch Then
                'Org Names do not match
                If pOrgAddressPotDup Then
                  If vPotentialAddNo = 0 Then vPotentialAddNo = vAddressNo
                End If
                vAddressNo = 0 'Don't do any more checks on this address
                vUseAddressNo = 0
                vAddFound = AddressFoundStatus.afsNoMatch
              End If
            End If
          End If

          'If a joint contact then check if exists already If found we don't need to add either individual
          If pJoint And vAddressNo > 0 Then
            vRS2 = mvEnv.Connection.GetRecordSet("SELECT surname, c.contact_number FROM contact_addresses ca, contacts c WHERE ca.address_number  = " & CStr(vAddressNo) & " AND ca.contact_number = c.contact_number AND contact_type = 'J'")
            vDone = False
            While vRS2.Fetch And vDone = False
              If StrComp(pJointCont.Surname, vRS2.Fields("surname").Value, CompareMethod.Text) = 0 Then
                vDone = True
                pContact.Init(vRS2.Fields("contact_number").LongValue, vAddressNo)
                vConFound = AddressFoundStatus.afsExactMatch
                vUseAddressNo = vAddressNo
                vAddFound = AddressFoundStatus.afsExactMatch
              End If
            End While
            vRS2.CloseRecordSet()
          End If

          If vAddFound <> AddressFoundStatus.afsNoMatch And vAddressNo > 0 And vConFound = AddressFoundStatus.afsNoMatch Then
            'Joint not found or individual
            If Len(pContact.Surname) > 0 Then
              'Look for individuals at this address
              vRS2 = mvEnv.Connection.GetRecordSet("SELECT surname, title, forenames, initials, c.contact_number FROM contact_addresses ca, contacts c WHERE ca.address_number  = " & CStr(vAddressNo) & " AND ca.contact_number = c.contact_number AND contact_type = 'C'")
              vDone = False
              While vRS2.Fetch And vDone = False
                If StrComp(pContact.Surname, vRS2.Fields("surname").Value, CompareMethod.Text) = 0 Then
                  If (Not pDedupTitle) Or (pDedupTitle And StrComp(pContact.TitleName, vRS2.Fields("title").Value, CompareMethod.Text) = 0) Or (pDedupTitle And Len(pContact.TitleName) = 0 Or Len(vRS2.Fields("title").Value) = 0) Then
                    If (Not pDedupForenameInitials) Or pContact.ForenameAndOrInitialsMatch(vRS2) Then 'ForenameAndOrInitialsMatch(vContact1, vRecSet2) Then
                      'We have found contact 1
                      If pJoint Then
                        pContact.Init(vRS2.Fields("contact_number").LongValue, vAddressNo)
                        If pContact2.ContactNumber > 0 Then vDone = True
                      Else
                        'Individual
                        pContact.Init(vRS2.Fields("contact_number").LongValue, vAddressNo)
                        vConFound = AddressFoundStatus.afsExactMatch
                        vDone = True
                      End If
                    End If
                  End If
                End If
                If pJoint Then
                  If StrComp(pContact2.Surname, vRS2.Fields("surname").Value, CompareMethod.Text) = 0 Then
                    If (Not pDedupTitle) Or (pDedupTitle And StrComp(pContact.TitleName, vRS2.Fields("title").Value, CompareMethod.Text) = 0) Then
                      If (Not pDedupForenameInitials) Or pContact2.ForenameAndOrInitialsMatch(vRS2) Then 'ForenameAndOrInitialsMatch(vContact2, vRecSet2) Then
                        'We have found contact 2
                        pContact2.Init((vRS2.Fields("contact_number").LongValue), vAddressNo)
                        If pContact.ContactNumber > 0 Then vDone = True
                      End If
                    End If
                  End If
                End If
                If vDone Then
                  vUseAddressNo = vAddressNo
                  vAddFound = AddressFoundStatus.afsExactMatch
                End If
              End While
              vRS2.CloseRecordSet()
            End If
          End If
          If vAddFound = AddressFoundStatus.afsNoMatch Then vAddressNo = 0
        End If
      End While
      vRS.CloseRecordSet()

      If vAddFound = AddressFoundStatus.afsNoMatch And vPotentialAddNo > 0 Then
        'Mark this as a potential duplicate
        vRS = mvEnv.Connection.GetRecordSet("SELECT name,o.organisation_number FROM organisation_addresses oa, organisations o WHERE oa.address_number  = " & CStr(vPotentialAddNo) & " AND oa.organisation_number = o.organisation_number")
        If vRS.Fetch Then
          pOrg.Init(vRS.Fields("organisation_number").LongValue)
          vOrgFound = AddressFoundStatus.afsPotentialMatch
          vUseAddressNo = vPotentialAddNo
          vAddFound = AddressFoundStatus.afsPotentialMatch
        End If
        vRS.CloseRecordSet()
      End If
      If vAddFound = AddressFoundStatus.afsNoMatch And Len(pOrg.Name) > 0 And vOrgFound = AddressFoundStatus.afsNoMatch Then
        'If not found address and loading an Organsiation, check if Org Name exists
        'Is this a new site for an existing organisation?
        vRS = mvEnv.Connection.GetRecordSet("SELECT name,organisation_number FROM organisations WHERE name " & mvEnv.Connection.DBLike(pOrg.Name, CDBField.FieldTypes.cftUnicode))
        vDone = False
        While vRS.Fetch And vDone = False
          If StrComp(pOrg.Name, vRS.Fields("name").Value, CompareMethod.Text) = 0 Then
            vDone = True
            pOrg.Init(vRS.Fields("organisation_number").LongValue)
            vOrgFound = AddressFoundStatus.afsExactMatch
          End If
        End While
        vRS.CloseRecordSet()
      End If

      If (Not pEmployeeLoad) And (pContactEntry And vConFound = AddressFoundStatus.afsNoMatch And vOrgFound <> AddressFoundStatus.afsNoMatch And vAddFound <> AddressFoundStatus.afsNoMatch) Then
        'Unless we are employee loading, if we are contact loading and not found the contact yet
        'but found the org and address check the site for any matching contacts
        If Len(pContact.Surname) > 0 Then
          'Look for individuals at this address
          vRS = mvEnv.Connection.GetRecordSet("SELECT surname, title, forenames, initials, c.contact_number FROM contact_addresses ca, contacts c WHERE ca.address_number  = " & CStr(vAddressNo) & " AND ca.contact_number = c.contact_number AND contact_type = 'C'")
          vDone = False
          While vRS.Fetch And vDone = False
            If StrComp(pContact.Surname, vRS.Fields("surname").Value, CompareMethod.Text) = 0 Then
              If StrComp(pContact.TitleName, vRS.Fields("title").Value, CompareMethod.Text) = 0 Then
                If (Not pDedupForenameInitials) Or (pContact.ForenameAndOrInitialsMatch(vRS)) Then
                  'We have found contact 1
                  If pJoint Then
                    pContact.Init(vRS.Fields("contact_number").LongValue)
                    If pContact2.ContactNumber > 0 Then vDone = True
                  Else
                    'Individual
                    pContact.Init(vRS.Fields("contact_number").LongValue)
                    vConFound = AddressFoundStatus.afsExactMatch
                    vDone = True
                  End If
                End If
              End If
            End If
            If pJoint Then
              If StrComp(pContact2.Surname, vRS.Fields("surname").Value, CompareMethod.Text) = 0 Then
                If StrComp(pContact2.TitleName, vRS.Fields("title").Value, CompareMethod.Text) = 0 Then
                  If (Not pDedupForenameInitials) Or (pContact2.ForenameAndOrInitialsMatch(vRS)) Then
                    'We have found contact 2
                    pContact2.Init(vRS.Fields("contact_number").LongValue)
                    If pContact.ContactNumber > 0 Then vDone = True
                  End If
                End If
              End If
            End If
          End While
          vRS.CloseRecordSet()
        End If
      End If
      If vUseAddressNo > 0 Then vAddressNo = vUseAddressNo
      If pDedupAddressOnly And vAddFound <> AddressFoundStatus.afsNoMatch Then
        'Only want the address de-dup, remove any exact match of Contacts/Orgs
        If pContactEntry Then
          If vConFound <> AddressFoundStatus.afsNoMatch Then
            'Contact at Org found - Force load the contact
            ' or Contact at an address
            vConFound = AddressFoundStatus.afsNoMatch
            pContact.Init()
            pContact2.Init()
            pJointCont.Init()
          End If
        Else
          If vOrgFound <> AddressFoundStatus.afsNoMatch Then
            pOrg.Init()
            vOrgFound = AddressFoundStatus.afsNoMatch
            vAddressNo = 0 'Can't use an existing address record for a new Organisation
            vAddFound = AddressFoundStatus.afsNoMatch
          End If
        End If
      End If
      If vAddFound = AddressFoundStatus.afsExactMatch And vAddressNo > 0 Then mvClassFields.Item(AddressFields.AddressNumber).IntegerValue = vAddressNo
    End Function

    Private Function IsCounty(ByVal pName As String) As Boolean
      'Copied from DataImportContactOrOrg
      Dim vFound As Boolean
      Dim vPos As Integer

      vPos = InStr(pName, ".")
      If vPos > 0 Then Mid(pName, vPos, 1) = " "
      If mvCountyDict.Contains(pName) Then vFound = True
      If Not vFound Then
        vPos = InStr(pName, " ")
        'See if just the second word is a valid county - i.e. pName = West Sussex, County = Sussex
        If vPos > 0 Then vFound = mvCountyDict.Contains(Right(pName, Len(pName) - vPos))
      End If
      If Not vFound Then
        If UCase(Right(pName, 5)) = "SHIRE" Then vFound = True
      End If
      Return vFound
    End Function

    Private mvCountriesDict As Hashtable
    Private mvCountriesDescDict As Hashtable
    Private mvCountyDict As StringList

    Private Sub SetupDictionaries()
      'Copied from DataImport ContactOrg
      'First set up the countries
      mvCountriesDescDict = New Hashtable
      mvCountriesDict = New Hashtable
      Dim vSQL As String = "SELECT country, country_desc FROM countries"
      Dim vRS As CDBRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
      Dim vKey As String = ""
      While vRS.Fetch
        Try
          vKey = vRS.Fields(2).Value
          mvCountriesDict.Add(vKey, vRS.Fields(1).Value) 'Find country by country
          vKey = vRS.Fields(1).Value
          mvCountriesDescDict.Add(vKey, vRS.Fields(2).Value) 'Find country by country_desc
        Catch ex As Exception
          RaiseError(DataAccessErrors.daeDuplicateValue, vKey, "Countries")
        End Try
      End While
      vRS.CloseRecordSet()
      'Now set up the counties
      mvCountyDict = New StringList( _
        "Anglesey,Angus,Argyll,Avon,Beds,Berks,Borders,Bucks,Cambs,Canary," & _
        "Channel,Cleveland,Clwyd,Clwys,Co,Cornwall,County,Cumbria,Devon,Dorset," & _
        "Dublin,Dyfed,E,East,Essex,Fife,Glamorgan,Glos,Grampian,Greater," & _
        "Guernsey,Gwent,Hants,Herts,Humberside,Isle,Isles,Jersey,Kent,Lancs," & _
        "Leics,Lincs,Lothian,Merseyside,Mid,Middlesex,Middx,Midlothian,N," & _
        "Norfolk,North,Northants,Northern,Northumberland,Notts,Orkney,Oxon,Pembs,Powys," & _
        "Rutland,S,Scotland,Shetland,Somerset,South,Staffs,Strathclyde,Suffolk,Surrey," & _
        "Sussex,Tameside,Tyne,Vale,W,Wales,Warks,West,Wilts,Worcs")
    End Sub

    Sub DoAddressMerge(ByVal pJob As JobSchedule, ByVal pConn As CDBConnection, ByVal pOContact As Integer, ByVal pDContact As Integer, ByVal pDAddress As Integer, ByVal pDelete As Boolean, ByVal pUpdateDates As Boolean)
      Dim vHistoryTables As String
      Dim vArray() As String
      Dim vTable As String
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vValidFrom As String
      Dim vValidTo As String
      Dim vContactPos As ContactPosition
      Dim vDT As CDBDataTable
      Dim vHistorical As Boolean
      Dim vDupContAddr As ContactAddress
      Dim vOrganisationTables As String 'Access using Organisation_Number
      Dim vOrganisationTableList As List(Of String)

      Dim vDupAddress As Address = New Address(Me.Environment)
      vDupAddress.Init(pDAddress)

      'vTables = vTables & "bankers_orders,batch_transactions,caf_voucher_transactions,contact_mailings,"
      'vTables = vTables & "communications,contact_positions,covenants,credit_card_authorities,direct_debits,"
      'vTables = vTables & "members,orders,order_details,receipts,subscriptions,thank_you_letters,"
      'vTables = vTables & "back_orders,back_order_details,batch_transaction_analysis,contact_header,"
      'vTables = vTables & "credit_customers,credit_sales,despatch_notes,invoices,proforma_invoices,proforma_invoices2,"
      'vTables = vTables & "proforma_invoice_details,selected_contacts"
      'vArray = Split(vTables, ",")

      vHistoryTables = "communications_log,communications_log_links,contact_addresses,contact_exports,contact_mailings,"
      vHistoryTables = vHistoryTables & "contacts,financial_history,gift_aid_donations,contact_mailing_documents"
      vOrganisationTables = "organisation_address_usages,organisation_addresses"
      vHistoryTables = vHistoryTables & "," & vOrganisationTables

      GetAddressMergeInfo(pConn, vHistoryTables)
      For Each vMergeInfo As AddressMergeInfo In mvMergeInfo
        '    If vArray(vIndex) = "proforma_invoices2" Then
        '      vTable = "proforma_invoices"
        '      vConAttr = "order_contact_number"
        '      vAddAttr = "order_address_number"
        '    Else
        '      vTable = vArray(vIndex)
        '      vConAttr = "contact_number"
        '      vAddAttr = "address_number"
        '    End If

        vTable = vMergeInfo.TableName
        Debug.Print(vTable)
        ChangeAddress(pJob, vTable, pOContact, pDContact, AddressNumber, pDAddress, vMergeInfo.ContactAttr, vMergeInfo.AddressAttr)
        If Me.AddressType = AddressTypes.ataOrganisation Then
          'Some tables uses Organisation_Number as Contact_Number e.g. batch_transactions. The following line will deal with these
          ChangeAddress(pJob, vTable, Me.OrganisationNumber, Me.OrganisationNumber, AddressNumber, pDAddress, vMergeInfo.ContactAttr, vMergeInfo.AddressAttr)
        End If
      Next
      If pOContact = pDContact And pUpdateDates Then
        vDT = New CDBDataTable
        vDT.FillFromSQLDONOTUSE(mvEnv, "SELECT address_number, valid_from, valid_to FROM contact_addresses WHERE address_number IN (" & pDAddress & "," & AddressNumber & ") AND contact_number = " & pOContact)
        If vDT.Rows.Count() < 2 Then RaiseError(DataAccessErrors.daeExpectedDataMissing) 'If the DataTable doesn't contain 2 rows then raise an error
        Dim vValidFrom1 As String = vDT.Rows.Item(0).Item("valid_from")
        Dim vValidFrom2 As String = vDT.Rows.Item(1).Item("valid_from")
        'Determine the Valid From
        vValidFrom = ""
        If vValidFrom1.Length > 0 And vValidFrom2.Length > 0 Then
          'If both addresses have Valid From set then use the earlier of the two
          If CDate(vValidFrom1) < CDate(vValidFrom2) Then
            vValidFrom = vValidFrom1
          Else
            vValidFrom = vValidFrom2
          End If
        ElseIf vValidFrom1.Length > 0 Then
          'If only one address has Valid From set then use that as long as it's a past date
          If CDate(vValidFrom1) < Date.Today Then vValidFrom = vValidFrom1
        ElseIf vValidFrom2.Length > 0 Then
          'If only one address has Valid From set then use that as long as it's a past date
          If CDate(vValidFrom2) < Date.Today Then vValidFrom = vValidFrom2
        End If
        Dim vValidTo1 As String = vDT.Rows.Item(0).Item("valid_to")
        Dim vValidTo2 As String = vDT.Rows.Item(1).Item("valid_to")
        'Determine the Valid To
        vValidTo = ""
        If vValidTo1.Length > 0 And vValidTo2.Length > 0 Then
          'If both addresses have Valid To set then use the later of the two
          If CDate(vValidTo1) > CDate(vValidTo2) Then
            vValidTo = vValidTo1
          Else
            vValidTo = vValidTo2
          End If
        End If
        'Determine the new setting of the Historical flag
        vHistorical = False
        If vValidTo.Length > 0 Then
          If CDate(vValidTo) < CDate(TodaysDate()) Then
            'ValidTo less then Today so it is historical
            vHistorical = True
          End If
        End If
        'Update this address with the Valid From/To dates
        vUpdateFields.Clear()
        vUpdateFields.Add("valid_from", CDBField.FieldTypes.cftDate, vValidFrom)
        vUpdateFields.Add("valid_to", CDBField.FieldTypes.cftDate, vValidTo)
        vUpdateFields.Add("historical", BooleanString(vHistorical))
        vWhereFields.Clear()
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pOContact)
        vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, AddressNumber)
        pConn.UpdateRecords("contact_addresses", vUpdateFields, vWhereFields, False)
      End If
      If AddressType = AddressTypes.ataOrganisation Then
        vContactPos = New ContactPosition(mvEnv)
        vContactPos.Init(pDContact, AddressNumber, OrganisationNumber, "", "", "", ContactPosition.CurrentSettingTypes.cstCurrent)
        If Not vContactPos.Existing Then
          vContactPos.Create(pDContact, AddressNumber, OrganisationNumber, "N", "Y", "", (TodaysDate()))
          vContactPos.Save()
        End If
      End If
      If pDelete Then
        'vTables = "communications_log,communications_log_links,contact_addresses,contact_exports,contact_mailings,"
        'vTables = vTables & "contacts,financial_history,gift_aid_donations"

        vArray = Split(vHistoryTables, ",")
        vOrganisationTableList = Split(vOrganisationTables, ",").ToList
        For Each vHistoryTable As String In vArray
          If Not vOrganisationTableList.Contains(vHistoryTable) Then
            ChangeAddress(pJob, vHistoryTable, pOContact, pDContact, AddressNumber, pDAddress, "contact_number", "address_number")
            If Me.AddressType = AddressTypes.ataOrganisation Then
              'Some tables uses Organisation_Number as Contact_Number e.g. financial_history. The following line will deal with these
              ChangeAddress(pJob, vHistoryTable, Me.OrganisationNumber, Me.OrganisationNumber, AddressNumber, pDAddress, "contact_number", "address_number")
            End If
          Else
            If Me.AddressType = AddressTypes.ataOrganisation Then
              'The Addresses will be for the same Organisation
              ChangeAddress(pJob, vHistoryTable, Me.OrganisationNumber, Me.OrganisationNumber, AddressNumber, pDAddress, "organisation_number", "address_number")
            End If
          End If
        Next
        If Not pJob Is Nothing Then pJob.InfoMessage = ProjectText.String31252 'Deleting Duplicate Address
        mvEnv.Connection.StartTransaction()
        vWhereFields.Clear()
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pDContact)
        vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, pDAddress)
        mvEnv.Connection.DeleteRecords("contact_address_usages", vWhereFields, False)
        If vDupAddress IsNot Nothing AndAlso vDupAddress.AddressNumber.Equals(pDAddress) Then
          vDupAddress.Delete(Me.Environment.User.UserID, True)
        End If
        mvEnv.Connection.CommitTransaction()
      Else
        If Not pJob Is Nothing Then pJob.InfoMessage = ProjectText.String31253 'Updating Duplicate Address to historical
        vDupContAddr = New ContactAddress(mvEnv)
        vDupContAddr.InitFromContactAndAddress(mvEnv, ContactAddress.ContactAddresssLinkTypes.caltContact, pDContact, pDAddress)
        vWhereFields.Clear()
        vUpdateFields.Clear()
        vUpdateFields.Add("historical", CDBField.FieldTypes.cftCharacter, "Y")
        'If ValidTo is null or dated in the future then set it to TodaysDate
        If Len(vDupContAddr.ValidTo) = 0 Then
          vUpdateFields.Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate)
        Else
          If CDate(vDupContAddr.ValidTo) > CDate(TodaysDate()) Then vUpdateFields.Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate)
        End If
        vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, pDAddress)
        mvEnv.Connection.UpdateRecords("contact_addresses", vUpdateFields, vWhereFields)

        'If setting duplicate address to historic and this is default address
        'Update contacts to use the other address as the default
        If pOContact = pDContact Then
          vWhereFields.Clear()
          vUpdateFields.Clear()
          vUpdateFields.Add("address_number", CDBField.FieldTypes.cftLong, AddressNumber)
          vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pOContact)
          vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, pDAddress)
          mvEnv.Connection.UpdateRecords("contacts", vUpdateFields, vWhereFields, False)
        End If
      End If
    End Sub

    Private Sub ChangeAddress(ByVal pJob As JobSchedule, ByVal pTable As String, ByVal pConTo As Integer, ByVal pConFrom As Integer, ByVal pAddTo As Integer, ByVal pAddFrom As Integer, ByVal pConAttr As String, ByVal pAddAttr As String)
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields

      If Not pJob Is Nothing Then pJob.InfoMessage = String.Format(ProjectText.String31251, ProperName(pTable)) 'Transferring: %s
      mvEnv.Connection.StartTransaction()
      vUpdateFields.Add(pAddAttr, CDBField.FieldTypes.cftLong, pAddTo)
      If Len(pConAttr) > 0 Then vWhereFields.Add(pConAttr, CDBField.FieldTypes.cftLong, pConFrom)
      vWhereFields.Add(pAddAttr, CDBField.FieldTypes.cftLong, pAddFrom)
      mvEnv.Connection.UpdateRecords(pTable, vUpdateFields, vWhereFields, False)
      'Delete any that are left behind - Why should any be left behind?
      mvEnv.Connection.DeleteRecords(pTable, vWhereFields, False)
      mvEnv.Connection.CommitTransaction()
    End Sub

    Private Sub GetAddressMergeInfo(ByVal pConn As CDBConnection, ByRef pHistoryTables As String)
      Dim vIgnoreTables As String
      Dim vTable As String
      Dim vAttr As String
      Dim vIgnore As Boolean
      Dim vRecordSet As CDBRecordSet
      Dim vMergeInfo As AddressMergeInfo

      If Not mvMergeInfoValid Then
        vIgnoreTables = "," & pHistoryTables & ",organisations,organisation_groups,addresses,addresses_report,contact_groups"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCollections) Then vIgnoreTables = vIgnoreTables & ",collection_points,unmanned_collections"
        vIgnoreTables = vIgnoreTables & ","
        vAttr = "address_number"
        'Create an array element for each instance of address number
        mvMergeInfo = New List(Of AddressMergeInfo)
        vRecordSet = pConn.GetRecordSet("SELECT table_name,primary_key FROM maintenance_attributes WHERE attribute_name = '" & vAttr & "' AND is_base_table_attribute = 'Y' ORDER BY table_name")
        While vRecordSet.Fetch()
          vTable = vRecordSet.Fields(1).Value
          vIgnore = False
          If Left(vTable, 5) = "temp_" Then vIgnore = True
          If Left(vTable, 4) = "ext_" Then vIgnore = True
          If InStr(vIgnoreTables, "," & vTable & ",") > 0 Then vIgnore = True
          If Not vIgnore Then
            vMergeInfo = New AddressMergeInfo
            vMergeInfo.TableName = vTable
            Select Case vTable
              Case "branch_history", "address_data"
                vMergeInfo.TableName = vTable
                vMergeInfo.ContactAttr = ""
                vMergeInfo.AddressAttr = "address_number"
              Case "proforma_invoices2"
                vMergeInfo.TableName = "proforma_invoices"
                vMergeInfo.ContactAttr = "order_contact_number"
                vMergeInfo.AddressAttr = "order_address_number"
              Case Else
                vMergeInfo.TableName = vTable
                vMergeInfo.ContactAttr = "contact_number"
                vMergeInfo.AddressAttr = "address_number"
            End Select
            mvMergeInfo.Add(vMergeInfo)
          End If
        End While
        vRecordSet.CloseRecordSet()
        vMergeInfo = New AddressMergeInfo
        vMergeInfo.TableName = "batch_transactions"
        vMergeInfo.ContactAttr = "mailing_contact_number"
        vMergeInfo.AddressAttr = "mailing_address_number"
        mvMergeInfo.Add(vMergeInfo)
        vMergeInfo = New AddressMergeInfo
        vMergeInfo.TableName = "caf_voucher_transactions"
        vMergeInfo.ContactAttr = "mailing_contact_number"
        vMergeInfo.AddressAttr = "mailing_address_number"
        mvMergeInfo.Add(vMergeInfo)
        vMergeInfo = New AddressMergeInfo
        vMergeInfo.TableName = "service_bookings"
        vMergeInfo.ContactAttr = "booking_contact_number"
        vMergeInfo.AddressAttr = "booking_address_number"
        mvMergeInfo.Add(vMergeInfo)
        mvMergeInfoValid = True
      End If
    End Sub

    Public Function ValidatePostcode(ByVal pUpdatePAF As Boolean, ByVal pUpdateAddress As Boolean, Optional ByVal pAddGridReference As Boolean = False, Optional ByRef pAddressChanged As Boolean = False, Optional ByVal pGetOrgName As Boolean = False, Optional ByVal pUpdateAdditionalData As Boolean = False) As Postcoder.ValidatePostcodeStatuses
      Dim vAddressString As String
      Dim vTown As String
      Dim vCounty As String
      Dim vPostcode As String
      Dim vStart As Integer
      Dim vIndex As Integer
      Dim vPos As Integer
      Dim vBuildingNumber As Boolean
      Dim vBuilding As String = ""
      Dim vStatus As Postcoder.ValidatePostcodeStatuses
      Dim vPaf As String
      Dim vDPS As String

      'Get the current address details
      vAddressString = Replace(mvClassFields.Item(AddressFields.Address).Value, vbCr, "")
      vTown = Town
      vCounty = County
      vPostcode = Postcode
      vDPS = DeliveryPointSuffix

      'Process to address attribute into 4 address lines
      Dim vAddress() As String = {"", "", "", ""} 'Initialise this with blank values to not have errors in  mvEnv.Postcoder.PostcodeAddress-PostcodeAddressGB
      vIndex = 0
      vStart = 1
      Do
        vPos = InStr(vStart, vAddressString, Chr(10))
        If vPos > 0 Then
          vAddress(vIndex) = Trim(Mid(vAddressString, vStart, vPos - vStart))
        Else
          vAddress(vIndex) = Trim(Mid(vAddressString, vStart))
        End If
        If Len(vAddress(vIndex)) > 0 Then vIndex = vIndex + 1
        vStart = vPos + 1
      Loop While vPos > 0 And vIndex < 4

      If (vPostcode.Length = 0) And mvEnv.Postcoder.PostcodeIfNoPostcodeGiven(Country) And mvEnv.Postcoder.QuickAddressType <> Postcoder.QuickAddressTypes.qatProOnDemand Then 'No Postcode so postcode the address
        vStatus = mvEnv.Postcoder.PostcodeAddress(vAddress, vTown, vCounty, vPostcode, Country, True)
      Else 'Got a postcode so try to validate the building
        If IsBuildingNumberCountry() Then
          If Len(BuildingNumber) > 0 Then
            vBuilding = BuildingNumber
            vBuildingNumber = True
          Else
            vBuilding = vAddress(0)
            vBuildingNumber = False
          End If
        Else
          PostcoderAddress.GetBuildingNumber(vAddress, vBuilding, vBuildingNumber)
        End If
        vStatus = mvEnv.Postcoder.ValidateBuilding(vBuildingNumber, vBuilding, vAddress, vTown, vCounty, vPostcode, Country)
        If (vStatus = Postcoder.ValidatePostcodeStatuses.vpsBuildingNotValidated) And mvEnv.Postcoder.PostcodeIfNoPostcodeGiven(Country) Then 'If we could not validate building use VP call to validate postcode
          vStatus = mvEnv.Postcoder.ValidatePostcode(vAddress, vTown, vCounty, vPostcode, Country)
        End If
      End If

      Select Case vStatus
        Case Postcoder.ValidatePostcodeStatuses.vpsAddressPostcoded, Postcoder.ValidatePostcodeStatuses.vpsAddressRePostcoded
          vPaf = "VB"
        Case Postcoder.ValidatePostcodeStatuses.vpsBuildingValidated
          vPaf = "VB"
        Case Postcoder.ValidatePostcodeStatuses.vpsPostcodeValidated
          vPaf = "VP"
        Case Postcoder.ValidatePostcodeStatuses.vpsQASBatchReportCode
          vPaf = Left(mvEnv.Postcoder.QASBatchReportCode, 2)
        Case Else
          vPaf = ""
      End Select
      If Len(vPaf) > 0 Then
        mvClassFields.Item(AddressFields.Paf).Value = vPaf
        SetAdditionalData(CStr(mvEnv.Postcoder.DPS), IntegerValue(mvEnv.Postcoder.Easting), IntegerValue(mvEnv.Postcoder.Northing), CStr(mvEnv.Postcoder.LEACode), CStr(mvEnv.Postcoder.LEAName))
        If vStatus = Postcoder.ValidatePostcodeStatuses.vpsAddressPostcoded Or vStatus = Postcoder.ValidatePostcodeStatuses.vpsAddressRePostcoded Then
          mvClassFields.Item(AddressFields.Postcode).Value = mvEnv.Postcoder.Postcode
          If mvClassFields.Item(AddressFields.Postcode).ValueChanged Then
            pAddressChanged = True
          Else
            vStatus = Postcoder.ValidatePostcodeStatuses.vpsBuildingValidated 'If we didn't actually change the postcode then just say VB
          End If
        End If
        If pUpdatePAF And Not pUpdateAddress Then Save(mvEnv.User.UserID)
      End If

      'If an address was returned see if it matches what we have got
      With mvEnv.Postcoder
        If .CheckAddress Then
          If Replace(.Address(pGetOrgName), vbCr, "") <> Replace(AddressText, vbCr, "") Then
            If Not (.Address(pGetOrgName) = "") And AddressText <> "" Then
              SetAddressField(.Address(pGetOrgName))
              pAddressChanged = True
            End If
          End If
          If .Town <> Town Then
            mvClassFields.Item(AddressFields.Town).Value = .Town
            pAddressChanged = True
          End If
          If .County <> County Then
            mvClassFields.Item(AddressFields.County).Value = .County
            pAddressChanged = True
          End If
          If .BuildingNumber <> BuildingNumber Then
            mvClassFields.Item(AddressFields.BuildingNumber).Value = .BuildingNumber
            pAddressChanged = True
          End If
          If .Postcode <> Postcode Then
            mvClassFields.Item(AddressFields.Postcode).Value = .Postcode
            pAddressChanged = True
          End If
        End If
      End With

      If Len(vPaf) > 0 Then
        If pAddGridReference And Len(mvEnv.Postcoder.Easting) > 0 And Len(mvEnv.Postcoder.Northing) > 0 Then CreateGridReference(IntegerValue(mvEnv.Postcoder.Easting), IntegerValue(mvEnv.Postcoder.Northing))
        If pUpdateAdditionalData And AddressNumber > 0 Then UpdateAdditionalData(True, mvEnv.Postcoder.LEACode, mvEnv.Postcoder.LEAName)
        If pUpdatePAF And pUpdateAddress Then
          Save(mvEnv.User.UserID)
        ElseIf pUpdateAdditionalData And AddressNumber > 0 Then
          mvClassFields.SetSaved()
          mvClassFields.Item(AddressFields.DeliveryPointSuffix).SetValueOnly = vDPS
          mvClassFields.Item(AddressFields.DeliveryPointSuffix).Value = mvEnv.Postcoder.DPS
          Save(mvEnv.User.UserID)
        End If
      End If
      ValidatePostcode = vStatus
    End Function

    Public Sub SetAdditionalData(ByVal pDPS As String, ByVal pEasting As Integer, ByVal pNorthing As Integer, ByVal pLeaCode As String, ByVal pLeaName As String)
      mvClassFields(AddressFields.DeliveryPointSuffix).Value = pDPS
      mvEasting = pEasting
      mvNorthing = pNorthing
      mvLEACode = pLeaCode
      mvLEAName = pLeaName
    End Sub

    Public Sub UpdateAdditionalData(ByVal pCheckExisting As Boolean, ByVal pLeaCode As String, ByVal pLeaName As String)
      Dim vInsertFields As New CDBFields
      Dim vAD As New AddressData

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAddressDPS) And mvEnv.GetConfigOption("qas_lea_data") Then
        If Len(pLeaCode) > 0 And Len(pLeaName) > 0 Then
          If pCheckExisting Then
            vAD.Init(mvEnv, AddressNumber)
          Else
            vAD.Init(mvEnv)
          End If
          If Not vAD.Existing Then vAD.Create(AddressNumber)
          vAD.SetLEAData(pLeaCode, pLeaName)
          vAD.Save()
        ElseIf pCheckExisting Then
          vAD.Init(mvEnv, AddressNumber)
          If vAD.Existing Then vAD.Delete()
        End If
      End If
    End Sub

    Public Sub CreateGridReference(ByVal pEasting As Integer, ByVal pNorthing As Integer)
      Dim vInsertFields As New CDBFields

      If pEasting > 0 And pNorthing > 0 Then
        If (mvEnv.GetConfigOption("opt_cd_create_grid_references") = True OrElse mvEnv.GetConfigOption("qas_delivery_point_suffix")) And Len(mvClassFields.Item(AddressFields.Postcode).Value) > 0 Then
          vInsertFields.Add("postcode", CDBField.FieldTypes.cftCharacter, mvClassFields.Item(AddressFields.Postcode).Value)
          vInsertFields.Add("easting", CDBField.FieldTypes.cftLong, pEasting)
          vInsertFields.Add("northing", CDBField.FieldTypes.cftLong, pNorthing)
          mvEnv.Connection.InsertRecord("postcode_grid_references", vInsertFields, True)
        End If
      End If
    End Sub

    Public Sub SetAmended(ByVal pAmendedOn As String, ByVal pAmendedBy As String)
      mvClassFields.Item(AddressFields.AmendedOn).Value = pAmendedOn
      mvClassFields.Item(AddressFields.AmendedBy).Value = pAmendedBy
      mvOverrideAmended = True
    End Sub

    Public Shared Function ValidatePostcodeFormat(ByVal pPostCode As String) As Boolean
      Dim vValid As Boolean
      Dim vPos As Integer
      Dim vNextChar As String

      vValid = True
      If Left(pPostCode, 1) < "A" Or Left(pPostCode, 1) > "Z" Then
        vValid = False
      Else
        If Len(pPostCode) > 4 Then
          vPos = InStr(pPostCode, " ")
          If vPos < 3 Or vPos > 5 Then
            vValid = False
          Else
            If Left(pPostCode, 4) = "BFPO" Then
              vNextChar = Mid(pPostCode, 6, 1)
              If (vNextChar < "1" Or vNextChar > "9") And vNextChar <> "S" Then vValid = False
            Else
              If Len(pPostCode) <> vPos + 3 Then
                vValid = False
              Else
                vNextChar = Mid(pPostCode, vPos + 1, 1)
                If (vNextChar < "0" Or vNextChar > "9") Then
                  vValid = False
                Else
                  vNextChar = Mid(pPostCode, vPos + 2, 1)
                  If vNextChar < "A" Or vNextChar > "Z" Then vValid = False
                  vNextChar = Mid(pPostCode, vPos + 3, 1)
                  If vNextChar < "A" Or vNextChar > "Z" Then vValid = False
                End If
              End If
            End If
          End If
        End If
      End If
      Return vValid
    End Function

    Public Sub SetAddressNumber(ByVal pAddressNumber As Integer, ByVal pAddressType As AddressTypes, ByVal pBranch As String)
      If pAddressType = AddressTypes.ataContact Then mvClassFields.Item(AddressFields.AddressType).Value = "C" Else mvClassFields.Item(AddressFields.AddressType).Value = "O"
      mvClassFields.Item(AddressFields.AddressNumber).IntegerValue = pAddressNumber
      mvClassFields.Item(AddressFields.Branch).Value = pBranch
    End Sub

    Public Sub SetCountryDescription(ByVal pCountryDescription As String, ByVal pUK As Boolean)
      mvCountryDescription = pCountryDescription
      mvUK = pUK
      mvCountryValid = True
    End Sub

    Public Sub PopulateAddressLines()
      'This method should only be used from the upgrade process as it forces the
      'Address line attributes to think they exist in the database
      'The upgrade process will already have checked this
      mvClassFields(AddressFields.AddressLine1).InDatabase = True
      mvClassFields(AddressFields.AddressLine2).InDatabase = True
      mvClassFields(AddressFields.AddressLine3).InDatabase = True
      mvClassFields(AddressFields.AddressLine4).InDatabase = True
      SetAddressField(mvClassFields(AddressFields.Address).Value)
    End Sub

  End Class
End Namespace
