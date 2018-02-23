Namespace Access
  Public Class TempBanksDataLoad

    Public Enum TempBanksDataLoadRecordSetTypes 'These are bit values
      tbdlrtAll = &HFFFFS
      'ADD additional recordset types here
      tbdlrtNamesAndAddress = 1
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum TempBanksDataLoadFields
      tbdlfAll = 0
      tbdlfRecordType
      tbdlfBranchOrPlace
      tbdlfSortCode
      tbdlfAddressLine1
      tbdlfAddressLine2
      tbdlfAddressLine3
      tbdlfPostcode
      tbdlfTelephoneNumber
      tbdlfSpecialCodes
      tbdlfBranchTitle
      tbdlfBankName
      tbdlfTown
      tbdlfDateLastAmended
      tbdlfRepeatAddrInd
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvAddressString As String = ""

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "temp_banks_data_load"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("record_type", CDBField.FieldTypes.cftInteger)
          .Add("branch_or_place")
          .Add("tbdl_sort_code")
          .Add("address_line1")
          .Add("address_line2")
          .Add("address_line3")
          .Add("tbdl_postcode")
          .Add("telephone_number")
          .Add("special_codes")
          .Add("branch_title")
          .Add("bank_name")
          .Add("tbdl_town")
          .Add("date_last_amended", CDBField.FieldTypes.cftLong)
          .Add("repeat_addr_ind")
        End With

      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As TempBanksDataLoadFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByRef pRSType As TempBanksDataLoadRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = TempBanksDataLoadRecordSetTypes.tbdlrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "tbdl")
      Else
        If (pRSType And TempBanksDataLoadRecordSetTypes.tbdlrtNamesAndAddress) = TempBanksDataLoadRecordSetTypes.tbdlrtNamesAndAddress Then
          vFields = "tbdl_sort_code,bank_name,branch_title,address_line1,address_line2,address_line3,tbdl_town,tbdl_postcode"
        End If
      End If
      Return vFields
    End Function

    Public Sub Init(ByRef pEnv As CDBEnvironment)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub InitFromRecordSet(ByRef pEnv As CDBEnvironment, ByRef pRecordSet As CDBRecordSet, ByRef pRSType As TempBanksDataLoadRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And TempBanksDataLoadRecordSetTypes.tbdlrtAll) = TempBanksDataLoadRecordSetTypes.tbdlrtAll Then
          .SetItem(TempBanksDataLoadFields.tbdlfRecordType, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfBranchOrPlace, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfSortCode, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfAddressLine1, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfAddressLine2, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfAddressLine3, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfPostcode, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfTelephoneNumber, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfSpecialCodes, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfBranchTitle, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfBankName, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfTown, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfDateLastAmended, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfRepeatAddrInd, vFields)
        ElseIf (pRSType And TempBanksDataLoadRecordSetTypes.tbdlrtNamesAndAddress) = TempBanksDataLoadRecordSetTypes.tbdlrtNamesAndAddress Then
          .SetItem(TempBanksDataLoadFields.tbdlfSortCode, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfAddressLine1, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfAddressLine2, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfAddressLine3, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfPostcode, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfBranchTitle, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfBankName, vFields)
          .SetItem(TempBanksDataLoadFields.tbdlfTown, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(TempBanksDataLoadFields.tbdlfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub
    Public Function DifferentFromBank(ByVal pBank As Bank) As Boolean
      'Return True if details on Temp Data Load record differ from Bank details
      Dim vDifferent As Boolean

      If BankName <> pBank.BankName Or BranchTitle <> pBank.BranchName Or AddressString <> pBank.Address Or Town <> pBank.Town Or AddressLine3 <> pBank.County Or Postcode <> pBank.Postcode Then
        vDifferent = True
      End If
      DifferentFromBank = vDifferent
    End Function
    Public Sub SetBankFromLoadData(ByRef pBank As Bank, ByVal pExisting As Boolean)
      Dim vBranchName As String

      If Not pExisting Then pBank.SortCode = SortCode
      vBranchName = BranchTitle
      If Len(vBranchName) = 0 Then vBranchName = BankName
      pBank.BankName = BankName
      pBank.BranchName = vBranchName 'BranchTitle
      pBank.Address = AddressString
      pBank.Town = Town
      pBank.County = AddressLine3
      pBank.Postcode = Postcode
    End Sub
    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AddressLine1() As String
      Get
        AddressLine1 = mvClassFields.Item(TempBanksDataLoadFields.tbdlfAddressLine1).Value
      End Get
    End Property

    Public ReadOnly Property AddressLine2() As String
      Get
        AddressLine2 = mvClassFields.Item(TempBanksDataLoadFields.tbdlfAddressLine2).Value
      End Get
    End Property

    Public ReadOnly Property AddressLine3() As String
      Get
        AddressLine3 = Mid(mvClassFields.Item(TempBanksDataLoadFields.tbdlfAddressLine3).Value, 1, 35)
      End Get
    End Property

    Public ReadOnly Property AddressString() As String
      Get
        Dim vPosition As Integer
        Dim vLastAddressItem As String

        'Derive Address String from Address Fields if not set
        If mvAddressString = "" Then
          mvAddressString = AddressLine1 & AddressLine2
          mvAddressString = Trim(mvAddressString)
          If mvAddressString.Length > 0 Then
            'If Last character is comma, remove
            vPosition = InStrRev(mvAddressString, ",")
            If vPosition = Len(mvAddressString) Then mvAddressString = Mid(mvAddressString, 1, vPosition - 1)
            'If Last item in Address string is the Town, remove
            vPosition = InStrRev(mvAddressString, ",")
            vLastAddressItem = UCase(Trim(Mid(mvAddressString, vPosition + 1)))
            vLastAddressItem = Replace(vLastAddressItem, "-", " ")
            If vLastAddressItem = Town Then
              If vPosition > 0 Then
                mvAddressString = Mid(mvAddressString, 1, vPosition - 1)
              Else
                mvAddressString = ""
              End If
            End If
            'Finally, replace all commas with newlines
            mvAddressString = mvAddressString.Replace(",", vbNewLine)
          End If
        End If
        AddressString = mvAddressString
      End Get
    End Property

    Public ReadOnly Property BankName() As String
      Get
        BankName = Mid(mvClassFields.Item(TempBanksDataLoadFields.tbdlfBankName).Value, 1, 30)
      End Get
    End Property

    Public ReadOnly Property BranchOrPlace() As String
      Get
        BranchOrPlace = mvClassFields.Item(TempBanksDataLoadFields.tbdlfBranchOrPlace).Value
      End Get
    End Property

    Public ReadOnly Property BranchTitle() As String
      Get
        BranchTitle = Mid(mvClassFields.Item(TempBanksDataLoadFields.tbdlfBranchTitle).Value, 1, 30)
      End Get
    End Property

    Public ReadOnly Property DateLastAmended() As Integer
      Get
        DateLastAmended = mvClassFields.Item(TempBanksDataLoadFields.tbdlfDateLastAmended).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Postcode() As String
      Get
        Postcode = mvClassFields.Item(TempBanksDataLoadFields.tbdlfPostcode).Value
      End Get
    End Property

    Public ReadOnly Property RecordType() As Integer
      Get
        RecordType = mvClassFields.Item(TempBanksDataLoadFields.tbdlfRecordType).IntegerValue
      End Get
    End Property

    Public ReadOnly Property RepeatAddrInd() As String
      Get
        RepeatAddrInd = mvClassFields.Item(TempBanksDataLoadFields.tbdlfRepeatAddrInd).Value
      End Get
    End Property

    Public ReadOnly Property SortCode() As String
      Get
        SortCode = mvClassFields.Item(TempBanksDataLoadFields.tbdlfSortCode).Value
      End Get
    End Property

    Public ReadOnly Property SpecialCodes() As String
      Get
        SpecialCodes = mvClassFields.Item(TempBanksDataLoadFields.tbdlfSpecialCodes).Value
      End Get
    End Property

    Public ReadOnly Property TelephoneNumber() As String
      Get
        TelephoneNumber = mvClassFields.Item(TempBanksDataLoadFields.tbdlfTelephoneNumber).Value
      End Get
    End Property

    Public ReadOnly Property Town() As String
      Get
        Town = mvClassFields.Item(TempBanksDataLoadFields.tbdlfTown).Value
      End Get
    End Property
  End Class
End Namespace
