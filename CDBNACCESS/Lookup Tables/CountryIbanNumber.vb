Namespace Access

  Public Class CountryIbanNumber
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum CountryIbanNumberFields
      AllFields = 0
      IbanCountry
      IbanCountryDesc
      BankIdStartPosition
      BankIdLength
      IbanNumberLength
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("iban_country")
        .Add("iban_country_desc")
        .Add("bank_id_start_position", CDBField.FieldTypes.cftInteger)
        .Add("bank_id_length", CDBField.FieldTypes.cftInteger)
        .Add("iban_number_length", CDBField.FieldTypes.cftInteger)
      End With

      mvClassFields.Item(CountryIbanNumberFields.IbanCountry).PrimaryKey = True
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "cin"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "country_iban_numbers"
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
    Public ReadOnly Property IbanCountry() As String
      Get
        Return mvClassFields(CountryIbanNumberFields.IbanCountry).Value
      End Get
    End Property
    Public ReadOnly Property IbanCountryDesc() As String
      Get
        Return mvClassFields(CountryIbanNumberFields.IbanCountryDesc).Value
      End Get
    End Property
    Public ReadOnly Property BankIdStartPosition() As Integer
      Get
        Return mvClassFields(CountryIbanNumberFields.BankIdStartPosition).IntegerValue
      End Get
    End Property
    Public ReadOnly Property BankIdLength() As Integer
      Get
        Return mvClassFields(CountryIbanNumberFields.BankIdLength).IntegerValue
      End Get
    End Property
    Public ReadOnly Property IbanNumberLength() As Integer
      Get
        Return mvClassFields(CountryIbanNumberFields.IbanNumberLength).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(CountryIbanNumberFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(CountryIbanNumberFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"

    Private Const MAX_IBAN_NUMBER_LENGTH As Integer = 34

    ''' <summary>Gets the Bank Code from the IBAN Number according to the IBAN Country.</summary>
    ''' <returns>The Bank Code portion of the IBAN Number, or an empty string of the IBAN Number is for a different country / insufficient length</returns>
    Friend Function GetBankCodeFromIban(ByVal pIbanNumber As String) As String
      Dim vBankCode As String = String.Empty
      If pIbanNumber.Length > 2 Then
        If pIbanNumber.Substring(0, 2).ToUpper = IbanCountry.ToUpper Then
          vBankCode = Substring(pIbanNumber, BankIdStartPosition - 1, BankIdLength)
        End If
      End If
      Return vBankCode
    End Function

    ''' <summary>Validate the IBAN Number according to the MOD 97-10 rule (according to ISO 7064).</summary>
    ''' <param name="pIbanNumber">IBAN number to validate</param>
    ''' <returns>True if IBAN number is valid Else False</returns>
    ''' <remarks>Checks the IBAN number as following:
    ''' Make sure Length is > max length (34 chars)
    ''' Make sure Length is = length specified in country_iban_number
    ''' Make Sure the following:
    ''' Move the four initial characters to the end of the string.
    ''' Replace each letter in the string with two digits, thereby expanding the string, where A=10, B=11, �, Z=35.
    ''' Interpret the string as a decimal integer and apply MOD 97-10 (according to ISO 7064).  
    ''' For the check digits to be correct, the remainder after calculating the modulus 97 must be 1. 
    ''' If the remainder is 1, the check digit test is passed and the IBAN might be valid</remarks>
    Friend Function ValidateIbanNumber(ByVal pIbanNumber As String) As Boolean
      Dim vValid As Boolean = False
      Try
        If pIbanNumber.Length > 0 Then
          Dim vAdjustedIban As StringBuilder
          vValid = True   'Assume valid
          pIbanNumber = pIbanNumber.Trim 'Trim the spaces before starting

          If pIbanNumber.Contains("*") Then
            Return True   'Do not validate 
          End If

          If pIbanNumber.Length > MAX_IBAN_NUMBER_LENGTH OrElse pIbanNumber.Length <> IbanNumberLength Then
            vValid = False
          ElseIf Substring(pIbanNumber, 0, 2).ToUpper <> IbanCountry.ToUpper Then
            vValid = False
          End If

          If vValid Then
            'Move the starting 4 characters to the end of the IBAN Number
            vAdjustedIban = New StringBuilder(pIbanNumber.Substring(4))
            vAdjustedIban.Append(pIbanNumber.Substring(0, 4))

            'Now change all the alphabets to numbers in the IBANNumber provided
            For Each vElement As Char In vAdjustedIban.ToString
              If Not IsNumeric(vElement) Then
                Dim vNumber As String = GetCharacterValue(vElement)
                vAdjustedIban.Replace(vElement, vNumber)
              End If
            Next

            Dim vRemainder As Integer
            While (vAdjustedIban.ToString.Length >= 7)
              vRemainder = Integer.Parse(vRemainder.ToString + vAdjustedIban.ToString.Substring(0, 7)) Mod 97
              vAdjustedIban = New StringBuilder(vAdjustedIban.ToString.Substring(7))
            End While

            vRemainder = Integer.Parse(vRemainder.ToString + vAdjustedIban.ToString) Mod 97

            If Not (vRemainder = 1) Then vValid = False
          End If
        End If
      Catch vEX As Exception
        vValid = False
      End Try

      Return vValid

    End Function

    Private Function GetCharacterValue(ByVal pAlphabet As String) As String
      Dim vAlphaNumbers(,) As String = New String(,) {{"A", "10"}, {"B", "11"}, {"C", "12"}, {"D", "13"}, {"E", "14"}, {"F", "15"}, {"G", "16"}, {"H", "17"}, _
                                                      {"I", "18"}, {"J", "19"}, {"K", "20"}, {"L", "21"}, {"M", "22"}, {"N", "23"}, {"O", "24"}, {"P", "25"}, _
                                                      {"Q", "26"}, {"R", "27"}, {"S", "28"}, {"T", "29"}, {"U", "30"}, {"V", "31"}, {"W", "32"}, {"X", "33"}, {"Y", "34"}, {"Z", "35"}}


      ' Loop over all elements.
      For vKey As Integer = 0 To vAlphaNumbers.GetUpperBound(0)
        If vAlphaNumbers(vKey, 0) = pAlphabet Then Return vAlphaNumbers(vKey, 1)
      Next
      Return ""
    End Function

#End Region

  End Class
End Namespace