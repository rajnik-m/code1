Namespace Access

  Public Class ContactIncentiveResponse
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum ContactIncentiveResponseFields
      AllFields = 0
      ContactNumber
      AddressNumber
      Source
      DateResponded
      DateFulfilled
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("address_number", CDBField.FieldTypes.cftLong)
        .Add("source")
        .Add("date_responded", CDBField.FieldTypes.cftDate)
        .Add("date_fulfilled", CDBField.FieldTypes.cftDate)
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return False
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "cir"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "contact_incentive_responses"
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
    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields(ContactIncentiveResponseFields.ContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AddressNumber() As Integer
      Get
        Return mvClassFields(ContactIncentiveResponseFields.AddressNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property Source() As String
      Get
        Return mvClassFields(ContactIncentiveResponseFields.Source).Value
      End Get
    End Property
    Public ReadOnly Property DateResponded() As String
      Get
        Return mvClassFields(ContactIncentiveResponseFields.DateResponded).Value
      End Get
    End Property
    Public ReadOnly Property DateFulfilled() As String
      Get
        Return mvClassFields(ContactIncentiveResponseFields.DateFulfilled).Value
      End Get
    End Property
#End Region

  End Class
End Namespace