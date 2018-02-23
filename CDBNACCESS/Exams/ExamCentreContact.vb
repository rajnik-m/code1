Namespace Access

  Public Class ExamCentreContact
    Inherits CARERecord
    Implements IRecordCreate

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum ExamCentreContactFields
      AllFields = 0
      ExamCentreContactId
      ExamCentreId
      ContactNumber
      AddressNumber
      ExamContactType
      CreatedBy
      CreatedOn
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("exam_centre_contact_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_centre_id", CDBField.FieldTypes.cftInteger)
        .Add("contact_number", CDBField.FieldTypes.cftInteger)
        .Add("address_number", CDBField.FieldTypes.cftInteger)
        .Add("exam_contact_type")
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)

        .Item(ExamCentreContactFields.ExamCentreContactId).PrimaryKey = True
        .Item(ExamCentreContactFields.ExamCentreContactId).PrefixRequired = True
        .SetControlNumberField(ExamCentreContactFields.ExamCentreContactId, "XCC")

        .Item(ExamCentreContactFields.ExamContactType).PrefixRequired = True
        .Item(ExamCentreContactFields.CreatedBy).PrefixRequired = True
        .Item(ExamCentreContactFields.CreatedOn).PrefixRequired = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "ecc"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "exam_centre_contacts"
      End Get
    End Property

'--------------------------------------------------
'Default constructor
'--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

'--------------------------------------------------
'IRecordCreate
'--------------------------------------------------
    Public Function CreateInstance(ByVal pEnv As CDBEnvironment) As CARERecord Implements IRecordCreate.CreateInstance
      Return New ExamCentreContact(mvEnv)
    End Function
'--------------------------------------------------
'Public property procedures
'--------------------------------------------------
    Public ReadOnly Property ExamCentreContactId() As Integer
      Get
        Return mvClassFields(ExamCentreContactFields.ExamCentreContactId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamCentreId() As Integer
      Get
        Return mvClassFields(ExamCentreContactFields.ExamCentreId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields(ExamCentreContactFields.ContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AddressNumber() As Integer
      Get
        Return mvClassFields(ExamCentreContactFields.AddressNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamContactType() As String
      Get
        Return mvClassFields(ExamCentreContactFields.ExamContactType).Value
      End Get
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(ExamCentreContactFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(ExamCentreContactFields.CreatedOn).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ExamCentreContactFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ExamCentreContactFields.AmendedOn).Value
      End Get
    End Property
#End Region

  End Class
End Namespace