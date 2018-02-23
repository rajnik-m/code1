Namespace Access

  Public Class TableNote
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum TableNoteFields
      AllFields = 0
      TableName
      AdministratorNotes
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("table_name")
        .Add("administrator_notes", CDBField.FieldTypes.cftMemo)

        .Item(TableNoteFields.TableName).PrimaryKey = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "tn"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "table_notes"
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
    Public ReadOnly Property TableName() As String
      Get
        Return mvClassFields(TableNoteFields.TableName).Value
      End Get
    End Property
    Public ReadOnly Property AdministratorNotes() As String
      Get
        Return mvClassFields(TableNoteFields.AdministratorNotes).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(TableNoteFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(TableNoteFields.AmendedOn).Value
      End Get
    End Property
#End Region

  End Class
End Namespace