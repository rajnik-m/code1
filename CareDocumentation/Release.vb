Imports CARE.Access
Imports CARE.Data

Public Class Release
  Inherits CARERecord

  Public Sub New(ByVal pEnv As CDBEnvironment)
    MyBase.New(pEnv)
  End Sub

  Protected Overrides Sub AddFields()
    With mvClassFields
      .Add("release")
      .Add("release_desc")
      .Add("destination_file_name")
      .Add("destination_dir")
      .Add("from_build", CDBField.FieldTypes.cftInteger)
      .Add("to_build", CDBField.FieldTypes.cftInteger)
      .Add("first_release_build", CDBField.FieldTypes.cftInteger)
    End With
  End Sub

  Protected Overrides ReadOnly Property DatabaseTableName() As String
    Get
      Return "releases"
    End Get
  End Property

  Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
    Get
      Return False
    End Get
  End Property

  Protected Overrides ReadOnly Property TableAlias() As String
    Get
      Return "r"
    End Get
  End Property

  
End Class
