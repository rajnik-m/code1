Namespace Access

  Public Class OrganisationSuppression
    Inherits ContactSuppression

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      mvClassFields.Add("organisation_number", CDBField.FieldTypes.cftLong).PrimaryKey = True
      MyBase.AddFields()
      mvClassFields.RemoveAt(1)
    End Sub

    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "organisation_suppressions"
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

  End Class
End Namespace
