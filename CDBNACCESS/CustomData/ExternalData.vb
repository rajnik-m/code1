Namespace Access

  Public Class ExternalData
    Inherits CARERecord

    Protected Overrides Sub AddFields()
      'Should not be called
    End Sub

    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return mvClassFields.DatabaseTableName
      End Get
    End Property

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return mvClassFields.ContainsKey("amended_on") AndAlso mvClassFields.ContainsKey("amended_by")
      End Get
    End Property

    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return mvClassFields.TableAlias
      End Get
    End Property

    '--------------------------------------------------
    'Default constructor
    '--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    Public Overrides ReadOnly Property NeedsMaintenanceInfo() As Boolean
      Get
        Return True
      End Get
    End Property

    Public Overrides Function GetAddRecordMandatoryParameters() As String
      Return GetUniqueKeyFieldNames()
    End Function

  End Class

End Namespace
