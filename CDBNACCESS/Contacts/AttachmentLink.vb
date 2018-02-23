Namespace Access

  Public Class AttachmentLink
    Inherits CARERecord

    Private Enum AttachmentLinkFields
      AllFields = 0
      AttachmentLinkId
      AttachmentLinkTable
      AttachmentLinkForeignId
      AttachmentId
      AmendedBy
      AmendedOn
    End Enum

    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("attachment_link_id", CDBField.FieldTypes.cftInteger)
        .Add("attachment_link_table")
        .Add("attachment_link_foreign_id")
        .Add("attachment_id", CDBField.FieldTypes.cftInteger)
        .Item(AttachmentLinkFields.AttachmentLinkId).PrimaryKey = True
        .SetControlNumberField(AttachmentLinkFields.AttachmentLinkId, "ATL")
        .Item(AttachmentLinkFields.AttachmentLinkId).PrefixRequired = True
        .Item(AttachmentLinkFields.AttachmentLinkTable).PrefixRequired = True
        .Item(AttachmentLinkFields.AttachmentLinkForeignId).PrefixRequired = True
        .Item(AttachmentLinkFields.AttachmentId).PrefixRequired = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property

    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "al"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "attachment_links"
      End Get
    End Property

    Private Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    Public ReadOnly Property AttachmentLinkId() As Integer
      Get
        Return mvClassFields(AttachmentLinkFields.AttachmentLinkId).IntegerValue
      End Get
    End Property

    Public Property AttachmentLinkTable() As String
      Get
        Return mvClassFields(AttachmentLinkFields.AttachmentLinkTable).Value
      End Get
      Private Set(value As String)
        mvClassFields(AttachmentLinkFields.AttachmentLinkTable).Value = value
      End Set
    End Property

    Public Property AttachmentLinkForeignId() As String
      Get
        Return mvClassFields(AttachmentLinkFields.AttachmentLinkForeignId).Value
      End Get
      Private Set(value As String)
        mvClassFields(AttachmentLinkFields.AttachmentLinkForeignId).Value = value
      End Set
    End Property

    Public Property AttachmentId() As Integer
      Get
        Return mvClassFields(AttachmentLinkFields.AttachmentId).IntegerValue
      End Get
      Private Set(value As Integer)
        mvClassFields(AttachmentLinkFields.AttachmentId).IntegerValue = value
      End Set
    End Property

    Private mvAttachment As Attachment = Nothing
    Public ReadOnly Property Attachment As Attachment
      Get
        If mvAttachment Is Nothing Then
          mvAttachment = Attachment.GetInstance(mvEnv, Me.AttachmentId)
        End If
        Return mvAttachment
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(AttachmentLinkFields.AmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(AttachmentLinkFields.AmendedOn).Value
      End Get
    End Property

    Public Shared Function GetInstance(pEnv As CDBEnvironment, pId As Integer) As AttachmentLink
      Dim vInstance As New AttachmentLink(pEnv)
      vInstance.Init()
      vInstance.InitWithPrimaryKey(New CDBFields(New CDBField(vInstance.mvClassFields(AttachmentLinkFields.AttachmentLinkId).Name, pId)))
      If Not vInstance.Existing Then
        vInstance = Nothing
      End If
      Return vInstance
    End Function

    Public Shared Function CreateInstance(pEnv As CDBEnvironment, pAttachmentId As Integer, pLinkTable As String, pForeignId As Integer) As AttachmentLink
      Dim vInstance As New AttachmentLink(pEnv)
      vInstance.Init()
      vInstance.AttachmentId = pAttachmentId
      vInstance.AttachmentLinkTable = pLinkTable
      vInstance.AttachmentLinkForeignId = pForeignId.ToString
      Return vInstance
    End Function

    Public Shared Function CreateInstance(pEnv As CDBEnvironment, pAttachmentId As Integer, pLinkTable As String, pForeignId As String) As AttachmentLink
      Dim vInstance As New AttachmentLink(pEnv)
      vInstance.Init()
      vInstance.AttachmentId = pAttachmentId
      vInstance.AttachmentLinkTable = pLinkTable
      vInstance.AttachmentLinkForeignId = pForeignId
      Return vInstance
    End Function

    Public Overrides Sub Delete(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer)
      Dim vTransactionStarted As Boolean = mvEnv.Connection.StartTransaction
      Try
        Dim vAttachmentRows As DataTable = New SQLStatement(mvEnv.Connection,
                                                            "attachment_link_id",
                                                            "attachment_links",
                                                            New CDBField("attachment_id",
                                                                         Me.AttachmentId)).GetDataTable
        If vAttachmentRows.Rows.Count < 2 Then
          Attachment.GetInstance(mvEnv, Me.AttachmentId).Delete()
        End If
        MyBase.Delete(pAmendedBy, pAudit, pJournalNumber)
        If vTransactionStarted Then
          mvEnv.Connection.CommitTransaction()
        End If
      Catch ex As Exception
        If vTransactionStarted Then
          mvEnv.Connection.RollbackTransaction()
        End If
        Throw
      End Try
    End Sub

    Public Overrides Sub Update(pParameterList As CDBParameters)
      Throw New NotSupportedException("An attachment link cannot be updated")
    End Sub
  End Class
End Namespace
