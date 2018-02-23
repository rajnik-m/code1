Namespace Access

  ''' <summary>
  ''' An miscelaneous array of bytes that can be attached to something
  ''' </summary>
  Public Class Attachment
    Inherits CARERecord

    Private mvContent As Byte() = Nothing

    ''' <summary>
    ''' The fields in the databse table
    ''' </summary>
    Private Enum AttachmentFields
      AllFields = 0
      AttachmentId
      Name
      Document
      AmendedBy
      AmendedOn
    End Enum

    ''' <summary>
    ''' Adds the fields.
    ''' </summary>
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("attachment_id", CDBField.FieldTypes.cftInteger)
        .Add("name")
        .Add("document", CDBField.FieldTypes.cftMemo)
        .Item(AttachmentFields.AttachmentId).PrimaryKey = True
        .SetControlNumberField(AttachmentFields.AttachmentId, "ATT")
        .Item(AttachmentFields.AttachmentId).PrefixRequired = True
        .Item(AttachmentFields.Name).PrefixRequired = True
        .Item(AttachmentFields.Document).PrefixRequired = True
      End With
    End Sub

    ''' <summary>
    ''' Gets a value indicating whether the class supports amended on and by.
    ''' </summary>
    ''' <value>
    ''' <c>true</c> if the class supports amended on and by; otherwise, <c>false</c>.
    ''' </value>
    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property

    ''' <summary>
    ''' Gets the table alias.
    ''' </summary>
    ''' <value>
    ''' The table alias fior this table.
    ''' </value>
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "a"
      End Get
    End Property

    ''' <summary>
    ''' Gets the name of the database table.
    ''' </summary>
    ''' <value>
    ''' The name of the database table used to store instances of this class.
    ''' </value>
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "attachments"
      End Get
    End Property

    ''' <summary>
    ''' Initializes a new instance of the <see cref="Attachment"/> class.
    ''' </summary>
    ''' <param name="pEnv">The envuironmemt to use when creating the instance.</param>
    ''' <remarks>Use <see cref="GetInstance"/> or <see cref="CreateInstance"/> as appropriate.</remarks>
    Private Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    ''' <summary>
    ''' Gets the attachment identifier.
    ''' </summary>
    ''' <value>
    ''' The unique identifier of this attachment.
    ''' </value>
    Public ReadOnly Property AttachmentId() As Integer
      Get
        Return mvClassFields(AttachmentFields.AttachmentId).IntegerValue
      End Get
    End Property

    ''' <summary>
    ''' Gets or sets the name.
    ''' </summary>
    ''' <value>
    ''' The name of this attachment.
    ''' </value>
    Public Property Name() As String
      Get
        Return mvClassFields(AttachmentFields.Name).Value
      End Get
      Set(value As String)
        mvClassFields(AttachmentFields.Name).Value = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the binary content.
    ''' </summary>
    ''' <value>
    ''' The binary content of the attachment.
    ''' </value>
    Public Property Document() As Byte()
      Get
        If mvContent Is Nothing Then
          Dim vSql As New SQLStatement(mvEnv.Connection,
                                       mvClassFields(AttachmentFields.Document).Name,
                                       Me.DatabaseTableName,
                                       New CDBField(mvClassFields(AttachmentFields.AttachmentId).Name,
                                                    Me.AttachmentId,
                                                     CDBField.FieldWhereOperators.fwoEqual))
          mvContent = CType(vSql.GetDataTable.Rows(0)(0), Byte())
        End If
        Return mvContent
      End Get
      Set(value As Byte())
        mvContent = value
      End Set
    End Property

    ''' <summary>
    ''' Gets the amendeding user.
    ''' </summary>
    ''' <value>
    ''' The the user that last updated the attachment.
    ''' </value>
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(AttachmentFields.AmendedBy).Value
      End Get
    End Property

    ''' <summary>
    ''' Gets the amendment date.
    ''' </summary>
    ''' <value>
    ''' The date that the last amendment to this attachment was made.
    ''' </value>
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(AttachmentFields.AmendedOn).Value
      End Get
    End Property

    ''' <summary>
    ''' Gets an existing attachment from the database.
    ''' </summary>
    ''' <param name="pEnv">The environment to use.</param>
    ''' <param name="pId">The identifier of the attachment to get.</param>
    ''' <returns>The requested attachment or a null reference if the requested attachemnt doesn't exist</returns>
    Public Shared Function GetInstance(pEnv As CDBEnvironment, pId As Integer) As Attachment
      Dim vInstance As New Attachment(pEnv)
      vInstance.Init()
      vInstance.InitWithPrimaryKey(New CDBFields(New CDBField(vInstance.mvClassFields(AttachmentFields.AttachmentId).Name, pId)))
      If Not vInstance.Existing Then
        vInstance = Nothing
      End If
      Return vInstance
    End Function

    ''' <summary>
    ''' Creates a new attachment.
    ''' </summary>
    ''' <param name="pEnv">The environment to use.</param>
    ''' <param name="pName">Name of the attachment to create.</param>
    ''' <param name="pContent">The binary data for this attachment.</param>
    ''' <returns>The new attachnment</returns>
    ''' <remarks>The created attachemnt will not exist on the database until it has been saved.</remarks>
    Public Shared Function CreateInstance(pEnv As CDBEnvironment, pName As String, pContent As Byte()) As Attachment
      Dim vInstance As New Attachment(pEnv)
      vInstance.Init()
      vInstance.Name = pName
      vInstance.Document = pContent
      Return vInstance
    End Function

    ''' <summary>
    ''' Saves the attachment.
    ''' </summary>
    ''' <param name="pAmendedBy">The amending user.</param>
    ''' <param name="pAudit">if set to <c>true</c> the update will be auditted.</param>
    ''' <param name="pJournalNumber">The journal number.</param>
    Public Overrides Sub Save(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer)
      Dim vTransactionStarted As Boolean = mvEnv.Connection.StartTransaction
      Try
        MyBase.Save(pAmendedBy, pAudit, pJournalNumber)
        SaveContent()
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

    ''' <summary>
    ''' Save the attachment
    ''' </summary>
    ''' <param name="pAmendedBy">The amending user</param>
    ''' <param name="pAudit">if set to <c>true</c> the update will be auditted.</param>
    ''' <param name="pJournalNumber">The journal number.</param>
    ''' <param name="pForceAmendmentHistory">if set to <c>true</c> the creation of amendment history is forced</param>
    Public Overrides Sub Save(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer, pForceAmendmentHistory As Boolean)
      Dim vTransactionStarted As Boolean = mvEnv.Connection.StartTransaction
      Try
        MyBase.Save(pAmendedBy, pAudit, pJournalNumber, pForceAmendmentHistory)
        SaveContent()
        If vTransactionStarted Then
          mvEnv.Connection.CommitTransaction()
        End If
      Catch ex As Exception
        If vTransactionStarted Then
          mvEnv.Connection.RollbackTransaction()
        End If
      End Try
    End Sub

    ''' <summary>
    ''' Saves the binary content.
    ''' </summary>
    ''' <remarks>This must be called after the base class save to save the binary content as the base class cannot handle it</remarks>
    ''' <exception cref="System.InvalidOperationException">Attempt to set the Document property of an email attachment updated no records</exception>
    Private Sub SaveContent()
      Dim vSql As New StringBuilder
      Dim vParamName As String = String.Empty
      Dim vParameter As IDbDataParameter = mvEnv.Connection.GetDBParameterFromByteArray(mvClassFields(AttachmentFields.Document).Name,
                                                                                        mvContent,
                                                                                        vParamName)
      vSql.AppendLine("UPDATE " & Me.DatabaseTableName & " ")
      vSql.AppendLine("SET    " & mvClassFields(AttachmentFields.Document).Name & " = " & vParamName & " ")
      vSql.AppendLine("WHERE  " & mvClassFields(AttachmentFields.AttachmentId).Name & " = " & mvClassFields(AttachmentFields.AttachmentId).Value)
      Using vCommand As IDbCommand = mvEnv.Connection.CreateCommand
        vCommand.CommandType = CommandType.Text
        vCommand.CommandText = vSql.ToString
        vCommand.Parameters.Add(vParameter)
        If vCommand.ExecuteNonQuery() <> 1 Then
          Throw New InvalidOperationException("Attempt to set the Document property of an email attachment updated no records")
        End If
      End Using
    End Sub

    ''' <summary>
    ''' Block updates of attachments.
    ''' </summary>
    ''' <param name="pParameterList">The parameter list.</param>
    ''' <remarks>Attachments are immutable and therefore updates are not allowed.</remarks>
    ''' <exception cref="System.NotSupportedException">Attachemnts cannot be changed</exception>
    Public Overrides Sub Update(pParameterList As CDBParameters)
      Throw New NotSupportedException("Attachemnts cannot be changed")
    End Sub
  End Class
End Namespace