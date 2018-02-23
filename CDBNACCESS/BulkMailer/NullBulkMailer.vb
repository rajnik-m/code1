Namespace Access.BulkMailer

  ''' <summary>
  ''' A dummy bulk mailer to satisfy the requirements of the system when no bulk
  ''' mailer is configured.
  ''' </summary>
  Public NotInheritable Class NullBulkMailer
    Inherits BulkMailer

    Public Sub New()
      MyBase.New(Nothing)
    End Sub
    ''' <summary>
    ''' The mailings available from this bulk mailer.
    ''' </summary>
    ''' <value>
    ''' A list of available mailings as <see cref="BulkMailing" /> instances.
    ''' </value>
    Protected Overrides ReadOnly Property AvailableMailings As System.Collections.Generic.List(Of BulkMailing)
      Get
        Return (New List(Of BulkMailing))
      End Get
    End Property

    ''' <summary>
    ''' A dummy mailing properties implmentation
    ''' </summary>
    ''' <param name="pId">The id of the mailing to get the properties of.</param>
    ''' <value>A <see cref="BulkMailingProperties" /> item for the mailing.</value>
    ''' <remarks>This method just throws a not supported exception.  Since <see cref="Mailings" /> is
    ''' always an empty list, this shouls never be called.</remarks>
    Public Overrides ReadOnly Property MailingProperties(ByVal pId As Integer) As BulkMailingProperties
      Get
        Throw New NotSupportedException()
      End Get
    End Property

    ''' <summary>
    ''' A dummy mail to list implementation.
    ''' </summary>
    ''' <param name="pContactsFilename">The name of a CSV file containing the contact data.</param>
    ''' <param name="pMailing">The ID of the mailing to send.</param>
    ''' <remarks>This method just throws a not supported exception.  Since <see cref="Mailings" /> is
    ''' always an empty list, this shouls never be called.</remarks>
    Protected Overrides Function MailToList(ByVal pContactsFilename As String, ByVal pMailing As Integer, ByVal pSendDate As Date, ByVal pMailingCode As String) As Integer
      Throw New NotSupportedException()
    End Function

    ''' <summary>
    ''' The statistics for a mailing.
    ''' </summary>
    ''' <value>A list of mailing statistics.</value>
    ''' <param name="MailingId">The ID of the mailing.</param>
    ''' <remarks>This method just throws a not supported exception.  Since <see cref="Mailings" /> is
    ''' always an empty list, this shouls never be called.</remarks>
    Public Overrides ReadOnly Property Statistics(ByVal MailingId As Integer) As BulkMailingStats
      Get
        Throw New NotSupportedException()
      End Get
    End Property

    ''' <summary>
    ''' The activity detail for a mailing.
    ''' </summary>
    ''' <value>
    ''' A list of <see cref="BulkMailerActvity" /> for the mailing
    ''' </value>
    ''' <param name="pMailingId">The ID of the mailing.</param>
    '''   <param name="pSince">The Earliest dtae and time to get activities for.</param>
    ''' <exception cref="System.NotSupportedException"></exception>
    Public Overrides ReadOnly Property ActivityDetail(ByVal pMailingId As Integer, ByVal pSince As Date) As List(Of BulkMailerActvity)
      Get
        Throw New NotSupportedException()
      End Get
    End Property

  End Class

End Namespace
