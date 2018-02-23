Namespace Access.BulkMailer

  ''' <summary>
  ''' A bulk mailer mailing
  ''' </summary>
  '''<remarks>A bulk mailer mailing is a template that has been set up in the bulk mailer 
  ''' to send to a mailing list.</remarks>
  Public Class BulkMailing

    Private mvMailingId As Integer
    Private mvMailing As String

    ''' <summary>
    ''' Initializes a new instance of the <see cref="BulkMailing" /> class.
    ''' </summary>
    ''' <param name="pMailingId">The mailing identifier.</param>
    ''' <param name="pMailing">The mailing name.</param>
    Public Sub New(pMailingId As Integer, pMailing As String)
      mvMailingId = pMailingId
      mvMailing = pMailing
    End Sub

    ''' <summary>
    ''' The identifier of the mailing.
    ''' </summary>
    Public ReadOnly Property MailingId As Integer
      Get
        Return mvMailingId
      End Get
    End Property

    ''' <summary>
    ''' The name of the mailing.
    ''' </summary>
    Public ReadOnly Property Mailing As String
      Get
        Return mvMailing
      End Get
    End Property
  End Class

End Namespace
