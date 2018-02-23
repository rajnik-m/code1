Namespace Access.BulkMailer
  ''' <summary>
  ''' An immutable class to provide statistics about a bulk mailer mailing.
  ''' </summary>
  Public Class BulkMailingStats

    Private mvSendDate As Date = Nothing
    Private mvNumberSent As Integer = 0
    Private mvBounced As Integer = 0
    Private mvOpened As Integer = 0
    Private mvClicked As Integer = 0

    ''' <summary>
    ''' Initializes a new instance of the <see cref="BulkMailingStats" /> class.
    ''' </summary>
    ''' <param name="pSendDate">The date that the emails were sent.</param>
    ''' <param name="pNumberSent">The number of emails sent.</param>
    ''' <param name="pBounced">The number of emails that were bounced.</param>
    ''' <param name="pOpened">The number of times the email was opened.</param>
    ''' <param name="pClicked">The number of click throughs.</param>
    Public Sub New(pSendDate As Date, pNumberSent As Integer, pBounced As Integer, pOpened As Integer, pClicked As Integer)
      mvSendDate = pSendDate
      mvNumberSent = pNumberSent
      mvBounced = pBounced
      mvOpened = pOpened
      mvClicked = pClicked
    End Sub

    ''' <summary>
    ''' The date that the emails were sent.
    ''' </summary>
    ''' <value>The email send date.</value>
    Public ReadOnly Property SendDate As Date
      Get
        Return mvSendDate
      End Get
    End Property

    ''' <summary>
    ''' The number of emails that were sent.
    ''' </summary>
    ''' <value>The total emails sent.</value>
    Public ReadOnly Property NumberSent As Integer
      Get
        Return mvNumberSent
      End Get
    End Property

    ''' <summary>
    ''' The number of emails that were bounced.
    ''' </summary>
    ''' <value>The bounce count.</value>
    Public ReadOnly Property Bounced As Integer
      Get
        Return mvBounced
      End Get
    End Property

    ''' <summary>
    ''' The number of times the email was opened.
    ''' </summary>
    ''' <value>The opened count.</value>
    Public ReadOnly Property Opened As Integer
      Get
        Return mvOpened
      End Get
    End Property

    ''' <summary>
    ''' The number of clicked throughs.
    ''' </summary>
    ''' <value>The click through count.</value>
    Public ReadOnly Property Clicked As Integer
      Get
        Return mvClicked
      End Get
    End Property
  End Class
End Namespace
