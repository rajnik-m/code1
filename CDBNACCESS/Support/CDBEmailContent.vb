''' <summary>
''' A class used to represent an archived email and it's attachemnts
''' </summary>
<Serializable()>
Public Class CDBEmailContent

  ''' <summary>
  ''' The body of the email.
  ''' </summary>
  ''' <value>
  ''' The email body content.
  ''' </value>
  Public Property Content As Byte()
  ''' <summary>
  ''' A value indicating whether the email body contains HTML.
  ''' </summary>
  ''' <value>
  ''' <c>true</c> if the body contains HTML; otherwise, <c>false</c>.
  ''' </value>
  Public Property IsBodyHtml As Boolean
  ''' <summary>
  ''' The attachments attached to this email.
  ''' </summary>
  ''' <value>
  ''' The attachments.
  ''' </value>
  Public Property Attachments As List(Of CDBEmailAttachment)

End Class
