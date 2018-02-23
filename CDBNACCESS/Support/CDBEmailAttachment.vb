''' <summary>
''' A class to represent a miscellaneous attachment
''' </summary>
<Serializable()>
Public Class CDBEmailAttachment

  ''' <summary>
  ''' The system wide identifier of the attachment.
  ''' </summary>
  ''' <value>
  ''' The attachment's identifier.
  ''' </value>
  ''' <remarks>This is intended to allow an existing <see cref="CDBEmailAttachment" /> to be attached to an email.  
  ''' In normal use this should be set to zero.</remarks>
  Property Id As Integer
  ''' <summary>
  ''' The name of the attachment.
  ''' </summary>
  ''' <value>
  ''' The attachment's name.
  ''' </value>
  ''' <remarks>Usually a filename for the attached data.</remarks>
  Property Name As String
  ''' <summary>
  ''' The content of the attachment.
  ''' </summary>
  ''' <value>
  ''' The attachment's content.
  ''' </value>
  Property Content As Byte()

End Class
