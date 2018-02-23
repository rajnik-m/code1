Imports System.Net

Public Class ProtXRequest
  Private mvAddress As Uri
  Private mvRequestFields As New SortedList(Of String, String)
  Private mvResponseFields As String
  Private mvTimeout As Integer

  Public Sub New(ByVal pUrl As String, ByVal pTimeOut As Integer)
    mvAddress = New Uri(pUrl)
    mvTimeout = pTimeOut
  End Sub

  Public Sub AddDigitalOrderField(ByVal pKey As String, ByVal pValue As String)
    mvRequestFields.Add(pKey, pValue)
  End Sub

  Public Function SendRequest() As String
    Try
      'Create the web vRequest  
      Dim vRequest As HttpWebRequest = DirectCast(WebRequest.Create(mvAddress), HttpWebRequest)

      'Set type to POST  
      vRequest.Method = "POST"
      vRequest.ContentType = "application/x-www-form-urlencoded"
      vRequest.Timeout = mvTimeout * 1000

      'Create a byte array of the data we want to send  
      Dim vByteData As Byte() = UTF8Encoding.UTF8.GetBytes(GetRequestRaw)

      'Set the content length in the vRequest headers  
      vRequest.ContentLength = vByteData.Length

      'Write data
      Using vPostStream As System.IO.Stream = vRequest.GetRequestStream
        vPostStream.Write(vByteData, 0, vByteData.Length)
      End Using
      'Get response  
      Using vResponse As HttpWebResponse = TryCast(vRequest.GetResponse, HttpWebResponse)
        'Get the response stream  
        Dim vReader As New System.IO.StreamReader(vResponse.GetResponseStream)
        'Console application output
        mvResponseFields = vReader.ReadToEnd()
      End Using
      Return ""
    Catch vEx As Exception
      Return vEx.Message
    End Try
  End Function

  Private Function GetRequestRaw() As String
    Dim vData As New StringBuilder()
    For Each vPair As KeyValuePair(Of String, String) In mvRequestFields
      If vPair.Value.Length > 0 Then
        If (vData.Length > 0) Then vData.Append("&")
        vData.Append(vPair.Key & "=" & System.Web.HttpUtility.UrlEncode(vPair.Value))
      End If
    Next
    Return vData.ToString
  End Function

  Public Function GetCardType(ByVal pCardType As String) As String
    Dim vCardType As String
    Select Case pCardType
      Case "V", "O"
        'VISA		
        vCardType = "VISA"
      Case "A"
        'Mastercard
        vCardType = "MC"
      Case "SW"
        'Switch Card
        vCardType = "SWITCH"
      Case "DE"
        'Delta Card
        vCardType = "DELTA"
      Case "AX"
        'American Express
        vCardType = "AMEX"
      Case "SO"
        vCardType = "SOLO"
      Case "DC"
        vCardType = "DC"
      Case "JCB"
        vCardType = "JCB"
      Case "M"
        vCardType = "MAESTRO"
      Case Else
        vCardType = ""
    End Select
    Return vCardType
  End Function

  Function GetResultField(ByVal pFieldName As String) As String
    Dim vItems() As String
    Dim vIndex As Integer
    Dim vItem As String
    Dim vReponse As String = String.Empty

    vItems = Split(mvResponseFields, Chr(13))
    For vIndex = LBound(vItems) To UBound(vItems)
      vItem = Replace(vItems(vIndex), Chr(10), "")
      If InStr(vItem, pFieldName & "=") = 1 Then
        ' found
        vReponse = Right(vItem, Len(vItem) - Len(pFieldName) - 1)
        Exit For
      End If
    Next
    Return vReponse
    'String value
    'Dim vValue As String = ""
    'If mvResponseFields.TryGetValue(pKey, vValue) Then
    '  Return vValue
    'Else
    '  Return pDefValue
    'End If
  End Function
End Class
