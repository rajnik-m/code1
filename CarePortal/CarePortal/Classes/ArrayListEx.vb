<CLSCompliant(True)> _
Public Class ArrayListEx
  '------------------------------------------------------------------------------------------------------------
  'This class exists in the CDBNUTILS class libarary and changes here should be reflected in that class as well
  '------------------------------------------------------------------------------------------------------------
  Inherits ArrayList

  Private mvIncompleteField As Boolean

  Sub New()
    MyBase.New()
  End Sub

  Sub New(ByVal pCSList As String)
    MyBase.New()
    Dim vList() As String
    vList = pCSList.Split(","c)
    For Each vString As String In vList
      MyBase.Add(vString)
    Next
  End Sub

  Sub New(ByVal pCSList As String, ByVal pCheckForCSV As Boolean)
    MyBase.New()
    Init(pCSList, pCheckForCSV, ",".ToCharArray)
  End Sub

  Sub New(ByVal pCSList As String, ByVal pCheckForSeparator As Boolean, ByVal pSeparator() As Char)
    MyBase.New()
    Init(pCSList, pCheckForSeparator, pSeparator)
  End Sub

  Sub Init(ByVal pCSList As String, ByVal pCheckForSeparator As Boolean, ByVal pSeparator() As Char)
    Dim vList() As String
    If pCheckForSeparator Then
      Dim vInQuotes As Boolean
      Dim vLastChar As Char = Nothing
      Dim vItem As New StringBuilder
      For Each vChar As Char In pCSList.ToCharArray
        Select Case vChar
          Case """"c
            If Not vInQuotes AndAlso vLastChar = """" Then ' Two quotes make a quote
              vLastChar = vChar
            End If
            vInQuotes = Not vInQuotes
          Case Else
            If vChar = pSeparator AndAlso Not vInQuotes Then
              MyBase.Add(vItem.ToString)
              vItem = New StringBuilder
            Else
              vItem.Append(vChar)
            End If
        End Select
        vLastChar = vChar
      Next
      If vItem.Length > 0 Then MyBase.Add(vItem.ToString)
      If vInQuotes Then mvIncompleteField = True
    Else
      vList = pCSList.Split(pSeparator)
      For Each vString As String In vList
        MyBase.Add(vString)
      Next
    End If
  End Sub

  Sub New(ByVal pSList As String, ByVal pSeparator() As Char)
    MyBase.New()
    Dim vList() As String
    vList = pSList.Split(pSeparator, StringSplitOptions.RemoveEmptyEntries)
    For Each vString As String In vList
      MyBase.Add(vString)
    Next
  End Sub

  Sub New(ByVal pSList As String, ByVal pSeparator() As Char, ByVal pSplitOption As StringSplitOptions)
    MyBase.New()
    Dim vList() As String
    vList = pSList.Split(pSeparator, pSplitOption)
    For Each vString As String In vList
      MyBase.Add(vString)
    Next
  End Sub

  Public ReadOnly Property IncompleteField() As Boolean
    Get
      Return mvIncompleteField
    End Get
  End Property

  Function CSList() As String
    Dim vList As New System.Text.StringBuilder
    For Each vString As String In Me
      If vList.Length > 0 Then vList.Append(",")
      vList.Append(vString)
    Next
    Return vList.ToString
  End Function

  Function CRLFList() As String
    Dim vList As New System.Text.StringBuilder
    For Each vString As String In Me
      If vList.Length > 0 Then vList.Append(ControlChars.CrLf)
      vList.Append(vString)
    Next
    Return vList.ToString
  End Function

  Function CSStringList() As String
    Dim vList As New System.Text.StringBuilder
    For Each vString As String In Me
      If vList.Length > 0 Then vList.Append(",")
      vList.Append("'")
      vList.Append(vString)
      vList.Append("'")
    Next
    Return vList.ToString
  End Function

  Function SSNonBlankList() As String
    Dim vList As New System.Text.StringBuilder
    Dim vAddSpace As Boolean
    For Each vString As String In Me
      If vString.Length > 0 Then
        If vAddSpace Then vList.Append(" ")
        vList.Append(vString)
        vAddSpace = True
      End If
    Next
    Return vList.ToString
  End Function

End Class
