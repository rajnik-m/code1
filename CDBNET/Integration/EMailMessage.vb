Public Class EMailMessage

  Private mvID As String
  Private mvSubject As String
  Private mvDateReceived As String
  Private mvOrigDisplayName As String
  Private mvOrigAddress As String
  Private mvRead As Boolean
  Private mvToList As String
  Private mvCCList As String
  Private mvAttachmentCount As Integer
  Private mvAttachmentCollection As ArrayList
  Private mvNoteText As String
  Private mvDeleteAfterSave As Boolean

  Friend Sub InitNewMessage(ByVal pSubject As String, ByVal pMessage As String, ByVal pToList As String, ByVal pCCList As String, ByVal pAttachmentCollection As ArrayList)
    mvSubject = pSubject
    mvNoteText = pMessage
    mvToList = pToList
    mvCCList = pCCList
    If pAttachmentCollection Is Nothing Then
      mvAttachmentCount = 0
      mvAttachmentCollection = New ArrayList
    Else
      mvAttachmentCount = pAttachmentCollection.Count
      mvAttachmentCollection = pAttachmentCollection
    End If
  End Sub

  Friend Sub InitFromMessage(ByVal pID As String, ByVal pSubject As String, ByVal pDateReceived As String, ByVal pOrigDisplayName As String, ByVal pOrigAddress As String, ByVal pRead As Boolean, ByVal pToList As String, ByVal pCCList As String, ByVal pAttachmentCount As Integer, ByVal pAttachmentColl As ArrayList)
    mvNoteText = ""
    mvID = pID
    mvSubject = pSubject
    mvDateReceived = pDateReceived
    mvOrigDisplayName = pOrigDisplayName
    mvOrigAddress = pOrigAddress
    mvRead = pRead
    mvToList = pToList
    mvCCList = pCCList
    mvAttachmentCount = pAttachmentCount
    mvAttachmentCollection = pAttachmentColl
  End Sub

  Public ReadOnly Property SenderAddress() As String
    Get
      If mvOrigAddress.StartsWith("SMTP:") Then
        Return mvOrigAddress.Substring(5)
      ElseIf mvOrigAddress.StartsWith("EX:/") Or mvOrigAddress.StartsWith("/") Then
        Return mvOrigDisplayName
      Else
        Return mvOrigAddress
      End If
    End Get
  End Property

  Public ReadOnly Property ID() As String
    Get
      Return mvID
    End Get
  End Property
  Public ReadOnly Property Subject() As String
    Get
      Return mvSubject
    End Get
  End Property
  Public ReadOnly Property DateReceived() As String
    Get
      Return mvDateReceived
    End Get
  End Property
  Public ReadOnly Property OrigDisplayName() As String
    Get
      Return mvOrigDisplayName
    End Get
  End Property
  Public Property OrigAddress() As String
    Get
      Return (mvOrigAddress)
    End Get
    Set(ByVal Value As String)        'Need to set this for Outlook items before saving
      mvOrigAddress = Value
    End Set
  End Property
  Public ReadOnly Property Read() As Boolean
    Get
      Return mvRead
    End Get
  End Property
  Public Property NoteText() As String
    Get
      Return mvNoteText
    End Get
    Set(ByVal Value As String)
      mvNoteText = Value
    End Set
  End Property
  Public Property ToList() As String
    Get
      Return mvToList
    End Get
    Set(ByVal Value As String)
      mvToList = Value
    End Set
  End Property
  Public Property CCList() As String
    Get
      Return mvCCList
    End Get
    Set(ByVal Value As String)
      mvCCList = Value
    End Set
  End Property
  Public ReadOnly Property AttachmentCount() As Integer
    Get
      Return mvAttachmentCount
    End Get
  End Property
  Public Property AttachmentCollection() As ArrayList
    Get
      Return mvAttachmentCollection
    End Get
    Set(ByVal Value As ArrayList)
      mvAttachmentCollection = Value
    End Set
  End Property
  Public ReadOnly Property AttachmentNameList() As String
    Get
      Dim vList As New System.Text.StringBuilder
      For Each vString As String In mvAttachmentCollection
        If vList.Length > 0 Then vList.Append(",")
        vList.Append(vString)
      Next
      Return vList.ToString
    End Get
  End Property
  Public Property DeleteAfterSave() As Boolean
    Get
      Return (mvDeleteAfterSave)
    End Get
    Set(ByVal Value As Boolean)
      mvDeleteAfterSave = Value
    End Set
  End Property
End Class
