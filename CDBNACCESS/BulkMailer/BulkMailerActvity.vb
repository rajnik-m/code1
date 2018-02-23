''' <summary>
''' The mailing statistics associated with an individual
''' </summary>
Public Class BulkMailerActvity

  Private mvEmailAddress As String
  Private mvBounced As Boolean
  Private mvUnsubscribed As Boolean
  Private mvOpened As Nullable(Of Date)
  Private mvClickedThrough As Nullable(Of Date)

  Public Sub New(ByVal pEmailAddress As String, ByVal pBounced As Boolean, ByVal pOpened As Nullable(Of Date), ByVal pClickedThrough As Nullable(Of Date), pUnsubscribed As Boolean)
    mvEmailAddress = pEmailAddress
    mvBounced = pBounced
    mvOpened = pOpened
    mvClickedThrough = pClickedThrough
    mvUnsubscribed = pUnsubscribed
  End Sub

  Public ReadOnly Property EmailAddress() As String
    Get
      Return mvEmailAddress
    End Get
  End Property

  Public ReadOnly Property IsBounced() As Boolean
    Get
      Return mvBounced
    End Get
  End Property

  Public ReadOnly Property IsUnsubscribed() As Boolean
    Get
      Return mvBounced
    End Get
  End Property

  Public ReadOnly Property Opened() As Nullable(Of Date)
    Get
      Return mvOpened
    End Get
  End Property

  Public ReadOnly Property ClickedThrough() As Nullable(Of Date)
    Get
      Return mvClickedThrough
    End Get
  End Property

End Class
