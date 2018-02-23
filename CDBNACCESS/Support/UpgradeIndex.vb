Namespace Access
  Public Class UpgradeIndex
    Public Key As String
    Private mvAttributes As New List(Of String)
    Private mvToBeCreated As Boolean
    Private mvToBeDeleted As Boolean
    Private mvUniqueIndex As Boolean
    Private mvPrimaryValue As Integer
    Private mvOverflowValue As Integer

    Public Property OverflowValue() As Integer
      Get
        OverflowValue = mvOverflowValue
      End Get
      Set(ByVal Value As Integer)
        mvOverflowValue = Value
      End Set
    End Property
    Public Property PrimaryValue() As Integer
      Get
        PrimaryValue = mvPrimaryValue
      End Get
      Set(ByVal Value As Integer)
        mvPrimaryValue = Value
      End Set
    End Property
    Public Property UniqueIndex() As Boolean
      Get
        UniqueIndex = mvUniqueIndex
      End Get
      Set(ByVal Value As Boolean)
        mvUniqueIndex = Value
      End Set
    End Property
    Public Property ToBeDeleted() As Boolean
      Get
        ToBeDeleted = mvToBeDeleted
      End Get
      Set(ByVal Value As Boolean)
        mvToBeDeleted = Value
      End Set
    End Property
    Public Property ToBeCreated() As Boolean
      Get
        ToBeCreated = mvToBeCreated
      End Get
      Set(ByVal Value As Boolean)
        mvToBeCreated = Value
      End Set
    End Property
    Public ReadOnly Property Attributes() As IList(Of String)
      Get
        Return mvAttributes
      End Get
    End Property
  End Class
End Namespace
