Public Class ReprintSelectedEventArgs
  Inherits EventArgs

  Public Sub New(pReprintType As String)
    MyBase.New()
    ReprintType = pReprintType
  End Sub

  Private mvReprintType As String = String.Empty
  Public Property ReprintType As String
    Get
      Return mvReprintType
    End Get
    Private Set(value As String)
      mvReprintType = value
    End Set
  End Property

End Class
