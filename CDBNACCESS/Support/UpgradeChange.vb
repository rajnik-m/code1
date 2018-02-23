Namespace Access
  Public Class UpgradeChange
    Public ChangeNumber As Integer
    Private mvChangeComment As String
    Private mvToBeApplied As Boolean
    Private mvRelease As String
    Private mvLogname As String
    Private mvDate As String
    Private mvDescription As String

    Public Sub SetVersionHistory(ByVal pRelease As String, ByVal pLogname As String, ByVal pDate As String, ByVal pDescription As String)
      mvRelease = pRelease
      mvLogname = If(String.IsNullOrWhiteSpace(pLogname), "dbinit", pLogname)
      mvDate = pDate
      mvDescription = pDescription
    End Sub

    Public ReadOnly Property VersionHistoryExists() As Boolean
      Get
        If mvDescription IsNot Nothing AndAlso mvDescription.Length > 0 Then
          VersionHistoryExists = True
        Else
          VersionHistoryExists = False
        End If

      End Get
    End Property
    Public Property ToBeApplied() As Boolean
      Get
        ToBeApplied = mvToBeApplied
      End Get
      Set(ByVal Value As Boolean)
        mvToBeApplied = Value
      End Set
    End Property
    Public Property ChangeComment() As String
      Get
        ChangeComment = mvChangeComment
      End Get
      Set(ByVal Value As String)
        mvChangeComment = Value
      End Set
    End Property
    Public ReadOnly Property ReleaseNumber() As String
      Get
        ReleaseNumber = mvRelease
      End Get
    End Property
    Public ReadOnly Property Logname() As String
      Get
        Logname = mvLogname
      End Get
    End Property
    Public ReadOnly Property ChangeDate() As String
      Get
        ChangeDate = mvDate
      End Get
    End Property
    Public ReadOnly Property ChangeDescription() As String
      Get
        ChangeDescription = mvDescription
      End Get
    End Property
  End Class
End Namespace
