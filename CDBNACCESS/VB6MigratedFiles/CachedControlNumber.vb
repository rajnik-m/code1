

Namespace Access
  Public Class CachedControlNumber

    Private mvEnv As CDBEnvironment
    Private mvControlNumberType As String
    Private mvBlockCount As Integer
    Private mvNextNumber As Integer
    Private mvMaxNumber As Integer

    Public Sub Init(ByRef pEnv As CDBEnvironment, ByRef pControlNumberType As String, ByRef pBlockCount As Integer)
      mvEnv = pEnv
      mvControlNumberType = pControlNumberType
      mvBlockCount = pBlockCount
      GetMoreNumbers()
    End Sub

    Public Sub CheckAvailable(ByRef pCount As Integer)
      If pCount > mvBlockCount Then mvBlockCount = pCount
      If mvMaxNumber - mvNextNumber < pCount Then GetMoreNumbers()
    End Sub

    Public Function NextControlNumber() As Integer
      If mvNextNumber = 0 Or mvNextNumber = mvMaxNumber Then GetMoreNumbers()
      NextControlNumber = mvNextNumber
      mvNextNumber = mvNextNumber + 1
    End Function

    Private Sub GetMoreNumbers()
      mvNextNumber = mvEnv.GetControlNumber(mvControlNumberType, mvBlockCount, True)
      mvMaxNumber = mvNextNumber + mvBlockCount
    End Sub

  End Class
End Namespace
