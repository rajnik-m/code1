Namespace Access

  Public Class Postcode
    Private mvAfterSpace As String = ""    'lhs of postcode including first character after space
    Private mvBeforeSpace As String   'lhs of postcode prior to space
    Private mvAlpha As String         'lhs of postcode - up to first numeric
    Private mvFullPC As String = ""      'Full postcode
    Private mvOutwardPC As String
    Private mvInwardPC As String
    Private mvMaxLength As Integer

    Public Sub New(ByVal pPostCode As String)
      'Given a Postcode, split into constituent parts for Branch/Geo Region assignment etc.
      Dim vSpacePos As Integer = (pPostCode.IndexOf(" ", 0) + 1)
      Dim vSpaceReqPos As Integer = 6 'If the space is before position 6 then keep it now
      mvMaxLength = 10
      If vSpacePos = 0 Then
        mvBeforeSpace = pPostCode
      Else
        mvFullPC = pPostCode
        mvOutwardPC = Substring(pPostCode, 0, vSpacePos - 1) 'All before the space
        mvBeforeSpace = Substring(pPostCode, 0, vSpacePos - 1)
        mvInwardPC = Substring(Substring(pPostCode, (vSpacePos - 1)).TrimStart(" "c).ToString, 0, 1) '1st character after space
        If mvInwardPC.Length > 0 Then
          If vSpacePos < vSpaceReqPos Then
            mvOutwardPC = mvOutwardPC & " " & mvInwardPC
          Else
            mvOutwardPC = mvOutwardPC & mvInwardPC
          End If
          mvAfterSpace = mvOutwardPC
        End If
      End If
      mvAlpha = mvBeforeSpace.TrimEnd(" "c)
      Dim vLength As Integer = mvAlpha.Length
      'SDT 25/5 Fixed problem where Postcode like W1J 6BD was leaving mvAlpha as W1J
      Dim vIndex As Integer = 1
      Do While vIndex + 1 <= vLength And Not IsNumeric(Substring(mvAlpha, vIndex, 1))
        vIndex = vIndex + 1
      Loop
      mvAlpha = Substring(mvAlpha, 0, vIndex)
    End Sub

    Public ReadOnly Property AfterSpace() As String
      Get
        Return mvAfterSpace
      End Get
    End Property
    Public ReadOnly Property BeforeSpace() As String
      Get
        Return mvBeforeSpace
      End Get
    End Property
    Public ReadOnly Property Alpha() As String
      Get
        Return mvAlpha
      End Get
    End Property
    Public ReadOnly Property FullPC() As String
      Get
        Return mvFullPC
      End Get
    End Property
    Public ReadOnly Property OutwardPC() As String
      Get
        Return mvOutwardPC
      End Get
    End Property
    Public ReadOnly Property InwardPC() As String
      Get
        Return mvInwardPC
      End Get
    End Property
    Public ReadOnly Property MaxLength() As Integer
      Get
        Return mvMaxLength
      End Get
    End Property

    Public ReadOnly Property Components() As CDBParameters
      Get
        Dim vParams As New CDBParameters
        If mvFullPC.Length > 0 Then vParams.Add(mvFullPC)
        If mvAfterSpace.Length > 0 Then
          If mvAfterSpace.Length <= mvMaxLength Then
            If Not vParams.Exists(mvAfterSpace) Then vParams.Add(mvAfterSpace)
          End If
        End If
        If mvBeforeSpace.Length <= mvMaxLength Then
          If Not vParams.Exists(mvBeforeSpace) Then vParams.Add(mvBeforeSpace)
        End If
        If mvAlpha.Length > 0 And mvAlpha.Length <= mvMaxLength Then
          If Not vParams.Exists(mvAlpha) Then vParams.Add(mvAlpha)
        End If
        Return vParams
      End Get
    End Property
  End Class

End Namespace