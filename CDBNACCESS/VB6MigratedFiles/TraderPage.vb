

Namespace Access
  Public Class TraderPage

    Public Enum GetSourceFromMailingType
      gsfmtAlways
      gsfmtMaybe
      gsfmtNever
    End Enum

    Public PageType As Integer
    Public PageCode As String
    Public First As Integer
    Public Last As Integer
    Public DefaultsSet As Boolean
    Public Summary As Boolean
    Public Menu As Boolean
    Public MenuCount As Integer
    Public FirstMenuIndex As Integer
    Public PageChanged As Boolean
    Public GetSourceFromMailing As GetSourceFromMailingType
  End Class
End Namespace
