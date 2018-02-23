Imports System.Runtime.CompilerServices

Public Module ExtensionMethods
  <Extension()>
  Public Function CaseInsensitiveCompare(ByVal pSource As String, ByVal pToCheck As String, ByVal pCompare As StringComparison) As Boolean
    Return pSource.IndexOf(pToCheck, pCompare) >= 0
  End Function
End Module
