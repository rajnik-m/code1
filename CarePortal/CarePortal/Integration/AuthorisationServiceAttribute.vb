Imports System.Runtime.CompilerServices
Imports CarePortal.ProcessPayment

Public Class AuthorisationServiceAttribute
  Inherits Attribute

  Public Sub New(pServiceName As String)
    mvServiceName = pServiceName
  End Sub

  Private mvServiceName As String

  Public ReadOnly Property ServiceName As String
    Get
      Return mvServiceName
    End Get
  End Property
End Class

Public Module AuthorisationServiceExtension
  <Extension()>
  Public Function GetServiceName(pValue As AuthorisationService) As String
    Dim vRtn As String = Nothing

    Try
      Dim vEnumType As Type = GetType(AuthorisationService)
      Dim vInfo() As Reflection.MemberInfo = vEnumType.GetMember(pValue.ToString)
      If vInfo IsNot Nothing AndAlso vInfo.Length > 0 Then
        Dim vAttributes() As Object = vInfo(0).GetCustomAttributes(GetType(AuthorisationServiceAttribute), False)
        If vAttributes IsNot Nothing AndAlso vAttributes.Length > 0 Then
          vRtn = TryCast(vAttributes(0), AuthorisationServiceAttribute).ServiceName
        End If
      End If
    Catch ex As Exception
      'Do Nothing
    End Try
    Return vRtn
  End Function
End Module
