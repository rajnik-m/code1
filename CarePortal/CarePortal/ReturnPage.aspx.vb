Public Class ReturnPage
  Inherits System.Web.UI.Page

  Public Enum TNSResult
    NONE = -1
    SUCCESSFUL = 0
    SESSION_EXPIRED = 2
    INVALID_FIELD_VALUES = 3
  End Enum

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Dim vResponse As IDictionary(Of String, String) = New Dictionary(Of String, String)()
    Dim vRedirectURL As String = String.Empty

   
    For Each key As String In Request.Form.AllKeys
      vResponse.Add(key, Request.Form(key))
    Next
    'Get the form's exit code if not using stored values
    Dim vFormExitCode As Integer = If(vResponse.ContainsKey("gatewayFormResponse"), CInt(vResponse("gatewayFormResponse").Substring(0, 1)), 0)

    'If addmembercc then do not read the redirect from the database 
    Dim vAddMemberCC As String = GetResponseValue(vResponse, "AddMemberCC")
    If vAddMemberCC.Length > 0 Then
      vRedirectURL = vAddMemberCC
    Else
      Dim vParams As New ParameterList(HttpContext.Current)
      vParams("CarePortal") = "Y"

      If Session("Trader") IsNot Nothing AndAlso String.Compare(Session("Trader").ToString, "Y", True) = 0 Then
        If Session("BatchCategory") IsNot Nothing Then
          vParams("BatchCategory") = Session("BatchCategory")
        Else
          'Raise error 
        End If
      Else
        If GetResponseValue(vResponse, "BatchCategory").Length > 0 Then
          vParams("BatchCategory") = GetResponseValue(vResponse, "BatchCategory")
        Else
          'Raise error
        End If
      End If

      Dim vResult As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMerchantDetails, vParams))

      If vResult IsNot Nothing Then
        vRedirectURL = vResult.Item("CardDetailsPageURL").ToString
      Else
        'Raise Error
      End If
    End If

      Select Case CInt(vFormExitCode)
        Case TNSResult.SUCCESSFUL
          Session("FormErrorContents") = vResponse
          Session("FormErrorCode") = vFormExitCode
        Case TNSResult.SESSION_EXPIRED, TNSResult.INVALID_FIELD_VALUES
          Session("FormErrorContents") = vResponse
          Session("FormErrorCode") = vFormExitCode
        Case Else
          Session("FormErrorContents") = TNSResult.NONE
          Session("FormErrorCode") = vFormExitCode
      End Select
      Response.Redirect(vRedirectURL)

  End Sub

  Private Function GetResponseValue(ByVal pResponse As IDictionary(Of String, String), ByVal pAttributeName As String) As String
    Dim vFormFieldValue As IDictionary(Of String, String) = TryCast(pResponse, IDictionary(Of String, String))
    Dim vFormatedKey As String = String.Empty

    For Each vKey As String In vFormFieldValue.Keys
      If vKey.Contains(pAttributeName) Then
        vFormatedKey = vKey
        Exit For
      End If
    Next

    If vFormatedKey.Length > 0 Then
      Return vFormFieldValue(vFormatedKey)
    Else
      Return ""
    End If
  End Function

End Class