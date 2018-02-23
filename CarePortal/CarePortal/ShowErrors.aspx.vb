Partial Public Class ShowErrors
  Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Dim vException As Exception
    Dim vLastPageNumber As Integer = 0

    If Not IsPostBack Then
      Dim vHeaderText As String = Utilities.GetCustomPageElement("head")
      If Not String.IsNullOrEmpty(vHeaderText) Then Page.Header.InnerHtml = vHeaderText
      BodyStart.Text = Utilities.GetCustomPageElement("bodystart")
      BodyEnd.Text = Utilities.GetCustomPageElement("bodyend")

      Dim vErrorId As Integer
      If Request.QueryString("EI") IsNot Nothing Then vErrorId = IntegerValue(Request.QueryString("EI"))
      Dim vErrorMessage As String = GetCustomConfigItem("CustomConfiguration/ErrorHandling/ErrorMessage")
      If String.IsNullOrEmpty(vErrorMessage) Then vErrorMessage = "An Error occurred processing the last request. Please report the error number {0} to the web site administrator"

      vException = DirectCast(Session("LastException"), Exception)
      If vException Is Nothing Then
        vException = DirectCast(Application("LastException"), Exception)
        If Application("LastPageNumber") IsNot Nothing Then vLastPageNumber = IntegerValue(Application("LastPageNumber").ToString)
        'check if it is a PortalException
        If TypeOf vException Is PortalAccessOrganisationException Then
          vErrorMessage = GetCustomConfigItem("CustomConfiguration/ErrorHandling/AccessOrganisationMessage")
          If String.IsNullOrEmpty(vErrorMessage) Then vErrorMessage = "You do not have access rights to update the organisation details"
        ElseIf TypeOf vException Is PortalAccessException Then
          vErrorMessage = GetCustomConfigItem("CustomConfiguration/ErrorHandling/AccessMessage")
          If String.IsNullOrEmpty(vErrorMessage) Then vErrorMessage = "You do not have access rights for this page"
        Else
          'We won't have logged this error in the database yet
          Try
            Dim vList As New ParameterList(HttpContext.Current)
            Dim vHTTPException As HttpException = TryCast(vException, HttpException)
            If vHTTPException IsNot Nothing Then
              vList("ErrorNumber") = vHTTPException.GetHttpCode
            Else
              vList("ErrorNumber") = 0
            End If
            vList("ErrorSource") = vException.Source
            vList("WebPageNumber") = vLastPageNumber
            vList("ErrorMessage") = vException.Message
            If vList("ErrorNumber").ToString = "404" Then
              vList("ErrorMessage") = vException.Message & vbCrLf & Request.QueryString("aspxerrorpath")
            End If
            vList("StackTrace") = vException.StackTrace
            'Record Error in Database.
            Dim vResult As DataRow = DataHelper.GetRowFromDataTable(DataHelper.AddErrorLog(vList))
            If vResult IsNot Nothing AndAlso vResult.Table.Columns.Contains("ErrorId") AndAlso vResult("ErrorId").ToString.Length > 0 Then
              vErrorId = IntegerValue(vResult("ErrorId").ToString)
            End If
          Catch ex As Exception
            'Give up cannot log the error
          End Try
        End If
      Else
        If Session("LastPageNumber") IsNot Nothing Then vLastPageNumber = IntegerValue(Session("LastPageNumber").ToString)
      End If
      lblError.Text = String.Format(vErrorMessage, vErrorId)

      If vException Is Nothing Then
        lblMessage.Text = "No Error Information is Available"
        lblSource.Text = Request.QueryString("aspxerrorpath")
        LocationRow.Visible = False
      Else
        If TypeOf vException Is HttpRequestValidationException Then
          lblMessage.Text = "HTML entry is not allowed in any fields." & vbCrLf & "Please make sure that your entries do not contain any angle brackets like < or >."
          lblSource.Visible = False
          SourceRow.Visible = False
          LocationRow.Visible = False
        Else
          Dim vShowFullMessage As Boolean = Debugger.IsAttached OrElse Request.Url.Host = "localhost"
          If vShowFullMessage Then
            lblMessage.Text = vException.Message
            lblSource.Text = vException.Source
            lblCallStack.Text = vException.StackTrace
          Else
            MessageRow.Visible = False
            SourceRow.Visible = False
            LocationRow.Visible = False
          End If
        End If
        If vLastPageNumber > 0 Then
          hyp.NavigateUrl = "default.aspx?pn=" & vLastPageNumber.ToString
        Else
          HyperlinkRow.Visible = False
        End If
      End If
    End If

  End Sub

End Class