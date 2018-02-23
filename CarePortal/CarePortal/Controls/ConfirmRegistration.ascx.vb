Partial Public Class ConfirmRegistration
  Inherits CareWebControl

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      If Not InWebPageDesigner() And Not IsPostBack Then
        Dim vEMail As String = Request.QueryString("UserName")
        If vEMail IsNot Nothing AndAlso vEMail.Length > 0 Then
          If Session("ConfirmRegistration") Is Nothing Then
            Dim vEP As New EncryptionProvider
            Dim vParams As New ParameterList(HttpContext.Current)
            vParams("UserName") = vEP.Decrypt(vEMail.Replace(vbCr, "").Replace(vbLf, ""))
            vParams("UseRegistrationData") = "Y"
            Dim vList As ParameterList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctContact, vParams)
            Session("ContactNumber") = vList("ContactNumber")
            Session("AddressNumber") = vList("AddressNumber")
            SetAuthentication(vList)
            Session("RegisteredUserName") = vParams("UserName")
            Session("UserContactNumber") = vList("ContactNumber")
            'Reload the page to read the authenticated vCookie value
            Session("ConfirmRegistration") = "Y"
            Response.Redirect(Request.RawUrl)
          End If
        Else
          Throw New CareException(String.Format("Invalid Registration URL {0}", Request.Url))
        End If
      End If
      If HTML.Length > 0 Then
        Dim vHTMLRow As New HtmlTableRow
        Dim vHTMLCell As New HtmlTableCell
        vHTMLCell.InnerHtml = HTML
        vHTMLRow.Cells.Add(vHTMLCell)
        tblContent.Rows.Add(vHTMLRow)
        tblContent.Attributes("Class") = GetClass()
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub
End Class