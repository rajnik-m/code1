Imports System.Web.Configuration
Imports System.IO

Partial Public Class AddFundraisingPage
  Inherits CareWebControl

  Private mvPageNumber As Integer           'This page number
  Private mvEditPageNumber As Integer       'The number of the page being edited
  Private mvFundraisingNumber As Integer    'The ContactFundraisingNumber for the page being edited

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctAddFundraisingPage, tblDataEntry, "TemplatePage", "Topic,SubTopic")
      If Request.QueryString("pn") IsNot Nothing AndAlso Request.QueryString("pn").Length > 0 Then mvPageNumber = IntegerValue(Request.QueryString("pn")) 'This is the current page
      If Request.QueryString("fpn") IsNot Nothing AndAlso Request.QueryString("fpn").Length > 0 Then mvEditPageNumber = IntegerValue(Request.QueryString("fpn")) 'This is the page we came from
      Dim vTable As DataTable = Nothing
      Dim vList As New ParameterList(HttpContext.Current)
      vList("WebPageNumber") = IIf(mvEditPageNumber > 0, mvEditPageNumber.ToString, InitialParameters("TemplatePage"))
      vTable = DataHelper.GetWebPageItems(vList)
      If vTable IsNot Nothing Then
        For Each vRow As DataRow In vTable.Rows
          Select Case vRow("WebPageItemName").ToString
            Case "ThankYouMessage"
              SetTextBoxText("ThankYouMessage", StripHTML(vRow("WebPageHtml").ToString))
            Case "PersonalMessage"
              SetTextBoxText("PersonalMessage", StripHTML(vRow("WebPageHtml").ToString))
            Case "PageTitle"
              SetTextBoxText("PageTitle", StripHTML(vRow("WebPageHtml").ToString))
          End Select
        Next
      End If
      If mvEditPageNumber > 0 Then
        'Editing an existing page so get the ThankYouMessage from the ContactFundraisingEvent
        Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactFundraisingEvents, vList))
        If vRow IsNot Nothing Then
          mvFundraisingNumber = IntegerValue(vRow("ContactFundraisingNumber").ToString)
          Dim vEventNumber As Integer = IntegerValue(vRow("EventNumber").ToString)
          SetDropDownText("EventNumber", vEventNumber.ToString)
          SetTextBoxText("FundraisingDescription", vRow("FundraisingDescription").ToString)
          SetTextBoxText("TargetDate", vRow("TargetDate").ToString)
          SetTextBoxText("TargetAmount", vRow("TargetAmount").ToString)
          SetTextBoxText("ThankYouMessage", vRow("ThankYouMessage").ToString)
          SetControlEnabled("EventNumber", False)
          SetControlEnabled("TargetDate", (vEventNumber = 0))
          SetControlEnabled("FundraisingDescription", (vEventNumber = 0))
        End If
        vRow = DataHelper.GetRowFromDataTable(DataHelper.SelectWebDataTable(CareNetServices.XMLWebDataSelectionTypes.wstPage, vList))
        If vRow IsNot Nothing Then
          Dim vFriendlyUrl As String = vRow("FriendlyUrl").ToString
          SetTextBoxText("FriendlyUrl", vFriendlyUrl)
          SetControlEnabled("FriendlyUrl", False)
          SetControlEnabled("MailTo", (vFriendlyUrl.Length > 0))
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Me.Page.Form.Enctype = "multipart/form-data"
    If Not IsPostBack Then
      If mvEditPageNumber = 0 Then
        SetControlEnabled("ClosePage", False)
      End If
    End If
  End Sub

  Protected Overrides Sub AddCustomValidator(ByVal pHTMLTable As HtmlTable)
    Dim vControl As Control = FindControlByName(tblDataEntry, "FriendlyUrl")
    If vControl IsNot Nothing Then
      AddCustomValidator(DirectCast(vControl.Parent, HtmlTableCell), "1", "Url is invalid or has already been used")
    End If
  End Sub

  Public Overrides Sub ServerValidate(ByVal sender As Object, ByVal args As ServerValidateEventArgs)
    Dim vValid As Boolean = True
    If GetTextBoxText("FriendlyUrl").Length > 0 Then
      If FindControlByName(Me, "FriendlyUrl").Visible = True Then vValid = ValidateFriendlyUrl()
    End If
    args.IsValid = vValid
  End Sub

  Public Overrides Sub ProcessSubmit()
    Dim vNewWebPageNumber As Integer
    Dim vList As New ParameterList(HttpContext.Current)

    'Add/Update WebPage
    vList("WebNumber") = WebConfigurationManager.AppSettings("WebNumber").ToString
    If mvEditPageNumber > 0 Then
      vNewWebPageNumber = mvEditPageNumber
      'Need to update the WebPage and the ContactFundraisingEvent
      Dim vCheckBox As CheckBox = TryCast(FindControlByName(Me, "ClosePage"), CheckBox)
      If vCheckBox IsNot Nothing AndAlso vCheckBox.Checked Then
        'Only want to update this flag if the check box has been checked
        vList("WebPageNumber") = mvEditPageNumber.ToString
        vList("PagePublished") = "N"
        DataHelper.UpdateWebItem(CareNetServices.XMLWebDataSelectionTypes.wstPage, vList)
      End If
    Else
      vList("WebPageNumber") = InitialParameters("TemplatePage")
      vList("WebPageName") = "User Sponsor Me Page"
      vList("AsCopy") = "Y"
      vList("EditWebPageNumber") = mvPageNumber.ToString
      Dim vFriendlyUrl As String = GetTextBoxText("FriendlyUrl")
      If vFriendlyUrl.Length > 0 Then
        If vFriendlyUrl.EndsWith("aspx") = False Then vFriendlyUrl &= ".aspx"
        vList("FriendlyUrl") = vFriendlyUrl
      End If
      Dim vReturnList As ParameterList = DataHelper.AddWebItem(CareNetServices.XMLWebDataSelectionTypes.wstPage, vList)
      vNewWebPageNumber = IntegerValue(vReturnList("WebPageNumber").ToString)
    End If

    'Add/Update FundraisingEvents
    vList = New ParameterList(HttpContext.Current)
    AddOptionalTextBoxValue(vList, "TargetDate")
    AddOptionalTextBoxValue(vList, "TargetAmount")
    AddOptionalTextBoxValue(vList, "FundraisingDescription")
    AddOptionalTextBoxValue(vList, "ThankYouMessage")
    AddDefaultParameters(vList)
    If mvEditPageNumber > 0 Then
      vList("ContactFundraisingNumber") = mvFundraisingNumber.ToString
      DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctFundraisingEvents, vList)
    Else
      vList("WebPageNumber") = vNewWebPageNumber
      vList("ContactNumber") = UserContactNumber()
      Dim vEventNumber As Integer = IntegerValue(GetDropDownValue("EventNumber"))
      If vEventNumber > 0 Then vList("EventNumber") = vEventNumber.ToString
      DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctFundraisingEvents, vList)
    End If

    'Now handle the image
    Dim vImageFileName As String = ""
    Dim vControl As HtmlInputFile = TryCast(FindControlByName(tblDataEntry, "Image"), HtmlInputFile)
    If vControl IsNot Nothing Then
      If vControl.PostedFile IsNot Nothing AndAlso vControl.PostedFile.ContentLength > 0 Then
        Dim vExtension As String = Path.GetExtension(vControl.PostedFile.FileName).ToLower
        If vExtension <> ".bmp" And vExtension <> ".jpg" And vExtension <> ".gif" Then
          Throw New Exception("Only BMP, JPG and GIF files may be loaded")
        End If
        Dim vFileName As String = Server.MapPath(String.Format("Images\SponsorMeImage{0}.tmp", vNewWebPageNumber))
        vControl.PostedFile.SaveAs(vFileName)
        Dim vFS As FileStream = Nothing
        Dim vBuffer(10) As Byte
        Try
          vFS = New FileStream(vFileName, FileMode.Open)
          vFS.Read(vBuffer, 0, 10)
          vFS.Close()
        Finally
          If vFS IsNot Nothing Then vFS.Close()
        End Try
        Dim vStr As String = System.Text.Encoding.ASCII.GetString(vBuffer)
        If vStr.Length < 10 OrElse (vStr.Substring(6, 4) <> "JFIF" AndAlso vStr.Substring(0, 3) <> "GIF" AndAlso vStr.Substring(0, 2) <> "BM") Then
          My.Computer.FileSystem.DeleteFile(vFileName)
          Throw New Exception("Image has an Invalid File Format")
        Else
          vImageFileName = Path.ChangeExtension(vFileName, vExtension)
          My.Computer.FileSystem.MoveFile(vFileName, vImageFileName)
        End If
      End If
    End If

    vList = New ParameterList(HttpContext.Current)
    vList("WebPageNumber") = vNewWebPageNumber
    Dim vTable As DataTable = DataHelper.GetWebPageItems(vList)
    If vTable IsNot Nothing Then
      For Each vRow As DataRow In vTable.Rows
        Select Case vRow("WebPageItemName").ToString
          Case "PersonalMessage"
            Dim vUpdateList As New ParameterList(HttpContext.Current)
            vUpdateList("WebPageItemNumber") = vRow("WebPageItemNumber")
            vUpdateList("WebPageHtml") = GetTextBoxText("PersonalMessage")
            DataHelper.UpdateWebItem(CareNetServices.XMLWebDataSelectionTypes.wstPageItem, vUpdateList)
          Case "PageTitle"
            Dim vUpdateList As New ParameterList(HttpContext.Current)
            vUpdateList("WebPageItemNumber") = vRow("WebPageItemNumber")
            vUpdateList("WebPageHtml") = GetTextBoxText("PageTitle")
            DataHelper.UpdateWebItem(CareNetServices.XMLWebDataSelectionTypes.wstPageItem, vUpdateList)
          Case "PersonalImage"
            If vImageFileName.Length > 0 Then
              Dim vUpdateList As New ParameterList(HttpContext.Current)
              vUpdateList("WebPageItemNumber") = vRow("WebPageItemNumber")
              Dim vImageText As String = vRow("WebPageHtml").ToString
              vUpdateList("WebPageHtml") = ConvertImageText(vImageText, vImageFileName)
              DataHelper.UpdateWebItem(CareNetServices.XMLWebDataSelectionTypes.wstPageItem, vUpdateList)
            End If
        End Select
      Next
    End If
    If mvSubmitItemNumber = 0 Then mvSubmitItemNumber = vNewWebPageNumber
  End Sub

  Private Function ConvertImageText(ByVal pSource As String, ByVal pFileName As String) As String
    pSource = pSource.ToLower
    Dim vSrcIndex As Integer
    Dim vEndSrc As Integer
    vSrcIndex = pSource.IndexOf("src=")
    If vSrcIndex >= 0 Then
      If pSource.Substring(vSrcIndex + 4, 1) = """" Then
        vEndSrc = pSource.IndexOf("""", vSrcIndex + 5)
        vSrcIndex += 5
      Else
        vEndSrc = pSource.IndexOf(" ", vSrcIndex + 4)
      End If
      Dim vSource As String = pSource.Substring(vSrcIndex, vEndSrc - vSrcIndex)
      Return pSource.Replace(Path.GetFileName(vSource), Path.GetFileName(pFileName))
    End If
    Return pSource
  End Function

  Private Function ValidateFriendlyUrl() As Boolean
    Dim vValid As Boolean = True
    Dim vFriendlyUrl As String = GetTextBoxText("FriendlyUrl")

    If mvEditPageNumber = 0 AndAlso vFriendlyUrl.Length > 0 Then
      'No need to re-validate the URL when editing a Web Page as it cannot be changed
      If vFriendlyUrl.EndsWith(".aspx") Then vFriendlyUrl = vFriendlyUrl.Substring(0, vFriendlyUrl.Length - 5)
      If vFriendlyUrl.StartsWith("http") Then vValid = False
      If vValid Then
        Dim vPatternSite As String = "\w*[\://]*\w+\.\w+\.\w+[/\w+]*[.\w+]*"
        Dim vRegEx As New Regex(vPatternSite)
        If vRegEx.IsMatch(vFriendlyUrl) Then
          vValid = False
        End If
      End If
      If vValid Then
        If vFriendlyUrl.Contains("www") Then vValid = False
        If vFriendlyUrl.Contains(".") Then vValid = False
        If vFriendlyUrl.Contains("/") Then vValid = False
        If vFriendlyUrl.Contains("\") Then vValid = False
        If vFriendlyUrl.Contains(" ") Then vValid = False
        If vFriendlyUrl.Length > 75 Then vValid = False 'Max characters is 80 which includes the ".aspx" at the end
      End If
      If vValid Then
        vFriendlyUrl &= ".aspx"
        Dim vList As New ParameterList(HttpContext.Current)
        vList("WebNumber") = WebConfigurationManager.AppSettings("WebNumber").ToString
        vList("FriendlyUrl") = vFriendlyUrl
        Dim vDT As DataTable = DataHelper.SelectWebDataTable(CareNetServices.XMLWebDataSelectionTypes.wstPages, vList)
        If vDT IsNot Nothing Then vValid = (vDT.Rows.Count = 0)
      End If
    End If

    Return vValid
  End Function

  Protected Overrides Sub FriendlyUrlChanged(ByVal pTextBox As System.Web.UI.WebControls.TextBox, ByVal pValue As String)
    Try
      Dim vUrlValid As Boolean = ValidateFriendlyUrl()
      Dim vHyperlink As HyperLink = TryCast(FindControl("MailTo"), HyperLink)
      If pValue.Length > 0 Then
        If vHyperlink IsNot Nothing Then
          If vUrlValid Then
            If pValue.EndsWith(".aspx") = False Then pValue &= ".aspx"
            Dim vUrl As String = Request.Url.AbsoluteUri
            Dim vPos As Integer = vUrl.LastIndexOf("/")
            If vPos >= 0 Then vUrl = vUrl.Substring(0, vPos + 1) & pValue
            vHyperlink.NavigateUrl = String.Format("MailTo:?subject=Sponsor Me&body=Please look at this page {0}", vUrl)
          End If
          vHyperlink.Enabled = vUrlValid
        End If
      Else
        vHyperlink.Enabled = False
      End If
    Catch vEX As Exception
      ProcessError(vEX)
    End Try
  End Sub
End Class