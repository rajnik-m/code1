Public Class DownloadSelection
  Inherits CareWebControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctDownloadSelection, tblDataEntry)
      SetLabelVisible("WarningMessage", False)
      If Me.FindControl("SearchDocument") IsNot Nothing Then CType(Me.FindControl("SearchDocument"), TextBox).MaxLength = 100
      If Request.QueryString("Document") IsNot Nothing Then CType(Me.FindControl("SearchDocument"), TextBox).Text = Request.QueryString("Document")
      If Not IsPostBack Then FindDocuments()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
    End Sub
  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      FindDocuments()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
  Private Sub FindDocuments()
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vQueryString As String = String.Empty
    If GetTextBoxText("SearchDocument").Length > 0 Then
      vList("SearchDocument") = "*" & GetTextBoxText("SearchDocument") & "*"
      vQueryString = "&Document=" & GetTextBoxText("SearchDocument")
    End If
    vList("SystemColumns") = "Y"
    vList("WebPageItemNumber") = Me.WebPageItemNumber
    If InitialParameters.ContainsKey("Topic") Then vList("Topic") = InitialParameters("Topic").ToString

    'Find List of valid view names for the logged-in web user
    Dim vViews As String = String.Empty
    vViews = FindValidViewsNamesForUser()
    If Not String.IsNullOrEmpty(vViews) Then
      vViews = "'" & vViews.Replace(",", "','") & "'"
      vList.Add("Views", vViews)
    End If

    Dim vBaseList As BaseDataList = CType(FindControlByName(tblDataEntry, "DocumentsData"), BaseDataList)
    Dim vDownloadPageNumber As String = String.Empty
    Dim vCount As Long = 0
    If vBaseList IsNot Nothing Then
      vCount = DataHelper.GetPagedFinderData(CareNetServices.XMLDataFinderTypes.xdftWebDocuments, vBaseList, Request, plcHolder, vList, IntegerValue(InitialParameters("ItemsPerPage").ToString), 0, False, vQueryString)
      If vCount > 0 Then
        If (Not InitialParameters.ContainsKey("DisplayFormat")) OrElse InitialParameters("DisplayFormat").ToString = "0" Then
          Dim vDataGrid As DataGrid = CType(vBaseList, DataGrid)
          vDataGrid.Visible = True
          vDownloadPageNumber = InitialParameters.OptionalValue("DownloadPage").ToString
          Dim vImagePath As String = String.Empty.ToString
          Dim vImagePos As Integer = -1

          If vDownloadPageNumber.Length > 0 And Not String.IsNullOrEmpty(InitialParameters("HyperlinkText").ToString()) Then
            Dim vDataSet As DataSet = TryCast(vDataGrid.DataSource, DataSet)
            If vDataSet IsNot Nothing AndAlso vDataSet.Tables.Contains("DataRow") Then
              Dim vColumn As New BoundColumn()
              vDataGrid.Columns.AddAt(0, vColumn)
              vBaseList.DataBind()
              For vRow As Integer = 0 To vDataGrid.Items.Count - 1
                If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
                  vDataGrid.Items(vRow).Cells(0).Text = String.Format("<input type=""button"" class=""Button"" value='{0}' onclick=""location.href='Default.aspx?pn={1}&WDN={2}'"">", InitialParameters("HyperlinkText").ToString(), vDownloadPageNumber, vDataSet.Tables("DataRow").Rows(vRow).Item("WebDocumentNumber"))
                Else
                  vDataGrid.Items(vRow).Cells(0).Text = String.Format("<a href='Default.aspx?pn={0}&WDN={1}'>{2}</a>", vDownloadPageNumber, vDataSet.Tables("DataRow").Rows(vRow).Item("WebDocumentNumber"), InitialParameters("HyperlinkText").ToString())
                End If
              Next
            End If
          End If
          For vColCount As Integer = 0 To vDataGrid.Columns.Count - 1
            Dim vBoundColumn As BoundColumn = DirectCast(vDataGrid.Columns(vColCount), BoundColumn)
            If vBoundColumn.DataField = "ImageName" Then
              vImagePos = vColCount
            End If
          Next
          For vRow As Integer = 0 To vDataGrid.Items.Count - 1
            If vImagePos >= 0 Then
              vImagePath = String.Format("Images/Downloads/{0}", vDataGrid.Items(vRow).Cells(vImagePos).Text)
              vDataGrid.Items(vRow).Cells(vImagePos).Text = GetImage(vImagePath, vDataGrid.Items(vRow).Cells(vImagePos).Text, "Images/Downloads/default.png", "DownloadImage")
            End If
          Next
        End If
      Else
        'Display Message 
        SetLabelVisible("WarningMessage")
        vBaseList.Visible = False
        DirectCast(FindControlByName(Me, "plcHolder"), PlaceHolder).Controls.Clear()
      End If
    End If
  End Sub

  Public Overrides Sub HandleDataListItemDataBound(ByVal e As System.Web.UI.WebControls.DataListItemEventArgs)
    If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
      Dim vDownloadPageNumber As String = InitialParameters.OptionalValue("DownloadPage").ToString
      Dim vDrv As DataRowView = CType(e.Item.DataItem, DataRowView)
      'Add the Download link at the end
      If vDownloadPageNumber.Length > 0 Then
        Dim vCount As Integer = e.Item.Controls.Count
        Dim vDownloadLink As New Literal
        If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
          vDownloadLink.Text = String.Format("<input type=""button"" class=""Button"" value='{0}' onclick=""location.href='Default.aspx?pn={1}&WDN={2}'"">", InitialParameters("HyperlinkText").ToString(), vDownloadPageNumber, vDrv.Row("WebDocumentNumber"))
        Else
          vDownloadLink.Text = String.Format("<a href='Default.aspx?pn={0}&WDN={1}'>{2}</a>", vDownloadPageNumber, vDrv.Row("WebDocumentNumber"), InitialParameters("HyperlinkText").ToString())
        End If
        If vCount > 0 Then e.Item.Controls(vCount - 1).Parent.Controls.Add(vDownloadLink)
      End If
    End If
  End Sub

  Private Function FindValidViewsNamesForUser() As String
    Dim vViews As String = String.Empty
    If HttpContext.Current.User.Identity.IsAuthenticated Then
      If Not TypeOf (HttpContext.Current.User.Identity) Is System.Security.Principal.WindowsIdentity Then
        Dim vIdentity As FormsIdentity = CType(HttpContext.Current.User.Identity, FormsIdentity)
        If vIdentity.Ticket.UserData.Length > 0 Then
          Dim vItems As String() = vIdentity.Ticket.UserData.Split("|"c)
          If vItems.Length > 4 Then 'Check if viewname exists in Userdata
            vViews = vItems(4).ToString()
          End If
        End If
      End If
    End If
    Return vViews
  End Function
  Private Sub SetLabelVisible(ByVal pMessageControl As String, Optional ByVal pVisible As Boolean = True)
    If FindControlByName(Me, pMessageControl) IsNot Nothing Then
      DirectCast(FindControlByName(Me, pMessageControl), Label).Visible = pVisible
    End If
  End Sub
End Class