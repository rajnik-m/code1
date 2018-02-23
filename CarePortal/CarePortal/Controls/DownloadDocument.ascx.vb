Imports System.IO

Public Class DownloadDocument
  Inherits CareWebControl

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctDownloadDocument, tblDataEntry)
      SetLabelmessage("WarningMessage1", False)
      SetLabelmessage("PageError", False)
      If Not ValidQueryParameters() Then
        SetLabelmessage("WarningMessage1")
      Else
        Dim vList As New ParameterList(HttpContext.Current)
        vList("WebDocumentNumber") = Request.QueryString("WDN")
        'Find List of valid view names for the logged-in web user
        Dim vViews As String = String.Empty
        vViews = FindValidViewsNamesForUser()
        If Not String.IsNullOrEmpty(vViews) Then
          If Not String.IsNullOrEmpty(vViews) Then
            vViews = "'" & vViews.Replace(",", "','") & "'"
            vList.Add("Views", vViews)
          End If
        End If
        Dim vRow As DataRow
        vRow = DataHelper.GetRowFromDataTable(GetDataTable(DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebDocuments, vList)))
        If vRow IsNot Nothing Then
          'Get the Downloadable documents directory from Configuration option
          Dim vLocation As String = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.web_documents_directory)
          If vLocation.Length > 0 Then
            'If Location has value in the Configuration
            If DownloadFile(vLocation, vRow.Item("FileName").ToString(), vRow.Item("MimeType").ToString, CLng(vRow.Item("WebDocumentNumber").ToString)) Then
              'If User is Login Add the User Contact Number and the Address Number for the Journal Entry
              If UserContactNumber() > 0 And UserAddressNumber() > 0 Then
                vList("ContactNumber") = UserContactNumber()
                vList("AddressNumber") = UserAddressNumber()
              End If
              'Update the Download Count and Last downloaded on
              vList("Downloaded") = "Y"
              DataHelper.UpdateWebDocument(vList)
              GoToSubmitPage()
            Else
              'If File does not Exists on the server raise error
              If FindControlByName(Me, "PageError") IsNot Nothing Then
                DirectCast(FindControlByName(Me, "PageError"), ITextControl).Text = String.Format("Document number {0} is not available.", CLng(vRow.Item("WebDocumentNumber").ToString))
                DirectCast(FindControlByName(Me, "PageError"), Label).Visible = True
              End If
            End If
          End If
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Function ValidQueryParameters() As Boolean
    If Not Request.QueryString("WDN") Is Nothing Then
      If Request.QueryString("WDN").Length <= 0 Then
        Return False
      Else
        Dim vValidNumber As Boolean = Integer.TryParse(Request.QueryString("WDN"), 0)
        If Not vValidNumber Then
          Return False
        ElseIf IntegerValue(Request.QueryString("WDN").ToString) <= 0 Then
          Return False
        Else
          Return True
        End If
      End If
    Else
      Return False
    End If
  End Function

  Private Function DownloadFile(ByVal pPath As String, ByVal pFileName As String, ByVal pMimeType As String, ByVal pWebDocumentNumber As Long) As Boolean
    Try
      If File.Exists(String.Format("{0}{1}", pPath, pFileName)) Then
        Dim vFs As FileStream
        vFs = File.Open(Path.Combine(pPath, pFileName), FileMode.Open)
        Dim vBytes(CInt(vFs.Length)) As Byte
        vFs.Read(vBytes, 0, CInt(vFs.Length))
        vFs.Close()
        Response.Clear()
        'Add a header to tell the client what the filename should be
        Response.AddHeader("Content-disposition", "attachment;filename=" & pFileName)
        'Next tell it thefile type-this should be the mime type
        Response.ContentType = pMimeType
        Response.BinaryWrite(vBytes)
        Response.Flush()
        Return True
      Else
        Return False
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Function

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

  Private Sub SetLabelmessage(ByVal pMessageControl As String, Optional ByVal pVisible As Boolean = True)
    If FindControlByName(Me, pMessageControl) IsNot Nothing Then
      DirectCast(FindControlByName(Me, pMessageControl), Label).Visible = pVisible
    End If
  End Sub
End Class