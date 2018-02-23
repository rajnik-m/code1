Partial Public Class ContactSelection
  Inherits CareWebControl

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctContactSelection, tblDataEntry)
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Public Overrides Sub ProcessSubmit()
    If DependantControls.Count > 0 Then
      Dim vContactNumber As Integer = IntegerValue(GetTextBoxText("ContactNumber"))
      If vContactNumber > 0 Then
        Dim vList As New ParameterList(HttpContext.Current)
        vList("ContactNumber") = vContactNumber.ToString
        Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vContactNumber))
        If vRow IsNot Nothing Then
          vList("AddressNumber") = vRow("AddressNumber").ToString
          For Each vCareWebControl As ICareChildWebControl In DependantControls
            vCareWebControl.SubmitChild(vList)
          Next
        End If
      End If
    End If
  End Sub

 Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    If Not IsPostBack Then
      If Session("CurrentContactNumber") IsNot Nothing AndAlso CInt(Session("CurrentContactNumber")) > 0 Then
        Dim vControl As Control = FindControlByName(Me, "ContactNumber")
        If vControl IsNot Nothing Then
          Dim vTextBox As TextBox = TryCast(vControl, TextBox)
          If vTextBox IsNot Nothing Then ContactNumberChanged(vTextBox, CInt(Session("CurrentContactNumber")))
        End If
      End If
    End If
  End Sub
End Class