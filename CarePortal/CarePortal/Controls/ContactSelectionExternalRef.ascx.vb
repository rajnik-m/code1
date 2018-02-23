Partial Public Class ContactSelectionExternalRef

  Inherits CareWebControl

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctContactSelectionExternalRef, tblDataEntry)
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Public Overrides Sub ProcessSubmit()
    If DependantControls.Count > 0 Then
      Dim vExternalReference As String = GetTextBoxText("ExternalReference")
      If vExternalReference.Length > 0 Then
        Dim vList As New ParameterList(HttpContext.Current)
        vList("DataSource") = InitialParameters("DataSource").ToString
        vList("ExternalReference") = vExternalReference
        Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.FindDataTable(CareNetServices.XMLDataFinderTypes.xdftContacts, vList))
        If vRow IsNot Nothing AndAlso vRow("OwnershipAccessLevel").ToString = "W" Then
          vList("ContactNumber") = vRow("ContactNumber").ToString
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
        Dim vControl As Control = FindControlByName(Me, "ExternalReference")
        If vControl IsNot Nothing Then
          Dim vList As New ParameterList(HttpContext.Current)
          vList("DataSource") = InitialParameters("DataSource").ToString
          Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactExternalReferences, IntegerValue(Session("CurrentContactNumber").ToString), vList))
          Dim vExternalReference As String = ""
          If vRow IsNot Nothing Then vExternalReference = vRow("ExternalReference").ToString
          Dim vTextBox As TextBox = TryCast(vControl, TextBox)
          If vTextBox IsNot Nothing Then ExternalReferenceChanged(vTextBox, vExternalReference)
        End If
      End If
    End If
  End Sub
End Class