Public Class AddOrganisation
  Inherits CareWebControl
  Implements ICareParentWebControl

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub
  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctAddOrganisation, tblDataEntry)
    Catch vEx As ThreadAbortException
      Throw vEx
    End Try
  End Sub
  Public Overrides Sub ProcessSubmit()
    Dim vOrganisationList As New ParameterList(HttpContext.Current)
    Dim vReturnList As New ParameterList(HttpContext.Current)
    vOrganisationList = GetAddOrganisationParameterList()
    vOrganisationList.Add("UserID", UserContactNumber.ToString)
    vReturnList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctOrganisation, vOrganisationList)
    'Adding position for user at the organisation
    AddContactPosition(vReturnList)
    ProcessChildControls(vReturnList)
  End Sub
  Public Sub ProcessChildControls(ByVal pList As ParameterList) Implements ICareParentWebControl.ProcessChildControls
    If Not InWebPageDesigner() Then
      Dim vSubmitParams As New StringBuilder
      With vSubmitParams
        .Append("&ON=")
        .Append(pList("ContactNumber"))
        .Append("&AN=")
        .Append(pList("AddressNumber"))
      End With
      GoToSubmitPage(vSubmitParams.ToString)
    End If
  End Sub
  Private Sub AddContactPosition(ByVal pList As ParameterList)
    Dim vList As New ParameterList(HttpContext.Current)
    vList("ContactNumber") = UserContactNumber().ToString
    vList("OrganisationNumber") = pList("ContactNumber")
    vList("AddressNumber") = pList("AddressNumber")
    vList("Position") = GetTextBoxText("Position")
    vList("ValidFrom") = TodaysDate()
    vList("UserID") = UserContactNumber.ToString
    DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctPosition, vList)
  End Sub
End Class