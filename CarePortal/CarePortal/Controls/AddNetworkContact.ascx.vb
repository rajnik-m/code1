Partial Public Class AddNetworkContact
  Inherits CareWebControl
  Implements ICareParentWebControl

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    InitialiseControls(CareNetServices.WebControlTypes.wctAddNetworkContact, tblDataEntry, "Notes", "DirectNumber,MobileNumber")
  End Sub

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Public Overrides Sub ProcessSubmit()
    Dim vReturnList As ParameterList = AddNewContact()
    'Assume contact has been added then need to add links
    Dim vLinkList As New ParameterList(HttpContext.Current)
    vLinkList("ContactNumber") = UserContactNumber()
    vLinkList("ContactNumber2") = vReturnList("ContactNumber")
    vLinkList("Relationship") = DefaultParameters("Relationship")
    vLinkList("ValidFrom") = Date.Today
    vLinkList("Notes") = GetTextBoxText("Notes")
    DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctLink, vLinkList)
    Dim vSuppList As New ParameterList(HttpContext.Current)
    vSuppList("ContactNumber") = vReturnList("ContactNumber")
    vSuppList("Suppression") = DefaultParameters("MailingSuppression")
    vSuppList("ValidFrom") = Date.Today.ToShortDateString
    vSuppList("ValidTo") = System.DateTime.Now.AddYears(100).ToShortDateString
    DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctSuppression, vSuppList)
    ProcessChildControls(vReturnList)
  End Sub

  Public Sub ProcessChildControls(ByVal pList As ParameterList) Implements ICareParentWebControl.ProcessChildControls
    SubmitChildControls(pList)
  End Sub
End Class