Partial Public Class SubmitEnquiry
  Inherits CareWebControl
  Implements ICareParentWebControl

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    InitialiseControls(CareNetServices.WebControlTypes.wctSubmitEnquiry, tblDataEntry, "Subject,Precis", "DirectNumber,MobileNumber")
  End Sub

  Public Overrides Sub ProcessSubmit()
    Dim vReturnList As ParameterList = AddNewContact()
    Dim vList As New ParameterList(HttpContext.Current)
    vList("SenderContactNumber") = vReturnList("ContactNumber")
    vList("SenderAddressNumber") = vReturnList("AddressNumber")
    vList("Dated") = Date.Today.ToShortDateString
    vList("Direction") = "I"
    AddOptionalTextBoxValue(vList, "DocumentSubject")
    AddOptionalTextBoxValue(vList, "Precis")
    AddDefaultParameters(vList)
    vList("AddresseeContactNumber") = vList("ContactNumber")
    vList("AddresseeAddressNumber") = vList("AddressNumber")
    vList.Remove("ContactNumber")
    vList.Remove("AddressNumber")
    DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctDocument, vList)
    ProcessChildControls(vReturnList)
  End Sub

  Public Sub ProcessChildControls(ByVal pList As ParameterList) Implements ICareParentWebControl.ProcessChildControls
    SubmitChildControls(pList)
  End Sub
End Class