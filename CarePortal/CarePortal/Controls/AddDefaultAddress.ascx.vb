Partial Public Class AddDefaultAddress
  Inherits CareWebControl

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    InitialiseControls(CareNetServices.WebControlTypes.wctAddNewDefaultAddress, tblDataEntry)
  End Sub

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Public Overrides Sub ProcessSubmit()
    Dim vList As New ParameterList(HttpContext.Current)
    vList("ContactNumber") = UserContactNumber()
    AddOptionalTextBoxValue(vList, "Address")
    AddOptionalTextBoxValue(vList, "Town")
    If Not vList.Contains("Town") Then vList.Add("Town", "#")
    AddOptionalTextBoxValue(vList, "County")
    AddOptionalTextBoxValue(vList, "Postcode")
    vList("Country") = GetDropDownValue("Country")
    vList("ValidFrom") = Date.Today.ToShortDateString
    vList("Default") = "Y"
    Dim vPAFStatus As String = GetLabelText("PafStatus")
    If vPAFStatus.Length > 0 Then vList("PafStatus") = vPAFStatus
    DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctAddresses, vList)
  End Sub

End Class