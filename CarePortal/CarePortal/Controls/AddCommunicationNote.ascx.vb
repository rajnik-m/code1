Partial Public Class AddCommunicationNote
  Inherits CareWebControl
  Implements ICareChildWebControl

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    mvNeedsParent = True
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctAddCommunicationNote, tblDataEntry)
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    SetDefaults(False)
  End Sub

  Public Overrides Sub ClearControls(ByVal pClearLabels As Boolean, ByVal pErrorLabel As Label)
    MyBase.ClearControls(pClearLabels, pErrorLabel)
    SetDefaults(True)
  End Sub

  Private Sub SetDefaults(ByVal pIgnorePostBack As Boolean)
    If Not IsPostBack OrElse pIgnorePostBack Then
      SetDropDownText("Direction", InitialParameters("Direction").ToString)
      If BooleanValue(DefaultParameters("DefaultTopicSubTopic").ToString) Then
        SetDropDownText("Topic", DefaultParameters("Topic").ToString, True)
        SetDropDownText("SubTopic", DefaultParameters("SubTopic").ToString)
      End If
    End If
  End Sub
  Public Sub SubmitChild(ByVal pList As ParameterList) Implements ICareChildWebControl.SubmitChild
    Dim vContactNumber As Integer
    Dim vAddressNumber As Integer
    If Me.ParentGroup = "SelectedContact" AndAlso Session("SelectedContactNumber") IsNot Nothing AndAlso Session("SelectedContactNumber").ToString.Length > 0 Then
      vContactNumber = IntegerValue(Session("SelectedContactNumber").ToString)
      vAddressNumber = IntegerValue(GetContactAddress(vContactNumber))
    Else
      vContactNumber = IntegerValue(pList("ContactNumber").ToString)
      vAddressNumber = IntegerValue(pList("AddressNumber").ToString)
    End If
    Dim vValue As String = GetDropDownValue("Direction")
    If vValue.Length = 0 Then vValue = "I"
    Dim vList As New ParameterList(HttpContext.Current)
    AddOptionalTextBoxValue(vList, "Dated")
    AddOptionalTextBoxValue(vList, "DocumentSubject")
    AddOptionalTextBoxValue(vList, "Precis")
    If vList.ContainsKey("Precis") And Not vList.ContainsKey("DocumentSubject") Then
      If vList("Precis").ToString.Length > 80 Then
        vList("DocumentSubject") = vList("Precis").ToString.Substring(0, 80)
      Else
        vList("DocumentSubject") = vList("Precis").ToString
      End If
    End If
    If vValue = "I" Then
      vList("AddresseeContactNumber") = UserContactNumber()
      vList("AddresseeAddressNumber") = UserAddressNumber()
      vList("SenderContactNumber") = vContactNumber
      vList("SenderAddressNumber") = vAddressNumber
    Else
      vList("SenderContactNumber") = UserContactNumber()
      vList("SenderAddressNumber") = UserAddressNumber()
      vList("AddresseeContactNumber") = vContactNumber
      vList("AddresseeAddressNumber") = vAddressNumber
    End If
    AddDefaultParameters(vList)
    vList.Remove("DefaultTopicSubTopic")
    vList("Direction") = vValue
    AddOptionalDropDownValue(vList, "Topic")
    AddOptionalDropDownValue(vList, "SubTopic")
    DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctDocument, vList)
  End Sub

End Class