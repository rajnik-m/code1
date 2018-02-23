Partial Public Class RecordStandardEmail
  Inherits CareWebControl
  Implements ICareParentWebControl

  Private mvTopic As String
  Private mvSubTopic As String
  Private mvDocumentType As String
  Private mvDocumentClass As String

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctRecordStandardEmail, tblDataEntry, "Subject,Precis", "DirectNumber,MobileNumber")
      'Now we must get the precis and subject of the appropriate standard document
      GetStandardDocumentData()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Sub GetStandardDocumentData()
    Dim vList As New ParameterList(HttpContext.Current)
    vList("StandardDocument") = InitialParameters("StandardDocument")
    Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtStandardDocuments, vList)
    If vTable IsNot Nothing Then
      Dim vRow As DataRow = vTable.Rows(0)
      Dim vTextBox As TextBox = TryCast(Me.FindControl("Subject"), TextBox)
      If vTextBox IsNot Nothing Then vTextBox.Text = vRow("Subject").ToString
      vTextBox = TryCast(Me.FindControl("Precis"), TextBox)
      If vTextBox IsNot Nothing Then vTextBox.Text = vRow("Precis").ToString
      mvTopic = vRow("Topic").ToString
      mvSubTopic = vRow("SubTopic").ToString
      mvDocumentType = vRow("DocumentType").ToString
      mvDocumentClass = vRow("DocumentClass").ToString
      If mvDocumentClass.Length = 0 Then mvDocumentClass = "DC"
    End If
  End Sub

  Public Overrides Sub ProcessSubmit()
    SendEmailMessage()
    Dim vReturnList As ParameterList = AddNewContact()
    Dim vDocList As New ParameterList(HttpContext.Current)
    vDocList("SenderContactNumber") = vReturnList("ContactNumber")
    vDocList("SenderAddressNumber") = vReturnList("AddressNumber")
    Dim vEMailAddress As String = InitialParameters("EMailAddress").ToString
    If vEMailAddress.Length > 0 Then
      Dim vList As ParameterList = GetContactFromEMailAddress(vEMailAddress)
      If vList.Contains("ContactNumber") Then
        'We have found the contact
        vDocList("AddresseeContactNumber") = vList("ContactNumber")
        vDocList("AddresseeAddressNumber") = vList("AddressNumber")
      Else
        Throw New CareException(String.Format("Failed to find Contact with EMail Address {0}", vEMailAddress))
      End If
    Else
      Throw New CareException(String.Format("EMail Address Undefined"))
    End If
    vDocList("UserID") = vDocList("SenderContactNumber")
    vDocList("Topic") = mvTopic
    vDocList("SubTopic") = mvSubTopic
    vDocList("DocumentType") = mvDocumentType
    vDocList("DocumentClass") = mvDocumentClass
    vDocList("Direction") = "O"
    vDocList("Dated") = Now.ToShortDateString
    vDocList("DocumentSubject") = GetTextBoxText("Subject")
    vDocList("Precis") = GetTextBoxText("Precis")
    vDocList("StandardDocument") = InitialParameters("StandardDocument")
    DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctDocument, vDocList)
    ProcessChildControls(vReturnList)
  End Sub

  Public Sub ProcessChildControls(ByVal pList As ParameterList) Implements ICareParentWebControl.ProcessChildControls
    SubmitChildControls(pList)
  End Sub

  Private Sub SendEmailMessage()
    If DefaultParameters.ContainsKey("SendEMail") AndAlso DefaultParameters("SendEMail").ToString = "Y" Then
      Dim vFromAddress As String = GetTextBoxText("EMailAddress")
      Dim vToAddress As String
      If DefaultParameters.ContainsKey("TestEMailAddress") AndAlso DefaultParameters("TestEMailAddress").ToString.Length > 0 Then
        vToAddress = DefaultParameters("TestEMailAddress").ToString
      Else
        vToAddress = InitialParameters("EMailAddress").ToString
      End If
      Dim vSubject As String = GetTextBoxText("Subject")
      Dim vPrecis As String = GetTextBoxText("Precis")
      SendEmail(vFromAddress, vToAddress, vSubject, vPrecis)
    End If
  End Sub

  'First we must find out if the contact exists
  'Dim vEMailAddress As String = GetTextBoxText( "EMailAddress")
  '      If vEMailAddress.Length > 0 Then
  'Dim vList As ParameterList = GetContactFromEMailAddress(vEMailAddress)
  'Dim vDocList As New ParameterList(HttpContext.Current)
  '        If vList.Contains("ContactNumber") Then
  'We have found the contact
  '          vDocList("SenderContactNumber") = vList("ContactNumber")
  '          vDocList("SenderAddressNumber") = vList("AddressNumber")
  '        Else
  'We must create the contact
  'Dim vContactList As New ParameterList(HttpContext.Current)
  '          vContactList("EMailAddress") = vEMailAddress
  '          AddOptionalTextBoxValue( vContactList, "Forenames")
  '          AddOptionalTextBoxValue( vContactList, "Surname")
  '          AddOptionalTextBoxValue( vContactList, "Address")
  '          AddOptionalTextBoxValue( vContactList, "Town")
  '          AddOptionalTextBoxValue( vContactList, "County")
  '          AddOptionalTextBoxValue( vContactList, "Postcode")
  '          vContactList("Source") = DefaultParameters("Source")
  'Dim vDDL As DropDownList = TryCast(FindControlByName(tblDataEntry, "Country"), DropDownList)
  '          If vDDL IsNot Nothing Then vContactList("Country") = vDDL.SelectedItem.Value
  'Dim vContactResultList As ParameterList = DataHelper.AddItem(LookupData.XMLMaintenanceControlTypes.xmctContact, vContactList)
  '          vDocList("SenderContactNumber") = vContactResultList("ContactNumber")
  '          vDocList("SenderAddressNumber") = vContactResultList("AddressNumber")
  '        End If
  '        vDocList("UserID") = vDocList("SenderContactNumber")
  '        vEMailAddress = InitialParameters("EMailAddress").ToString
  '        If vEMailAddress.Length > 0 Then
  '          vList = GetContactFromEMailAddress(vEMailAddress)
  '          If vList.Contains("ContactNumber") Then
  'We have found the contact
  '            vDocList("AddresseeContactNumber") = vList("ContactNumber")
  '            vDocList("AddresseeAddressNumber") = vList("AddressNumber")
  '          End If
  '        End If
  '        vDocList("Topic") = mvTopic
  '        vDocList("SubTopic") = mvSubTopic
  '        vDocList("DocumentType") = mvDocumentType
  '        vDocList("DocumentClass") = mvDocumentClass
  '        vDocList("Direction") = "O"
  '        vDocList("Dated") = Now.ToShortDateString
  '        vDocList("DocumentSubject") = GetTextBoxText( "Subject")
  '        vDocList("Precis") = GetTextBoxText( "Precis")
  '        vDocList("StandardDocument") = InitialParameters("StandardDocument")
  'Dim vResultList As ParameterList = DataHelper.AddItem(LookupData.XMLMaintenanceControlTypes.xmctDocument, vDocList)

End Class