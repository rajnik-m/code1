Friend Class FDEAddTransactionDetails
  Inherits CareFDEControl

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pEditing)
    mvSupportsContactData = True
    mvSupportsAddressData = True
    mvSupportsReferenceData = True
  End Sub

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pInitialSettings As String, ByVal pDefaultSettings As String, ByVal pFDEPageNumber As Integer, ByVal pSequenceNumber As Integer, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pInitialSettings, pDefaultSettings, pFDEPageNumber, pSequenceNumber, pEditing)
    mvSupportsContactData = True
    mvSupportsAddressData = True
    mvSupportsReferenceData = True
  End Sub

  Friend Overrides Sub RefreshContactData(ByVal pContactInfo As CDBNETCL.ContactInfo)
    MyBase.RefreshContactData(pContactInfo)
    mvContactInfo.SelectedAddressNumber = mvContactInfo.AddressNumber
    Dim vTLB As TextLookupBox = epl.FindPanelControl(Of TextLookupBox)("MailingContactNumber", False)
    If vTLB IsNot Nothing Then vTLB.text = pContactInfo.ContactNumber.ToString
    Dim vAddrTLB As TextLookupBox = epl.FindPanelControl(Of TextLookupBox)("MailingAddressNumber", False)
    If vAddrTLB IsNot Nothing Then vAddrTLB.FillComboFromContactData(pContactInfo)
  End Sub

  Friend Overrides Sub RefreshAddressData(ByVal pAddressNumber As Integer)
    MyBase.RefreshAddressData(pAddressNumber)
    mvContactInfo.SelectedAddressNumber = pAddressNumber
  End Sub

  Friend Overrides Sub SetDefaults()
    epl.FillDeferredCombos(epl)
    MyBase.SetDefaults()
    RaiseValueChangedEvent(epl, "Source", epl.GetValue("Source"))
  End Sub

  Friend Overrides Sub SetReferenceMandatory(ByVal pMandatory As Boolean)
    If pMandatory Then
      epl.PanelInfo.PanelItems("Reference").Mandatory = True
    Else
      epl.PanelInfo.PanelItems("Reference").Mandatory = False
      epl.SetErrorField("Reference", "")
    End If
  End Sub

  Friend Overrides Function BuildParameterList(ByRef pList As CDBNETCL.ParameterList) As Boolean
    Dim vValid As Boolean = MyBase.BuildParameterList(pList)
    If mvContactInfo IsNot Nothing Then
      pList.IntegerValue("PayerContactNumber") = mvContactInfo.ContactNumber
      pList.IntegerValue("PayerAddressNumber") = mvContactInfo.SelectedAddressNumber
      pList.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber
      pList.IntegerValue("AddressNumber") = mvContactInfo.SelectedAddressNumber
    End If
    Return vValid
  End Function

End Class
