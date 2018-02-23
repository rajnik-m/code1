Friend Class FDEAddressDisplay
  Inherits CareFDEControl

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pEditing)
    mvSupportsContactData = True
  End Sub

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pInitialSettings As String, ByVal pDefaultSettings As String, ByVal pFDEPageNumber As Integer, ByVal pSequenceNumber As Integer, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pInitialSettings, pDefaultSettings, pFDEPageNumber, pSequenceNumber, pEditing)
    mvSupportsContactData = True
  End Sub

  Friend Overrides Sub RefreshContactData(ByVal pContactInfo As CDBNETCL.ContactInfo)
    MyBase.RefreshContactData(pContactInfo)
    Dim vTLB As TextLookupBox = epl.FindPanelControl(Of TextLookupBox)("AddressNumber")
    If vTLB IsNot Nothing Then vTLB.FillComboFromContactData(pContactInfo)
  End Sub

End Class
