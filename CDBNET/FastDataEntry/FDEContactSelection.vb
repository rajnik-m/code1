Friend Class FDEContactSelection
  Inherits CareFDEControl

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pEditing)
    mvSupportsSelectionChanged = True
  End Sub

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pInitialSettings As String, ByVal pDefaultSettings As String, ByVal pFDEPageNumber As Integer, ByVal pSequenceNumber As Integer, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pInitialSettings, pDefaultSettings, pFDEPageNumber, pSequenceNumber, pEditing)
    mvSupportsSelectionChanged = True
  End Sub

  Public Overrides Property Enabled As Boolean
    Get
      Return MyBase.Enabled
    End Get
    Set(ByVal value As Boolean)
      epl.FindPanelControl(Of TextLookupBox)("ContactNumber").EnabledProperty = value
    End Set
  End Property


End Class
