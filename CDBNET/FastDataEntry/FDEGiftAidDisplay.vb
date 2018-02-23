Friend Class FDEGiftAidDisplay
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

    Dim vCombo As ComboBox = epl.FindPanelControl(Of ComboBox)("DeclarationNumber")
    Dim vList As New ParameterList(True)
    vList("FastDataEntry") = "Y"
    Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactGiftAidDeclarations, pContactInfo.ContactNumber, vList)
    Dim vDT As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
    If pContactInfo.ContactType = ContactInfo.ContactTypes.ctOrganisation Then
      If vDT Is Nothing Then vDT = DataHelper.NewDataTable(vDataSet.Tables("Column"))
      Dim vRow As DataRow = vDT.NewRow
      vRow.Item("DeclarationNumber") = "0"
      vRow.Item("Summary") = "Organisation"
      vDT.Rows.Add(vRow)
    End If
    With vCombo
      .DisplayMember = "Summary"
      .ValueMember = "DeclarationNumber"
      .DataSource = vDT
    End With
    Dim vControl As Control = FindControl(epl, "Edit", False)
    If vControl IsNot Nothing Then
      vControl.Enabled = (mvContactInfo.ContactType = ContactInfo.ContactTypes.ctContact)
    End If
  End Sub
End Class
