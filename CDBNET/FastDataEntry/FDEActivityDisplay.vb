Friend Class FDEActivityDisplay
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
    Dim vControl As Control = FindControl(epl, "ActivityValue")
    Dim vUseCombo As Boolean = (TypeOf (vControl) Is ComboBox)
    Dim vList As New ParameterList(True)

    If mvDefaultSettings.Length > 0 Then
      Dim vParamList As New ParameterList
      vParamList.FillFromValueList(mvDefaultSettings)
      If vParamList.ContainsKey("ActivityGroup") Then vList("ActivityGroup") = vParamList("ActivityGroup")
    End If
    vList("Current") = "Y"
    Dim vDS As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactCategories, pContactInfo.ContactNumber, vList)

    If vUseCombo Then
      Dim vDT As DataTable = DataHelper.GetTableFromDataSet(vDS)
      Dim vCombo As ComboBox = DirectCast(vControl, ComboBox)
      With vCombo
        .DisplayMember = "ActivityDesc"
        .ValueMember = "ActivityCode"
        .DataSource = vDT
      End With
    Else
      Dim vGrid As DisplayGrid = DirectCast(vControl, DisplayGrid)
      vGrid.Populate(vDS)
    End If

  End Sub
End Class
