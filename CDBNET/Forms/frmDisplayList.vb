Friend Class frmDisplayList
  Inherits CDBNETCL.frmDisplayList

  Protected Sub New()
    MyBase.new()
  End Sub

  Public Sub New(ByVal pListUsage As ListUsages)
    MyBase.New(pListUsage)
  End Sub

  Public Sub New(ByVal pListUsage As ListUsages, ByVal pDataSource As DashboardDataSource, ByVal pList As ParameterList)
    MyBase.New(pListUsage, pDataSource, pList)
  End Sub
  Public Sub New(ByVal pListUsage As ListUsages, ByVal pList As ParameterList)
    MyBase.New(pListUsage, pList)
  End Sub
  Protected Overrides Sub ProcessNew()
    MyBase.ProcessNew()
    If mvListUsage = ListUsages.FastDataEntry Then
      Dim vList As ParameterList
      Dim vFrmAP As New frmApplicationParameters(CareServices.FunctionParameterTypes.fptAddFastDataEntryPage, Nothing, Nothing)
      If vFrmAP.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        vList = vFrmAP.ReturnList
        Dim vFrmFDE As New frmFastDataEntry(vList("FdePageTitle"), vList("FdePageName"))
        If vFrmFDE.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
          'Need to update lstAvailable
          Dim vLookupItem As New LookupItem(vFrmFDE.PageNumber.ToString, vList("FdePageName"))
          AddItemToAvailableList(vLookupItem)
        End If
      End If
    ElseIf mvListUsage = ListUsages.TraderMaintenance Then
      Dim vForm As New frmFPApplication(0)
      vForm.ShowDialog()
    End If
  End Sub

  Protected Overrides Sub EditRecord(ByVal pLookupItem As LookupItem)
    MyBase.EditRecord(pLookupItem)
    Dim vList As New ParameterList(True, True)
    If mvListUsage = ListUsages.FastDataEntry Then
      vList.IntegerValue("FdePageNumber") = IntegerValue(pLookupItem.LookupCode)
      Dim vDT As DataTable = DataHelper.GetFastDataEntryData(CareNetServices.XMLFastDataEntryTypes.fdePages, vList)
      If vDT IsNot Nothing AndAlso vDT.Rows.Count > 0 Then
        Dim vFrmFDE As New frmFastDataEntry(vDT.Rows(0), True)
        vFrmFDE.ShowDialog(Me)
      End If
    ElseIf mvListUsage = ListUsages.TraderMaintenance Then
      Dim vForm As New frmFPApplication(IntegerValue(pLookupItem.LookupCode))
      vForm.ShowDialog()
    End If
  End Sub

  Protected Overrides Sub DeleteRecord(ByVal pLookupItem As CDBNETCL.LookupItem)
    MyBase.DeleteRecord(pLookupItem)
    Dim vList As New ParameterList(True, True)
    If mvListUsage = ListUsages.FastDataEntry Then
      vList.IntegerValue("FdePageNumber") = IntegerValue(pLookupItem.LookupCode)
      DataHelper.DeleteFastDataEntryData(CareNetServices.XMLFastDataEntryTypes.fdePages, vList)
    End If
  End Sub

  Private Sub InitializeComponent()
    Me.SuspendLayout()
    '
    'frmDisplayList
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.ClientSize = New System.Drawing.Size(592, 517)
    Me.Name = "frmDisplayList"
    Me.ResumeLayout(False)

  End Sub
End Class