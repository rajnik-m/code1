Public Class FinancialSubMenu
  Inherits BaseFinancialMenu

  Public Shadows Event MenuSelected(ByVal pItem As FinancialMenuItems, ByVal pDataRow As DataRow)

  Public Sub New(ByVal pParent As MaintenanceParentForm, ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo)
    MyBase.New(pParent, pDST, pContactInfo)
  End Sub

  Public Sub SetContext(ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pDataRow As DataRow, ByVal pContactInfo As ContactInfo, ByVal pReadOnly As Boolean)
    SetContext(pDST, pDataRow, pContactInfo, pReadOnly, False)
  End Sub

  Public Sub SetContext(ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pDataRow As DataRow, ByVal pContactInfo As ContactInfo, ByVal pReadOnly As Boolean, ByVal pMultiSelect As Boolean)
    mvDataType = pDST
    mvDataRow = pDataRow
    mvContactInfo = pContactInfo
    mvReadOnly = pReadOnly
    mvMultiSelect = pMultiSelect
    mvSubMenu = True
  End Sub
 
  Private Sub FinancialSubMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Opening
    SetVisibleItems(e)
  End Sub

  Protected Overrides Sub MenuHandler(ByVal pMenuItem As ToolStripMenuItem, ByVal pItem As FinancialMenuItems)
    Dim vCursor As New BusyCursor
    Try
      RaiseEvent MenuSelected(pItem, mvDataRow)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
End Class
