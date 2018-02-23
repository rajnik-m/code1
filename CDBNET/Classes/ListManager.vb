Public Class ListManager
  Private WithEvents mvListManager As frmListManager
  Private mvBrowserMenu As BrowserMenu

  Public Sub New(ByVal pCriteriaSet As Integer, ByVal pCriteriaDesc As String)         'Use from campaign manager
    mvListManager = New frmListManager(True, pCriteriaSet, pCriteriaDesc)
    mvBrowserMenu = New BrowserMenu(Nothing)
    mvListManager.BrowserMenuStrip = mvBrowserMenu
  End Sub

  Public Sub New(ByVal pSelectionSet As Integer, ByVal pNewSelectionSet As Boolean)
    mvListManager = New frmListManager(True, pSelectionSet, pNewSelectionSet)            'Use from non-campaign manager
    mvBrowserMenu = New BrowserMenu(Nothing)
    mvListManager.BrowserMenuStrip = mvBrowserMenu
  End Sub

  Public Sub Show()
    mvListManager.Show()
  End Sub

  Public Sub ShowDialog()
    mvListManager.ShowDialog()
  End Sub

  Private Sub mvListManager_ContactSelected(ByVal sender As Object, ByVal pContactNumber As Integer) Handles mvListManager.ContactSelected
    FormHelper.ShowContactCardIndex(pContactNumber)
  End Sub

  Private Sub mvListManager_SetBrowserMenuContext(ByVal sender As Object, ByVal pEntityType As CDBNETCL.HistoryEntityTypes, ByVal pItemNumber As Integer) Handles mvListManager.SetBrowserMenuContext
    mvBrowserMenu.EntityType = pEntityType
    mvBrowserMenu.ItemNumber = pItemNumber
  End Sub
End Class
