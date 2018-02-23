Public Class frmAuthoriseExpenses
  Private mvExpensesDataTable As DataTable
  Private mvDataChanged As Boolean
  Public Sub New(ByVal pEventNumber As Integer)

    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pEventNumber)
  End Sub

  Private Sub InitialiseControls(ByVal pEventNumber As Integer)
    SetControlTheme()
    Dim vProductDataSet As DataSet = DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventAuthoriseExpenses, pEventNumber)
    If Not vProductDataSet Is Nothing Then
      mvExpensesDataTable = DataHelper.GetTableFromDataSet(vProductDataSet)
      dgr.Populate(vProductDataSet)
      If dgr.RowCount > 0 Then dgr.SelectRow(-1)
      cmdAuthorise.Enabled = False
      cmdAuthoriseAll.Enabled = False
    End If
  End Sub

  Private Sub cmdAuthorise_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAuthorise.Click
    Try
      If dgr.RowCount > 0 Then
        Dim vSelectedRow As DataRow = mvExpensesDataTable.Rows(dgr.CurrentDataRow)
        mvDataChanged = True
        If vSelectedRow.Item("AuthorisedOn").ToString = String.Empty And vSelectedRow.Item("AuthorisedBy").ToString = String.Empty Then
          vSelectedRow.Item("AuthorisedOn") = Today.ToString(AppValues.DateFormat)
          vSelectedRow.Item("AuthorisedBy") = AppValues.Logname
        Else
          vSelectedRow.Item("AuthorisedOn") = String.Empty
          vSelectedRow.Item("AuthorisedBy") = String.Empty
        End If
      End If
    Catch vException As CareException
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdAuthoriseAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAuthoriseAll.Click
    Try
      If dgr.RowCount > 0 Then
        Dim vRowCounter As Integer
        mvDataChanged = True
        For vRowCounter = 0 To mvExpensesDataTable.Rows.Count - 1
          Dim vSelectedRow As DataRow = mvExpensesDataTable.Rows(vRowCounter)
          If vSelectedRow.Item("AuthorisedOn").ToString = String.Empty And vSelectedRow.Item("AuthorisedBy").ToString = String.Empty Then
            vSelectedRow.Item("AuthorisedOn") = Today.ToString(AppValues.DateFormat)
            vSelectedRow.Item("AuthorisedBy") = AppValues.Logname
          Else
            vSelectedRow.Item("AuthorisedOn") = String.Empty
            vSelectedRow.Item("AuthorisedBy") = String.Empty
          End If
        Next
      End If
    Catch vException As CareException
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Try
      Me.Close()
    Catch vException As CareException
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub frmAuthoriseExpenses_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
    Try
      If mvDataChanged Then
        If ConfirmSave() Then
          e.Cancel = Not ProcessUpdate()
        End If
      End If
    Catch vException As CareException
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Try
      If mvDataChanged Then
        If ConfirmChanges() Then
          If ProcessUpdate() Then
            mvDataChanged = False 'allow the form to close
            Me.Close()
          End If
        Else
          mvDataChanged = False 'allow the form to close
          Me.Close()
        End If
      End If
    Catch vException As CareException
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function ProcessUpdate() As Boolean
    Dim vList As New ParameterList(True)
    For Each vRow As DataRow In mvExpensesDataTable.Rows
      If vRow("AuthorisedOn").ToString().Length > 0 Then
        vList("EventPersonnelNumber") = vRow("EventPersonnelNumber").ToString
        vList("ContactNumber") = vRow("ContactNumber").ToString
        vList("AddressNumber") = vRow("AddressNumber").ToString
        vList("AuthorisedOn") = vRow("AuthorisedOn").ToString
        vList("AuthorisedBy") = vRow("AuthorisedBy").ToString
        DataHelper.UpdateEventPersonnel(vList)
      End If
    Next
    mvDataChanged = False
    Return True
  End Function

  
  Private Sub dgr_RowSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgr.RowSelected
    Try
      If dgr.RowCount > 0 Then
        cmdAuthorise.Enabled = True
        cmdAuthoriseAll.Enabled = True
      End If
    Catch vException As CareException
      DataHelper.HandleException(vException)
    End Try
  End Sub

End Class