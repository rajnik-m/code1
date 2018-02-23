Public Class frmLicenceMaintenance

#Region "Private Members"

  Private mvLicenceCount As Integer
  Private mvUserFound As Boolean
  Private mvRowChanged As Boolean 'prevent the selection of the first occurance of a user if multiple entries exist

#End Region

#Region "Constructor"

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub

#End Region

#Region "Private and Protected Methods"

  Private Sub InitialiseControls()
    SetControlTheme()

    cboUser.DisplayMember = "LogName"
    cboUser.ValueMember = "LogName"
    cboUser.DataSource = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtUsers)

    Dim vList As New ParameterList(True)
    cboModuleName.DisplayMember = "ModuleDesc"
    cboModuleName.ValueMember = "Module"
    cboModuleName.DataSource = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtModuleNames, vList)
    cboModuleName.SelectedValue = "CD" 'CARE

    If dgr.RowCount > 0 Then dgr.SelectRow(0)
  End Sub

  Private Sub DisplayGrid()
    Dim vActive As Integer
    Dim vlist As New ParameterList(True)
    vlist("Module") = cboModuleName.SelectedValue.ToString()
    dgr.Populate(DataHelper.GetSystemModuleUsersData(vlist))

    For vIndex As Integer = 0 To dgr.RowCount - 1
      If dgr.GetValue(vIndex, "NamedUser") = "Y" Then
        dgr.SetValue(vIndex, "NamedUser", "Yes")
      Else
        dgr.SetValue(vIndex, "NamedUser", String.Empty)
      End If
      If dgr.GetValue(vIndex, "Active") = "Y" Then
        dgr.SetValue(vIndex, "Active", "Yes")
        vActive += 1
      Else
        dgr.SetValue(vIndex, "Active", String.Empty)
      End If
    Next
    lblActiveUsers.Text = String.Format(ControlText.LblCurrentActiveUsers, vActive.ToString)

    If dgr.RowCount > 0 Then dgr.SelectRow(0)
  End Sub

#End Region

#Region "Control Events"

  Private Sub cmdDeDup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeDup.Click
    Try
      Dim vList As New ParameterList(True)
      If ShowQuestion(QuestionMessages.QmDeleteDuplicateSMUsers, MessageBoxButtons.YesNoCancel) = System.Windows.Forms.DialogResult.Yes Then
        DataHelper.DeleteDuplicateSystemModuleUsers(vList)
      End If
      DisplayGrid()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    Me.Close()
  End Sub

  Private Sub dgr_RowSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgr.RowSelected
    Try
      mvRowChanged = True
      cboUser.SelectedValue = dgr.GetValue(dgr.ActiveRow, "Logname")
      chkNamedUser.Checked = CBool(IIf(dgr.GetValue(pRow, "NamedUser") = "Yes", True, False))
      mvUserFound = True
      mvRowChanged = False
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cboModuleName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboModuleName.SelectedIndexChanged
    Try
      DisplayGrid()

      Dim vList As New ParameterList(True)
      vList("Module") = cboModuleName.SelectedValue.ToString
      mvLicenceCount = DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctModuleLicences, vList)
      lblNoOfLicences.Text = String.Format(ControlText.LblNoOfLicences, mvLicenceCount.ToString)
      mvUserFound = False

      If dgr.RowCount > 0 Then
        dgr.SelectRow(0)
      Else
        chkNamedUser.Checked = False
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub chkNamedUser_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNamedUser.CheckedChanged
    Try
      'Dim vFound As Boolean
      Dim vList As New ParameterList(True)
      Dim vNamedUser As Boolean
      If dgr.RowCount > 0 Then vNamedUser = CBool(IIf(dgr.GetValue(dgr.ActiveRow, "NamedUser") = "Yes", True, False))
      If dgr.RowCount = 0 OrElse
        (cboUser.SelectedValue IsNot Nothing AndAlso
         dgr.GetValue(dgr.ActiveRow, "Logname") <> cboUser.SelectedValue.ToString) OrElse
        vNamedUser <> chkNamedUser.Checked Then
        vList("Module") = cboModuleName.SelectedValue.ToString()
        vList("Logname") = cboUser.SelectedValue.ToString()
        If chkNamedUser.Checked Then
          Dim vCount As Integer = DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctSystemModuleUsers, vList)
          If (vCount >= mvLicenceCount) Then
            ShowWarningMessage(InformationMessages.ImLicencesInUse, cboUser.SelectedValue.ToString)
            chkNamedUser.Checked = False
          Else
            'if an entry is present in the grid update it else add a new one
            If mvUserFound Then
              vList("NamedUser") = CBoolYN(chkNamedUser.Checked)
              vList.Add("SystemStartTime", dgr.GetValue(dgr.ActiveRow, "StartTime"))
              DataHelper.UpdateSystemModuleUser(vList)
            Else
              DataHelper.AddSystemModuleUser(vList)
            End If
            DisplayGrid()
          End If
        ElseIf mvUserFound Then
          vList("NamedUser") = CBoolYN(chkNamedUser.Checked)
          vList.Add("SystemStartTime", dgr.GetValue(dgr.ActiveRow, "StartTime"))
          DataHelper.UpdateSystemModuleUser(vList)
          DisplayGrid()
        End If
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
      'Revert back on failure as it will throw record not found exception 
      'if the user tries again
      chkNamedUser.Checked = Not chkNamedUser.Checked
    End Try
  End Sub

  Private Sub dgr_RowDoubleClicked(ByVal sender As Object, ByVal pRow As Integer) Handles dgr.RowDoubleClicked
    Try
      If ShowQuestion(String.Format(QuestionMessages.QmDeleteInactiveSMUsers, dgr.GetValue(dgr.ActiveRow, "Logname")), MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
        Dim vList As New ParameterList(True)
        vList("Module") = cboModuleName.SelectedValue.ToString
        vList("Logname") = cboUser.SelectedValue.ToString
        vList("SystemStartTime") = dgr.GetValue(dgr.ActiveRow, "StartTime")
        DataHelper.DeleteInactiveSystemModuleUsers(vList)
        DisplayGrid()
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
    Try
      DisplayGrid()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cboUser_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboUser.SelectedIndexChanged
    Try
      If mvRowChanged Then
        mvUserFound = True
      Else
        mvUserFound = False
        'if the row exists in the grid then select it
        For vIndex As Integer = 0 To dgr.RowCount - 1
          If cboUser.SelectedValue.ToString = dgr.GetValue(vIndex, "Logname") Then
            dgr.SelectRow(vIndex)
            mvUserFound = True
            Exit For
          End If
        Next
        If Not mvUserFound Then chkNamedUser.Checked = False
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

#End Region

End Class