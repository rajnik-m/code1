Public Class frmGenMMerge
  Private mvMailingInfo As MailingInfo
  Private mvMergeType As Integer

  Private mvCurrentRecords As Boolean
  Private mvCommonRecords As Boolean
  Private mvNewRecords As Boolean
  Private mvList As ParameterList = Nothing

  Public Event ShowGeneralMailingForm(ByVal pGeneralMailing As CareNetServices.MailingTypes)

  Public Sub New(ByVal pMailingInfo As MailingInfo, ByVal pList As ParameterList)
    mvMailingInfo = pMailingInfo
    mvList = pList
    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub
  Private Sub InitialiseControls()
    SetControlTheme()
  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Dim vStoreGvRev As Integer
    Dim vMergeDesc As String = String.Empty
    Dim vResult As Integer

    Try
      vStoreGvRev = mvMailingInfo.Revision

      Select Case mvMergeType
        Case 0
          vMergeDesc = "A and B"
        Case 1
          vMergeDesc = "A or B"
        Case 2
          vMergeDesc = "B and not in A"
        Case 3
          vMergeDesc = "not in A and B"
        Case 4
          vMergeDesc = "A and not in B"
      End Select

      If mvList IsNot Nothing Then
        mvList("MergerMode") = vMergeDesc
        If Not mvList.Contains("CriteriaSet") Then mvList.IntegerValue("CriteriaSet") = mvMailingInfo.CriteriaSet
        If Not mvList.Contains("SelectionSetNumber") Then mvList.IntegerValue("SelectionSetNumber") = mvMailingInfo.SelectionSet
        If Not mvList.Contains("Revision") Then mvList.IntegerValue("Revision") = mvMailingInfo.Revision
        If Not mvList.Contains("ApplicationName") Then mvList("ApplicationName") = mvMailingInfo.MailingTypeCode

        mvMailingInfo.Revision = mvMailingInfo.ProcessSelection(mvList)
      End If
      'mvMailingInfo.Revision = mvMailingInfo.ProcessSelection(mvMailingInfo.CriteriaSet, mvMailingInfo.SelectionSet, mvMailingInfo.Revision, vMergeDesc, mvMailingInfo.MailingTypeCode)

      If mvMailingInfo.Revision = 0 Then
        mvMailingInfo.Revision = vStoreGvRev
      Else
        mvMailingInfo.SelectionCount = mvMailingInfo.GetMailingSelectionCount(mvMailingInfo.SelectionSet, mvMailingInfo.Revision, mvMailingInfo.MailingTypeCode)
        ShowInformationMessage(InformationMessages.ImContactSelected, mvMailingInfo.SelectionCount.ToString)
        If mvMailingInfo.SelectionCount > 0 Then
          Dim vFrmGenMail As New frmGenMGen(mvMailingInfo.MailingTypeCode, mvMailingInfo, mvMailingInfo.SelectionSet)
          vResult = vFrmGenMail.ShowDialog()
        Else
          mvMailingInfo.Revision = vStoreGvRev
        End If
      End If
      Me.Close()
      Exit Sub
    Catch vCareException As CareException
      If vCareException.ErrorNumber = 1049 Then
        ShowInformationMessage(vCareException.Message)
      Else
        DataHelper.HandleException(vCareException)
      End If
    Catch ex As Exception
      DataHelper.HandleException(ex)
    End Try
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Try

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
    Me.Close()
  End Sub

  Private Sub frmGenMMerge_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Me.Text = mvMailingInfo.Caption & " - List Merge Options"
    optMergerAll.Checked = True
    mvCurrentRecords = True
    mvCommonRecords = True
    mvNewRecords = True
  End Sub

  Private Sub lbl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbl1.Click
    Try
      'A
      If mvCurrentRecords = True Then
        If optAAndC.Checked = True Then   'A and C
          optCOnly.Checked = True         'C only
          mvCurrentRecords = False
          mvCommonRecords = False
          mvNewRecords = True
        Else
          optBOnly.Checked = True   'B only
          mvCurrentRecords = False
          mvCommonRecords = True
          mvNewRecords = False
        End If
      Else
        If optBOnly.Checked = True Then 'B only
          optAOnly.Checked = True      'A only
          mvCurrentRecords = True
          mvCommonRecords = False
          mvNewRecords = False
        End If
        optAAndC.Checked = True         'A and C
        mvCurrentRecords = True
        mvCommonRecords = False
        mvNewRecords = True
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub lbl2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbl2.Click
    Try
      'B
      If mvCommonRecords = True Then
        optAAndC.Checked = True    'a and c
        mvCurrentRecords = True
        mvCommonRecords = False
        mvNewRecords = True
      Else
        If mvCommonRecords = False Then
          If optAAndC.Checked = True Then    'a and c
            optMergerAll.Checked = True  'all
            mvCurrentRecords = True
            mvCommonRecords = True
            mvNewRecords = True
          Else
            optBOnly.Checked = True        'b
            mvCurrentRecords = False
            mvCommonRecords = True
            mvNewRecords = False
          End If
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub lbl3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbl3.Click
    Try
      'C
      If mvNewRecords = True Then
        If optAAndC.Checked = True Then      'a and c
          optAOnly.Checked = True           'a
          mvCurrentRecords = True
          mvCommonRecords = False
          mvNewRecords = False
        Else
          optBOnly.Checked = True           'b
          mvCurrentRecords = False
          mvCommonRecords = True
          mvNewRecords = False
        End If
      ElseIf optAOnly.Checked = True Then      'a
        optAAndC.Checked = True           'a and c
        mvCurrentRecords = True
        mvCommonRecords = False
        mvNewRecords = True
      Else
        optCOnly.Checked = True           'c
        mvCurrentRecords = False
        mvCommonRecords = False
        mvNewRecords = True
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub frmGenMMerge_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

  End Sub

  Private Sub optMergerAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMergerAll.CheckedChanged, optAAndC.CheckedChanged, optAOnly.CheckedChanged, optBOnly.CheckedChanged, optCOnly.CheckedChanged
    Try
      'reset "All"
      lbl1.ForeColor = Color.LightGray
      lbl2.ForeColor = Color.LightGray
      lbl3.ForeColor = Color.LightGray

      Select Case DirectCast(sender, RadioButton).Name
        Case "optMergerAll"
          mvMergeType = 1
          lbl1.ForeColor = Color.Black
          lbl2.ForeColor = Color.Black
          lbl3.ForeColor = Color.Black
          lblDesc.Text = ControlText.LblDescSelectAll
        Case "optAOnly"
          mvMergeType = 4
          lbl1.ForeColor = Color.Black
          lblDesc.Text = ControlText.LblDescExcludeCriteria
        Case "optBOnly"
          mvMergeType = 0
          lbl2.ForeColor = Color.Black
          lblDesc.Text = ControlText.LblDescIncludeCriteria
        Case "optCOnly"
          mvMergeType = 2
          lbl3.ForeColor = Color.Black
          lblDesc.Text = ControlText.LblDescExcludeCurrentSet
        Case "optAAndC"
          mvMergeType = 3
          lbl1.ForeColor = Color.Black
          lbl3.ForeColor = Color.Black
          lblDesc.Text = ControlText.LblDescSelectCurrentAndCriteria
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub lblA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblA.Click
    lbl1_Click(sender, e)
  End Sub

  Private Sub lblB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblB.Click
    lbl2_Click(sender, e)
  End Sub

  Private Sub lblC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblC.Click
    lbl3_Click(sender, e)
  End Sub
End Class