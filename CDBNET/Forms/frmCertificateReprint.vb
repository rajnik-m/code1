Public Class frmCertificateReprint
  Inherits ThemedForm

  Private mvCertificateId As Integer = 0
  Private mvReprintType As String = String.Empty
  Private mvPanelInfo As New EditPanelInfo(CareServices.FunctionParameterTypes.fptExamCertificateReprint)

  Private WithEvents ctlActionQueue As RadioButton = Nothing
  Private WithEvents ctlActionImmediate As RadioButton = Nothing
  Private WithEvents ctlDestination As TextBox = Nothing
  Private WithEvents cmdBrowse As Button = Nothing
  Private WithEvents ctlAutomerge As CheckBox = Nothing

  <Obsolete("Designer Only", True)>
  Public Sub New()
    MyBase.New()
    If Not DesignMode Then
      Throw New NotSupportedException("The default constructor of frmCertificateReprint can only be called from the designer.")
    End If
    InitializeComponent()
    Me.epl.Height = Me.Height - Me.bpl.Height
    Me.SetControlTheme()
  End Sub

  Public Sub New(pCertificateId As Integer, pReprintType As String)
    MyBase.New()
    InitializeComponent()
    Me.epl.Height = Me.Height - Me.bpl.Height
    Me.SetControlTheme()
    mvCertificateId = pCertificateId
    mvReprintType = pReprintType
  End Sub

  Private Sub ctlActionQueue_CheckedChanged(sender As Object, e As EventArgs) Handles ctlActionQueue.CheckedChanged, ctlActionImmediate.CheckedChanged
    SetControlState()
  End Sub

  Private Sub cmdOk_Click(sender As Object, e As EventArgs) Handles cmdOk.Click
    Try
      Dim vResult As Byte() = ExamsDataHelper.ReprintCertificate(mvCertificateId,
                                                                 mvReprintType,
                                                                 If(Me.ctlActionQueue IsNot Nothing,
                                                                    Me.ctlActionQueue.Checked,
                                                                    False))
      If Me.ctlActionQueue IsNot Nothing AndAlso Me.ctlActionQueue.Checked Then
        Dim vResultData As DataSet = DataHelper.GetDataSetFromResult(Encoding.UTF8.GetString(vResult))
        If vResultData IsNot Nothing Then
          If vResultData.Tables.Contains("Result") Then
            ShowInformationMessage("Certificate reprint request {0} queued.",
                                   CStr(vResultData.Tables("Result").Rows(0)("ContactExamCertReprintId")))
          End If
        End If
      Else
        Using vResultFile As New FileStream(Me.ctlDestination.Text, FileMode.Create)
          vResultFile.Write(vResult, 0, vResult.Length)
        End Using
        ShowInformationMessage("Data has been saved in {0}.",
                               ctlDestination.Text)
        If Me.ctlAutomerge.Checked Then
          Call New ExamCertificateMergeEngine(Me.ctlDestination.Text).ProduceDocuments()
        End If
      End If
    Catch vEx As Exception
      ShowError(vEx.Source, vEx.Message, vEx.StackTrace)
    End Try
  End Sub

  Private Sub cmdBrowse_Click(sender As Object, e As EventArgs) Handles cmdBrowse.Click
    Dim vFileDialog As New SaveFileDialog
    vFileDialog.AddExtension = True
    vFileDialog.OverwritePrompt = True
    vFileDialog.InitialDirectory = AppValues.ConfigurationValue(AppValues.ConfigurationValues.default_mailing_directory, "c:\contacts")
    vFileDialog.Filter = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|Output files (*.out)|*.out|All files (*.*)|*.*"
    vFileDialog.FileName = Me.ctlDestination.Text
    vFileDialog.ShowDialog()
    Me.ctlDestination.Text = vFileDialog.FileName
  End Sub

  Private Sub ctlDestination_EnabledChanged(sender As Object, e As EventArgs) Handles ctlDestination.EnabledChanged
    cmdBrowse.Enabled = DirectCast(sender, Control).Enabled
  End Sub

  Private Sub ctlDestination_TextChanged(sender As Object, e As EventArgs) Handles ctlDestination.TextChanged
    SetControlState()
  End Sub

  Private Sub SetControlState()
    cmdOk.Enabled = Not String.IsNullOrWhiteSpace(If(Me.ctlDestination IsNot Nothing, Me.ctlDestination.Text, "")) Or
                    Not (Me.ctlDestination IsNot Nothing AndAlso Me.ctlDestination.Enabled)
    If Me.ctlDestination IsNot Nothing Then
      Me.ctlDestination.Enabled = ctlActionImmediate.Checked And Not AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ma_auto_name_mailing_files, False)
    End If
    If Me.ctlAutomerge IsNot Nothing Then
      Me.ctlAutomerge.Enabled = Me.ctlActionImmediate.Checked
    End If
  End Sub

  Private Sub frmCertificateReprint_Load(sender As Object, e As EventArgs) Handles Me.Load

    Me.epl.Init(mvPanelInfo)

    Me.ctlActionQueue = DirectCast(FindControl(Me.epl, "QueueReprint_Y", False), RadioButton)
    Me.ctlActionImmediate = DirectCast(FindControl(Me.epl, "QueueReprint_N", False), RadioButton)
    Me.ctlDestination = DirectCast(FindControl(Me.epl, "Destination", False), TextBox)
    Me.cmdBrowse = DirectCast(FindControl(Me.epl, "Browse", False), Button)
    Me.ctlAutomerge = DirectCast(FindControl(Me.epl, "AutoMerge", False), CheckBox)

    If ctlDestination IsNot Nothing Then
      DirectCast(ctlDestination.Tag, PanelItem).Mandatory = False
    End If

    If ctlActionQueue IsNot Nothing Then
      ctlActionQueue.Checked = True
    ElseIf ctlActionImmediate IsNot Nothing Then
      ctlActionQueue.Checked = True
    End If

    If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ma_auto_name_mailing_files, False) Then
      ctlDestination.Text = AppValues.GetMailingFileName()
    End If

    Me.Height = (From vControl As Control In epl.Controls.Cast(Of Control)()
                 Select vControl.Top).Min +
                (From vControl As Control In epl.Controls.Cast(Of Control)()
                 Select vControl.Bottom).Max +
                Me.bpl.Height +
                (Me.Height - Me.ClientSize.Height)

    Me.Width = (From vControl As Control In epl.Controls.Cast(Of Control)()
                Select vControl.Left).Min +
                (From vControl As Control In epl.Controls.Cast(Of Control)()
                 Select vControl.Right).Max +
               (Me.Width - Me.ClientSize.Width)

    Me.MinimumSize = New Size(Me.Size.Width, Me.Size.Height)

    SetControlState()
  End Sub

  Private Sub frmCertificateReprint_Resize(sender As Object, e As EventArgs) Handles Me.Resize
    Me.epl.Height = Me.ClientSize.Height - Me.bpl.Height
  End Sub
End Class