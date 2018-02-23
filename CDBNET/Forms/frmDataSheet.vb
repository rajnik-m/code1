Public Class frmDataSheet

  Public Enum DataSheetTypes
    dstActivities
    dstRelationships
  End Enum

  Dim mvDataSheetType As DataSheetTypes
  Dim mvGroupCode As String

  Public Sub New(ByVal pContactInfo As ContactInfo, ByVal pDataSheetType As DataSheetTypes, ByVal pGroupCode As String, ByVal pDescription As String, ByVal pTable As DataTable, ByVal pSource As String)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pContactInfo, pDataSheetType, pGroupCode, pDescription, pTable, pSource)
  End Sub

  Private Sub InitialiseControls(ByVal pContactInfo As ContactInfo, ByVal pDataSheetType As DataSheetTypes, ByVal pGroupCode As String, ByVal pDescription As String, ByVal pTable As DataTable, Optional ByVal pSource As String = "")
#If DEBUG Then
    Static mvDebugTest As New DebugTest(Me)       'Include this statement so we can keep track of memory leakage of forms
#End If
    SetControlColors()
    mvDataSheetType = pDataSheetType
    mvGroupCode = pGroupCode
    Select Case pDataSheetType
      Case DataSheetTypes.dstActivities
        If pDescription.Length = 0 Then
          Me.Text = GetInformationMessage(ControlText.frmActivityDataSheet, pContactInfo.ContactName)
        Else
          Me.Text = GetInformationMessage(ControlText.frmActivityDataSheet, pDescription)
        End If
        rds.Visible = False
        ads.Init(pContactInfo, pGroupCode, pTable, pSource)
        If ads.Source.Length = 0 Then
          Dim vPanelItem As New PanelItem(txtSource, "source")
          txtSource.TotalWidth = txtSource.Width
          txtSource.SetBounds(txtSource.Location.X, txtSource.Location.Y, 100, txtSource.Size.Height)
          txtSource.Init(vPanelItem, False, True)
        Else
          pnlSource.Visible = False
        End If
      Case DataSheetTypes.dstRelationships
        If pDescription.Length = 0 Then
          Me.Text = GetInformationMessage(ControlText.frmRelationshipDataSheet, pContactInfo.ContactName)
        Else
          Me.Text = GetInformationMessage(ControlText.frmRelationshipDataSheet, pDescription)
        End If
        ads.Visible = False
        rds.Init(pContactInfo, pGroupCode, pTable)
        pnlSource.Visible = False
    End Select
  End Sub

  Private Sub SetControlColors()
    Me.BackColor = DisplayTheme.FormBackColor
    bpl.UseTheme()
  End Sub

  Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Dim vValid As Boolean

    Select Case mvDataSheetType
      Case DataSheetTypes.dstActivities
        erp.SetError(txtSource, "")
        vValid = ads.ValidateActivities
        If vValid AndAlso ads.Source.Length = 0 Then
          If Not txtSource.IsValid Then
            erp.SetError(txtSource, GetInformationMessage(InformationMessages.imInvalidValue))
            vValid = False
          End If
          If vValid AndAlso txtSource.Text.Length = 0 Then
            erp.SetError(txtSource, GetInformationMessage(InformationMessages.imFieldMandatory))
            vValid = False
          End If
        End If
        If vValid Then
          ads.SaveActivities(txtSource.Text)
          Me.Close()
          Me.DialogResult = Windows.Forms.DialogResult.OK
        End If
      Case DataSheetTypes.dstRelationships
        If rds.ValidateRelationships Then
          rds.SaveRelationships()
          Me.Close()
          Me.DialogResult = Windows.Forms.DialogResult.OK
        End If
    End Select
  End Sub

  Private Sub frmDataSheet_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    bpl.RepositionButtons()
  End Sub

  Private Sub frmDataSheet_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
    If rds.Visible Then rds.SetHeight()
    If ads.Visible Then ads.SetHeight()
  End Sub
End Class