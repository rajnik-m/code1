Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Public Class frmDocumentDistributor
  Private mvOriginalImage As Image
  Private mvSelectedContact As Integer
  Private mvDocumentNumber As Integer
  Private mvButtonWidth As Integer = 230
  Private mvButtonHeight As Integer = 30
  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub

#Region "Initialise"

  Private Sub InitialiseControls()
    Try
      SetControlTheme()
      tab.SetItemSizes()
      Me.MdiParent = MDIForm
      cboZoom.Text = ControlText.CboHundredPercent

    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub PopulateTypes()
    Dim vParams As New ParameterList(True)
    vParams("DocumentSource") = "S"
    Dim vTypeDataSet As DataSet = DataHelper.GetLookupDataSet(CareNetServices.XMLLookupDataTypes.xldtDocumentTypes, vParams)
    cboType.DisplayMember = "DocumentTypeDesc"
    cboType.ValueMember = "DocumentType"
    cboType.DataSource = vTypeDataSet.Tables("DataRow")
    Dim vX As Integer = pnlType.Bounds.X
    Dim vY As Integer = pnlType.Bounds.Y
    For vRowCounter As Integer = 0 To vTypeDataSet.Tables("DataRow").Rows.Count - 1
      Dim vToggleButton As New CheckBox
      vToggleButton.Text = vTypeDataSet.Tables("DataRow").Rows(vRowCounter).Item("DocumentTypeDesc").ToString()
      vToggleButton.Name = vTypeDataSet.Tables("DataRow").Rows(vRowCounter).Item("DocumentType").ToString()
      vToggleButton.Appearance = Appearance.Button
      vToggleButton.SetBounds(vX, vY, mvButtonWidth, mvButtonHeight)
      vToggleButton.BackColor = Button.DefaultBackColor
      pnlType.Controls.Add(vToggleButton)
      AddHandler vToggleButton.CheckedChanged, AddressOf Types_Click
      vY = vY + mvButtonHeight
    Next
    cboType.SelectedIndex = 0
    Toggle(CType(pnlType.Controls(0), CheckBox), True)
  End Sub

  Private Sub PopulatePostPoints()
    Dim vParams As New ParameterList(True)
    Dim vPostPointDataSet As DataSet = DataHelper.GetLookupDataSet(CareNetServices.XMLLookupDataTypes.xldtPostPoints, vParams)
    cboPostPoint.DataSource = vPostPointDataSet.Tables("DataRow")
    cboPostPoint.DisplayMember = "PostPointDesc"
    cboPostPoint.ValueMember = "PostPoint"
    Dim vX As Integer = pnlPostPoint.Bounds.X
    Dim vY As Integer = pnlPostPoint.Bounds.Y
    For vRowCounter As Integer = 0 To vPostPointDataSet.Tables("DataRow").Rows.Count - 1
      Dim vToggleButton As New CheckBox
      vToggleButton.Text = vPostPointDataSet.Tables("DataRow").Rows(vRowCounter).Item("PostPointDesc").ToString()
      vToggleButton.Name = vPostPointDataSet.Tables("DataRow").Rows(vRowCounter).Item("PostPoint").ToString()
      vToggleButton.Appearance = Appearance.Button
      vToggleButton.SetBounds(vX, vY, mvButtonWidth, mvButtonHeight)
      vToggleButton.BackColor = Button.DefaultBackColor
      pnlPostPoint.Controls.Add(vToggleButton)
      AddHandler vToggleButton.CheckedChanged, AddressOf PostPoint_Click
      vY = vY + mvButtonHeight
    Next
    cboPostPoint.SelectedIndex = 0
    Toggle(CType(pnlPostPoint.Controls(0), CheckBox), True)
  End Sub

  Private Sub PopulateRecipients()
    Dim vX As Integer = pnlType.Bounds.X
    Dim vY As Integer = pnlType.Bounds.Y
    Dim vParams As New ParameterList(True)
    Dim vRecipientTable As New DataTable()
    vRecipientTable.Columns.Add("ContactNumber")
    vRecipientTable.Columns.Add("ContactName")
    Dim vRecipientDataSet As DataSet = DataHelper.GetLookupDataSet(CareNetServices.XMLLookupDataTypes.xldtUserContactNames, vParams)
    cboRecipient.DataSource = vRecipientDataSet.Tables("DataRow")
    cboRecipient.DisplayMember = "CONTACT_NAME"
    cboRecipient.ValueMember = "contact_number"
    For vRowCounter As Integer = 0 To vRecipientDataSet.Tables("DataRow").Rows.Count - 1
      'Adding Toggle Buttons
      Dim vToggleButton As New CheckBox
      vToggleButton.Text = vRecipientDataSet.Tables("DataRow").Rows(vRowCounter).Item("CONTACT_NAME").ToString()
      vToggleButton.Name = vRecipientDataSet.Tables("DataRow").Rows(vRowCounter).Item("contact_number").ToString()
      vToggleButton.Appearance = Appearance.Button
      vToggleButton.SetBounds(vX, vY, mvButtonWidth, mvButtonHeight)
      vToggleButton.BackColor = Button.DefaultBackColor
      pnlRecipient.Controls.Add(vToggleButton)
      AddHandler vToggleButton.CheckedChanged, AddressOf Recipient_Click
      vY = vY + mvButtonHeight
    Next
    cboRecipient.SelectedIndex = 0
    Toggle(CType(pnlRecipient.Controls(0), CheckBox), True)
  End Sub
#End Region

#Region "Click Events"

  Private Sub PostPoint_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      'Code below will deselect other buttons.
      For vRowCounter As Integer = 0 To pnlPostPoint.Controls.Count - 1
        If pnlPostPoint.Controls.Item(vRowCounter).GetType().Name = "CheckBox" Then
          Toggle(CType(pnlPostPoint.Controls.Item(vRowCounter), CheckBox), False)
        End If
      Next
      'Code below will select current buttons.
      Toggle(CType(sender, CheckBox), True)
      cboPostPoint.SelectedValue = CType(sender, CheckBox).Name
      If LoadImage("", cboPostPoint.SelectedValue.ToString, 0, 0, 0) = False Then
        Me.Close()
      End If
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub Types_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      'Code below will deselect other buttons.
      For vRowCounter As Integer = 0 To pnlType.Controls.Count - 1
        If pnlType.Controls.Item(vRowCounter).GetType().Name = "CheckBox" Then
          Toggle(CType(pnlType.Controls.Item(vRowCounter), CheckBox), False)
        End If
      Next
      'Code below will select current buttons.
      Toggle(CType(sender, CheckBox), True)
      cboType.SelectedValue = CType(sender, CheckBox).Name
      If LoadImage(cboType.SelectedValue.ToString, "", 0, 0, 0) = False Then
        Me.Close()
      End If
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub Recipient_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      'Code below will deselect other buttons.
      For vRowCounter As Integer = 0 To pnlType.Controls.Count - 1
        If pnlRecipient.Controls.Item(vRowCounter).GetType().Name = "CheckBox" Then
          Toggle(CType(pnlRecipient.Controls.Item(vRowCounter), CheckBox), False)
        End If
      Next
      'Code below will select current buttons.
      Toggle(CType(sender, CheckBox), True)
      cboRecipient.SelectedValue = CType(sender, CheckBox).Name
      If LoadImage("", "", IntegerValue(cboRecipient.SelectedValue), 0, 0) = False Then
        Me.Close()
      End If
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    Me.Close()
  End Sub

  Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
    Try
      FormHelper.EditDocument(mvDocumentNumber)
      If LoadImage("", "", 0, 0, 0) = False Then
        Me.Close()
      End If
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Function LoadImage(ByVal pDocumentType As String, ByVal pPostPoint As String, ByVal pRecipient As Integer, ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer) As Boolean
    Dim vList As New ParameterList(True)
    Dim vImageFile As String = ""
    Dim vSuccess As Boolean = True
    Try
      If pDocumentType.Length > 0 Then vList("DocumentType") = pDocumentType
      If pPostPoint.Length > 0 Then vList("PostPoint") = pPostPoint
      If pRecipient > 0 Then vList("RecipientNumber") = pRecipient.ToString()
      If mvSelectedContact > 0 Then vList("ContactNumber") = mvSelectedContact.ToString()
      If Not cboAddresses.SelectedValue Is Nothing Then vList("AddressNumber") = cboAddresses.SelectedValue.ToString()
      If mvDocumentNumber > 0 Then vList("DocumentNumber") = mvDocumentNumber.ToString
      Dim vResultList As New ParameterList()
      vResultList = DataHelper.GetDocumentNumberForDistribution(vList)

      mvDocumentNumber = IntegerValue(vResultList("DocumentNumber"))
      vImageFile = DataHelper.GetDocumentFileForDistributor(mvDocumentNumber, ".jpg")

      If vImageFile.Length > 0 Then
        Dim vStream As New FileStream(vImageFile, IO.FileMode.Open, IO.FileAccess.Read)
        mvOriginalImage = Image.FromStream(vStream)
        pctDocument.Image = Image.FromStream(vStream)
        vStream.Dispose()
        DataHelper.DeleteTempFile(vImageFile)
        ' set Slider Attributes
        ZoomSlider.Enabled = True
        cboZoom.Enabled = True
        ZoomSlider.Minimum = 1
        ZoomSlider.Maximum = 18
        ZoomSlider.SmallChange = 1
        ZoomSlider.UseWaitCursor = False
        cmdRotate.Enabled = True
        ZoomSlider.Value = 10
        grpBox.Text = ControlText.GrpBoxDocumentNumber & mvDocumentNumber.ToString
      Else
        ZoomSlider.Enabled = False
        cboZoom.Enabled = False
      End If
    Catch vEx As CareException
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enNoMorePostToBeDistributed
          ShowInformationMessage(vEx.Message)
          vSuccess = False
        Case CareException.ErrorNumbers.enFailedToFindDocumentDefaults, CareException.ErrorNumbers.enFailedToFindNextRecipientInTheQueue, CareException.ErrorNumbers.enFailedToRetrievePostPointInformation, _
          CareException.ErrorNumbers.enNoPostPointOrRecipientHaveBeenDefined, CareException.ErrorNumbers.enNoRecipientsWereFoundForThePostPoint, CareException.ErrorNumbers.enExternalFilenameInvalid
          ShowInformationMessage(vEx.Message)
          vSuccess = True
        Case Else
          DataHelper.HandleException(vEx)
      End Select
    End Try
    Return vSuccess
  End Function

  Private Sub cmdRotate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRotate.Click
    Try
      If mvOriginalImage IsNot Nothing Then
        ZoomSlider.Value = 10
        cboZoom.SelectedIndex = ZoomSlider.Value - 1
        mvOriginalImage.RotateFlip(RotateFlipType.Rotate270FlipXY)
        pctDocument.Image = mvOriginalImage
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdContactFinder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdContactFinder.Click
    Try
      Dim vList As New ParameterList(True)
      Dim vContactName As String
      Dim vAddressLine As String
      Me.SendToBack()
      mvSelectedContact = ShowFinder(CareServices.XMLDataFinderTypes.xdftContacts, vList, Me.ParentForm, True, True)
      If mvSelectedContact > 0 Then
        Dim vContactInfo As New ContactInfo(mvSelectedContact)
        vContactName = vContactInfo.ContactName
        Dim vAddressDataTable As DataTable = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddresses, mvSelectedContact)
        cboAddresses.DataSource = vAddressDataTable
        cboAddresses.DisplayMember = "AddressLine"
        cboAddresses.ValueMember = "AddressNumber"
        vAddressLine = vAddressDataTable.Select("Default= 'Yes'")(0).Item("AddressLine").ToString()
        cboAddresses.SelectedValue = vAddressDataTable.Select("Default= 'Yes'")(0).Item("AddressNumber").ToString()
        txtName.Text = vContactName
      End If
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdOrganisationFinder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOrganisationFinder.Click
    Try
      Dim vList As New ParameterList(True)
      Dim vContactName As String
      Dim vAddressLine As String
      Me.SendToBack()
      mvSelectedContact = ShowFinder(CareServices.XMLDataFinderTypes.xdftOrganisations, vList, Me.ParentForm, True, True)
      If mvSelectedContact > 0 Then
        Dim vContactInfo As New ContactInfo(mvSelectedContact)
        vContactName = vContactInfo.ContactName
        Dim vAddressDataTable As DataTable = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddresses, mvSelectedContact)
        cboAddresses.DataSource = vAddressDataTable
        cboAddresses.DisplayMember = "AddressLine"
        cboAddresses.ValueMember = "AddressNumber"
        vAddressLine = vAddressDataTable.Select("Default= 'Yes'")(0).Item("AddressLine").ToString()
        cboAddresses.SelectedValue = vAddressDataTable.Select("Default= 'Yes'")(0).Item("AddressNumber").ToString()
        txtName.Text = vContactName
      End If
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub

#End Region

#Region "PictureBox Events"
  Private Sub ZoomSlider_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZoomSlider.Scroll
    Try
      If ZoomSlider.Value > 10 Then
        pctDocument.Image = PictureBoxZoom(mvOriginalImage, New Size(ZoomSlider.Value - 10, ZoomSlider.Value - 10), New Rectangle(0, 0, mvOriginalImage.Width, mvOriginalImage.Height))
      ElseIf ZoomSlider.Value = 10 Then
        pctDocument.Image = PictureBoxZoom(mvOriginalImage, New Size(1, 1), New Rectangle(0, 0, mvOriginalImage.Width, mvOriginalImage.Height))
      Else
        pctDocument.Image = PictureBoxZoomIn(mvOriginalImage, New Size(10 - ZoomSlider.Value, 10 - ZoomSlider.Value), New Rectangle(0, 0, mvOriginalImage.Width, mvOriginalImage.Height))
      End If
      cboZoom.SelectedIndex = ZoomSlider.Value - 1
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Public Function PictureBoxZoom(ByVal pImage As Image, ByVal pSize As Size, ByVal pSourceArea As Rectangle) As Image
    Dim vBitmap As Bitmap = New Bitmap(pImage, CInt(pImage.Width * pSize.Width), CInt(pImage.Height * pSize.Height))
    Dim vGraphics As Graphics = Graphics.FromImage(vBitmap)
    vGraphics.InterpolationMode = InterpolationMode.High
    Return vBitmap
  End Function

  Public Function PictureBoxZoomIn(ByVal pImage As Image, ByVal pSize As Size, ByVal pSourceArea As Rectangle) As Image
    Dim vBitmap As Bitmap = New Bitmap(pImage, CInt(pImage.Width / pSize.Width), CInt(pImage.Height / pSize.Height))
    Dim vGraphics As Graphics = Graphics.FromImage(vBitmap)
    vGraphics.InterpolationMode = InterpolationMode.High
    Return vBitmap
  End Function
#End Region

  Private Sub cboType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboType.SelectedIndexChanged
    Try
      If cboType.Items.Count > 0 Then
        For vRowCounter As Integer = 0 To pnlType.Controls.Count - 1
          If pnlType.Controls.Item(vRowCounter).GetType().Name = "CheckBox" Then
            Toggle(CType(pnlType.Controls.Item(vRowCounter), CheckBox), False)
          End If
        Next
        If pnlType.Controls.Count > 0 Then
          Toggle(CType(pnlType.Controls.Find(cboType.SelectedValue.ToString, True)(0), CheckBox), True)
        End If
      End If
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cboPostPoint_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPostPoint.SelectedIndexChanged
    Try
      If cboPostPoint.Items.Count > 0 Then
        For vRowCounter As Integer = 0 To pnlPostPoint.Controls.Count - 1
          If pnlPostPoint.Controls.Item(vRowCounter).GetType().Name = "CheckBox" Then
            Toggle(CType(pnlPostPoint.Controls.Item(vRowCounter), CheckBox), False)
          End If
        Next
        If pnlPostPoint.Controls.Count > 0 Then Toggle(CType(pnlPostPoint.Controls.Find(cboPostPoint.SelectedValue.ToString, True)(0), CheckBox), True)
      End If
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cboRecipient_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRecipient.SelectedIndexChanged
    Try
      If cboRecipient.Items.Count > 0 Then
        For vRowCounter As Integer = 0 To pnlRecipient.Controls.Count - 1
          If pnlRecipient.Controls.Item(vRowCounter).GetType().Name = "CheckBox" Then
            Toggle(CType(pnlRecipient.Controls.Item(vRowCounter), CheckBox), False)
          End If
        Next
        If pnlRecipient.Controls.Count > 0 Then Toggle(CType(pnlRecipient.Controls.Find(cboRecipient.SelectedValue.ToString, True)(0), CheckBox), True)
      End If
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub Toggle(ByVal pButton As CheckBox, ByVal pSelect As Boolean)
    If pSelect Then
      pButton.BackColor = Color.Black
      pButton.ForeColor = Color.White
    Else
      pButton.BackColor = Button.DefaultBackColor
      pButton.ForeColor = Button.DefaultForeColor
    End If
  End Sub

  Private Sub frmDocumentDistributor_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
    Try
      pctDocument.SetBounds(pctDocument.Bounds.X, pctDocument.Bounds.Y, pnlPictureBox.Size.Width - 10, pnlPictureBox.Size.Height - 10)
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Try
      Dim vList As New ParameterList(True)
      vList("DocumentNumber") = mvDocumentNumber.ToString()
      DataHelper.DeleteItem(CareServices.XMLMaintenanceControlTypes.xmctDocument, vList)
      If LoadImage("", "", 0, 0, 0) = False Then
        Me.Close()
      End If
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub frmDocumentDistributor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Try
      'Setting Background Images for Finder Buttons
      cmdContactFinder.BackgroundImage = AppHelper.ImageProvider.NewImageList32.Images.Item("ContactFinder")
      cmdContactFinder.AutoSize = True
      cmdContactFinder.BackgroundImageLayout = ImageLayout.Center
      cmdContactFinder.SetBounds(cmdContactFinder.Bounds.X, cmdContactFinder.Bounds.Y, 40, 40)

      cmdOrganisationFinder.BackgroundImage = AppHelper.ImageProvider.NewImageList32.Images.Item("OrganisationFinder")
      cmdOrganisationFinder.AutoSize = True
      cmdOrganisationFinder.BackgroundImageLayout = ImageLayout.Center
      cmdOrganisationFinder.SetBounds(cmdOrganisationFinder.Bounds.X, cmdOrganisationFinder.Bounds.Y, 40, 40)

      PopulateTypes()
      PopulatePostPoints()
      PopulateRecipients()
      grpBox.Text = ControlText.GrpBoxDocumentNumber & mvDocumentNumber.ToString
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub
  Public Function InitDocumentDistributor() As Boolean
    Return LoadImage("", "", 0, 0, 0)
  End Function

  Private Sub cboZoom_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboZoom.SelectionChangeCommitted
    Try
      If cboZoom.SelectedIndex > 9 Then
        pctDocument.Image = PictureBoxZoom(mvOriginalImage, New Size(cboZoom.SelectedIndex - 9, cboZoom.SelectedIndex - 9), New Rectangle(0, 0, mvOriginalImage.Width, mvOriginalImage.Height))
      ElseIf cboZoom.SelectedIndex = 9 Then
        pctDocument.Image = PictureBoxZoom(mvOriginalImage, New Size(1, 1), New Rectangle(0, 0, mvOriginalImage.Width, mvOriginalImage.Height))
      Else
        pctDocument.Image = PictureBoxZoomIn(mvOriginalImage, New Size(9 - cboZoom.SelectedIndex, 9 - cboZoom.SelectedIndex), New Rectangle(0, 0, mvOriginalImage.Width, mvOriginalImage.Height))
      End If
      ZoomSlider.Value = cboZoom.SelectedIndex + 1
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub


End Class