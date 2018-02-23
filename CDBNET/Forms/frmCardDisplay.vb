Public Class frmCardDisplay
  Inherits CDBNET.frmCardSet

#Region " Windows Form Designer generated code "

  Public Sub New()
    MyBase.New()
    'This call is required by the Windows Form Designer.
    InitializeComponent()
    'Add any initialization after the InitializeComponent() call
    InitialiseControls()
  End Sub

  'Form overrides dispose to clean up the component list.
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    '
    'frmCardDisplay
    '
    Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
    Me.ClientSize = New System.Drawing.Size(696, 392)
    Me.Name = "frmCardDisplay"

  End Sub

#End Region

  Private Sub InitialiseControls()
    splTop.Panel1Collapsed = True
    splBottom.Panel1Collapsed = True
  End Sub

  Public Overrides Sub Init(ByVal pDataSet As DataSet, ByVal pType As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo, Optional ByVal pRetainPage As Boolean = False)
    MyBase.Init(pDataSet, pType, pContactInfo, pRetainPage)
    If pType = CareServices.XMLContactDataSelectionTypes.xcdtContactJournals Then
      If pContactInfo.ContactNumber > 0 Then
        Me.Text = String.Format(ControlText.FrmJournalFor, pContactInfo.ContactName)
      Else
        Me.Text = ControlText.FrmMyJournal
      End If
    End If
  End Sub

End Class
