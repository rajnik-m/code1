
Public Class frmHTMLEditor
  Dim mvEditor As HtmlEditorControl.CareHtmlEditor.HtmlEditorControl

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    mvEditor = New HtmlEditorControl.CareHtmlEditor.HtmlEditorControl
    mvEditor.Dock = DockStyle.Fill
    Me.pnl.Controls.Add(mvEditor)
  End Sub

  Public Property HTMLText() As String
    Get
      Return mvEditor.InnerHtml
    End Get
    Set(ByVal pValue As String)
      mvEditor.InnerHtml = pValue
    End Set
  End Property

  Public WriteOnly Property ImageTable() As DataTable
    Set(ByVal pValue As DataTable)
      mvEditor.ImageTable = pValue
    End Set
  End Property

  Public WriteOnly Property PageTable() As DataTable
    Set(ByVal pValue As DataTable)
      mvEditor.PageTable = pValue
    End Set
  End Property

  Public WriteOnly Property DocumentTable() As DataTable
    Set(ByVal pValue As DataTable)
      mvEditor.DocumentTable = pValue
    End Set
  End Property

  Public WriteOnly Property BaseImageUrl() As String
    Set(ByVal pValue As String)
      mvEditor.BaseImageUrl = pValue
    End Set
  End Property

  Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
    mvEditor.HtmlContentsEdit()
    mvEditor.Focus()
  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Me.Close()
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.Close()
  End Sub

  Private Sub frmHTMLEditor_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    bpl.RepositionButtons()
  End Sub
End Class