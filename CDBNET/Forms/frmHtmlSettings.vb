Imports System.Xml.Linq

Public Class frmHtmlSettings


  Private Property Options As New HtmlSettingsOptions
  ''' <summary>
  ''' Initialises a new instance of the Settings dialog for the specifed Html.  When the users presses Ok, a JavaScript command is invoked on the Web page and the return value of the Javascript command is passed back in the pReturnValue object
  ''' </summary>
  ''' <param name="pHtmlSettingsOption">An instance of the frmHtmlSettings.HtmlSettingsOptions</param>
  ''' <remarks></remarks>
  ''' 
  Public Sub New(pHtmlSettingsOption As HtmlSettingsOptions)

    InitializeComponent()

    Me.Options = pHtmlSettingsOption
    If Me.Options.Settings Is Nothing Then Me.Options.Settings = New XDocument()

    Initialise()

  End Sub

  Private Sub Initialise()
    Dim sb As New StringBuilder()

    'Need this as sometimes setting the DocumentText (further down) does nothing
    Me.HtmlBrowser.Navigate("about:blank")
    Dim doc As HtmlDocument = Me.HtmlBrowser.Document
    doc.Write(String.Empty)

#If DEBUG Then
    Me.HtmlBrowser.ScriptErrorsSuppressed = False
#Else
    Me.HtmlBrowser.ScriptErrorsSuppressed = True
#End If

    Me.HtmlBrowser.DocumentText = Me.Options.Html

  End Sub

  Private Sub HtmlBrowser_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles HtmlBrowser.DocumentCompleted
    If Not String.IsNullOrWhiteSpace(Me.Options.LoadSettingsCommand) AndAlso e.Url.Fragment = "" Then
      Dim vRtn As Object = Me.HtmlBrowser.Document.InvokeScript(Me.Options.LoadSettingsCommand, New Object() {Me.Options.Settings.ToString()})
    End If
  End Sub


  Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
    Try
      If Not String.IsNullOrWhiteSpace(Me.Options.SaveSettingsCommand) Then
        Dim vReturn As Object = HtmlBrowser.Document.InvokeScript(Me.Options.SaveSettingsCommand)
        If vReturn IsNot Nothing AndAlso TypeOf vReturn Is String Then
          Me.Options.Settings = XDocument.Parse(vReturn.ToString())
        End If

      End If
    Catch ex As Exception

    Finally
      Me.DialogResult = System.Windows.Forms.DialogResult.OK
    End Try
  End Sub
  ''' <summary>
  ''' A class that is passed to the constructor of frmHtmlSettings.  This class contains various data elements that define the behaviour of the frmHtmlSettings dialog
  ''' </summary>
  ''' <remarks></remarks>
  Public Class HtmlSettingsOptions
    ''' <summary>
    ''' The Html that will be displayed by frmHtmlSettings in its WebBrowser control
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Html As String
    ''' <summary>
    ''' The Javascript command that will  be called when the users selects Ok in frmHtmlSettings.  This must be a parameterless function that can be called on the page rendered by the Html property
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The return of the Javascript function will be stored in the HtmlSettingsOptions.SettingsObject</remarks>
    Public Property SaveSettingsCommand As String
    ''' <summary>
    ''' The object that contains all the settings required by the Html page.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This object is passed as a parameter when the function detailed in the LoadSettingsCommand is invoked</remarks>
    Public Property Settings As XDocument
    ''' <summary>
    ''' The Javascript function that will  be invoked to pass the settings to the Html page.  This function take a single parameter and must be callable on the page rendered by the Html property
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LoadSettingsCommand As String
  End Class

End Class