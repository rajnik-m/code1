Public Class frmCardAuthorisation
  Inherits ThemedForm

  Private cardAuthoriser As WebBasedCardAuthoriser
  Private serverParameters As ParameterList
  Private contactNumber As Integer
  Private addressNumber As Integer
  Private transactionType As String
  Private transactionAmount As Integer
  Private batchCategory As String

  <Obsolete("For designer use only")>
  Public Sub New()
    InitializeComponent()
    If Not Me.DesignMode Then
      Throw New NotSupportedException("The default constructor is only supported in design mode")
    End If
  End Sub

  Public Sub New(contactNumber As Integer,
                 addressNumber As Integer,
                 transactionType As String,
                 transactionAmount As Integer,
                 batchCategory As String,
                 parameters As ParameterList)
    InitializeComponent()
    Me.contactNumber = contactNumber
    Me.addressNumber = addressNumber
    Me.transactionType = transactionType
    Me.transactionAmount = transactionAmount
    Me.batchCategory = batchCategory
    Me.serverParameters = parameters
  End Sub

  Private Sub frmCardAuthorisation_Load(sender As Object, e As EventArgs) Handles Me.Load
    Me.cardAuthoriser = WebBasedCardAuthoriser.GetInstance(Me.browser)
    AddHandler Me.cardAuthoriser.ProcessingComplete, AddressOf Me.CardAuthorisationComplete
    Me.InitCardAuthorisation()
  End Sub

  Private Sub frmCardAuthorisation_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    If Not Me.cardAuthoriser.IsAuthorised AndAlso
       Settings.ConfirmCancel Then
      e.Cancel = Not ConfirmCancel()
    Else
      e.Cancel = False
    End If

  End Sub

  Private Sub InitCardAuthorisation()
    Me.cardAuthoriser.RequestAuthorisation(Me.contactNumber,
                                           Me.addressNumber,
                                           Me.transactionType,
                                           CInt(Me.transactionAmount),
                                           Me.batchCategory,
                                           String.Empty)
  End Sub

  Private Sub CardAuthorisationComplete(sender As Object, e As EventArgs)
    If Me.cardAuthoriser.IsAuthorised AndAlso
       Not Me.cardAuthoriser.IsCancelled Then
      Me.cardAuthoriser.SetServerValues(Me.serverParameters)
      Me.Close()
    ElseIf Me.cardAuthoriser.IsCancelled Then
      Me.Close()
      Me.InitCardAuthorisation()
    Else
      Me.InitCardAuthorisation()
    End If
  End Sub

  Public ReadOnly Property isAuthorised As Boolean
    Get
      Return If(Me.cardAuthoriser IsNot Nothing,
                Me.cardAuthoriser.IsAuthorised,
                False)
    End Get
  End Property

  Public ReadOnly Property isCancelled As Boolean
    Get
      Return If(Me.cardAuthoriser IsNot Nothing,
                Me.cardAuthoriser.IsCancelled,
                False)
    End Get
  End Property
End Class