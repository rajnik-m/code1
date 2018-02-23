Imports System.Reflection
Imports System.Collections.Specialized
Imports System.Web

Public MustInherit Class WebBasedCardAuthoriser

  Shared Sub New()
    Dim configuredType As String = AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_cc_authorisation_type)
    For Each candidateProviderClass As System.Type In From vAssembly In AppDomain.CurrentDomain.GetAssemblies()
                                       From vType In vAssembly.GetTypes()
                                       Where vType.IsSubclassOf(GetType(WebBasedCardAuthoriser))
                                       Select vType
      For Each vAttribute As AuthorisationMethodAttribute In candidateProviderClass.GetCustomAttributes(GetType(AuthorisationMethodAttribute), True)
        If vAttribute.AuthorisationMethod.Equals(configuredType, StringComparison.InvariantCultureIgnoreCase) Then
          If ProviderClass Is Nothing Then
            ProviderClass = candidateProviderClass
#If DEBUG Then
          Else
            Debug.Assert(False, String.Format("Multiple card authorisation classes found for {0}.", configuredType))
#End If
          End If
        End If
      Next vAttribute
    Next candidateProviderClass
  End Sub

  Public Shared Function GetInstance(browser As WebBrowser) As WebBasedCardAuthoriser
    Dim provider As WebBasedCardAuthoriser = Nothing
    If ProviderClass IsNot Nothing Then
      Try
        Dim constructor As ConstructorInfo = ProviderClass.GetConstructor(BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, {GetType(WebBrowser)}, Nothing)
        provider = DirectCast(constructor.Invoke({browser}), WebBasedCardAuthoriser)
      Catch vEx As Exception
        If vEx.GetType Is GetType(TargetInvocationException) AndAlso
           vEx.InnerException IsNot Nothing Then
          vEx = vEx.InnerException
        End If
        Throw vEx
      End Try
    End If
    Return provider
  End Function

  Public Shared ReadOnly Property IsAvailable() As Boolean
    Get
      Return ProviderClass IsNot Nothing
    End Get
  End Property

  Protected Sub New(browser As WebBrowser)
    Me.Browser = browser
  End Sub

  Public MustOverride Sub RequestAuthorisation(contactNumber As Integer,
                                               addressNumber As Integer,
                                               transactionType As String,
                                               transactionAmount As Integer,
                                               batchCategory As String,
                                               merchantRetailNumber As String)
  Public MustOverride Sub SetServerValues(list As ParameterList)

  Private browserControl As WebBrowser = Nothing

  Protected Property Browser As WebBrowser
    Get
      Return browserControl
    End Get
    Private Set(value As WebBrowser)
      If browserControl IsNot Nothing Then
        RemoveHandler browserControl.Navigated, AddressOf Me.Browser_Navigated
      End If
      browserControl = value
      AddHandler browserControl.Navigated, AddressOf Me.Browser_Navigated
    End Set
  End Property

  Private Shared Property ProviderClass As Type = Nothing

  Private isAuthorisedField As Boolean = False
  Public Property IsAuthorised As Boolean
    Get
      Return isAuthorisedField
    End Get
    Protected Set(value As Boolean)
      isAuthorisedField = value
    End Set
  End Property

  Private isCancelledField As Boolean = False
  Public Property IsCancelled As Boolean
    Get
      Return isCancelledField
    End Get
    Protected Set(value As Boolean)
      isCancelledField = value
    End Set
  End Property

  Public Event ProcessingComplete(ByVal sender As Object, ByVal e As EventArgs)

  Protected Sub RaiseProcesesingComplete()
    RaiseEvent ProcessingComplete(Me, New EventArgs)
  End Sub

  Private Sub Browser_Navigated(sender As Object, e As WebBrowserNavigatedEventArgs)
    Me.OnWebBrowserNavigated()
  End Sub

  Protected MustOverride Sub OnWebBrowserNavigated()

  Protected ReadOnly Property ReturnParameter(name As String) As String
    Get
      Dim result As String = String.Empty
      If Me.Browser.Url IsNot Nothing Then
        Dim parameterValue As String = HttpUtility.ParseQueryString(Me.Browser.Url.Query).Get(name)
        If parameterValue IsNot Nothing Then
          result = Uri.UnescapeDataString(parameterValue)
        End If
      End If
      Return If(result Is Nothing, String.Empty, result)
    End Get
  End Property

  Private tokenField As String = String.Empty
  Public Property Token As String
    Get
      Return tokenField
    End Get
    Set(value As String)
      tokenField = value
    End Set
  End Property

  Private createTokenField As Boolean = False
  Public Property CreateToken As Boolean
    Get
      Return createTokenField
    End Get
    Set(value As Boolean)
      createTokenField = value
    End Set
  End Property

  Private cardTypeField As String = String.Empty
  Public Property CardType As String
    Get
      Return cardTypeField
    End Get
    Protected Set(value As String)
      cardTypeField = value
    End Set
  End Property

  Private cardDigitsField As String = String.Empty
  Public Property CardDigits As String
    Get
      Return cardDigitsField
    End Get
    Protected Set(value As String)
      cardDigitsField = value
    End Set
  End Property

  Private cardExpiryField As String = String.Empty
  Public Property CardExpiry As String
    Get
      Return cardExpiryField
    End Get
    Protected Set(value As String)
      cardExpiryField = value
    End Set
  End Property

  Private statusField As String = String.Empty
  Public Property Status As String
    Get
      Return statusField
    End Get
    Protected Set(value As String)
      statusField = value
    End Set
  End Property

  Private statusDetailField As String = String.Empty
  Public Property StatusDetail As String
    Get
      Return statusDetailField
    End Get
    Protected Set(value As String)
      statusDetailField = value
    End Set
  End Property

  Private vendorCodeField As String = String.Empty
  Public Property VendorCode As String
    Get
      Return vendorCodeField
    End Get
    Protected Set(value As String)
      vendorCodeField = value
    End Set
  End Property

  Private authCodeField As String = String.Empty
  Public Property AuthCode As String
    Get
      Return authCodeField
    End Get
    Protected Set(value As String)
      authCodeField = value
    End Set
  End Property

End Class

Public Class AuthorisationMethodAttribute
  Inherits Attribute
  ''' <summary>
  ''' Initializes a new instance of the <see cref="AuthorisationMethodAttribute"/> class.
  ''' </summary>
  ''' <param name="authMethod">The p task.</param>
  Public Sub New(authMethod As String)
    authorisationMethodField = authMethod
  End Sub

  Private authorisationMethodField As String
  ''' <summary>
  ''' Gets the authorisation method associated with the task.
  ''' </summary>
  ''' <value>
  ''' The type of the task.
  ''' </value>
  Public ReadOnly Property AuthorisationMethod As String
    Get
      Return authorisationMethodField
    End Get
  End Property

End Class