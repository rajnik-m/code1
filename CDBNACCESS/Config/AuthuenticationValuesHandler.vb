Imports System.Configuration

Namespace Config

  ''' <summary>
  ''' A class to get Authentication Values which are passed in Web Service calls. Currently usded by Postcoders QASOnDemand and AFD Evolution
  ''' </summary>
  Public Class AuthenticationValuesHandler
    Inherits ConfigurationElement

    Private Shared username As New ConfigurationProperty("username", GetType(String), Nothing, ConfigurationPropertyOptions.IsRequired)
    Private Shared password As New ConfigurationProperty("password", GetType(String), Nothing, ConfigurationPropertyOptions.IsRequired)
    Private Shared propertiesField As New ConfigurationPropertyCollection

    Shared Sub New()
      propertiesField.Add(username)
      propertiesField.Add(password)
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="AuthenticationValuesHandler"/> class.
    ''' </summary>
    ''' <remarks>Prevent anything outside this assembly instanceating this class.</remarks>
    Sub New()
    End Sub

    ''' <summary>
    ''' Gets the collection of properties.
    ''' </summary>
    ''' <returns>The <see cref="T:System.Configuration.ConfigurationPropertyCollection" /> of properties for the element.</returns>
    Protected Overrides ReadOnly Property Properties As ConfigurationPropertyCollection
      Get
        Return propertiesField
      End Get
    End Property

    ''' <summary>
    ''' Gets the username.
    ''' </summary>
    ''' <value>
    ''' The username/serial number.
    ''' </value>
    <ConfigurationProperty("username", IsRequired:=True)>
    Public ReadOnly Property UsernameValue As String
      Get
        Return DirectCast(MyBase.Item(Username), String)
      End Get
    End Property

    ''' <summary>
    ''' Gets the password.
    ''' </summary>
    ''' <value>
    ''' The password.
    ''' </value>
    <ConfigurationProperty("password", IsRequired:=True)>
    Public ReadOnly Property PasswordValue As String
      Get
        Return DirectCast(MyBase.Item(password), String)
      End Get
    End Property

  End Class

End Namespace
