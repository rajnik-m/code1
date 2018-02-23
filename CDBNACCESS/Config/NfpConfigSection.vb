Imports System.Configuration
Imports Advanced.Extensibility

Namespace Config

  ''' <summary>
  ''' The NFP configuration section definition.
  ''' </summary>
  Public Class NfpConfigSection
    Inherits ConfigurationSection

    Private Shared databasesField As New ConfigurationProperty("databases", GetType(DatabasesCollectionHandler))
    Private Shared extensionAssembliesField As New ConfigurationProperty("extensionAssemblies", GetType(RegisteredAssemblyCollection))
    Private Shared qaAuthenticationField As New ConfigurationProperty("QAAuthenticationValues", GetType(AuthenticationValuesHandler))
    Private Shared BankFinderAuthenticationField As New ConfigurationProperty("BankFinderAuthenticationValues", GetType(AuthenticationValuesHandler))
    Private Shared propertiesField As New ConfigurationPropertyCollection

    Shared Sub New()
      propertiesField.Add(databasesField)
      propertiesField.Add(extensionAssembliesField)
      propertiesField.Add(qaAuthenticationField)
      propertiesField.Add(BankFinderAuthenticationField)
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
    ''' Gets the colection of <see cref="DbConfig" /> database definition classes.
    ''' </summary>
    ''' <value>
    ''' The database definitions.
    ''' </value>
    Public ReadOnly Property Databases As DatabasesCollectionHandler
      Get
        Return DirectCast(MyBase.Item(databasesField), DatabasesCollectionHandler)
      End Get
    End Property

    Public ReadOnly Property ExtensionAssemblies As RegisteredAssemblyCollection
      Get
        Return DirectCast(MyBase.Item(extensionAssembliesField), RegisteredAssemblyCollection)
      End Get
    End Property

    Public ReadOnly Property QAAuthenticationValues As AuthenticationValuesHandler
      Get
        Return DirectCast(MyBase.Item(qaAuthenticationField), AuthenticationValuesHandler)
      End Get
    End Property
    Public ReadOnly Property BankFinderAuthenticationValues As AuthenticationValuesHandler
      Get
        Return DirectCast(MyBase.Item(BankFinderAuthenticationField), AuthenticationValuesHandler)
      End Get
    End Property
  End Class

End Namespace