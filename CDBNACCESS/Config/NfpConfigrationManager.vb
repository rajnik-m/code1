Imports System.Configuration
Imports System.Web.Configuration
Imports Advanced.Extensibility
Imports System.Linq

Namespace Config

  ''' <summary>
  ''' A helper class to simplify the use of the configuration data in the web.config file.
  ''' </summary>
  Public Class NfpConfigrationManager

    ''' <summary>
    ''' Gets the colection of <see cref="DbConfig" /> database definition classes from the config file.
    ''' </summary>
    ''' <value>
    ''' The database definitions.
    ''' </value>
    Public Shared ReadOnly Property Databases As DatabasesCollectionHandler
      Get
        Dim Result As DatabasesCollectionHandler = Nothing
        If WebConfigurationManager.GetWebApplicationSection("nfpConfig") IsNot Nothing Then
          Result = DirectCast(WebConfigurationManager.GetWebApplicationSection("nfpConfig"), NfpConfigSection).Databases
        ElseIf ConfigurationManager.GetSection("nfpConfig") IsNot Nothing Then
          Result = DirectCast(WebConfigurationManager.GetWebApplicationSection("nfpConfig"), NfpConfigSection).Databases
        End If
        Return Result
      End Get
    End Property

    Public Shared ReadOnly Property ExtensionAssemblies As RegisteredAssemblyCollection
      Get
        Dim Result As RegisteredAssemblyCollection = Nothing
        If WebConfigurationManager.GetWebApplicationSection("nfpConfig") IsNot Nothing Then
          Result = DirectCast(WebConfigurationManager.GetWebApplicationSection("nfpConfig"), NfpConfigSection).ExtensionAssemblies
        ElseIf ConfigurationManager.GetSection("nfpConfig") IsNot Nothing Then
          Result = DirectCast(WebConfigurationManager.GetWebApplicationSection("nfpConfig"), NfpConfigSection).ExtensionAssemblies
        End If
        Return Result
      End Get
    End Property
    Public Shared ReadOnly Property QAAuthenticationValues As AuthenticationValuesHandler
      Get
        Dim Result As AuthenticationValuesHandler = Nothing
        If WebConfigurationManager.GetWebApplicationSection("nfpConfig") IsNot Nothing Then
          Result = DirectCast(WebConfigurationManager.GetWebApplicationSection("nfpConfig"), NfpConfigSection).QAAuthenticationValues
        ElseIf ConfigurationManager.GetSection("nfpConfig") IsNot Nothing Then
          Result = DirectCast(WebConfigurationManager.GetWebApplicationSection("nfpConfig"), NfpConfigSection).QAAuthenticationValues
        End If
        Return Result
      End Get
    End Property
    Public Shared ReadOnly Property BankFinderAuthenticationValues As AuthenticationValuesHandler
      Get
        Dim Result As AuthenticationValuesHandler = Nothing
        If WebConfigurationManager.GetWebApplicationSection("nfpConfig") IsNot Nothing Then
          Result = DirectCast(WebConfigurationManager.GetWebApplicationSection("nfpConfig"), NfpConfigSection).BankFinderAuthenticationValues
        ElseIf ConfigurationManager.GetSection("nfpConfig") IsNot Nothing Then
          Result = DirectCast(WebConfigurationManager.GetWebApplicationSection("nfpConfig"), NfpConfigSection).BankFinderAuthenticationValues
        End If
        Return Result
      End Get
    End Property
    Public Shared Function RegisterWebExtensionAssembly(pAssemblyInfo As Reflection.AssemblyName) As Boolean

      Dim vRegistrar = New Config.Extensibility.ExtensionAssemblyRegistrar("..\Web.Config")

      vRegistrar.Register(pAssemblyInfo)

    End Function

    Public Shared Function UnregisterWebExtensionAssembly(pAssemblyInfo As Reflection.AssemblyName) As Boolean

      Dim vRegistrar = New Config.Extensibility.ExtensionAssemblyRegistrar("..\Web.Config")

      vRegistrar.Unregister(pAssemblyInfo)

    End Function

    Public Shared Function RegisterJobSchedulerExtensionAssembly(pAssemblyInfo As Reflection.AssemblyName) As Boolean

      Dim vRegistrar As New Config.Extensibility.ExtensionAssemblyRegistrar("JobProcessor.Exe.Config")

      vRegistrar.Register(pAssemblyInfo)

    End Function
    Public Shared Function UnregisterJobSchedulerExtensionAssembly(pAssemblyInfo As Reflection.AssemblyName) As Boolean

      Dim vRegistrar As New Config.Extensibility.ExtensionAssemblyRegistrar("JobProcessor.Exe.Config")

      vRegistrar.Unregister(pAssemblyInfo)

    End Function

  End Class
End Namespace
