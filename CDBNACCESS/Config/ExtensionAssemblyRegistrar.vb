Imports System.Configuration
Imports Advanced.Extensibility
Imports System.Web.Configuration
Imports CARE.Config

Namespace Config.Extensibility
  Public Class ExtensionAssemblyRegistrar

    Private vConfigPathField As String

    Public Sub New(pConfigPath As String)
      Me.ConfigurationFile = pConfigPath
    End Sub

    Public Property ConfigurationFile As String
      Get
        Return vConfigPathField
      End Get
      Set(value As String)
        vConfigPathField = value
      End Set
    End Property

    Private Function OpenConfiguration() As Configuration

      Dim vRtn As Configuration = Nothing
      If WebConfigurationManager.GetWebApplicationSection("nfpConfig") IsNot Nothing Then
        vRtn = OpenCurrentConfiguration()
      Else
        vRtn = OpenFileConfiguration()
      End If

      Return vRtn

    End Function

    Private Function OpenCurrentConfiguration() As Configuration
      Return WebConfigurationManager.OpenWebConfiguration("~")
    End Function

    Private Function OpenFileConfiguration() As Configuration
      Dim vConfigFileMap As New ExeConfigurationFileMap()
      Dim vPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)
      vPath = String.Format("{0}\{1}", vPath, ConfigurationFile)
      vConfigFileMap.ExeConfigFilename = vPath
      Return ConfigurationManager.OpenMappedExeConfiguration(vConfigFileMap, ConfigurationUserLevel.None)
    End Function

    Public Sub Register(pAssemblyInfo As Reflection.AssemblyName)
      Try
        Dim vConfig As Configuration = OpenConfiguration()
        If vConfig Is Nothing Then Throw New NullReferenceException(String.Format("Configuration could not be located."))
        Dim vSection As ConfigurationSection = vConfig.GetSection("nfpConfig")
        If vSection Is Nothing Then Throw New NullReferenceException(String.Format("Incorrect Configuration. NFP Config section missing."))
        Dim vReg As RegisteredAssemblyInfo = RegisteredAssemblyInfo.CreateInstance(pAssemblyInfo)
        Dim vExtensionAssemblies As RegisteredAssemblyCollection = TryCast(vSection, NfpConfigSection).ExtensionAssemblies
        vExtensionAssemblies.Add(vReg)
        vConfig.Save(ConfigurationSaveMode.Full)
      Catch ex As Exception
        Throw New Exception(String.Format("Failed to register Assembly {0}.  See inner exception for details.", pAssemblyInfo.Name), ex)
      End Try
    End Sub

    Public Sub Unregister(pAssemblyInfo As Reflection.AssemblyName)
      Try
        Dim vConfig As Configuration = OpenConfiguration()
        If vConfig Is Nothing Then Throw New NullReferenceException(String.Format("Configuration could not be located."))
        Dim vSection As ConfigurationSection = vConfig.GetSection("nfpConfig")
        If vSection Is Nothing Then Throw New NullReferenceException(String.Format("Incorrect Configuration. NFP Config section missing."))
        Dim vReg As RegisteredAssemblyInfo = RegisteredAssemblyInfo.CreateInstance(pAssemblyInfo)
        Dim vExtensionAssemblies As RegisteredAssemblyCollection = TryCast(vSection, NfpConfigSection).ExtensionAssemblies
        For Each vInfo As RegisteredAssemblyInfo In vExtensionAssemblies
          If vInfo.FilePath = vReg.FilePath Then
            vExtensionAssemblies.Remove(vInfo)
            Exit For
          End If
        Next
        vConfig.Save(ConfigurationSaveMode.Full)
      Catch ex As Exception
        Throw New Exception(String.Format("Failed to unregister Assembly {0}.  See inner exception for details.", pAssemblyInfo.Name), ex)
      End Try
    End Sub

  End Class

End Namespace