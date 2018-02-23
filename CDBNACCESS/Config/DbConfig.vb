Imports System.Configuration

Namespace Config

  ''' <summary>
  ''' A class to describe an available database.
  ''' </summary>
  Public Class DbConfig
    Inherits ConfigurationElement

    Private Shared nameField As New ConfigurationProperty("name", GetType(String), Nothing, ConfigurationPropertyOptions.IsKey Or ConfigurationPropertyOptions.IsRequired)
    Private Shared descriptionField As New ConfigurationProperty("description", GetType(String), Nothing, ConfigurationPropertyOptions.IsRequired)
    Private Shared connectionStringNameField As New ConfigurationProperty("connectionStringName", GetType(String), Nothing, ConfigurationPropertyOptions.IsRequired)
    Private Shared clientCodeField As New ConfigurationProperty("clientCode", GetType(String), Nothing, ConfigurationPropertyOptions.IsRequired)
    Private Shared initialiseDatabaseFromField As New ConfigurationProperty("initialiseDatabaseFrom", GetType(String), String.Empty)
    Private Shared sqlLogQueueNameField As New ConfigurationProperty("sqlLogQueueName", GetType(String), String.Empty)
    Private Shared sqlLoggingField As New ConfigurationProperty("sqlLogging", GetType(CDBConnection.SQLLoggingModes), CDBConnection.SQLLoggingModes.None)
    Private Shared propertiesField As New ConfigurationPropertyCollection

    Shared Sub New()
      propertiesField.Add(nameField)
      propertiesField.Add(descriptionField)
      propertiesField.Add(connectionStringNameField)
      propertiesField.Add(clientCodeField)
      propertiesField.Add(initialiseDatabaseFromField)
      propertiesField.Add(sqlLogQueueNameField)
      propertiesField.Add(sqlLoggingField)
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="DbConfig"/> class.
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
    ''' Gets the database name.
    ''' </summary>
    ''' <value>
    ''' The dataabase name.
    ''' </value>
    <ConfigurationProperty("name", IsRequired:=True)>
    Public ReadOnly Property Name As String
      Get
        Return DirectCast(MyBase.Item(nameField), String)
      End Get
    End Property

    ''' <summary>
    ''' Gets the database description.
    ''' </summary>
    ''' <value>
    ''' The database description.
    ''' </value>
    <ConfigurationProperty("description", IsRequired:=True)>
    Public ReadOnly Property Description As String
      Get
        Return DirectCast(MyBase.Item(descriptionField), String)
      End Get
    End Property

    ''' <summary>
    ''' Gets the connection string name.
    ''' </summary>
    ''' <value>
    ''' The connection string name.
    ''' </value>
    <ConfigurationProperty("connectionStringName", IsRequired:=True)>
    Public ReadOnly Property ConnectionStringName As String
      Get
        Return DirectCast(MyBase.Item(connectionStringNameField), String)
      End Get
    End Property

    ''' <summary>
    ''' Gets the client code.
    ''' </summary>
    ''' <value>
    ''' The client code.
    ''' </value>
    <ConfigurationProperty("clientCode", IsRequired:=True)>
    Public ReadOnly Property ClientCode As String
      Get
        Return DirectCast(MyBase.Item(clientCodeField), String)
      End Get
    End Property

    ''' <summary>
    ''' Gets the initialise location.
    ''' </summary>
    ''' <value>
    ''' The initialise location.
    ''' </value>
    ''' <remarks>This is the UNC path of the folder containing the administration files to create a new database</remarks>
    <ConfigurationProperty("initialiseDatabaseFrom", IsRequired:=False, DefaultValue:="")>
    Public ReadOnly Property InitialiseDatabaseFrom As String
      Get
        Return DirectCast(MyBase.Item(initialiseDatabaseFromField), String)
      End Get
    End Property

    ''' <summary>
    ''' Gets the name of the SQL log queue.
    ''' </summary>
    ''' <value>
    ''' The name of the SQL log queue.
    ''' </value>
    <ConfigurationProperty("sqlLogQueueName", IsRequired:=False, DefaultValue:="")>
    Public ReadOnly Property SqlLogQueueName As String
      Get
        Return DirectCast(MyBase.Item(sqlLogQueueNameField), String)
      End Get
    End Property

    ''' <summary>
    ''' Gets the SQL logging mode.
    ''' </summary>
    ''' <value>
    ''' The SQL logging mode.
    ''' </value>
    <ConfigurationProperty("sqlLogging", IsRequired:=False, DefaultValue:=CDBConnection.SQLLoggingModes.None)>
    Public ReadOnly Property SqlLogging As CDBConnection.SQLLoggingModes
      Get
        Return DirectCast(MyBase.Item(sqlLoggingField), CDBConnection.SQLLoggingModes)
      End Get
    End Property
  End Class

End Namespace
