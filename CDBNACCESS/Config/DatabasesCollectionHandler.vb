Imports System.Configuration
Imports System.Web
Imports System.Xml

Namespace Config

  ''' <summary>
  ''' The configuration section handler for the databases collection.
  ''' </summary>
  <ConfigurationCollection(GetType(DbConfig),
                           addItemName:="database",
                           CollectionType:=ConfigurationElementCollectionType.BasicMap)>
  Public Class DatabasesCollectionHandler
    Inherits ConfigurationElementCollection

    Private Shared propertiesField As New ConfigurationPropertyCollection

    Sub New()
    End Sub

    ''' <summary>
    ''' When overridden in a derived class, creates a new <see cref="T:System.Configuration.ConfigurationElement" />.
    ''' </summary>
    ''' <returns>
    ''' A new <see cref="T:System.Configuration.ConfigurationElement" />.
    ''' </returns>
    Protected Overrides Function CreateNewElement() As ConfigurationElement
      Return New DbConfig
    End Function

    ''' <summary>
    ''' Gets the element key for a specified configuration element when overridden in a derived class.
    ''' </summary>
    ''' <param name="element">The <see cref="T:System.Configuration.ConfigurationElement" /> to return the key for.</param>
    ''' <returns>
    ''' An <see cref="T:System.Object" /> that acts as the key for the specified <see cref="T:System.Configuration.ConfigurationElement" />.
    ''' </returns>
    Protected Overrides Function GetElementKey(element As ConfigurationElement) As Object
      Return DirectCast(element, DbConfig).Name
    End Function

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
    ''' Gets the type of the <see cref="T:System.Configuration.ConfigurationElementCollection" />.
    ''' </summary>
    ''' <returns>The <see cref="T:System.Configuration.ConfigurationElementCollectionType" /> of this collection.</returns>
    Public Overrides ReadOnly Property CollectionType As ConfigurationElementCollectionType
      Get
        Return ConfigurationElementCollectionType.BasicMap
      End Get
    End Property

    ''' <summary>
    ''' Gets the name used to identify this collection of elements in the configuration file when overridden in a derived class.
    ''' </summary>
    ''' <returns>The name of the collection; otherwise, an empty string. The default is an empty string.</returns>
    Protected Overrides ReadOnly Property ElementName As String
      Get
        Return "database"
      End Get
    End Property

    ''' <summary>
    ''' Gets or sets a property or attribute of this configuration element.
    ''' </summary>
    ''' <returns>The specified property, attribute, or child element.</returns>
    '''   <param name="prop">The property to access. </param>
    Default Public Overloads Property Item(prop As Integer) As DbConfig
      Get
        Return DirectCast(MyBase.BaseGet(prop), DbConfig)
      End Get
      Set(value As DbConfig)
        If MyBase.BaseGet(prop) IsNot Nothing Then
          MyBase.BaseRemoveAt(prop)
        End If
        MyBase.BaseAdd(prop, value)
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets a property or attribute of this configuration element.
    ''' </summary>
    ''' <returns>The specified property, attribute, or child element.</returns>
    '''   <param name="prop">The property to access. </param>
    Default Public Overloads ReadOnly Property Item(prop As String) As DbConfig
      Get
        Return DirectCast(MyBase.BaseGet(prop), DbConfig)
      End Get
    End Property

  End Class

End Namespace