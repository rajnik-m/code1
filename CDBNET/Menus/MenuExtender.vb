Imports System.ComponentModel

<ProvideProperty("Tag", GetType(MenuItem)), _
  ProvideProperty("RadioGroup", GetType(MenuItem))> _
  Public Class MenuExtender
  Inherits Component
  Implements IExtenderProvider

#Region " Component Designer generated code "

  Public Sub New(ByVal Container As System.ComponentModel.IContainer)
    MyClass.New()

    'Required for Windows.Forms Class Composition Designer support
    Container.Add(Me)
  End Sub

  Public Sub New()
    MyBase.New()

    'This call is required by the Component Designer.
    InitializeComponent()

    'Add any initialization after the InitializeComponent() call

  End Sub

  'Component overrides dispose to clean up the component list.
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  'Required by the Component Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Component Designer
  'It can be modified using the Component Designer.
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    components = New System.ComponentModel.Container
  End Sub

#End Region

  Private mvAutoRadioCheck As Boolean = True


#Region " Extended Properties Implementation "

  Private Props As Hashtable = New Hashtable
  Private Groups As Hashtable = New Hashtable

  Public Function CanExtend(ByVal extendee As Object) As Boolean Implements System.ComponentModel.IExtenderProvider.CanExtend
    If TypeOf extendee Is MenuItem Then
      Return True
    End If
    Return False
  End Function

  <ExtenderProvidedProperty(), _
  Category("Appearance"), _
  DefaultValue(GetType(Object), Nothing), _
  Description("User defined data associated with the control."), _
  TypeConverter(GetType(StringConverter))> _
  Public Function GetTag(ByVal target As MenuItem) As Object
    If Props.Contains(target) Then
      Return CType(Props(target), Object)
    End If
    Props.Add(target, Nothing)
    Return Nothing
  End Function

  Public Sub SetTag(ByVal target As MenuItem, ByVal value As Object)
    If Props.Contains(target) Then
      Props.Remove(target)
    End If
    Props.Add(target, value)
  End Sub

  <ExtenderProvidedProperty(), _
  Category("Behavior"), _
  DefaultValue(GetType(Integer), "-1"), _
  Description("Assign an Integer greater than -1 to group MenuItems by number." & vbCrLf & "-1 indicates no group."), _
  TypeConverter(GetType(RadioGroupConverter))> _
  Public Function GetRadioGroup(ByVal target As MenuItem) As Integer
    If TypeOf target.Parent Is MainMenu OrElse target.MenuItems.Count > 0 Then
      Groups.Remove(target)
    End If
    If Groups.Contains(target) Then
      Return CType(Groups(target), Integer)
    End If
    Groups.Add(target, -1)
    Return -1
  End Function

  Public Sub SetRadioGroup(ByVal target As MenuItem, ByVal value As Integer)
    If value < -1 Then Return
    If TypeOf target.Parent Is MainMenu OrElse target.MenuItems.Count > 0 Then
      Throw New SystemException("Top Level and Parent MenuItems cannot be checked.")
    End If
    If Groups.Contains(target) Then
      RemoveHandler target.Click, AddressOf CheckMe
      Groups.Remove(target)
    End If
    Groups.Add(target, value)
    AddHandler target.Click, AddressOf CheckMe
  End Sub

  Private Class RadioGroupConverter
    Inherits TypeConverter

    Public Overloads Overrides Function CanConvertFrom(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal sourceType As System.Type) As Boolean
      If sourceType Is GetType(String) Then Return True
      Return MyBase.CanConvertFrom(context, sourceType)
    End Function

    Public Overloads Overrides Function ConvertFrom(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object) As Object
      If TypeOf value Is String Then
        If value.Equals(String.Empty) OrElse value.ToString = "" Then
          Return -1
        Else
          Return Integer.Parse(value.ToString)
        End If
      End If
      Return MyBase.ConvertFrom(context, culture, value)
    End Function

    Public Overloads Overrides Function CanConvertTo(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal destinationType As System.Type) As Boolean
      If destinationType Is GetType(String) Then Return True
      Return MyBase.CanConvertTo(context, destinationType)
    End Function

    Public Overloads Overrides Function ConvertTo(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object, ByVal destinationType As System.Type) As Object
      If destinationType Is GetType(String) Then
        If value.ToString = "-1" Then Return String.Empty
        Return value.ToString
      End If
      Return MyBase.ConvertTo(context, culture, value, destinationType)
    End Function

  End Class

#End Region

  <DefaultValue(True)> _
  Public Property AutoRadioCheck() As Boolean
    Get
      Return mvAutoRadioCheck
    End Get
    Set(ByVal Value As Boolean)
      mvAutoRadioCheck = Value
    End Set
  End Property

  Private Sub CheckMe(ByVal sender As Object, ByVal e As EventArgs)
    Dim Item As MenuItem = DirectCast(sender, MenuItem)
    If Item.Checked Then Return
    For Each de As DictionaryEntry In Groups
      If de.Value.Equals(Groups.Item(Item)) Then
        CType(de.Key, MenuItem).Checked = False
      End If
    Next
    If mvAutoRadioCheck Then Item.RadioCheck = True
    Item.Checked = True
  End Sub
End Class
