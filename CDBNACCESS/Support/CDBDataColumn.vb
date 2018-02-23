Namespace Access

  Public Class CDBDataColumn

    Private mvName As String
    Private mvType As CDBField.FieldTypes
    Private mvIndex As Integer
    Private mvAttributeName As String

    Friend Sub New(ByVal pName As String, ByVal pType As CDBField.FieldTypes, ByVal pIndex As Integer)
      mvName = pName
      mvType = pType
      mvIndex = pIndex
    End Sub

    Public Overrides Function ToString() As String
      ToString = mvName + " : " + FieldType().ToString()
    End Function

    Public Property AttributeName() As String
      Get
        Return mvAttributeName
      End Get
      Set(ByVal pValue As String)
        Dim vPos As Integer = pValue.IndexOf(".")
        If vPos >= 0 Then
          mvAttributeName = pValue.Substring(vPos + 1)
        Else
          mvAttributeName = pValue
        End If
      End Set
    End Property

    Public ReadOnly Property Index() As Integer
      Get
        Return mvIndex
      End Get
    End Property

    Public Property Name() As String
      Get
        Return mvName
      End Get
      Set(ByVal value As String)
        mvName = value
      End Set
    End Property

    Public Property FieldType() As CDBField.FieldTypes
      Get
        Return mvType
      End Get
      Set(ByVal pValue As CDBField.FieldTypes)
        mvType = pValue
      End Set
    End Property

    Friend Sub SetStandardColumnName()
      mvName = ProperName(mvName)
    End Sub
  End Class
End Namespace