Namespace Access

  Public Class CDBParameters
    Inherits CollectionList(Of CDBParameter)

    Public Sub New()
      MyBase.New(1)
    End Sub

    Public Overloads Function Add(ByVal pName As String) As CDBParameter
      Dim vParam As New CDBParameter(pName)
      MyBase.Add(pName, vParam)
      Return vParam
    End Function

    Public Overloads Function Add(ByVal pName As String, ByVal pValue As String) As CDBParameter
      Dim vParam As New CDBParameter(pName, pValue)
      MyBase.Add(pName, vParam)
      Return vParam
    End Function

    Public Overloads Function Add(ByVal pName As String, ByVal pValue As Integer) As CDBParameter
      Dim vParam As New CDBParameter(pName, pValue)
      MyBase.Add(pName, vParam)
      Return vParam
    End Function

    Public Overloads Function Add(ByVal pName As String, ByVal pValue As Double) As CDBParameter
      Dim vParam As New CDBParameter(pName, CDBField.FieldTypes.cftNumeric, pValue.ToString)
      MyBase.Add(pName, vParam)
      Return vParam
    End Function

    Public Overloads Function Add(ByVal pName As String, ByVal pType As CDBField.FieldTypes) As CDBParameter
      Dim vParam As New CDBParameter(pName, pType)
      MyBase.Add(pName, vParam)
      Return vParam
    End Function

    Public Overloads Function Add(ByVal pName As String, ByVal pType As CDBField.FieldTypes, ByVal pValue As String) As CDBParameter
      Dim vParam As New CDBParameter(pName, pType, pValue)
      MyBase.Add(pName, vParam)
      Return vParam
    End Function

    Public Overloads Function Add(ByVal pParameter As CDBParameter) As CDBParameter
      MyBase.Add(pParameter.Name, pParameter)
      Return pParameter
    End Function

    Public Overloads Function Insert(ByVal pIndex As Integer, ByVal pName As String) As CDBParameter
      Dim vParam As New CDBParameter(pName)
      MyBase.Insert(pIndex, pName, vParam)
      Return vParam
    End Function

    Public Function Exists(ByVal pKey As String) As Boolean
      Return ContainsKey(pKey)
    End Function

    Public Function HasValue(ByVal pKey As String) As Boolean
      Return ContainsKey(pKey) AndAlso Item(pKey).Value.Length > 0
    End Function

    Public Sub InitFromUniqueList(ByVal pList As String)
      If pList.Length > 0 Then
        Dim vItems() As String = pList.Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
        For Each vItem As String In vItems
          Add(vItem, vItem)
        Next
      End If
    End Sub

    Public Sub InitFromUniqueList(ByVal pNameList As String, ByVal pValueList As String, ByVal pSeparator As String)
      If pNameList.Length > 0 AndAlso pValueList.Length > 0 Then
        Dim vNameItems() As String = pNameList.Split(pSeparator.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
        Dim vValueItems() As String = pValueList.Split(pSeparator.ToCharArray, StringSplitOptions.None)
        For vIndex As Integer = 0 To vNameItems.GetUpperBound(0)
          Add(vNameItems(vIndex), vValueItems(vIndex))
        Next
      End If
    End Sub

    Public Sub InitKeysFromUniqueList(ByVal pList As String)
      If pList.Length > 0 Then
        Dim vItems() As String = pList.Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
        For Each vItem As String In vItems
          Add(vItem, "")
        Next
      End If
    End Sub

    Public Sub InitFromSQLAttributes(ByVal pSQL As String)
      Dim vSQL As String
      Dim vPos As Integer
      Dim vPos2 As Integer
      Dim vItems() As String
      Dim vIndex As Integer
      Dim vParam As CDBParameter
      Dim vInBrackets As Boolean

      vSQL = pSQL.Replace(vbCr, "").Replace(vbLf, "")
      vPos = (vSQL.IndexOf("/*", 0) + 1)
      If vPos > 0 Then
        vPos2 = (vSQL.IndexOf("*/", 0) + 1)
        If vPos2 > 0 Then
          vSQL = vSQL.Substring(0, vPos - 1) & vSQL.Substring(vPos2 + 3 - 1)
        End If
      End If
      vPos = (vSQL.ToUpper().IndexOf("SELECT ", 0) + 1)
      If vPos > 0 Then
        vSQL = vSQL.Substring(vPos + 7 - 1)
      Else
        vSQL = pSQL
      End If
      vPos = (vSQL.ToUpper().IndexOf("FROM ", 0) + 1)
      If vPos > 0 Then vSQL = vSQL.Substring(0, vPos - 1)
      vItems = vSQL.Split(","c)
      For vIndex = 0 To vItems.GetUpperBound(0)
        If vInBrackets Then
          vPos = (vItems(vIndex).ToUpper().IndexOf(" AS ", 0) + 1)
          If vPos > 0 Then vInBrackets = False
        Else
          vPos = (vItems(vIndex).IndexOf("(", 0) + 1)
          If vPos > 0 Then
            vPos = (vItems(vIndex).IndexOf(")", 0) + 1)
            If vPos = 0 Then
              vInBrackets = True
            End If
          End If
        End If
        If Not vInBrackets Then
          vPos = (vItems(vIndex).ToUpper().IndexOf(" AS ", 0) + 1)
          If vPos > 0 Then vItems(vIndex) = vItems(vIndex).Substring(vPos + 4 - 1)
          vPos = (vItems(vIndex).IndexOf(".", 0) + 1)
          If vPos > 0 Then vItems(vIndex) = vItems(vIndex).Substring(vPos)
          vItems(vIndex) = vItems(vIndex).Trim(" "c)
          vParam = New CDBParameter(vItems(vIndex))
          MyBase.Add(vParam.Name, vParam)
        End If
      Next
    End Sub

    Public ReadOnly Property InList() As String
      Get
        Dim vList As New StringBuilder
        Dim vAddSeparator As Boolean
        For Each vParam As CDBParameter In Me
          If vAddSeparator Then vList.Append(",")
          vList.Append("'")
          vList.Append(vParam.Name)
          vList.Append("'")
          vAddSeparator = True
        Next
        Return vList.ToString
      End Get
    End Property

    Public ReadOnly Property ItemList() As String
      Get
        Return ItemList(",", True)
      End Get
    End Property
    Public ReadOnly Property ItemList(ByVal pSeparator As String, ByVal pIncludeBlankItems As Boolean) As String
      Get
        Dim vList As New StringBuilder
        Dim vAddSeparator As Boolean
        For Each vParam As CDBParameter In Me
          If pIncludeBlankItems OrElse (Not pIncludeBlankItems AndAlso vParam.Name.Length > 0) Then
            If vAddSeparator Then vList.Append(pSeparator)
            vList.Append(vParam.Name)
            vAddSeparator = True
          End If
        Next
        Return vList.ToString
      End Get
    End Property

    Public ReadOnly Property StandardColumnNameList() As String
      Get
        Dim vList As New StringBuilder
        Dim vAddSeparator As Boolean
        For Each vParam As CDBParameter In Me
          If vAddSeparator Then vList.Append(",")
          Dim vName As String = vParam.Name
          Dim vPos As Integer = vName.IndexOf(".")
          If vPos > 0 Then vName = vName.Substring(vPos + 1)
          vList.Append(ProperName(vName))
          vAddSeparator = True
        Next
        Return vList.ToString
      End Get
    End Property

    Public Function OptionalValue(ByVal pKey As String, ByVal pValue As String) As String
      If ContainsKey(pKey) Then
        Return Item(pKey).Value.ToString
      Else
        Return pValue
      End If
    End Function

    Public Function OptionalValue(ByVal pKey As String, ByVal pValue As Double) As String
      If ContainsKey(pKey) Then
        Return Item(pKey).Value.ToString
      Else
        Return pValue.ToString
      End If
    End Function

    Public Function OptionalValue(ByVal pKey As String, ByVal pValue As Integer) As Integer
      If ContainsKey(pKey) Then
        Return IntegerValue(Item(pKey).Value.ToString)
      Else
        Return pValue
      End If
    End Function

    Public Function ParameterExists(ByVal pKey As String) As CDBParameter
      If ContainsKey(pKey) Then
        Return Item(pKey)
      Else
        Return New CDBParameter(pKey)
      End If
    End Function

    Public Function NameList() As List(Of String)
      Dim vList As New List(Of String)
      For Each vParam As CDBParameter In Me
        vList.Add(vParam.Name)
      Next
      Return vList
    End Function

    Public Sub SetProperNames()
      For Each vParam As CDBParameter In Me
        If vParam.Name.Contains("_") OrElse Char.IsLower(vParam.Name(0)) Then
          Dim vNewName As String = ProperName(vParam.Name)
          Me.ChangeKey(vParam, vParam.Name, vNewName)
          vParam.Name = vNewName
        End If
      Next
    End Sub
  End Class
End Namespace