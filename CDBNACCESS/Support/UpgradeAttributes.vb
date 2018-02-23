Public Class UpgradeAttributes
  Public mvCol As Collection
  Private mvUpgradeTable As UpgradeTable

  Public WriteOnly Property UpgradeTable() As UpgradeTable
    Set(ByVal value As UpgradeTable)
      mvUpgradeTable = value
    End Set
  End Property
  Public ReadOnly Property Count() As Integer
    Get
      Count = mvCol.Count
    End Get
  End Property
  Public ReadOnly Property Item(ByVal pIndexKey As String) As UpgradeAttribute
    Get
      Item = CType(mvCol.Item(pIndexKey), UpgradeAttribute)
    End Get
  End Property
  Public ReadOnly Property Item(ByVal pIndexKey As Integer) As UpgradeAttribute
    Get
      Item = CType(mvCol.Item(pIndexKey), UpgradeAttribute)
    End Get
  End Property
  Public ReadOnly Property Exists(ByVal pIndexKey As String) As Boolean
    Get
      Return mvCol.Contains(pIndexKey)
    End Get
  End Property

  Public Sub New()
    mvCol = New Collection
  End Sub
  Public Sub Remove(ByVal pIndexKey As String)
    mvCol.Remove(pIndexKey)
  End Sub
  Public Sub Remove(ByVal pIndexKey As Integer)
    mvCol.Remove(pIndexKey)
  End Sub
  Public Function Add(ByVal pEnv As CDBEnvironment, ByVal pKey As String, ByVal pDataType As String, ByVal pParameter1 As String, ByVal pParameter2 As String, ByVal pParameter3 As String, ByVal pNullable As UpgradeAttribute.NullOptions) As UpgradeAttribute
    Dim vNewMember As UpgradeAttribute
    Dim vParmCount As Integer
    Dim vField As New CDBField(pKey)

    vNewMember = New UpgradeAttribute
    If pDataType.ToUpper = "NUMERIC" Then pDataType = "N"
    vField.FieldType = CDBField.GetFieldType(pDataType.ToUpper)
    pDataType = pEnv.Connection.NativeDataType(vField).Replace("()", "")
    If Not pDataType.Equals("varchar(max)", StringComparison.InvariantCultureIgnoreCase) AndAlso
       Not pDataType.Equals("varbinary(max)", StringComparison.InvariantCultureIgnoreCase) AndAlso
       pDataType.IndexOf("(") > 0 Then
      pDataType = Mid(pDataType, 1, pDataType.IndexOf("("))
    End If
    DBSetup.GetNativeDataType(pEnv.Connection, pKey, pDataType, vParmCount, False)
    With vNewMember
      .Key = pKey
      If pDataType = "char" Or pDataType = "nlschar" Then pDataType = ReplaceString(pDataType, "char", "character")
      If StrComp(pDataType, "number", vbTextCompare) = 0 And pParameter1 = "38" Then
        pDataType = "integer"
        pParameter1 = ""
      End If
      .DataType = pDataType
      .Parameter1 = IntegerValue(pParameter1)
      .Parameter2 = IntegerValue(pParameter2)
      .Parameter3 = IntegerValue(pParameter3)
      .ParameterCount = vParmCount
      .Nullable = pNullable
      .ToBeCreated = False
      .StructureModified = False
      .DefaultValue = ""
      .BeforeAttribute = ""
      .RangeCheck = ""
      .RangeCheckOnly = False
      .ToBeDeleted = False
    End With
    mvCol.Add(vNewMember, pKey)

    'return the object created
    Return vNewMember
    vNewMember = Nothing
  End Function
  Public Function AddFromChange(ByVal pEnv As CDBEnvironment, ByVal pKey As String, ByVal pDataType As String, ByVal pParameter1 As String, ByVal pParameter2 As String, ByVal pParameter3 As String, ByVal pNullable As String, ByVal pToBeCreated As Boolean, ByVal pStructureModified As Boolean, ByVal pBeforeAttrName As String, ByVal pDefaultValue As String, ByVal pRangeCheck As String) As UpgradeAttribute
    'create a new object
    Dim vNewMember As UpgradeAttribute
    Dim vParmCount As Integer
    Dim vDataMod As DataMod = New DataMod()
    Dim vDataValue As DataValue
    Dim vAddToInsert As Boolean
    Dim vNullable As UpgradeAttribute.NullOptions
    DBSetup.GetNativeDataType(pEnv.Connection, pKey, pDataType, vParmCount, , mvUpgradeTable.Key)
    If pDataType = "char" Or pDataType = "nlschar" Then
      pDataType = ReplaceString(pDataType, "char", "character")
    ElseIf Left(pDataType, 8) = "varchar2" Then
      If Mid(pDataType, 9, 1) = "(" Then pParameter1 = Math.Round(Val(Mid(pDataType, 10))).ToString
      pDataType = "varchar2"
    ElseIf Left(pDataType, 7) = "varchar" AndAlso
           Not pDataType.Equals("varchar(max)", StringComparison.InvariantCultureIgnoreCase) Then
      If Mid(pDataType, 8, 1) = "(" Then pParameter1 = Math.Round(Val(Mid(pDataType, 9))).ToString
      pDataType = "varchar"
    End If
    If pNullable = "NULLS" Then
      vNullable = UpgradeAttribute.NullOptions.noNullsAllowed
    Else
      vNullable = UpgradeAttribute.NullOptions.noNullsInvalid
    End If
    If vNullable = UpgradeAttribute.NullOptions.noNullsInvalid Then
      'add the new, mandatory attribute to any existing INSERTs to be performed on the table

      For vCtr As Integer = 1 To mvUpgradeTable.DataMods.Count
        vAddToInsert = True
        If vDataMod.ChangeType = DataModTypes.dmtInsert Then
          For Each vDataValue In mvUpgradeTable.DataMods.Item(vCtr).DataValues
            If vDataValue.Key = pKey Then vAddToInsert = False
          Next
          If vAddToInsert Then vDataMod.Add(pEnv.Connection, pKey, pDataType, pDefaultValue, "N")
        End If
      Next
    End If

    vNewMember = New UpgradeAttribute
    With vNewMember
      .Key = pKey
      .DataType = pDataType
      .Parameter1 = IntegerValue(pParameter1)
      .Parameter2 = IntegerValue(pParameter2)
      .Parameter3 = IntegerValue(pParameter3)
      .ParameterCount = vParmCount
      .Nullable = vNullable
      .StructureModified = pStructureModified
      .ToBeCreated = pToBeCreated
      .BeforeAttribute = pBeforeAttrName
      .DefaultValue = pDefaultValue
      .RangeCheck = pRangeCheck
      .RangeCheckOnly = False
      .ToBeDeleted = False
    End With
    mvCol.Add(vNewMember, pKey)

    If Not mvUpgradeTable.ToBeCreated Then
      mvUpgradeTable.StructureModified = True
      If mvUpgradeTable.ToBeDropped Then mvUpgradeTable.ToBeDropped = False
    End If
    'return the object created
    AddFromChange = vNewMember
  End Function
  
End Class
