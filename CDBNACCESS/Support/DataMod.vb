Namespace Access
  Public Class DataMod
    Private mvCol As Collection
    Private mvChangeType As DataModTypes


    Public ReadOnly Property DataValues() As Collection
      Get
        DataValues = mvCol
      End Get
    End Property
    Public ReadOnly Property Item(ByVal pIndexKey As Integer) As DataValue
      Get
        Item = CType(mvCol(pIndexKey), DataValue)
      End Get
    End Property
    Public ReadOnly Property Item(ByVal pIndexKey As String) As DataValue
      Get
        Item = CType(mvCol(pIndexKey), DataValue)
      End Get
    End Property
    Public ReadOnly Property Count() As Integer
      Get
        Count = mvCol.Count
      End Get
    End Property

    Public Property ChangeType() As DataModTypes
      Get
        ChangeType = mvChangeType
      End Get
      Set(ByVal Value As DataModTypes)
        mvChangeType = Value
      End Set
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
    Public Sub Add(ByVal pConn As CDBConnection, ByVal pKey As String, ByVal pDataType As String, ByVal pAttrValue As String, ByVal pKeyValue As String)
      Dim vNewMember As DataValue
      Dim vParmCount As Integer

      vNewMember = New DataValue
      'set the properties passed into the method
      If pKeyValue.Length < 1 Then pKeyValue = "N"
      With vNewMember
        .Key = pKey & pKeyValue
        .Attr = pKey
        DBSetup.GetNativeDataType(pConn, pKey, pDataType, vParmCount)
        .DataType = pDataType
        .AttrValue = pAttrValue
        .KeyValue = CBool(IIf(pKeyValue = "Y", True, False))
      End With
      mvCol.Add(vNewMember, pKey & pKeyValue)
      vNewMember = Nothing
    End Sub
    Public Function IsReplacedBy(ByVal pDataMod As DataMod) As Boolean
      Dim vIndex As Integer
      Dim vMatches As Boolean

      With pDataMod
        If ChangeType = .ChangeType Then
          vMatches = True
          
          For vIndex = 1 To Count
            If Item(vIndex).Attr = .Item(vIndex).Attr Then
              If Item(vIndex).KeyValue = True And .Item(vIndex).KeyValue = True Then
                If Item(vIndex).AttrValue <> .Item(vIndex).AttrValue Then
                  vMatches = False
                  Exit For
                End If
              End If
            Else
              vMatches = False
              Exit For
            End If
          Next
          If vMatches Then mvCol = pDataMod.DataValues
        End If
      End With
      Return vMatches
    End Function
    Public Function WhereFields(ByVal pConn As CDBConnection) As CDBFields
      Dim vDataValue As DataValue
      Dim vWhereFields As New CDBFields
      Dim vType As CDBField.FieldTypes
      Dim vField As CDBField

      For Each vDataValue In mvCol
        'vDataValue = CType(mvCol.Item(vCtr), DataValue)
        vType = vDataValue.FieldType
        If vDataValue.KeyValue Then
          vField = vWhereFields.Add(vDataValue.Attr, vType, vDataValue.AttrValue)
          If pConn.IsSpecialColumn(vDataValue.Attr) Then vField.SpecialColumn = True
        End If
      Next
      Return vWhereFields
    End Function

    Public Function Fields(ByVal pConn As CDBConnection) As CDBFields
      Dim vDataValue As DataValue
      Dim vFields As New CDBFields
      Dim vType As CDBField.FieldTypes
      Dim vField As CDBField

      For Each vDataValue In mvCol
        'vDataValue = CType(mvCol.Item(vkey), DataValue)
        vType = vDataValue.FieldType
        If vDataValue.KeyValue And ChangeType = DataModTypes.dmtUpdate Then
          'do nothing
        ElseIf vFields.Exists(vDataValue.Attr) Then
          'Field is in as value and as key?
          Debug.Print(vDataValue.Key & " Field is in as value and as key?")
        Else
          vField = vFields.Add(vDataValue.Attr, vType, vDataValue.AttrValue)
          If pConn.IsSpecialColumn(vDataValue.Attr) Then vField.SpecialColumn = True
        End If
      Next
      Return vFields
    End Function
    Public Function ChangeApplied(ByVal pTable As UpgradeTable) As Boolean
      Dim vWhereFields As CDBFields
      Dim vCount As Long

      vWhereFields = WhereFields(pTable.Connection)
      If vWhereFields.Count > 0 Then
        vCount = pTable.Connection.GetCount(pTable.Key, vWhereFields)
        If ChangeType = DataModTypes.dmtInsert And vCount > 0 Then
          ChangeApplied = True
        ElseIf ChangeType = DataModTypes.dmtDelete And vCount = 0 Then
          ChangeApplied = True
        End If
      End If
    End Function
  End Class

  Public Enum DataModTypes
    dmtInsert
    dmtUpdate
    dmtDelete
  End Enum
End Namespace
