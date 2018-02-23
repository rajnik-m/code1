Imports System.Linq
Imports CARE

Public Class CARERecordFactory

  Private Shared mvInstance As CARERecordFactory

  ''' <summary>
  ''' This method is non-private to allow people to use as many instances as they want where needed.
  ''' There should really be no need and you should used the Instance property, but there is no reason why a factory should be a singleton
  ''' </summary>
  Public Sub New()

  End Sub

  Public Shared Property Instance As CARERecordFactory
    Get
      If mvInstance Is Nothing Then CARERecordFactory.Instance = New CARERecordFactory()
      Return mvInstance
    End Get
    Private Set(value As CARERecordFactory)
      mvInstance = value
    End Set
  End Property

  Public Function GetTableMaintenanceData(pEnv As CDBEnvironment, pTableName As String, Optional pWhere As CDBFields = Nothing) As IEnumerable(Of TableMaintenanceData)
    'Create a factory method to pass to the CARERecordFactory to generate the TableMaintenance class.  It is needed because the constructor is not a standard CARERecord constructor, and takes the TM table name in the constructor
    Dim vTMInstantiator As Func(Of TableMaintenanceData) = Function() (New TableMaintenanceData(pEnv, pTableName))
    Return CARERecordFactory.SelectList(Of TableMaintenanceData)(pEnv, pWhere, vTMInstantiator)
  End Function

  Public Function GetList(Of T As CARERecord)(pEnv As CDBEnvironment, pWhere As CDBFields) As IList(Of T)
    Dim vReturn As New List(Of T)
    Dim vResult As IEnumerable(Of T) = Me.GetEnumerable(Of T)(pEnv, pWhere)
    If vResult IsNot Nothing Then vReturn = vResult.ToList()
    Return vReturn
  End Function
  Public Function GetEnumerable(Of T As CARERecord)(pEnv As CDBEnvironment, pWhere As CDBFields) As IEnumerable(Of T)
    Dim vFactory As Func(Of T) = Function() (TryCast(Activator.CreateInstance(GetType(T), pEnv), T))
    Return Me.GetEnumerable(Of T)(pEnv, pWhere, vFactory)
  End Function

  Public Function GetEnumerable(Of T As CARERecord)(pEnv As CDBEnvironment, pWhere As CDBFields, pInstantiator As Func(Of T)) As IEnumerable(Of T)
    Return CARERecordFactory.SelectList(Of T)(pEnv, pWhere, pInstantiator)
  End Function

  Public Function GetInstance(Of T As CARERecord)(pEnv As CDBEnvironment, pWhere As CDBFields) As T
    'Refactored code to use generic method call
    Return CARERecordFactory.SelectInstance(Of T)(pEnv, pWhere)
  End Function

  Public Function GetInstanceByPrimaryKey(Of T As CARERecord)(pEnv As CDBEnvironment, pPrimaryKeyValue As String) As T
    Dim vResult As T = Nothing
    Dim vDummy As T = TryCast(Activator.CreateInstance(GetType(T), pEnv), T)
    If vDummy IsNot Nothing Then
      vResult = CARERecordFactory.SelectInstanceByPrimaryKey(Of T)(pEnv, pPrimaryKeyValue, vDummy.ClassFields)
    End If
    Return vResult
  End Function

  Public Shared Function SelectList(Of T As {IDbLoadable, IDbSelectable})(pEnv As CDBEnvironment, pWhere As CDBFields) As IEnumerable(Of T)
    Dim vInstantiator As Func(Of T) = Function() (DirectCast(Activator.CreateInstance(GetType(T), pEnv), T))
    Dim vResults As IEnumerable(Of T) = CARERecordFactory.SelectList(Of T)(pEnv, pWhere, vInstantiator)
    Return vResults
  End Function

  Public Shared Function SelectList(Of T As {IDbLoadable, IDbSelectable})(pEnv As CDBEnvironment, pWhere As CDBFields, pInstantiator As Func(Of T)) As IEnumerable(Of T)
    Dim vInitialiser As New Action(Of T, DataRow)(Sub(vInst As IDbLoadable, vDR As DataRow) vInst.LoadFromRow(vDR))
    Dim vResults As IEnumerable(Of T) = CARERecordFactory.SelectList(Of T)(pEnv, pWhere, pInstantiator, vInitialiser)
    Return vResults
  End Function

  Public Shared Function SelectInstance(Of T As {IDbLoadable, IDbSelectable})(pEnv As CDBEnvironment, pWhere As CDBFields, Optional ByVal pOrderBy As String = "") As T
    Dim vInstantiator As Func(Of T) = Function() (DirectCast(Activator.CreateInstance(GetType(T), pEnv), T))
    Return CARERecordFactory.SelectInstance(Of T)(pEnv, pWhere, vInstantiator, pOrderBy)
  End Function
  Public Shared Function SelectInstance(Of T As {IDbLoadable, IDbSelectable})(pEnv As CDBEnvironment, pWhere As CDBFields, pInstantiator As Func(Of T), Optional ByVal pOrderBy As String = "") As T
    Dim vInitialiser As New Action(Of T, DataRow)(Sub(vInst As IDbLoadable, vDR As DataRow) vInst.LoadFromRow(vDR))
    Return CARERecordFactory.SelectInstance(Of T)(pEnv, pWhere, pInstantiator, vInitialiser, pOrderBy)
  End Function


  Public Shared Function SelectList(Of T As {IDbLoadable, IDbSelectable})(pEnv As CDBEnvironment, pWhere As CDBFields, pInstantiator As Func(Of T), pInitialiser As Action(Of T, DataRow)) As IEnumerable(Of T)
    Dim vResults As New List(Of T)

    Dim p As T = pInstantiator()

    Dim vSQL As New SQLStatement(pEnv.Connection, p.DbFieldNames, p.DbAliasedTableName, pWhere)

    Dim vDT As DataTable = vSQL.GetDataTable()
    For Each vDR As DataRow In vDT.Rows
      Dim vNewItem As T = pInstantiator()
      pInitialiser(vNewItem, vDR)
      vResults.Add(vNewItem)
    Next

    Return vResults
  End Function

  Public Shared Function SelectInstance(Of T As {IDbLoadable, IDbSelectable})(pEnv As CDBEnvironment, pWhere As CDBFields, pInstantiator As Func(Of T), pInitialiser As Action(Of T, DataRow), Optional ByVal pOrderBy As String = "") As T
    Dim vResults As T

    Dim vDummy As T = pInstantiator()

    Dim vSQL As New SQLStatement(pEnv.Connection, vDummy.DbFieldNames, vDummy.DbAliasedTableName, pWhere, pOrderBy)

    vSQL.MaxRows = 1

    Dim vDT As DataTable = vSQL.GetDataTable()

    If vDT.Rows.Count > 0 Then
      vResults = GenerateInstance(Of T)(pInstantiator, pInitialiser, vDT.Rows(0))
    End If

    Return vResults
  End Function

  Private Shared Function GenerateInstance(Of T As IDbLoadable)(pInstantiator As Func(Of T), pInitialiser As Action(Of T, DataRow), pRow As DataRow) As T
    Dim vNewItem As T = pInstantiator()
    pInitialiser(vNewItem, pRow)
    Return vNewItem
  End Function

  Public Shared Function SelectInstanceByPrimaryKey(Of T As {IDbLoadable, IDbSelectable})(pEnv As CDBEnvironment, pPrimaryKeyValue As String, pClassFields As ClassFields) As T
    Dim vResult As T = Nothing
    pClassFields.GetUniquePrimaryKey().SetValue = pPrimaryKeyValue
    Dim vFields As CDBFields = pClassFields.WhereFields()
    vResult = CARERecordFactory.SelectInstance(Of T)(pEnv, vFields)
    Return vResult
  End Function

End Class

''' <summary>
''' Implement this interface if your class can be loaded from a DataRow
''' </summary>
''' <remarks>See the CARERecordFactory for more information about how this is used.
''' The name of this interface isn't great, but it's the least of the bad ones we could come up with.  Please feel free to change
''' </remarks>
Public Interface IDbLoadable
  Sub LoadFromRow(pRow As DataRow)
End Interface

''' <summary>
''' Implement this interface if you want your class to be loaded by the CARERecordFactory.
''' </summary>
''' <remarks>See the CARERecordFactory for more information about how this is used.
''' The name of this interface isn't great, but it's the least of the bad ones we could come up with.  Please feel free to change
''' </remarks>
Public Interface IDbSelectable
  ReadOnly Property DbFieldNames As String
  ReadOnly Property DbAliasedTableName As String
End Interface