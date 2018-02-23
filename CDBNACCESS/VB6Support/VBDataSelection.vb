Namespace Access

  Partial Class DataSelection
    Public Sub New()

    End Sub

    Public Sub Init(ByVal pEnv As CDBEnvironment, ByVal pDataSelectionType As DataSelection.DataSelectionTypes, ByVal pParams As CDBParameters)
      Init(pEnv, pDataSelectionType, pParams, DataSelection.DataSelectionListType.dsltUser, DataSelection.DataSelectionUsages.dsuCare, "")
    End Sub

    Public Sub Init(ByVal pEnv As CDBEnvironment, ByVal pDataSelectionType As DataSelection.DataSelectionTypes, ByVal pParams As CDBParameters, ByVal pListType As DataSelection.DataSelectionListType, ByVal pUsage As DataSelection.DataSelectionUsages)
      Init(pEnv, pDataSelectionType, pParams, pListType, pUsage, "")
    End Sub

  End Class


  Public Class VBDataSelection
    Private mvDS As DataSelection

    Public Sub New()
    End Sub

    Public Sub Init(ByVal pEnv As CDBEnvironment, ByVal pDataSelectionType As DataSelection.DataSelectionTypes)
      mvDS = New DataSelection(pEnv, pDataSelectionType, DataSelection.DataSelectionListType.dsltUser)
    End Sub

    Public Sub Init(ByVal pEnv As CDBEnvironment, ByVal pDataSelectionType As DataSelection.DataSelectionTypes, ByVal pParams As CDBParameters)
      mvDS = New DataSelection(pEnv, pDataSelectionType, pParams, DataSelection.DataSelectionListType.dsltUser, DataSelection.DataSelectionUsages.dsuCare)
    End Sub

    Public Sub Init(ByVal pEnv As CDBEnvironment, ByVal pDataSelectionType As DataSelection.DataSelectionTypes, ByVal pParams As CDBParameters, ByVal pListType As DataSelection.DataSelectionListType, ByVal pUsage As DataSelection.DataSelectionUsages)
      mvDS = New DataSelection(pEnv, pDataSelectionType, pParams, pListType, pUsage)
    End Sub

    Public Function DataTable() As CDBDataTable
      Return mvDS.DataTable
    End Function

    Public Sub AddParameter(ByVal pName As String, ByVal pFieldType As CDBField.FieldTypes, ByVal pValue As String)
      mvDS.AddParameter(pName, pFieldType, pValue)
    End Sub

  End Class

End Namespace
