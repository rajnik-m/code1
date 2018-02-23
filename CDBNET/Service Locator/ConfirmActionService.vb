Imports CDBNETCL.My.Resources

Public Class ConfirmActionService
  Implements IConfirmActionService

  Public Function ConfirmCancel() As Boolean Implements IConfirmActionService.ConfirmCancel
    Return Utilities.ConfirmCancel()
  End Function

  Public Function ConfirmDelete() As Boolean Implements IConfirmActionService.ConfirmDelete
    Return Utilities.ConfirmDelete()
  End Function

  Public Function ConfirmInsert() As Boolean Implements IConfirmActionService.ConfirmInsert
    Return Utilities.ConfirmInsert()
  End Function

  Public Function ConfirmUpdate() As Boolean Implements IConfirmActionService.ConfirmUpdate
    Utilities.ConfirmUpdate()
  End Function

  Public Function ConfirmSave() As Boolean Implements IConfirmActionService.ConfirmSave
    Return Utilities.ConfirmSave()
  End Function

End Class
