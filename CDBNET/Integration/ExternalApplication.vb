Imports System.IO

Public MustInherit Class ExternalApplication

  Public Enum DocumentActions
    daUnknown
    daCreating
    daEditing
    daPrinting
    daViewing
  End Enum

  Protected mvAction As DocumentActions
  Protected mvFileName As String
  Protected mvDocumentNumber As Integer
  Protected mvExtension As String

  Public Event ActionComplete(ByVal pAction As DocumentActions, ByVal pFilename As String)

  Public MustOverride Sub ProcessAppActive()
  Protected MustOverride Sub DoEditDocument()
  Protected MustOverride Sub DoViewDocument()
  Protected MustOverride Sub DoPrintDocument()
  Protected MustOverride Sub DoEditNewDocument(ByVal pList As ParameterList)

  Public MustOverride Sub EditNewStandardDocument(ByVal pList As ParameterList, ByVal pExtension As String, ByVal pMailMerge As Boolean)
  Public MustOverride Sub MergeStandardDocument(ByVal pStandardDocument As String, ByVal pExtension As String, ByVal pMergeFileName As String, Optional ByVal pInstantPrint As Boolean = False)
  Public Sub EditDocument(ByVal pDocumentNumber As Integer, ByVal pExtension As String)
    Try
      GetDocumentFile(pDocumentNumber, pExtension)
      mvAction = DocumentActions.daEditing
      DoEditDocument()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Public Sub ViewDocument(ByVal pFileInfo As FileInfo)
    Try
      mvDocumentNumber = 0
      pFileInfo.Attributes = FileAttributes.ReadOnly
      mvFileName = pFileInfo.FullName
      mvExtension = pFileInfo.Extension
      mvAction = DocumentActions.daViewing
      DoViewDocument()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Public Sub ViewDocument(ByVal pDocumentNumber As Integer, ByVal pExtension As String)
    Try
      GetDocumentFile(pDocumentNumber, pExtension, True)
      mvAction = DocumentActions.daViewing
      DoViewDocument()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Public Sub PrintDocument(ByVal pDocumentNumber As Integer, ByVal pExtension As String)
    Try
      GetDocumentFile(pDocumentNumber, pExtension, True)
      mvAction = DocumentActions.daPrinting
      DoPrintDocument()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Public Sub EditNewDocument(ByVal pList As ParameterList, ByVal pExtension As String)
    Try
      mvFileName = DataHelper.GetTempFile(pExtension)
      mvAction = DocumentActions.daCreating
      DoEditNewDocument(pList)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Protected Sub ProcessActionComplete()
    Select Case mvAction
      Case DocumentActions.daCreating
        'Don't know document number yet?
      Case DocumentActions.daEditing
        DataHelper.AddDocumentHistory(CareServices.XMLDocumentHistoryActions.xdhaEdited, mvDocumentNumber)
      Case DocumentActions.daPrinting
        DataHelper.AddDocumentHistory(CareServices.XMLDocumentHistoryActions.xdhaPrinted, mvDocumentNumber)
        Dim vFileInfo As New FileInfo(mvFileName)
        vFileInfo.Attributes = FileAttributes.Normal
        vFileInfo.Delete()
      Case DocumentActions.daViewing
        If mvDocumentNumber > 0 Then
          DataHelper.AddDocumentHistory(CareServices.XMLDocumentHistoryActions.xdhaViewed, mvDocumentNumber)
          Dim vFileInfo As New FileInfo(mvFileName)
          vFileInfo.Attributes = FileAttributes.Normal
          vFileInfo.Delete()        'Can get IO exception here if file still in use by WORD
        End If
    End Select
    RaiseEvent ActionComplete(mvAction, mvFileName)
  End Sub
  Protected Sub GetDocumentFile(ByVal pDocumentNumber As Integer, ByVal pExtension As String, Optional ByVal pReadOnly As Boolean = False)
    mvDocumentNumber = pDocumentNumber
    mvExtension = pExtension
    mvFileName = DataHelper.GetDocumentFile(mvDocumentNumber, mvExtension)
    If pReadOnly Then
      Dim vFileInfo As New FileInfo(mvFileName)
      vFileInfo.Attributes = FileAttributes.ReadOnly
    End If
  End Sub

End Class
