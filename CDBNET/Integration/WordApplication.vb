Option Strict Off
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices

Public Class WordApplication
  Inherits ExternalApplication

  'Private mvWord As Word.Application
  'Private mvWordEvents As Word.ApplicationEvents2_Event
  Private mvWord As Object
  Private mvMergeFileName As String
  Private mvWordActive As Boolean

  Public Overrides Sub ProcessAppActive()
    If mvWordActive Then
      If WordActive() Then
        Dim vFileInfo As New FileInfo(mvFileName)
        Select Case mvAction
          Case DocumentActions.daCreating, DocumentActions.daEditing, DocumentActions.daPrinting, DocumentActions.daViewing
            For Each vDoc As Object In mvWord.Documents     'Word.Document
              If vDoc.Name = vFileInfo.Name Then
                Select Case mvAction
                  Case DocumentActions.daCreating, DocumentActions.daEditing
                    vDoc.Close(-1)                              'WdSaveOptions.wdSaveChanges
                  Case DocumentActions.daViewing, DocumentActions.daPrinting
                    vDoc.Close(0)                               'WdSaveOptions.wdDoNotSaveChanges
                End Select
                Exit For
              End If
            Next
        End Select
        mvWordActive = False
        mvWord.Quit()
      End If
      mvWord = Nothing
      MyBase.ProcessActionComplete()
    End If
  End Sub

  Protected Overrides Sub DoEditDocument()
    InitWord(DocumentActions.daEditing)
    mvWord.Documents.Open(CType(mvFileName, Object))
    mvWord.Visible = True
    mvWord.Activate()
  End Sub

  Protected Overrides Sub DoViewDocument()
    InitWord(DocumentActions.daViewing)
    mvWord.Documents.Open(CType(mvFileName, Object), ConfirmConversions:=False, ReadOnly:=True)
    mvWord.Visible = True
    mvWord.Activate()
  End Sub

  Protected Overrides Sub DoPrintDocument()
    InitWord(DocumentActions.daPrinting)
    mvWord.Visible = True
    Dim vDocument As Object = mvWord.Documents.Open(CType(mvFileName, Object), ConfirmConversions:=False, ReadOnly:=True)   'Word.Document
    vDocument.PrintOut(Background:=False)
  End Sub

  Protected Overrides Sub DoEditNewDocument(ByVal pList As ParameterList)
    InitWord(DocumentActions.daCreating)
    Dim vDocument As Object = mvWord.Documents.Add()       'Word.Document
    If Not pList Is Nothing Then
      InsertLine(vDocument, pList("Addressee"), 1)
      InsertLine(vDocument, pList("AddresseeAddress").Trim(New Char() {ChrW(13), ChrW(10)}), 2)
      InsertLine(vDocument, pList("Dated"), 3)
      InsertLine(vDocument, ControlText.docOurReference & " ", 0)
      InsertLine(vDocument, pList("OurReference"), 2)
      InsertLine(vDocument, ControlText.docYourReference & " ", 0)
      InsertLine(vDocument, pList("TheirReference"), 3)
      InsertLine(vDocument, pList("Salutation"), 8)
      InsertLine(vDocument, ControlText.docYoursSincerely, 4)
      InsertLine(vDocument, pList("SignatureName"), 1)
      InsertLine(vDocument, pList("SignaturePosition"), 1)
    End If
    vDocument.SaveAs(CType(mvFileName, Object))
    mvWord.Visible = True
    mvWord.Activate()
  End Sub

  Public Overrides Sub EditNewStandardDocument(ByVal pList As ParameterList, ByVal pExtension As String, ByVal pMailMerge As Boolean)
    Try
      Dim vStandardDocument As String = pList("StandardDocument")
      mvFileName = DataHelper.GetStandardDocumentFile(vStandardDocument, pExtension)
      Dim vRelatedContactNumber As Integer = 0
      If pList.ContainsKey("RelatedContactNumber") Then vRelatedContactNumber = pList.IntegerValue("RelatedContactNumber")
      If pMailMerge Then mvMergeFileName = DataHelper.GetDocumentMergeData(vStandardDocument, pList.IntegerValue("AddresseeContactNumber"), pList.IntegerValue("AddresseeAddressNumber"), _
                                           pList("OurReference"), pList("TheirReference"), pList("Dated"), pList.IntegerValue("SenderContactNumber"), vRelatedContactNumber)
      InitWord(DocumentActions.daCreating)
      Dim vDocument As Object = mvWord.Documents.Open(CType(mvFileName, Object))   'Word.Document
      If pMailMerge Then
        With vDocument.MailMerge
          .MainDocumentType = 0                         'Word.WdMailMergeMainDocType.wdFormLetters = 0
          .OpenDataSource(mvMergeFileName, 4, 0)        'Word.WdOpenFormat.wdOpenFormatText = 4
          '.OpenDataSource(Name:="tmp237.DOC", ConfirmConversions:=False, ReadOnly:=False, LinkToSource:=True, _
          'AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:="", _
          'WritePasswordDocument:="", WritePasswordTemplate:="", Revert:=False, _
          'Format:=wdOpenFormatAuto, Connection:="", SQLStatement:="", SQLStatement1:="", SubType:=wdMergeSubTypeOther)
          .Destination = 0                              'Word.WdMailMergeDestination.wdSendToNewDocument = 0
          .SuppressBlankLines = True
          .Execute(Pause:=False)
        End With
        vDocument.Close(False)
        mvWord.ActiveDocument.SaveAs(CType(mvFileName, Object))
        Dim vFileInfo As New FileInfo(mvMergeFileName)
        vFileInfo.Delete()
      End If
      mvWord.Visible = True
      mvWord.Activate()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub InitWord(ByVal pAction As DocumentActions)
    If mvWord Is Nothing Then
      'mvWord = New Word.Application     'Create a Quit event handler.
      'mvWordEvents = CType(mvWord, Word.ApplicationEvents2_Event)
      'AddHandler mvWordEvents.Quit, AddressOf QuitHandler
      mvWord = CreateObject("Word.Application")
      If mvWord Is Nothing Then Throw New CareException(CareException.ErrorNumbers.enCannotRunWord)
      Dim vType As System.Type = mvWord.GetType
      'Dim vEventInfo As EventInfo

      'Dim vEvents As EventInfo() = vType.GetEvents()
      'For Each vEventInfo In vEvents
      '  ShowInformationMessage(String.Format("Found Word Event: {0}", vEventInfo.Name))
      'Next

      'vEventInfo = vType.GetEvent("ApplicationEvents4_Event_Quit")
      'If vEventInfo Is Nothing Then vEventInfo = vType.GetEvent("ApplicationEvents3_Event_Quit")
      'If vEventInfo Is Nothing Then vEventInfo = vType.GetEvent("ApplicationEvents2_Event_Quit")
      'If vEventInfo Is Nothing Then
      '  ShowInformationMessage(InformationMessages.imCannotAccessWordQuitEvent)
      'Else
      '  vEventInfo.AddEventHandler(mvWord, [Delegate].CreateDelegate(vEventInfo.EventHandlerType, Me, "QuitHandler"))
      'End If

    End If
      mvAction = pAction
      mvWordActive = True
  End Sub

  Private Sub InsertLine(ByVal pDocument As Object, ByVal pString As String, ByVal pCount As Integer)
    With pDocument.Content
      .InsertAfter(pString)
      For vIndex As Integer = 1 To pCount
        .InsertParagraphAfter()
      Next
    End With
  End Sub

  'Private Sub QuitHandler()
  '  If mvWordActive Then
  '    mvWordActive = False
  '    mvWord = Nothing
  '    MyBase.ProcessActionComplete()
  '  End If
  'End Sub

  Private Function WordActive() As Boolean
    Try
      If Not mvWord Is Nothing AndAlso mvWord.Visible Then WordActive = True
    Catch vException As System.Runtime.InteropServices.COMException
      mvWordActive = False
      Debug.Print(vException.ToString)
    End Try
  End Function

  Public Overrides Sub MergeStandardDocument(ByVal pStandardDocument As String, ByVal pExtension As String, ByVal pMergeFileName As String, Optional ByVal pInstantPrint As Boolean = False)
    Try
      mvFileName = DataHelper.GetStandardDocumentFile(pStandardDocument, pExtension)
      InitWord(DocumentActions.daCreating)
      Dim vDocument As Object = mvWord.Documents.Open(CType(mvFileName, Object))   'Word.Document
      With vDocument.MailMerge
        .MainDocumentType = 0                         'Word.WdMailMergeMainDocType.wdFormLetters = 0
        .OpenDataSource(pMergeFileName, 4, 0)        'Word.WdOpenFormat.wdOpenFormatText = 4
        '.OpenDataSource(Name:="tmp237.DOC", ConfirmConversions:=False, ReadOnly:=False, LinkToSource:=True, _
        'AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:="", _
        'WritePasswordDocument:="", WritePasswordTemplate:="", Revert:=False, _
        'Format:=wdOpenFormatAuto, Connection:="", SQLStatement:="", SQLStatement1:="", SubType:=wdMergeSubTypeOther)
        .Destination = 0                              'Word.WdMailMergeDestination.wdSendToNewDocument = 0
        .SuppressBlankLines = True
        .Execute(Pause:=False)
      End With
      vDocument.Close(False)
      If pInstantPrint Then
        mvWord.ActiveDocument.PrintOut(Background:=False)
      Else
        mvWord.Visible = True
        mvWord.Activate()
      End If
      Dim vFileInfo As New FileInfo(pMergeFileName)
      vFileInfo.Delete()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
End Class
