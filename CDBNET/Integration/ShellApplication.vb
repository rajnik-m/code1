Imports System.IO

Public Class ShellApplication
  Inherits ExternalApplication

  Private WithEvents mvProcess As Process
  Private mvStartInfo As ProcessStartInfo
  Private mvProcessActive As Boolean
  Private mvActivatingApp As Boolean

  Public Overrides Sub ProcessAppActive()
    If mvProcessActive And Not mvActivatingApp Then
      mvActivatingApp = True
      If Not mvProcess.HasExited Then mvProcess.WaitForExit(1000)
      mvProcess.Refresh()
      If Not mvProcess.HasExited Then
        'We have been reactivated but the process we started is still running
        ShowWarningMessage(InformationMessages.imExtAppActive)
      End If
      If Not mvProcess.HasExited Then AppActivate(mvProcess.Id)
      mvActivatingApp = False
    End If
  End Sub

  Protected Overrides Sub DoEditDocument()
    mvProcess = Process.Start(mvFileName)
    If Not mvProcess Is Nothing Then
      mvProcessActive = True
      mvProcess.EnableRaisingEvents = True
    End If
  End Sub

  Protected Overrides Sub DoViewDocument()
    mvProcess = Process.Start(mvFileName)
    If Not mvProcess Is Nothing Then
      mvProcessActive = True
      mvProcess.EnableRaisingEvents = True
    End If
  End Sub

  Protected Overrides Sub DoPrintDocument()
    Dim vCanPrint As Boolean

    mvStartInfo = New ProcessStartInfo(mvFilename)
    Dim vVerb As String
    For Each vVerb In mvStartInfo.Verbs
      If String.Compare(vVerb, "print", True) = 0 Then
        mvStartInfo.Verb = vVerb
        mvProcessActive = True
        mvProcess = Process.Start(mvStartInfo)
        If Not mvProcess Is Nothing Then
          mvProcess.EnableRaisingEvents = True
          mvProcess.WaitForInputIdle(1000)
        End If
        vCanPrint = True
        Exit For
        'ElseIf String.Compare(vVerb, "printto", True) = 0 Then
        '  mvPrintVerb = vVerb
        '  Exit For
      End If
    Next
    If Not vCanPrint Then ShowWarningMessage(InformationMessages.imCannotPrintDocument, mvExtension)
  End Sub

  Protected Overrides Sub DoEditNewDocument(ByVal pList As ParameterList)
    If Not pList Is Nothing Then
      Dim vFileInfo As New FileInfo(mvFileName)
      Dim vSW As StreamWriter = vFileInfo.CreateText
      InsertLine(vSW, pList("Addressee"), 1)
      InsertLine(vSW, pList("AddresseeAddress").Trim(New Char() {ChrW(13), ChrW(10)}), 2)
      InsertLine(vSW, pList("Dated"), 3)
      InsertLine(vSW, ControlText.docOurReference & " ", 0)
      InsertLine(vSW, pList("OurReference"), 2)
      InsertLine(vSW, ControlText.docYourReference & " ", 0)
      InsertLine(vSW, pList("TheirReference"), 3)
      InsertLine(vSW, pList("Salutation"), 8)
      InsertLine(vSW, ControlText.docYoursSincerely, 4)
      InsertLine(vSW, pList("SignatureName"), 1)
      InsertLine(vSW, pList("SignaturePosition"), 1)
      vSW.Close()
    End If
    DoEditDocument()
  End Sub

  Private Sub InsertLine(ByVal pSW As StreamWriter, ByVal pString As String, ByVal pCount As Integer)
    With pSW
      .Write(pString)
      For vIndex As Integer = 1 To pCount
        .WriteLine()
      Next
    End With
  End Sub

  Private Sub mvProcess_Exited(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvProcess.Exited
    If mvProcessActive Then
      mvProcessActive = False
      'Debug.WriteLine("Process Complete at " & Date.Now.ToLongTimeString)
      If MDIForm.InvokeRequired Then
        MDIForm.Invoke(New MethodInvoker(AddressOf MyBase.ProcessActionComplete))
      Else
        MyBase.ProcessActionComplete()
      End If
    End If
  End Sub

  Public Overrides Sub EditNewStandardDocument(ByVal pList As ParameterList, ByVal pExtension As String, ByVal pMailMerge As Boolean)
    Try
      Dim vStandardDocument As String = pList("StandardDocument")
      mvFileName = DataHelper.GetStandardDocumentFile(vStandardDocument, pExtension)
      DoEditDocument()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Public Overrides Sub MergeStandardDocument(ByVal pStandardDocument As String, ByVal pExtension As String, ByVal pMergeFileName As String, Optional ByVal pInstantPrint As Boolean = False)

  End Sub
End Class
