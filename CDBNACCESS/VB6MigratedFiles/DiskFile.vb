

Namespace Access
  Public Class DiskFile

    Public Enum FileOpenModes
      fomOutput
      fomInput
      fomAppend
    End Enum

    Private mvFileOpen As Boolean
    Private mvFileHandle As Integer
    Private mvLineCount As Integer
    Private mvFileName As String
    Private mvFileMode As FileOpenModes
    Private mvLine As String

    Protected Overrides Sub Finalize()
      CloseFile()
      MyBase.Finalize()
    End Sub

    Public Sub OpenFile(ByRef pFileName As String, ByRef pMode As FileOpenModes)
      mvFileHandle = FreeFile()
      Select Case pMode
        Case FileOpenModes.fomOutput
          FileOpen(mvFileHandle, pFileName, OpenMode.Output)
        Case FileOpenModes.fomInput
          FileOpen(mvFileHandle, pFileName, OpenMode.Input)
        Case FileOpenModes.fomAppend
          FileOpen(mvFileHandle, pFileName, OpenMode.Append)
        Case Else
          RaiseError(DataAccessErrors.daeActionFailed)
      End Select
      mvFileOpen = True
      mvFileMode = pMode
      mvLineCount = 0
      mvFileName = pFileName
    End Sub

    Public Sub CloseFile()
      If mvFileOpen Then FileClose(mvFileHandle)
      mvFileOpen = False
      mvLineCount = 0
    End Sub

    Public Sub ReadToEndOfFile()
      Dim vLine As String
      mvLine = ""
      While Not EOF(mvFileHandle)
        vLine = LineInput(mvFileHandle)
        If Len(mvLine) > 0 Then mvLine = mvLine & vbCrLf
        mvLine = mvLine & vLine
      End While
    End Sub

    Public Sub ReadLine()
      mvLine = LineInput(mvFileHandle)
    End Sub

    Public ReadOnly Property EndOfFile() As Boolean
      Get
        EndOfFile = EOF(mvFileHandle)
      End Get
    End Property

    Public ReadOnly Property CurrentLine() As String
      Get
        CurrentLine = mvLine
      End Get
    End Property

    Public ReadOnly Property CurrentRow() As Integer
      Get
        CurrentRow = mvLineCount
      End Get
    End Property

    Public ReadOnly Property IsOpen() As Boolean
      Get
        IsOpen = mvFileOpen
      End Get
    End Property

    'Public Sub ReOpenFile()
    '  CloseFile
    '  OpenFile mvFileName, mvFileMode
    'End Sub

    'UPGRADE_NOTE: PrintLine was upgraded to PrintLine_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public Sub PrintLine_Renamed(Optional ByRef pString As String = "")
      PrintLine(mvFileHandle, pString)
      mvLineCount = mvLineCount + 1
    End Sub

    Public Sub PrintString(Optional ByRef pString As String = "")
      Print(mvFileHandle, pString)
    End Sub

    Public Sub PrintMailMergeItem(ByRef pString As String, Optional ByRef pMailMergeType As MailMergeInformation.MailMergeTypes = MailMergeInformation.MailMergeTypes.mmtWord, Optional ByRef pFirst As Boolean = False)
      If pMailMergeType = MailMergeInformation.MailMergeTypes.mmtWordPerfect Then
        PrintString(pString & Chr(18) & Chr(10))
      Else
        If pFirst Then
          PrintString(Chr(34) & pString & Chr(34))
        Else
          PrintString("," & Chr(34) & pString & Chr(34))
        End If
      End If
    End Sub

    Public Sub PrintNewLine()
      PrintLine(mvFileHandle, "")
      mvLineCount = mvLineCount + 1
    End Sub
  End Class
End Namespace
