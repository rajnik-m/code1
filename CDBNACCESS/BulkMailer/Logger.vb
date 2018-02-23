Imports System.IO

Public Class Logger
  Implements IDisposable

  Private mvLogFile As StreamWriter = Nothing
  Private mvEnvironment As CDBEnvironment = Nothing
  Private mvLogName As String = String.Empty

  Public Sub New(pEnvironment As CDBEnvironment, pLogName As String)
    mvEnvironment = pEnvironment
    mvLogName = pLogName
  End Sub

  Public Sub LogMessage(pMessage As String)
    LogFile.WriteLine("{0}: - {1}", TodaysDateAndTime, pMessage)
  End Sub

  Public Sub LogMessage(pMessage As String, pArguments() As String)
    LogFile.WriteLine("{0}: - {1}", TodaysDateAndTime, String.Format(pMessage, pArguments))
  End Sub

  Public Sub WriteBlankLine(pLines As Integer)
    For vLine As Integer = 1 To pLines
      LogFile.WriteLine()
    Next
  End Sub

  Private ReadOnly Property LogFile As StreamWriter
    Get
      If mvLogFile Is Nothing Then
        mvLogFile = New StreamWriter(mvEnvironment.GetLogFileName(mvLogName & "_" & Date.Now.ToString("yyyyMMdd_HHmmss") & ".log"))
      End If
      Return mvLogFile
    End Get
  End Property

#Region "IDisposable Support"
  Private disposedValue As Boolean

  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not Me.disposedValue Then
      If disposing Then
        If mvLogFile IsNot Nothing Then
          mvLogFile.Close()
          mvLogFile.Dispose()
        End If
      End If
    End If
    Me.disposedValue = True
  End Sub

  Public Sub Dispose() Implements IDisposable.Dispose
    Dispose(True)
    GC.SuppressFinalize(Me)
  End Sub
#End Region

End Class
