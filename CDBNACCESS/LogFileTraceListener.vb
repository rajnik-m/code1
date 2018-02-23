Imports System.Globalization
Imports System.IO

Namespace Access

  ''' <summary>
  ''' A trace listener to write to a rotatable log file.
  ''' </summary>
  Public Class LogFileTraceListener
    Inherits TraceListener

    Private mvLogFile As LogFileWriter = Nothing

    ''' <summary>
    ''' Initializes a new instance of the <see cref="LogFileTraceListener" /> class.
    ''' </summary>
    ''' <param name="pParameters">The parameters string from the XML.</param>
    ''' <remarks>This constructor is designed to be used by the app</remarks>
    Public Sub New(pParameters As String)
      Dim vParameters() As String = pParameters.Split(","c)
      mvLogFile = New LogFileWriter(vParameters(0).Trim, Integer.Parse(vParameters(1).Trim))
      mvLogFile.AutoFlush = True
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="LogFileTraceListener" /> class.
    ''' </summary>
    ''' <param name="pFilename">The filename of the log.</param>
    ''' <param name="pMaxBYtes">The maximum size of the log in bytes.</param>
    Public Sub New(pFilename As String, pMaxBYtes As Integer)
      mvLogFile = New LogFileWriter(pFilename, pMaxBYtes)
      mvLogFile.AutoFlush = True
    End Sub

    ''' <summary>
    ''' When overridden in a derived class, writes the specified message to the listener you create in the derived class.
    ''' </summary>
    ''' <param name="message">A message to write.</param>
    Public Overloads Overrides Sub Write(message As String)
      SyncLock mvLogFile
        mvLogFile.Write(message)
      End SyncLock
    End Sub

    ''' <summary>
    ''' When overridden in a derived class, writes a message to the listener you create in the derived class, followed by a line terminator.
    ''' </summary>
    ''' <param name="message">A message to write.</param>
    ''' <remarks>This implementation adds the data and time to the message before writing it.</remarks>
    Public Overloads Overrides Sub WriteLine(message As String)
      SyncLock mvLogFile
        mvLogFile.Write(String.Format(CultureInfo.InvariantCulture, "{0}: {1}{2}", Date.Now.ToString("dd-MM-yyyy HH:mm:ss.ttt", CultureInfo.InvariantCulture), message, Environment.NewLine))
      End SyncLock
    End Sub

    ''' <summary>
    ''' Gets a value indicating whether the trace listener is thread safe.
    ''' </summary>
    ''' <returns>true if the trace listener is thread safe; otherwise, false. The default is false.</returns>
    Public Overrides ReadOnly Property IsThreadSafe As Boolean
      Get
        Return True
      End Get
    End Property

    ''' <summary>
    ''' Releases the unmanaged resources used by the <see cref="T:System.Diagnostics.TraceListener" /> and optionally releases the managed resources.
    ''' </summary>
    ''' <param name="disposing">true to release both managed and unmanaged resources; false to release only unmanaged resources.</param>
    Protected Overrides Sub Dispose(disposing As Boolean)
      mvLogFile.Dispose()
      mvLogFile = Nothing
      MyBase.Dispose(disposing)
    End Sub

  End Class

End Namespace