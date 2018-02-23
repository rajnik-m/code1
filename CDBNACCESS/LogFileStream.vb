Imports System.IO
Imports System.Text
Imports System.Globalization

Namespace Access

  ''' <summary>
  ''' A <see cref="System.IO.StreamWriter" /> for writing to a log file.
  ''' </summary>
  ''' <remarks>The log file is created with a base log file name and a maximum size.  When data is written to the stream,
  ''' the size of the file is checked to ensure that the write will not take it over the maximum size.  If the write 
  ''' would take it over the maximum size, the current log file is backed up to a file in the same folder, with the 
  ''' same extension, named as the original log file is with the currernt date and time appended.  The original log file
  ''' is then emptied and the requested write is performed.</remarks>
  Public Class LogFileWriter
    Inherits StreamWriter

    Private mvMaxBytes As Integer
    Private mvFilename As String

    ''' <summary>
    ''' Initializes a new instance of the <see cref="LogFileWriter" /> class.
    ''' </summary>
    ''' <param name="path">The path of the active log file.</param>
    ''' <param name="maxBytes">The maximum size of the log file in bytes.</param>
    Public Sub New(path As String, maxBytes As Integer)
      MyBase.New(path, True, New UTF8Encoding)
      mvFilename = System.IO.Path.GetFullPath(path)
      mvMaxBytes = maxBytes
    End Sub

    ''' <summary>
    '''   Writes a subarray of characters to the stream.
    ''' </summary>
    ''' <param name="buffer">A character array containing the data to write.</param>
    ''' <param name="index">The index into <paramref name="buffer" /> at which to begin writing.</param>
    ''' <param name="count">The number of characters to read from <paramref name="buffer" />.</param>
    ''' <exception cref="T:System.ArgumentNullException">
    '''   <paramref name="buffer" /> is null. 
    ''' </exception>
    ''' <exception cref="T:System.ArgumentException">
    '''   The buffer length minus <paramref name="index" /> is less than <paramref name="count" />. 
    ''' </exception>
    ''' <exception cref="T:System.ArgumentOutOfRangeException">
    '''   <paramref name="index" /> or <paramref name="count" /> is negative or <paramref name="count" /> is greater
    '''   than the maximum file size.
    '''  </exception>
    '''  <exception cref="T:System.IO.IOException">
    '''   An I/O error occurs. 
    ''' </exception>
    ''' <exception cref="T:System.ObjectDisposedException">
    '''   <see cref="P:System.IO.StreamWriter.AutoFlush" /> is true or the <see cref="T:System.IO.StreamWriter" /> 
    '''   buffer is full, and current writer is closed. 
    ''' </exception>
    ''' <exception cref="T:System.NotSupportedException">
    '''   <see cref="P:System.IO.StreamWriter.AutoFlush" /> is true or the <see cref="T:System.IO.StreamWriter" /> 
    '''   buffer is full, and the contents of the buffer cannot be written to the underlying fixed size stream because 
    '''   the <see cref="T:System.IO.StreamWriter" /> is at the end the stream.
    ''' </exception>
    Public Overrides Sub Write(buffer() As Char, index As Integer, count As Integer)
      If buffer IsNot Nothing Then
        If count > mvMaxBytes Then
          Throw New ArgumentOutOfRangeException("count")
        Else
          If BaseStream.Position + System.Math.Min(count, buffer.GetLength(0)) > mvMaxBytes Then
            Flush()
            File.Copy(mvFilename, Path.GetDirectoryName(mvFilename) & Path.PathSeparator & Path.GetFileNameWithoutExtension(mvFilename) & Date.Now.ToString("_yyyyMMdd_HHmmss.", CultureInfo.InvariantCulture) & Path.GetExtension(mvFilename))
            BaseStream.SetLength(0)
          End If
          MyBase.Write(buffer, index, count)
        End If
      Else
        Throw New ArgumentNullException("buffer")
      End If
    End Sub

  End Class

End Namespace