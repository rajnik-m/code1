Imports System.IO
Imports System.Linq

''' <summary>
''' Class to enable the processing of raw exam certificate data into the system
''' </summary>
Public Class CertificateDataProcessor
  Implements IDisposable

  Private dataSourceField As IDataReader = Nothing
  Private environmentField As CDBEnvironment = Nothing
  Private runTypeField As ExamUnitCertRunType = Nothing

  ''' <summary>
  ''' Initializes a new instance of the <see cref="CertificateDataProcessor"/> class.
  ''' </summary>
  ''' <param name="environemnt">The <see cref="CDBEnvironment" /> associated with this session.</param>
  ''' <param name="dataSource">The <see cref="IDataReader" /> containing the raw data.</param>
  ''' <param name="runType">The <see cref="ExamUnitCertRunType" /> being used for this certificate run.</param>
  Public Sub New(environemnt As CDBEnvironment,
                 dataSource As IDataReader,
                 runType As ExamUnitCertRunType)
    Me.Environment = environemnt
    Me.DataSource = dataSource
    Me.RunType = runType
  End Sub

  ''' <summary>
  ''' Gets the data source.
  ''' </summary>
  ''' <value>
  ''' The <see cref="IDataReader" /> containing the raw data.
  ''' </value>
  Public Property DataSource As IDataReader
    Get
      Return Me.dataSourceField
    End Get
    Private Set(value As IDataReader)
      Me.dataSourceField = value
    End Set
  End Property

  ''' <summary>
  ''' Gets the environment.
  ''' </summary>
  ''' <value>
  ''' The <see cref="CDBEnvironment" /> associated with this session.
  ''' </value>
  Public Property Environment As CDBEnvironment
    Get
      Return Me.environmentField
    End Get
    Private Set(value As CDBEnvironment)
      Me.environmentField = value
    End Set
  End Property

  ''' <summary>
  ''' Gets or sets the CSV writer.
  ''' </summary>
  ''' <value>
  ''' The <see cref="CsvWriter" /> to be used when outputting the processed data.
  ''' </value>
  Private Property CsvWriter As CsvWriter = Nothing

  ''' <summary>
  ''' Gets or sets the stream writer.
  ''' </summary>
  ''' <value>
  ''' The <see cref="StreamWriter" /> to which the processed data will be written in CSV format.
  ''' </value>
  Public Property StreamWriter As StreamWriter
    Get
      Return CsvWriter.StreamWriter
    End Get
    Set(value As StreamWriter)
      CsvWriter = New CsvWriter(value)
    End Set
  End Property

  ''' <summary>
  ''' Gets the run type.
  ''' </summary>
  ''' <value>
  ''' The <see cref="ExamUnitCertRunType" /> being used for this certificate run.
  ''' </value>
  Public Property RunType As ExamUnitCertRunType
    Get
      Return runTypeField
    End Get
    Private Set(value As ExamUnitCertRunType)
      runTypeField = value
    End Set
  End Property

  ''' <summary>
  ''' Processes the data in the data source.
  ''' </summary>
  ''' <remarks>
  ''' The data contained in the <see cref="DataSource" /> will have the certificate number, certificate number prefix and certificate number
  ''' columns overwritten with the correct, system generated values and the data will be written to the database.  In addition, if the data source
  ''' has a column named standard_document, it will be overwritten with the standard document code from the <see cref="RunType" />.   Also, if 
  ''' <see cref="StreamWriter" /> is set, the processed data will written in CSV format to the <see cref="System.IO.StreamWriter" /> object
  ''' specified.
  ''' </remarks>
  Public Function Process() As Boolean

    Dim startedTransaction As Boolean = Me.Environment.Connection.StartTransaction()
    Dim invalidCertificateCount As Integer = 0
    Dim goodCertificateCount As Integer = 0

    Try
      Using certificateData As New DataTable()
        certificateData.Load(Me.DataSource)
        Using certificateLog As New StreamWriter(Environment.GetLogFileName(String.Format("certificate_import_{0:ddMMyyyy_HHmmss}.txt", Date.Now)))
          certificateLog.WriteLine("Exam Certificate Load - {0:dd/MM/yyyy HH:mm:ss}", Date.Now)
          certificateLog.WriteLine()
          If (Not certificateData.Columns.Contains("contact_number") AndAlso
              Not certificateData.Columns.Contains("contact number")) OrElse
             Not certificateData.Columns.Contains("exam_student_unit_header_id") Then
            Throw New InvalidOperationException("Certificate data must contain columns ""contact_number"" and ""exam_student_unit_header_id"".")
          End If
          If Not certificateData.Columns.Contains("certificate_number_prefix") Then
            certificateData.Columns.Add("certificate_number_prefix")
          End If
          If Not certificateData.Columns.Contains("certificate_number") Then
            certificateData.Columns.Add("certificate_number")
          End If
          If Not certificateData.Columns.Contains("certificate_number_suffix") Then
            certificateData.Columns.Add("certificate_number_suffix")
          End If

          Dim contactNumberColumnName As String = If(certificateData.Columns.Contains("contact_number"), "contact_number", "contact number")

          Dim vRun As ExamCertRun = ExamCertRun.CreateInstance(Me.Environment, Me.RunType)
          vRun.Save()

          For Each data As DataRow In certificateData.Rows
            For Each column As DataColumn In certificateData.Columns
              If column.ColumnName.Equals("standard_document", StringComparison.InvariantCultureIgnoreCase) Then
                data(column.ColumnName) = Me.RunType.Document
              End If
            Next column

            If Not String.IsNullOrWhiteSpace(CStr(data(contactNumberColumnName))) AndAlso
               Not String.IsNullOrWhiteSpace(CStr(data("exam_student_unit_header_id"))) Then
              Dim vCertificate As ContactExamCert = ContactExamCert.CreateInstance(Me.Environment,
                                                                                     CInt(data(contactNumberColumnName)),
                                                                                     CInt(data("exam_student_unit_header_id")),
                                                                                     vRun,
                                                                                     Me.GetAttributes(data))
              vCertificate.Save()
              data("certificate_number_prefix") = vCertificate.CertificateNumberPrefix
              data("certificate_number") = vCertificate.CertificateNumber
              data("certificate_number_suffix") = vCertificate.CertificateNumberSuffix
              Me.WriteCsvData(data)
              goodCertificateCount += 1
            Else
              invalidCertificateCount += 1
              If String.IsNullOrWhiteSpace(CStr(data(contactNumberColumnName))) Then
                certificateLog.WriteLine("Certificate record {0}: Mandatory value in column {1} is blank", goodCertificateCount + invalidCertificateCount, contactNumberColumnName)
              End If
              If String.IsNullOrWhiteSpace(CStr(data("exam_student_unit_header_id"))) Then
                certificateLog.WriteLine("Certificate record {0}: Mandatory value in column exam_student_unit_header_id is blank", goodCertificateCount + invalidCertificateCount)
              End If
            End If
          Next data
          certificateLog.WriteLine("{0}{1} certificate{2} loaded successfully{3}.",
                                   If(invalidCertificateCount > 0, vbCrLf, String.Empty),
                                   goodCertificateCount,
                                   If(goodCertificateCount <> 1, "s", String.Empty),
                                   If(invalidCertificateCount > 0, String.Format(", {0} failed", invalidCertificateCount), String.Empty))
        End Using
      End Using
    Catch ex As Exception
      If startedTransaction Then
        Me.Environment.Connection.RollbackTransaction()
      End If
      Throw
    End Try
    If startedTransaction Then
      Me.Environment.Connection.CommitTransaction()
    End If
    Return invalidCertificateCount = 0
  End Function

  ''' <summary>
  ''' Gets the attributes.
  ''' </summary>
  ''' <param name="data">The <see cref="DataRow" /> to get the attributes from.</param>
  ''' <returns>An <see cref="IEnumerable" /> containing a key value pair for each of the arbitrary attributes on the passed <see cref="DataRow" />.</returns>
  Private Function GetAttributes(data As DataRow) As IEnumerable(Of KeyValuePair(Of String, String))
    Return From column As DataColumn In data.Table.Columns.OfType(Of DataColumn)()
           Where Me.IsAttributeColumn(column.ColumnName)
           Select New KeyValuePair(Of String, String)(column.ColumnName, CStr(data(column.ColumnName)))
  End Function

  ''' <summary>
  ''' Determines whether the specified column is an arbitrary attribute column.
  ''' </summary>
  ''' <param name="columnName">Name of the column.</param>
  ''' <returns><c>true</c> if the column is arbitrary data; otherwise, <c>false</c></returns>
  Private Function IsAttributeColumn(columnName As String) As Boolean
    Return Not (columnName.Equals("certificate_number_prefix", StringComparison.InvariantCultureIgnoreCase) Or
                columnName.Equals("certificate_number", StringComparison.InvariantCultureIgnoreCase) Or
                columnName.Equals("certificate_number_suffix", StringComparison.InvariantCultureIgnoreCase) Or
                columnName.Equals("standard_document", StringComparison.InvariantCultureIgnoreCase))
  End Function

  ''' <summary>
  ''' Writes data to the output <see cref="CsvWriter" /> if on has been specified.
  ''' </summary>
  ''' <param name="data">The <see cref="DataRow" /> containing the data to be written.</param>
  Private Sub WriteCsvData(data As DataRow)
    If CsvWriter IsNot Nothing Then
      CsvWriter.Write(data)
    End If
  End Sub

#Region "IDisposable Support"
  Private disposedValue As Boolean

  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not Me.disposedValue Then
      If disposing Then
        If Me.CsvWriter IsNot Nothing Then
          Me.CsvWriter.Dispose()
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

Public Class CsvWriter
  Implements IDisposable

  ''' <summary>
  ''' Gets or sets the stream writer.
  ''' </summary>
  ''' <value>
  ''' The stream writer that the CSV data is to be written to.
  ''' </value>
  Public Property StreamWriter As StreamWriter = Nothing
  ''' <summary>
  ''' Gets or sets a value indicating whether column headers have been written.
  ''' </summary>
  ''' <value>
  '''   <c>true</c> if a header record has been written; otherwise, <c>false</c>.
  ''' </value>
  Private Property AreHeadersWritten As Boolean = False

  ''' <summary>
  ''' Initializes a new instance of the <see cref="CsvWriter"/> class.
  ''' </summary>
  ''' <param name="stream">The <see cref="StreamWriter" /> to write the CSV data to.</param>
  Public Sub New(stream As StreamWriter)
    Me.StreamWriter = stream
  End Sub

  ''' <summary>
  ''' Writes the specified data to the output stream.
  ''' </summary>
  ''' <param name="data">The <see cref="DataRow" /> containing the data to be written.</param>
  Public Sub Write(data As DataRow)
    If Not Me.AreHeadersWritten Then
      Me.WriteHeaders(data.Table.Columns)
    End If
    Me.WriteData(From dataItem As Object In data.ItemArray
                 Select CStr(dataItem))
  End Sub

  ''' <summary>
  ''' Writes the headers.
  ''' </summary>
  ''' <param name="columns">The <see cref="DataColumnCollection" /> specifying the columns to write headers for.</param>
  Private Sub WriteHeaders(columns As DataColumnCollection)
    Me.WriteData(From column As DataColumn In columns.OfType(Of DataColumn)()
                 Select column.ColumnName)
    AreHeadersWritten = True
  End Sub

  ''' <summary>
  ''' Writes the data.
  ''' </summary>
  ''' <param name="data">The <see cref="DataRow" /> contain the data to be written.</param>
  Private Sub WriteData(data As IEnumerable(Of String))
    If Me.StreamWriter IsNot Nothing Then
      Dim dataRecord As New StringBuilder
      For Each dataItem As String In data
        dataRecord.Append(String.Format("""{0}"",", dataItem))
      Next dataItem
      dataRecord.Remove(dataRecord.Length - 1, 1)
      Me.StreamWriter.WriteLine(dataRecord.ToString)
    End If
  End Sub

#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not Me.disposedValue Then
      If disposing Then
        If Me.StreamWriter IsNot Nothing Then
          Me.StreamWriter = Nothing
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
