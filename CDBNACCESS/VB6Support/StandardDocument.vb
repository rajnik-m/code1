Namespace Access

  Partial Public Class StandardDocument

    Public Function ProcessBulkEmail(ByVal pMergeFileName As String, ByVal pFromAddress As String, ByVal pFromName As String, ByVal pMailing As String, ByVal pNoMailingHistory As Boolean, ByVal pErrorCount As Integer) As Integer
      'TODO VB6 CONVERSION ProcessBulkEmail
      '  Dim vBaseDocumentName As String
      '  Dim vBaseDocument As String
      '  Dim vExternalApp As ExternalApplication
      '  Dim vFile As New CSVFile
      '  Dim vMergeFields() As String
      '  Dim vItems() As String
      '  Dim vNewDocument As String
      '  Dim vIndex As Short
      '  Dim vValue As String
      '  Dim vEMailAddress As String
      '  Dim vDocumentFile As New DiskFile
      '  Dim vCount As Integer
      '  Dim vSMTPInterface As New SMTPInterface
      '  Dim vInitError As Integer
      '  Dim vContactNumber As Integer
      '  Dim vErrorNumber As Integer
      '  Dim vUpdateFields As CDBFields
      '  Dim vWhereFields As CDBFields

      '  mvEnv.AllExternalApplications = True
      '  If mvEnv.ExternalApplications.Exists(ExternalApplicationCode) Then
      '    vExternalApp = mvEnv.ExternalApplications(ExternalApplicationCode)
      '  Else
      '    RaiseError(DataAccessErrors.daeExternalApplicationNotFound, ExternalApplicationCode)
      '  End If
      '  vInitError = vSMTPInterface.Init(mvEnv, pFromAddress, pFromName)
      '  If vInitError = 0 Then 'Make sure the server is initialised
      '    vBaseDocumentName = mvEnv.GetDocument(CDBEnvironment.GetDocumentLocations.gdlStandardDocument, StandardDocumentCode, vExternalApp.ExternalStorage, vExternalApp.Extension)
      '    vDocumentFile.OpenFile(vBaseDocumentName, DiskFile.FileOpenModes.fomInput)
      '    vDocumentFile.ReadToEndOfFile()
      '    vBaseDocument = vDocumentFile.CurrentLine
      '    vDocumentFile.CloseFile()
      '    gvSystem.KillFile(vBaseDocumentName)

      '    vFile.OpenFile(pMergeFileName)
      '    vFile.ReadLine()
      '    ReDim vMergeFields(vFile.NumberOfFields)
      '    For vIndex = 1 To vFile.NumberOfFields
      '      vMergeFields(vIndex) = vFile.Item(vIndex)
      '    Next
      '    Do
      '      vFile.ReadLine()
      '      If vFile.EndOfFile = False Then
      '        vNewDocument = vBaseDocument
      '        For vIndex = 1 To vFile.NumberOfFields
      '          vNewDocument = Replace(vNewDocument, "&lt;&lt;" & vMergeFields(vIndex) & "&gt;&gt;", vFile.Item(vIndex))
      '          vNewDocument = Replace(vNewDocument, "<<" & vMergeFields(vIndex) & ">>", vFile.Item(vIndex))
      '          If vMergeFields(vIndex) = "EMail" Then vEMailAddress = vFile.Item(vIndex)
      '          If vMergeFields(vIndex) = "Contact Number" Then vContactNumber = CInt(vFile.Item(vIndex))
      '        Next
      '        If vContactNumber > 0 Or pNoMailingHistory = True Then
      '          vErrorNumber = vSMTPInterface.SendMail(Subject, vNewDocument, vEMailAddress)
      '        Else
      '          RaiseError(DataAccessErrors.daeAttributesNotDefined, "Contact Number") 'Must have contact number if doing Mailing History
      '        End If
      '        If vErrorNumber > 0 Then
      '          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMailingError) = True And pNoMailingHistory = False Then
      '            vWhereFields = New CDBFields
      '            vUpdateFields = New CDBFields
      '            vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, vContactNumber)
      '            vWhereFields.Add("mailing_number", CDBField.FieldTypes.cftLong, "(SELECT MAX(mailing_number) FROM mailing_history WHERE mailing='" & pMailing & "')")
      '            vUpdateFields.Add("error_number", CDBField.FieldTypes.cftLong, vErrorNumber)
      '            mvEnv.Connection.UpdateRecords("contact_emailings", vUpdateFields, vWhereFields, False)
      '          End If
      '          pErrorCount = pErrorCount + 1
      '        Else
      '          vCount = vCount + 1
      '        End If
      '      End If
      '    Loop While vFile.EndOfFile = False
      '    ProcessBulkEmail = vCount
      '    vFile.CloseFile()
      '  Else
      '    RaiseError(DataAccessErrors.daeFailedToInitSMTP, VB6.Format(vInitError))
      '  End If
    End Function
  End Class

End Namespace
