Imports System.Data.DataTableExtensions

Public Class DocumentTransporter

  Private mvEnv As CDBEnvironment = Nothing
  Private mvDestination As String = String.Empty
  Private mvDocuments As IList(Of Integer) = Nothing
  Private mvLogFile As LogFile = Nothing

  Public Sub New(pEnv As CDBEnvironment, pDestination As String, pDocuments As IList(Of Integer))
    If pEnv Is Nothing Then
      RaiseError(DataAccessErrors.daeEnvNotInitialised)
    End If
    mvEnv = pEnv
    If pDestination Is Nothing Then
      RaiseError(DataAccessErrors.daeInvalidParameter)
    End If
    mvDestination = pDestination
    If pDocuments Is Nothing OrElse pDocuments.Count <= 0 Then
      RaiseError(DataAccessErrors.daeInvalidParameter)
    End If
    mvDocuments = pDocuments
    mvLogFile = New LogFile(mvEnv.GetConfig("default_logfile_directory", "c:\contacts\logfiles"), TypeName(Me), False, True, True)
  End Sub

  Public Function MoveDocuments() As String
    Dim vSucceeded As Integer = 0
    Dim vFailed As Integer = 0
    mvLogFile.WriteStandardHeaderOrFooter(True)
    mvLogFile.WriteLine(String.Format("Moving external documents to {0}.", mvDestination))
    For Each vDocument As Integer In mvDocuments
      Try
        Dim vCommsLog As New CommunicationsLog(mvEnv)
        vCommsLog.InitWithPrimaryKey(New CDBFields({New CDBField("communications_log_number", vDocument, CDBField.FieldWhereOperators.fwoEqual)}))
        If vCommsLog.IsHeldExternally Then
          mvLogFile.WriteLine(String.Format("Moving document {0} from {1}...", vCommsLog.CommunicationsLogNumber, vCommsLog.ExternalDocumentName))
          vCommsLog.RelocateExternalDocument(mvDestination)
          vCommsLog.Save()
          vSucceeded += 1
          mvLogFile.WriteLine("...Complete.")
          mvLogFile.WriteBlankLine()
        Else
          vFailed += 1
          mvLogFile.WriteLine(String.Format("Document {0} does not appear to be stored externally and so will not be moved.", vCommsLog.CommunicationsLogNumber))
          mvLogFile.WriteBlankLine()
        End If
      Catch vEx As Exception
        vFailed += 1
        mvLogFile.WriteLine(String.Format("...Failed.  Error was ""{0}"".", vEx.Message))
        mvLogFile.WriteBlankLine()
      End Try
    Next vDocument
    Dim vMessage As String = String.Format("Attempted to move {0:#,##0} documents; {1:#,##0} succeeded, {2:#,##0} failed.", mvDocuments.Count, vSucceeded, vFailed)
    mvLogFile.WriteLine(vMessage)
    mvLogFile.WriteBlankLine()
    mvLogFile.WriteStandardHeaderOrFooter(False)
    Return vMessage
  End Function

End Class
