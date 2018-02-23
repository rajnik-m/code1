Imports System.IO

Namespace Access.BulkMailer

  ''' <summary>
  ''' This class is used to get an instance of a bulk mailer interface that is appropriate to the
  ''' current environment.  It is the only supported way of obtaining a new <see cref="BulkMailer" />
  ''' instance.  New bulk mailer interface implementations should be added here as required.
  ''' </summary>
  Public Class BulkMailerFactory

    ''' <summary>
    ''' Gets a new instance of an <see cref="BulkMailer" /> appropriate to the current environemt.
    ''' </summary>
    ''' <param name="pEnvironment">The current <see cref="CDBEnvironment" />.</param>
    ''' <returns>An appropriate <see cref="BulkMailer" /> instance.</returns>
    Public Shared Function GetBulkMailerInstance(pEnvironment As CDBEnvironment) As BulkMailer
      Dim vBulkMailer As BulkMailer = Nothing
      If String.IsNullOrWhiteSpace(pEnvironment.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBulkMailerLoginId)) Or String.IsNullOrWhiteSpace(pEnvironment.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBulkMailerPassword)) Then
        vBulkMailer = New NullBulkMailer()
      Else
        Try
          vBulkMailer = New DotMailer(pEnvironment.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBulkMailerLoginId), pEnvironment.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlBulkMailerPassword), pEnvironment)
        Catch vEx As FileNotFoundException
          RaiseError(DataAccessErrors.daeDotMailerSdkNotFound)
        End Try
      End If
      Return vBulkMailer
    End Function

  End Class

End Namespace
