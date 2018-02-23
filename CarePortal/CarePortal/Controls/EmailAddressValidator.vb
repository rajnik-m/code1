Imports System.Net.Mail

Public Class EmailAddressValidator
  Inherits BaseValidator

  Protected Overrides Function EvaluateIsValid() As Boolean
    Dim validAddress As Boolean = True
    Try
      Utilities.ValidateEmailAddress(GetControlValidationValue(ControlToValidate))
    Catch ex As CareException
      validAddress = False
    End Try
    Return validAddress
  End Function

End Class
