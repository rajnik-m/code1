Public Class TelephoneNumberValidator
  Inherits BaseValidator

  Protected Overrides Function EvaluateIsValid() As Boolean
    Return Utilities.ValidateTelephoneNumber(GetControlValidationValue(ControlToValidate))
  End Function
End Class
