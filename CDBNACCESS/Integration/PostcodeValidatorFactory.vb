Namespace Access.PostcodeValidation

  Public Class PostcodeValidatorFactory

    Friend Shared Function GetValidator(ByVal pEnv As CDBEnvironment, ByVal iso3DefaultCountryCode As String) As QASProOnDemandPostcodeValidator
      Dim Url As String = pEnv.GetConfig("qas_pro_ondemand_url")
      If Not String.IsNullOrWhiteSpace(Url) Then
        Dim vQASPostcodeValidator As QASProOnDemandPostcodeValidator = QASProOnDemandPostcodeValidator.GetInstance(New Uri(Url),
                                                                                                                   pEnv.GetConfig("qas_delivery_point_suffix"),
                                                                                                                   pEnv.GetConfig("qas_grid_references"),
                                                                                                                   pEnv.GetConfig("qas_lea_data"),
                                                                                                                   iso3DefaultCountryCode)
        Return vQASPostcodeValidator
      Else
        Return Nothing
      End If
    End Function
  End Class

End Namespace

