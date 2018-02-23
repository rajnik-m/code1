Imports System.Configuration

Public Class AppValues

  Public Shared ReadOnly Property DefaultCountryCode() As String
    Get
      Dim vDefaultCountry As String = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.option_country)
      If vDefaultCountry.Length = 0 Then
        If My.Application.UICulture.Name.ToLower.Contains("ch") Then
          vDefaultCountry = "CH"
        ElseIf My.Application.UICulture.Name.ToLower.Contains("nl") Then
          vDefaultCountry = "NL"
        Else
          vDefaultCountry = "UK"
        End If
      End If
      Return vDefaultCountry
    End Get
  End Property

  Public Shared ReadOnly Property DateFormat() As String
    Get
      If My.Application.Culture.DateTimeFormat.ShortDatePattern.Contains("yyyy") Then
        Return My.Application.Culture.DateTimeFormat.ShortDatePattern
      Else
        Return My.Application.Culture.DateTimeFormat.ShortDatePattern.Replace("yy", "yyyy")
      End If
    End Get
  End Property

  Public Shared Function TodaysDate() As String
    Return System.DateTime.Today.ToString(DateFormat)
  End Function

End Class