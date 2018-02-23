Imports System.Web.UI.WebControls

Public Class DateValidator
  Inherits BaseValidator

  'http://www.codeproject.com/aspnet/datevalidator.asp

  Protected Overrides Function EvaluateIsValid() As Boolean
    Try
      Dim vDate As String = DirectCast(Me.FindControl(Me.ControlToValidate), TextBox).Text
      If vDate = String.Empty Then Return True
      Dim vResult As Date
      Return DateTime.TryParse(vDate, vResult)
    Catch vEx As Exception
      Return False
    End Try
    Return True
  End Function

  Protected Overrides Function ControlPropertiesValid() As Boolean
    Return TypeOf (Me.FindControl(Me.ControlToValidate)) Is TextBox
  End Function


  Protected Overrides Sub AddAttributesToRender(ByVal writer As System.Web.UI.HtmlTextWriter)
    MyBase.AddAttributesToRender(writer)
    If RenderUplevel Then
      writer.AddAttribute("evaluationfunction", "ValidateDate")
    End If
  End Sub

  'Protected Overrides Sub OnPreRender(ByVal e As System.EventArgs)
  '  If EnableClientScript Then ClientScript()
  '  MyBase.OnPreRender(e)
  'End Sub

  End Class
